"""
BOM 비교 엔진
- 이전 BOM 데이터와 새 BOM 데이터를 비교하여 변경사항 감지
- 행 추가/삭제, 필드값 변경, 컬러 헤더/값 변경, Master 정보 변경 추적
"""
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from models import BomRow


def _normalize_val(s: str) -> str:
    """비교 전 값 정규화: private-use Unicode 제거, 공백 정리"""
    if not s:
        return ""
    # Private Use Area 문자 제거 (\ue000-\uf8ff) — PDF 체크마크 등
    s = re.sub(r"[\ue000-\uf8ff]", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _is_data_bleed(color_val: str, material_name: str) -> bool:
    """
    컬러 값이 실제 컬러가 아니라 Material Name 텍스트 블리드인지 판별.
    긴 Material Name이 Excel 셀 경계를 넘어 컬러 컬럼으로 흘러가는 현상.
    패턴: "Traceability Label/Joker Tag / /", "Price Ticket / /" 등
    """
    if not color_val or not material_name:
        return False
    # " / /" 로 끝나는 경우: 블리드 가능성 높음
    stripped = color_val.rstrip(" /").strip()
    if stripped and stripped in material_name:
        return True
    return False


@dataclass
class RowDiff:
    """양쪽 BOM에 모두 존재하지만 내용이 다른 행"""
    old_row: BomRow
    new_row: BomRow
    changed_fields: Dict[str, Tuple[str, str]] = field(default_factory=dict)
    changed_colors: Dict[str, Tuple[str, str]] = field(default_factory=dict)


@dataclass
class BomDiff:
    """BOM 비교 결과"""
    added_rows: List[BomRow] = field(default_factory=list)
    removed_rows: List[BomRow] = field(default_factory=list)
    modified_rows: List[RowDiff] = field(default_factory=list)
    added_colors: List[str] = field(default_factory=list)
    removed_colors: List[str] = field(default_factory=list)
    master_changes: Dict[str, Tuple[str, str]] = field(default_factory=dict)

    @property
    def has_changes(self) -> bool:
        return bool(
            self.added_rows or self.removed_rows or self.modified_rows
            or self.added_colors or self.removed_colors or self.master_changes
        )

    def summary_lines(self) -> List[str]:
        """GUI 로그용 요약 텍스트 생성"""
        lines = []
        if not self.has_changes:
            lines.append("변경사항 없음")
            return lines

        if self.master_changes:
            lines.append(f"  Master 정보 변경: {len(self.master_changes)}건")
            for k, (old, new) in self.master_changes.items():
                lines.append(f"    - {k}: '{old}' → '{new}'")

        if self.added_colors:
            lines.append(f"  컬러 추가: {len(self.added_colors)}개")
            for c in self.added_colors:
                short = c.replace('\n', ' / ')
                lines.append(f"    + {short}")

        if self.removed_colors:
            lines.append(f"  컬러 삭제: {len(self.removed_colors)}개")
            for c in self.removed_colors:
                short = c.replace('\n', ' / ')
                lines.append(f"    - {short}")

        if self.added_rows:
            lines.append(f"  자재 추가: {len(self.added_rows)}행")
            for r in self.added_rows:
                lines.append(f"    + [{r.category}] {r.product} / {r.material_name} / {r.usage}")

        if self.removed_rows:
            lines.append(f"  자재 삭제: {len(self.removed_rows)}행")
            for r in self.removed_rows:
                lines.append(f"    - [{r.category}] {r.product} / {r.material_name} / {r.usage}")

        if self.modified_rows:
            lines.append(f"  변경된 행: {len(self.modified_rows)}행")
            for rd in self.modified_rows:
                desc = f"    ~ [{rd.new_row.category}] {rd.new_row.product} / {rd.new_row.material_name}"
                lines.append(desc)
                for fname, (ov, nv) in rd.changed_fields.items():
                    lines.append(f"        {fname}: '{ov}' → '{nv}'")
                for ch, (ov, nv) in rd.changed_colors.items():
                    short_h = ch.replace('\n', ' / ')
                    lines.append(f"        [{short_h}]: '{ov}' → '{nv}'")

        return lines


def _row_key(r: BomRow) -> Tuple[str, str, str]:
    return (
        (r.product or "").strip(),
        (r.material_name or "").strip(),
        (r.usage or "").strip(),
    )


def _row_key_loose(r: BomRow) -> Tuple[str, str]:
    return (
        (r.product or "").strip(),
        (r.material_name or "").strip(),
    )


COMPARE_FIELDS = [
    ("category", "Category"),
    ("product", "Product"),
    ("material_name", "Material Name"),
    ("supplier_article_number", "Supplier Article Number"),
    ("usage", "Usage"),
    ("quality_details", "Quality Details"),
    ("supplier", "Supplier"),
]


def _compare_row_fields(old_row: BomRow, new_row: BomRow) -> Dict[str, Tuple[str, str]]:
    changes = {}
    for attr, label in COMPARE_FIELDS:
        old_val = _normalize_val(getattr(old_row, attr, "") or "")
        new_val = _normalize_val(getattr(new_row, attr, "") or "")
        # Category: Excel 템플릿에 없는 경우가 많아 old가 비어있으면 비교 스킵
        if attr == "category" and not old_val:
            continue
        if old_val != new_val:
            changes[label] = (old_val, new_val)
    return changes


def _compare_row_colors(
    old_row: BomRow,
    new_row: BomRow,
    matched_color_pairs: List[Tuple[str, str]],
    added_color_headers: List[str],
    removed_color_headers: List[str],
) -> Dict[str, Tuple[str, str]]:
    """
    컬러 값 비교.
    matched_color_pairs: [(old_header, new_header), ...] 매칭된 헤더 쌍
    added/removed: 한쪽에만 존재하는 헤더
    """
    changes = {}
    old_mat = (old_row.material_name or "")
    new_mat = (new_row.material_name or "")

    # 매칭된 컬러 헤더의 값 비교
    for old_h, new_h in matched_color_pairs:
        old_val = _normalize_val((old_row.colors or {}).get(old_h, ""))
        new_val = _normalize_val((new_row.colors or {}).get(new_h, ""))
        # 데이터 블리드 필터: Material Name 잔여 텍스트가 컬러 셀에 들어간 경우 무시
        if _is_data_bleed(old_val, old_mat):
            old_val = ""
        if _is_data_bleed(new_val, new_mat):
            new_val = ""
        if old_val != new_val:
            changes[new_h] = (old_val, new_val)
    # 새로 추가된 컬러에 값이 있으면 변경으로 표시
    for h in added_color_headers:
        new_val = _normalize_val((new_row.colors or {}).get(h, ""))
        if _is_data_bleed(new_val, new_mat):
            new_val = ""
        if new_val:
            changes[h] = ("", new_val)
    # 삭제된 컬러에 값이 있었으면 변경으로 표시
    for h in removed_color_headers:
        old_val = _normalize_val((old_row.colors or {}).get(h, ""))
        if _is_data_bleed(old_val, old_mat):
            old_val = ""
        if old_val:
            changes[h] = (old_val, "")
    return changes


def compare_boms(
    old_rows: List[BomRow],
    old_color_headers: List[str],
    old_master: Dict[str, str],
    new_rows: List[BomRow],
    new_color_headers: List[str],
    new_master: Dict[str, str],
) -> BomDiff:
    """
    이전 BOM과 새 BOM을 비교하여 BomDiff를 반환.

    매칭 전략:
    1차: (product, material_name, usage) 정확 매칭
    2차: 매칭 안 된 행끼리 (product, material_name) 느슨 매칭
    """
    diff = BomDiff()

    # --- Master 비교 ---
    master_keys = set(list(old_master.keys()) + list(new_master.keys()))
    for k in master_keys:
        ov = _normalize_val(old_master.get(k, "") or "")
        nv = _normalize_val(new_master.get(k, "") or "")
        if ov != nv:
            diff.master_changes[k] = (ov, nv)

    # --- 컬러 헤더 비교 (정규화하여 매칭) ---
    def _norm_header(h: str) -> str:
        """컬러 헤더 비교용 정규화 — 공백/줄바꿈 무시"""
        return re.sub(r"\s+", "", h).lower()

    old_norm_map = {_norm_header(h): h for h in old_color_headers}
    new_norm_map = {_norm_header(h): h for h in new_color_headers}
    old_norm_set = set(old_norm_map.keys())
    new_norm_set = set(new_norm_map.keys())

    diff.added_colors = [new_norm_map[n] for n in new_norm_set - old_norm_set]
    diff.removed_colors = [old_norm_map[n] for n in old_norm_set - new_norm_set]

    # 매칭된 컬러 헤더 쌍 구축
    matched_color_pairs = []
    for n in old_norm_set & new_norm_set:
        matched_color_pairs.append((old_norm_map[n], new_norm_map[n]))

    # --- 행 매칭 (1차: 정확 키) ---
    old_by_key: Dict[Tuple, List[BomRow]] = {}
    for r in old_rows:
        k = _row_key(r)
        old_by_key.setdefault(k, []).append(r)

    new_by_key: Dict[Tuple, List[BomRow]] = {}
    for r in new_rows:
        k = _row_key(r)
        new_by_key.setdefault(k, []).append(r)

    matched_old = set()
    matched_new = set()

    all_keys = set(list(old_by_key.keys()) + list(new_by_key.keys()))
    for k in all_keys:
        old_list = old_by_key.get(k, [])
        new_list = new_by_key.get(k, [])

        pairs = min(len(old_list), len(new_list))
        for i in range(pairs):
            or_ = old_list[i]
            nr_ = new_list[i]
            matched_old.add(id(or_))
            matched_new.add(id(nr_))

            field_changes = _compare_row_fields(or_, nr_)
            color_changes = _compare_row_colors(or_, nr_, matched_color_pairs, diff.added_colors, diff.removed_colors)
            if field_changes or color_changes:
                diff.modified_rows.append(RowDiff(
                    old_row=or_, new_row=nr_,
                    changed_fields=field_changes,
                    changed_colors=color_changes,
                ))

    # --- 2차: 느슨 매칭 (product, material_name) ---
    unmatched_old = [r for r in old_rows if id(r) not in matched_old]
    unmatched_new = [r for r in new_rows if id(r) not in matched_new]

    loose_old: Dict[Tuple, List[BomRow]] = {}
    for r in unmatched_old:
        k = _row_key_loose(r)
        loose_old.setdefault(k, []).append(r)

    loose_new: Dict[Tuple, List[BomRow]] = {}
    for r in unmatched_new:
        k = _row_key_loose(r)
        loose_new.setdefault(k, []).append(r)

    still_unmatched_old_ids = set(id(r) for r in unmatched_old)
    still_unmatched_new_ids = set(id(r) for r in unmatched_new)

    loose_keys = set(list(loose_old.keys()) + list(loose_new.keys()))
    for k in loose_keys:
        ol = loose_old.get(k, [])
        nl = loose_new.get(k, [])
        pairs = min(len(ol), len(nl))
        for i in range(pairs):
            or_ = ol[i]
            nr_ = nl[i]
            still_unmatched_old_ids.discard(id(or_))
            still_unmatched_new_ids.discard(id(nr_))

            field_changes = _compare_row_fields(or_, nr_)
            color_changes = _compare_row_colors(or_, nr_, matched_color_pairs, diff.added_colors, diff.removed_colors)
            if field_changes or color_changes:
                diff.modified_rows.append(RowDiff(
                    old_row=or_, new_row=nr_,
                    changed_fields=field_changes,
                    changed_colors=color_changes,
                ))

    # --- 남은 행 = 추가/삭제 ---
    diff.removed_rows = [r for r in old_rows if id(r) in still_unmatched_old_ids]
    diff.added_rows = [r for r in new_rows if id(r) in still_unmatched_new_ids]

    return diff
