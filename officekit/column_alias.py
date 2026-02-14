from __future__ import annotations

from collections.abc import Iterable

# 영어/한글 컬럼명을 상호 호환하기 위한 별칭 그룹
COLUMN_ALIAS_GROUPS: tuple[tuple[str, ...], ...] = (
    ("ID", "아이디"),
    ("Name", "이름"),
    ("Department", "부서"),
    ("Category", "분류", "카테고리"),
    ("Qty", "수량"),
    ("UnitPrice", "단가"),
    ("OrderDate", "주문일자", "주문일"),
    ("Email", "이메일"),
    ("Total", "합계"),
)

_ALIAS_TO_GROUP: dict[str, tuple[str, ...]] = {}
for group in COLUMN_ALIAS_GROUPS:
    for alias in group:
        _ALIAS_TO_GROUP[alias] = group


def alias_candidates(name: str) -> tuple[str, ...]:
    """요청한 이름의 별칭 후보 목록을 반환한다."""
    return _ALIAS_TO_GROUP.get(name, (name,))


def resolve_existing_column(columns: Iterable[object], requested: str) -> str | None:
    """요청한 열과 별칭 중 실제로 존재하는 열 이름을 반환한다."""
    names = [str(c) for c in columns]
    name_set = set(names)
    for candidate in alias_candidates(requested):
        if candidate in name_set:
            return candidate
    return None


def require_existing_column(columns: Iterable[object], requested: str, role: str = "열") -> str:
    """요청한 열(또는 별칭)이 없으면 예외를 발생시키고, 있으면 실제 열명을 반환한다."""
    resolved = resolve_existing_column(columns, requested)
    if resolved is not None:
        return resolved
    names = [str(c) for c in columns]
    raise KeyError(f"{role} '{requested}'을(를) 찾을 수 없습니다. 사용 가능한 열: {names}")


def add_alias_keys(mapping: dict[str, object]) -> dict[str, object]:
    """컨텍스트 딕셔너리에 별칭 키를 추가해 템플릿 치환 호환성을 높인다."""
    expanded = dict(mapping)
    for key, value in list(mapping.items()):
        for alias in alias_candidates(key):
            expanded.setdefault(alias, value)
    return expanded
