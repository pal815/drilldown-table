# -*- coding: utf-8 -*-
"""IR 덤프 vs 기준 픽스처 의미 비교. 사용: py compare_ir.py <dump.json> <fixture.json>
   숫자는 1e-6 허용오차, 문자열/None은 정확 일치. 일치하면 'MATCH'+exit0, 아니면 차이 출력+exit1."""
import json
import sys


def load(p):
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def as_num(x):
    if isinstance(x, bool):
        return None
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def diff(a, b, path, out):
    if isinstance(a, dict) and isinstance(b, dict):
        for k in sorted(set(a) | set(b), key=str):
            if k not in a:
                out.append(f"{path}.{k}: DUMP에 없음 (기준={b[k]!r})")
            elif k not in b:
                out.append(f"{path}.{k}: DUMP에 추가됨 ({a[k]!r})")
            else:
                diff(a[k], b[k], f"{path}.{k}", out)
    elif isinstance(a, list) and isinstance(b, list):
        if len(a) != len(b):
            out.append(f"{path}: 리스트 길이 {len(a)} != {len(b)}")
        else:
            for i, (x, y) in enumerate(zip(a, b)):
                diff(x, y, f"{path}[{i}]", out)
    else:
        na, nb = as_num(a), as_num(b)
        if na is not None and nb is not None:
            if abs(na - nb) > 1e-6:
                out.append(f"{path}: {a} != {b}")
        elif a != b:
            out.append(f"{path}: {a!r} != {b!r}")


def main():
    if len(sys.argv) != 3:
        print("usage: py compare_ir.py <dump.json> <fixture.json>")
        sys.exit(2)
    dump, fix = load(sys.argv[1]), load(sys.argv[2])
    out = []
    diff(dump, fix, "", out)
    if not out:
        print("MATCH")
        sys.exit(0)
    print(f"DIFF ({len(out)}개):")
    for line in out[:30]:
        print("  ", line)
    if len(out) > 30:
        print(f"   ... (+{len(out) - 30} more)")
    sys.exit(1)


if __name__ == "__main__":
    main()
