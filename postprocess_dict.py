import re
import json
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# ── 용어 규칙 (최우선) ────────────────────────────────────────────────
TERM_RULES = {
    "Dập Tự Động": "압착기",
    "dập tự động": "압착기",
    "Dập tự động": "압착기",
    "DAP TU DONG": "압착기",
}

# ── 목적어+동사 → 동사(명사형) 처리 대상 동사 목록 ─────────────────────
ACTION_VERBS = [
    '확인', '수행', '준비', '진행', '관리', '적용', '설정', '실시',
    '사용', '기록', '교체', '점검', '조정', '검사', '실행', '처리',
    '보관', '표시', '부착', '제거', '측정', '삽입', '연결', '분리',
    '청소', '유지', '보전', '작성', '등록', '배치', '정리', '고정',
    '투입', '장착', '탈착', '조립', '해체', '세척', '건조', '포장',
    '표기', '기입', '저장', '삭제', '선택', '결정', '승인', '반영',
]

def to_concise(text: str) -> str:
    """간결체 변환: 동사 명사화, 경어체 제거"""
    t = text.strip()
    if len(t) < 3:
        return t

    # 1) 목적어 조사 + 특정 동사 + 어미 → 동사만 남기기
    #    예: "~을 확인합니다." → "~확인."
    for v in ACTION_VERBS:
        for ending in ['합니다', '한다', '하다']:
            t = re.sub(rf'([을를])\s*{v}{ending}\.?$', f' {v}.', t)
        # ~을/를 [v]해야 합니다 → ~[v]해야 함.
        t = re.sub(rf'([을를])\s*{v}해야\s*합니다\.?$', f' {v}해야 함.', t)

    # 2) 일반 경어체 → 간결체
    endings = [
        # 의무/명령
        (r'해야\s*합니다\.?$',   '해야 함.'),
        (r'해야\s*한다\.?$',     '해야 함.'),
        (r'하여야\s*합니다\.?$', '해야 함.'),
        (r'되어야\s*합니다\.?$', '되어야 함.'),
        (r'하십시오\.?$',        '할 것.'),
        (r'하세요\.?$',          '할 것.'),
        (r'해주세요\.?$',        '할 것.'),
        (r'해야\s*함\.?$',       '해야 함.'),
        # 서술어
        (r'습니다\.?$',          '음.'),
        (r'ㅂ니다\.?$',          '음.'),
        (r'합니다\.?$',          '함.'),
        (r'입니다\.?$',          '임.'),
        (r'됩니다\.?$',          '됨.'),
        (r'있습니다\.?$',        '있음.'),
        (r'없습니다\.?$',        '없음.'),
        (r'한다\.?$',            '함.'),
        (r'된다\.?$',            '됨.'),
        (r'있다\.?$',            '있음.'),
        (r'없다\.?$',            '없음.'),
        (r'이다\.?$',            '임.'),
        # 연결어미가 문말에 오는 경우
        (r'하여$',               '함.'),
        (r'하며$',               '함.'),
    ]

    for pattern, replacement in endings:
        new_t = re.sub(pattern, replacement, t)
        if new_t != t:
            t = new_t
            break   # 하나만 적용 후 종료

    # 3) 이중 마침표 정리
    t = re.sub(r'\.{2,}$', '.', t)
    t = re.sub(r'\.\s+\.$', '.', t)

    # 4) 불필요한 ~을 사용하여/~를 이용하여 → 활용
    t = re.sub(r'([을를])\s*사용하여', ' 활용', t)
    t = re.sub(r'([을를])\s*이용하여', ' 활용', t)

    return t


def main():
    src = "C:\\hex_edit_project\\translation_dict.json"
    dst = "C:\\hex_edit_project\\translation_dict.json"

    with open(src, 'r', encoding='utf-8') as f:
        orig = json.load(f)

    result = {}
    changed = 0
    samples = []   # before/after 샘플

    for vn, ko in orig.items():
        concise = to_concise(ko)
        result[vn] = concise
        if concise != ko:
            changed += 1
            if len(samples) < 15:
                samples.append((vn, ko, concise))

    # 용어 규칙 최우선 덮어쓰기
    for term, ko in TERM_RULES.items():
        result[term] = ko

    # 저장
    with open(dst, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"총 {len(result)}개 항목 중 {changed}개 간결체 변환 완료\n")
    print("── 변환 샘플 (before → after) ──────────────────────────────")
    for vn, before, after in samples:
        print(f"  VN : {vn[:60]}")
        print(f"  전 : {before}")
        print(f"  후 : {after}")
        print()

if __name__ == "__main__":
    main()
