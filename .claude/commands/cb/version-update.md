---
description: Crewboard 스캐폴딩을 최신 릴리스(또는 로컬 문서)로 갱신한다
---

인자: $ARGUMENTS  (선택: 로컬 CREWBOARD.md 경로 — 폐쇄망 수동 다운로드 파일)

docs/GUIDE.md 의 §0.8 업데이트 프로토콜을 따른다.

1. `.claude/CREWBOARD-VERSION` 의 `version:` 확인
   (파일 없으면 사람에게 현재 버전을 묻고 마커 생성 후 계속)
2. 새 문서 확보
   - 인자 있음(오프라인): 인자 경로의 파일 사용
   - 인자 없음(온라인): source 저장소 최신 릴리스 에셋 다운로드
     `gh release download --repo <source> --pattern "CREWBOARD.md" -O /tmp/CREWBOARD-new.md`
     (gh 불가 시 raw URL WebFetch 폴백)
3. 버전 비교 — `.` 분리 후 major → minor 정수 비교 (float 변환 금지, 예: 1.10 > 1.9).
   새 버전 ≤ 현재이면 "현재 버전이 최신입니다" 보고 후 종료
4. changelog 에서 현재 버전 초과 변경분만 추려 요약 제시
5. [사람 게이트] §0.8.1 분류표 기준 재생성·병합 대상의 diff 표시 → 승인
6. 적용 (재생성/병합/보존 — §0.8.1 준수) → `CREWBOARD-VERSION` 갱신
7. 오프라인 모드면 인자로 받은 입력 파일 삭제
8. 적용 파일 목록 보고 + 커밋 제안 (push 는 사람)
