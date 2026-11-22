# 검토결과 입력항목

| 필드 | 설명 |
|---|---|
| Event_ID | 사건 고유 식별자 |
| Review_Type | 비교검토 또는 무후보 표본검토 |
| Original_Labels | 원자료에서 표준화한 라벨 |
| Candidate_Labels | 프로그램이 제안한 후보라벨 |
| Final_Labels | 검토 후 확정한 C01∼C08 라벨 |
| Decision | 유지·추가·삭제·변경·보류·제외 |
| Evidence_Text | 판정에 사용한 사건서술 근거 |
| Reason | 최종 판정 사유 |
| Reviewer_ID | 비식별 검토자 식별자 |
| Review_Status | 검토완료·미검토·보류·제외 |
| Review_Date | 검토일(YYYY-MM-DD) |

복수라벨은 세미콜론으로 구분한다. 예: `C01;C06`.

