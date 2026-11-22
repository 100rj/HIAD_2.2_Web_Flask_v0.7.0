# HIAD 2.2 Web Flask v0.7.0 연구 재현성 부속자료

이 폴더는 공개 프로그램 저장소와 함께 제공하는 연구 재현성 부속자료이다. 설정파일, 검토 입력양식, 실행 매니페스트, 합성 예제 입·출력자료 및 SHA-256 해시목록을 포함한다.

## 폴더 구성

- `config/`: v0.7.0 설정파일의 비식별 예시
- `review_templates/`: 비교검토·무후보 검토 입력양식과 필드 설명
- `manifests/`: 합성 예제 실행 매니페스트와 실제 계산한 SHA-256 목록
- `examples/input/`: 공개 가능한 합성 입력자료
- `examples/review/`: 합성사건에 대한 예제 검토결과
- `examples/output/`: 합성 입력으로 생성한 예제 출력자료
- `docs/`: 공개범위와 GitHub 업로드 절차
- `tools/`: SHA-256을 다시 계산하는 보조 프로그램

## 주의사항

이 폴더의 예제자료와 해시는 연구 본 데이터의 결과가 아니라 **합성 예제의 형식·무결성 검증자료**이다. 예제 출력은 별도 재현성 빌더에서 작성한 공개용 합성 예시이므로, v0.7.0 웹 프로그램의 본 실행결과로 표현해서는 안 된다. HIAD 2.2 원자료, 실제 검토자의 식별정보, API 인증정보, SQLite 실행 DB 및 비밀설정은 포함하지 않는다. 논문에서 보고한 본 실행의 파일별 해시값을 이 예제 해시로 대체해서도 안 된다.

## 해시 재계산

저장소 최상위에서 다음과 같이 실행한다.

```bash
python reproducibility_materials/tools/generate_sha256.py reproducibility_materials/examples/output
```

운영체제나 프로그램 구현에 따라 출력 위치가 다를 수 있으므로, 공개된 `manifests/SHA256SUMS.example.txt`와 비교할 때에는 파일 내용의 SHA-256을 기준으로 확인한다.
