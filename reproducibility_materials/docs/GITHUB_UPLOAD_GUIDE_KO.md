# GitHub 웹브라우저 업로드 절차

1. 제공된 ZIP 파일을 Windows에서 압축 해제한다.
2. 공개 저장소 `https://github.com/100rj/HIAD_2.2_Web_Flask_v0.7.0`에 접속한다.
3. `Add file`을 누르고 `Upload files`를 선택한다.
4. 압축을 푼 폴더 중 `reproducibility_materials` 폴더 전체를 업로드 영역으로 끌어 놓는다.
5. 업로드 목록에 `reproducibility_materials/config`, `review_templates`, `manifests`, `examples`, `docs`, `tools` 경로가 모두 표시되는지 확인한다.
6. 커밋 메시지에 `Add reproducibility materials and synthetic examples`를 입력한다.
7. `Commit directly to the main branch`를 선택하고 `Commit changes`를 누른다.
8. 저장소 첫 화면에서 `reproducibility_materials` 폴더가 보이는지 확인한다.

업로드 전에 `data` 폴더, HIAD 원본 `.xlsx`, `hiad_web.sqlite3`, `.env`, 실제 검토자 식별정보가 포함되지 않았는지 반드시 확인한다. ZIP 파일 하나만 저장소에 올리지 말고, 압축을 해제한 폴더와 파일을 올려야 웹에서 내용을 확인할 수 있다.
