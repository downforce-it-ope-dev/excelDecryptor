# Excel Password Remover

암호화된 엑셀 파일을 비밀번호로 복호화하는 프로그램입니다.

## 포함된 실행 방식

- `app.py`: 브라우저로 사용하는 Flask 웹앱
- `desktop_app.py`: 사용자가 바로 실행하기 쉬운 데스크톱 앱

## 데스크톱 앱 실행

```cmd
"C:\Users\downforceITkkt\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" "C:\Users\downforceITkkt\Documents\Codex\2026-04-23-node-js-view-async-function-fnimportexcel-2\desktop_app.py"
```

## 웹앱 실행

```cmd
"C:\Users\downforceITkkt\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" "C:\Users\downforceITkkt\Documents\Codex\2026-04-23-node-js-view-async-function-fnimportexcel-2\app.py"
```

브라우저에서 `http://127.0.0.1:5000` 으로 접속하면 됩니다.

## 지원 형식

- `.xls`
- `.xlsx`
- `.xlsm`
- `.xlsb`

## 주의

- 비밀번호가 틀리면 복호화에 실패합니다.
- 일부 오래된 `.xls` 파일은 암호화 방식에 따라 실패할 수 있습니다.
- 배포용으로는 데스크톱 앱을 exe로 빌드한 뒤 installer로 감싸는 방식을 권장합니다.

## 자동 업데이트 방식

- 프로그램 실행 시 `update_config.json`의 `manifest_url`을 확인합니다.
- 해당 주소의 `latest.json`에 더 높은 버전이 있으면 새 exe를 내려받아 교체한 뒤 다시 실행합니다.

예시 `latest.json`

```json
{
  "version": "1.1.0",
  "download_url": "https://example.com/excel-decryptor/ExcelDecryptor.exe"
}
```
