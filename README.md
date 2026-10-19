# 업무메모

업무지시를 입력하고, `결재_미마감_업무지시_*.xlsx`의 미마감 업무를 찾아 완료 처리하는 Windows 프로그램.

## 사용
1. `dist/업무메모.exe`를 `바탕화면\업무지시` 폴더에 넣고 더블클릭
2. 처음 실행 때 바탕화면 '업무메모' 아이콘을 만들지 물어봄
3. 업무 입력 → Ctrl+Enter 저장 / 찾기(Ctrl+F) → 더블클릭 또는 Enter로 완료

- 실행 환경: Windows 10/11 (.NET Framework 4.8 기본 탑재)
- 서명되지 않은 exe라 처음 실행 때 "Windows의 PC 보호" 경고가 뜰 수 있음 → 추가 정보 → 실행

## 데이터 파일 (exe와 같은 폴더, 저장소에는 올리지 않음)
| 파일 | 내용 |
|---|---|
| `업무지시.txt` | 메모 업무. 한 줄 = 한 건, `[ ]` 진행중 / `[v]` 완료 |
| `완료기록.txt` | 결재(엑셀) 항목 완료 기록. `시각<TAB>키<TAB>내용` |
| `결재_미마감_업무지시_*.xlsx` | 매일 15:48 자동 작업이 생성. 입력한 업무는 가장 최근 파일 맨 아래에 추가, 완료는 노란색으로 표시 |
| `엑셀반영대기.txt` | 엑셀이 열려 있어 못 쓴 건. 엑셀을 닫으면 자동 반영 |
| `입력창_오류기록.txt` | 오류 기록 |

자세한 규칙은 [CLAUDE.md](CLAUDE.md) 참고.

## 폴더
- `src/` — 소스(`업무메모.cs`)와 아이콘
- `build/compile.ps1` — 빌드 스크립트 (PowerShell 7 + .NET Framework 4.8 참조 어셈블리)
- `dist/` — 빌드된 exe
- `legacy/` — 예전 PowerShell 버전 입력창과 바로가기 생성 bat

## 빌드
```
pwsh build/compile.ps1 -Src src/업무메모.cs -Out dist/업무메모.exe -Ico src/업무메모.ico -RefDir <net48 참조 어셈블리 폴더>
```
참조 어셈블리: NuGet `Microsoft.NETFramework.ReferenceAssemblies.net48` 패키지의 `build/.NETFramework/v4.8`
