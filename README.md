[README.md](https://github.com/user-attachments/files/32735317/README.md)
# 보건의료데이터 중개 포털 (프로토타입)

https://kdyoung100.github.io/Korean-Healthcare-Data-Portal/

흩어져 있는 보건의료 데이터셋(공공데이터·임상데이터)을 한 곳에서 검색하고,
데이터셋별 범위·분석환경·신청 방법을 확인한 뒤 실제 신청 창구
(K-CURE, 보건의료 빅데이터 통합 플랫폼, 의료데이터 중심병원 등)로 연결해 주는 정적 웹사이트입니다.

- 서버 없이 **GitHub Pages**로 서비스합니다.
- 데이터는 **엑셀 2개**로 관리하고, 푸시하면 GitHub Actions가 자동으로 `data.json`을 만들어 배포합니다.

## 주요 기능

| 기능 | 설명 |
|---|---|
| 데이터셋 검색 | 이름·기관·키워드·설명·테이블명 검색, 관련도/최신 업데이트/조회수/이름순 정렬 |
| 검색 필터 | 구분(임상/공공), 주제(암·감염병 등 13개), 공급 기관, 수집 기간 |
| 보기 전환 | 카드형 목록 ↔ 표 형식(데이터셋명·기관명·수집기간·최종 업데이트·키워드) |
| 대화형 AI 검색 | 자연어 질문 → 관련 데이터셋 카드 + AI 답변(OpenRouter) + 웹 검색 결과 |
| 상세페이지 | 데이터셋 정보표, 신청 창구 링크, 자료집 기반 상세 안내, 같은 기관의 관련 데이터셋 |
| 메인 화면 | 관련 사이트 바로가기, 리소스 탐색(4×2 카드), 주제별 데이터 탐색, 주목 데이터셋 |

## 폴더 구성

```
index.html                         # 사이트 전체 (HTML/CSS/JS 한 파일)
chatbot-icon.png                   # 챗봇 버튼 이미지
build.py                           # 엑셀 2개 → data.json 생성
requirements.txt                   # build.py 의존성 (openpyxl)
자료입력양식.xlsx                  # 원본 ① 상세페이지 본문 (자료집)
중개포털_데이터_리스트_*.xlsx      # 원본 ② 검색 목록 (데이터셋)
자료집ID.json                      # 발급한 자료집ID 기록 — 반드시 커밋
cloudflare-worker.js               # 웹검색·AI 답변 프록시 (Cloudflare에 수동 배포)
.github/workflows/deploy.yml       # 빌드 & GitHub Pages 배포
업데이트.bat / 업데이트.command     # 로컬에서 data.json 다시 만들기 (Windows / Mac)
```

`data.json`은 빌드 산출물이라 **커밋하지 않습니다**(`.gitignore`에 포함). 배포 때 Actions가 새로 만듭니다.

## 데이터 수정 방법

1. 엑셀을 수정합니다.
   - 검색 목록(기관명, 데이터명, 수집기간, 키워드, 신청절차 등) → `중개포털_데이터_리스트_*.xlsx`의 **데이터셋** 시트
   - 상세페이지 본문 → `자료입력양식.xlsx`
2. 두 파일은 **자료집ID**(예: `G26`) 열로 연결됩니다. 데이터셋 행의 자료집ID 칸이 비어 있으면 상세페이지가
   "준비 중"으로 표시됩니다. 제목이 바뀌어도 연결은 끊기지 않습니다.
3. `main`에 푸시하면 자동으로 빌드·배포됩니다. 푸시 전에 확인하려면 아래 "로컬 미리보기"를 보세요.

파일명에 날짜가 들어간 데이터 리스트가 여러 개 있으면 **날짜가 가장 큰 파일**을 씁니다.
자세한 ID 규칙은 `자료집ID_안내.md`를 참고하세요.

## 로컬 미리보기

```bash
pip install -r requirements.txt
python build.py              # data.json 생성 (또는 업데이트.bat / 업데이트.command 더블클릭)
python -m http.server        # http://localhost:8000 접속
```

`index.html`을 더블클릭해 `file://`로 열면 `data.json`을 읽지 못해 동작하지 않습니다. 반드시 위처럼 로컬 서버로 여세요.

## 배포 (GitHub Pages)

처음 한 번만 설정합니다.

1. GitHub 저장소 **Settings → Pages → Build and deployment → Source**를 **GitHub Actions**로 선택합니다.
2. `main` 브랜치에 푸시합니다.

그 뒤로는 엑셀, `자료집ID.json`, `build.py`, `index.html`, `chatbot-icon.png` 중 하나가 바뀐 채로
`main`에 푸시될 때마다 자동 배포됩니다. **Actions** 탭에서 직접 실행할 수도 있습니다(Run workflow).
빌드 중 새 자료집ID가 발급되면 Actions가 `자료집ID.json`을 자동으로 커밋합니다.

빌드 결과(연결된 데이터셋 수, 새로 발급된 ID, 경고)는 Actions 실행 화면 맨 위 요약에서 볼 수 있습니다.

## AI 검색 · 웹 검색 (Cloudflare Worker)

`cloudflare-worker.js`는 git 푸시로 배포되지 **않습니다**. 수정했다면 Cloudflare 대시보드의
Worker 코드 편집기에 붙여 넣고 직접 배포해야 합니다.

- `GET /?q=` — 웹 검색 결과(DuckDuckGo)를 가져와 JSON으로 돌려줍니다.
- `POST /chat` — OpenRouter로 AI 답변을 중계합니다. API 키는 Cloudflare Secret `OPENROUTER_KEY`에만 있고
  `index.html`이나 git에는 들어가지 않습니다.
- 사용하는 모델은 `index.html`의 `CHAT_MODEL`(기본)과 `CHAT_MODEL_FALLBACK`(예비)입니다.
  무료 모델은 수시로 없어지므로 답변이 안 나오면 `AI검색_안내.md`의 모델 교체 절차를 참고하세요.

## 보안 주의

- `*.env` 파일(API 키 보관용)은 커밋하지 않습니다(`.gitignore`에 포함).
- 공개 저장소라면 엑셀 원본도 누구나 볼 수 있습니다. 비공개 정보가 들어 있지 않은지 확인하고 올리세요.
