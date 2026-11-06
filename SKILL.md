---
name: my_pptx
description: "빨강 주도 투자자 발표(Red Pitch Deck) 스타일의 한국어 PowerPoint(.pptx)를 만들거나 고칠 때 사용한다. 표지·섹션·마무리는 전면 빨강, 내지는 흰 바탕 + 코너 빨강 블록. 레이아웃 6유형 함수 라이브러리(brand.js), 브랜드 규칙 검사기(brand_lint.py), PowerPoint 렌더링 QA(render.ps1)를 포함한다. '슬라이드', '발표자료', 'PPT', '덱', 'pptx' 요청에서 트리거한다."
---

# my_pptx — Red Pitch Deck PPT 스킬

응답과 슬라이드 내용은 기본 **한국어**. 디자인 값의 원본은 [references/brandguideline.md](references/brandguideline.md) — **처음 쓰기 전에 반드시 읽는다.** 사용자가 다른 색·스타일을 명시하면 그 요청이 가장 우선이다.

## 구성

| 파일 | 역할 |
|------|------|
| `scripts/brand.js` | pptxgenjs 래퍼. 레이아웃 A–F 함수, 장식 모티프, 한글 자동 줄바꿈, 넘침·노트·출처 누락 경고 |
| `scripts/brand_lint.py` | 완성된 .pptx의 브랜드 규칙 자동 검사 (표준 라이브러리만 사용) |
| `scripts/render.ps1` | PowerPoint COM으로 슬라이드 PNG + 한 장 미리보기(`_contact.png`) 생성 |
| `examples/sample_deck.js` | 6유형을 모두 쓰는 예제 (출처 있는 수치만 사용) |

## 작업 순서

1. `references/brandguideline.md` 읽기.
2. 슬라이드 계획: 1장 표지(A), 마지막 마무리(F)는 빨강. 내지는 B·C·D·E를 섞는다 (5장 이상이면 3종 이상, 흰 슬라이드 3–4장마다 E 1장).
3. 작업 폴더 준비 (아래 "환경")에서 덱 스크립트를 작성하고 `node deck.js out.pptx`로 생성. **콘솔 경고(넘침·노트 없음·출처 없음)는 모두 해결한다.**
4. 검사: `PYTHONUTF8=1 python scripts/brand_lint.py out.pptx` → ERROR 0 이어야 한다. (WARN은 의도된 것인지 확인. 10pt 출처 캡션 WARN은 정상)
5. 렌더링: `powershell -ExecutionPolicy Bypass -File scripts/render.ps1 -Pptx out.pptx -OutDir png` → `_contact.png`로 전체를 보고, 의심 슬라이드는 `slideNN.png`를 크게 본다. 텍스트 넘침·어색한 줄바꿈·겹침·대비 확인.
6. 고친 뒤 3–5 반복. 수정 후엔 반드시 다시 렌더링.
7. 최종 파일 경로, lint/렌더링 결과, 확인하지 못한 부분을 사실대로 보고.

기존 .pptx를 읽거나 XML 수준으로 편집할 때는 환경에 있는 기본 `pptx` 스킬(markitdown, unzip→XML 편집→zip, validate)을 함께 쓴다.

## brand.js API

```js
const { createDeck } = require("<skill>/scripts/brand.js");
const deck = createDeck({ title: "덱 제목" });

await deck.cover({ brand, title, subtitle, icon, satellites: [아이콘 최대 3], notes });          // A
await deck.halfPanel({ side: "left"|"right", image?, panelIcon, panelText, title,
                       items: [{ title, desc, icon }] /*≤4*/, footnote?, notes });              // B
await deck.textList({ title, items /*≤4*/, card?: { icon, title, desc }, notes });              // C
await deck.stats({ title, chart?: { type: "bar"|"line"|"doughnut"|"pie", labels, values,
                   name?, series?, showValue?, format?: "0.0", min? },
                   points?: [{ title, desc }] /*≤3*/, stats?: [{ value, label }] /*≤4*/,
                   source /*필수*/, notes });                                                    // D
await deck.section({ title, subtitle?, items?: [{ title, desc, icon }] /*≤4, 없으면 큰 제목형*/, notes }); // E
await deck.closing({ heading = "Thank You,", message, footer?, heading2?, items? /*≤3*/, icon?, notes }); // F
const s = deck.blank({ bg, notes });  // 자유 구성: deck.helpers.{txt, rect, oval, badge, xMark, cornerBlock, dotLine}
await deck.save("out.pptx");          // 경고 목록 출력 + 반환
```

- `icon`은 `react-icons/fa` 이름 문자열 (예: `FaRobot`, `FaChartLine`).
- `\n`으로 직접 줄을 끊어도 되고, 안 끊으면 `wrapKo`가 단어 경계에서 고르게 끊는다.
- 항목 수가 한도를 넘으면 잘리고 경고가 나온다 → 슬라이드를 나눈다.
- `image`는 사용자가 `assets/`에 넣은 파일만 사용. 없으면 빨강 패널 + 아이콘이 대신 들어간다.

## 디자인 핵심 (요약)

- 색: `F40009` · `B30000` · `FDECEC` · `FFFFFF` · `111111` · `5F5F5F` (+ `F2F2F2`, `D9D9D9`). 그 외 금지. 빨강 위 글자는 흰색만. 슬라이드당 빨강 "주인공" 1곳.
- 글꼴: 한글 `Malgun Gothic`, 영문·숫자 `Arial`. 제목 28–32pt Bold, 본문 12–14pt(최소 11pt), 캡션 10pt.
- 제목 밑줄·전폭 헤더바·가장자리 줄무늬·카드 한쪽 테두리 금지. 모든 슬라이드에 시각 요소 1개 이상.
- 통계·수치는 출처가 있는 값만. 특정 기업 로고·사진·사칭 금지.
- 슬라이드마다 발표자 노트 1–2문장.

## 환경 (Windows 기준)

- 의존 패키지는 **현재 작업 폴더의 node_modules**에서 찾는다: 작업 폴더에서 `npm install pptxgenjs react react-dom react-icons sharp` 후 그 폴더에서 `node` 실행.
- Node가 없으면 nodejs.org 휴대용 zip(`node-v22.x-win-x64.zip`)을 임시 폴더에 풀어 사용.
- **스킬 폴더나 덱 스크립트가 한글 경로에 있으면 node가 비정상 종료할 수 있다 (실측: 0xC0000005).** 이 경우 `scripts/`와 덱 스크립트를 영문 경로(예: 임시 폴더)로 복사해서 실행한다.
- 디스크 여유가 작으면 `node_modules`·PNG는 임시 폴더에 두고 최종 .pptx와 소스만 프로젝트에 저장.
- 한글이 든 코드는 셸 heredoc으로 넘기지 말고 파일로 저장해 실행. Python은 `PYTHONUTF8=1`.
- LibreOffice 환경이면 render.ps1 대신 `soffice --headless --convert-to pdf` 후 이미지로 변환해 같은 방식으로 확인한다.
