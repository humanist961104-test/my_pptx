---
name: my_pptx
description: "빨강 주도 투자자 발표(Red Pitch Deck) 스타일의 한국어 PowerPoint(.pptx)를 만들거나 고칠 때 사용한다. 표지·섹션·마무리는 전면 빨강, 내지는 흰 바탕 + 코너 빨강 블록. 색상 토큰(F40009 등), 글꼴(Malgun Gothic/Arial), 6가지 레이아웃 유형, 장식 모티프, QA 체크리스트를 포함한다. '슬라이드', '발표자료', 'PPT', '덱', 'pptx' 요청에서 트리거한다."
---

# my_pptx — 나만의 Red Pitch Deck PPT 스킬

응답과 슬라이드 내용은 기본 **한국어**. 디자인 값은 모두 [references/brandguideline.md](references/brandguideline.md)에 있다. **PPT를 만들기 전에 반드시 먼저 읽는다.** 사용자가 요청에서 다른 색·스타일을 명시하면 그 요청이 가장 우선이다.

## 디자인 핵심 (요약)

- 스타일: 빨강 주도. 표지·섹션·마무리 = **전면 빨강**, 내지 = **흰 바탕 + 코너 빨강 블록**.
- 색: `F40009`(RED) · `B30000`(RED_DEEP) · `FDECEC`(RED_TINT) · `FFFFFF` · `111111` · `5F5F5F`. 그 외 색 금지. 빨강 배경 위 글자는 흰색만. 슬라이드당 빨강 "주인공"은 1곳.
- 글꼴: 한글 `Malgun Gothic`, 영문·숫자 `Arial`. 제목 28–32pt Bold, 본문 12–14pt(최소 11pt), 표지 타이틀 54–60pt.
- 레이아웃: `LAYOUT_16x9`(10"×5.625"), 바깥 여백 0.5", 간격 0.3". 유형 A 표지 / B 인물형 / C 본문 텍스트형 / D 통계형 / E 섹션형 / F 마무리. **내지 5장 이상이면 3종 이상 섞는다.**
- 모든 슬라이드에 사진·아이콘·차트·도형 중 시각 요소 1개 이상. 제목 밑줄, 전폭 헤더바, 가장자리 줄무늬, 카드 한쪽 테두리 금지.
- 장식 모티프(코너 블록, 작은 검정 블록, × 마커, 점선, 아이콘 원)는 슬라이드당 1–2개, 덱 전체 일관.
- 한글은 글자 단위로 줄바꿈되므로 카드 안 설명은 `\n`으로 직접 끊고 한 줄 9–13자 이내.
- 공식 로고·제품 사진은 사용자가 `assets/`에 넣은 것만 사용. 임의 다운로드 금지. 특정 기업 사칭 금지.
- 통계·수치는 출처가 있는 값만. 근거 없는 수치 금지.
- 슬라이드마다 발표자 노트(`addNotes`) 1–2문장.

## 작업 순서

1. `references/brandguideline.md` 읽기.
2. 슬라이드 수·주제·유형 배치 계획 (표지/마무리 빨강, 내지 유형 섞기).
3. pptxgenjs 스크립트 작성 → `.pptx` 생성. 시작 코드는 guideline 9장 템플릿 사용.
4. 기존 `.pptx`를 읽거나 편집할 때는 환경에 있는 `pptx` 스킬(unzip → XML 편집 → zip, `markitdown`)을 병행한다.
5. 구조 검증(`validate.py` 등 가능한 도구) 후, PowerPoint COM(Windows)이나 LibreOffice로 PNG 렌더링해 **모든 슬라이드를 눈으로 확인** (텍스트 넘침·줄바꿈·겹침·대비).
6. 문제 수정 후 5 반복. 수정 후엔 반드시 다시 렌더링.
7. 최종 파일 경로, 검증 결과, 확인하지 못한 부분을 사실대로 보고.

## 환경 팁 (Windows)

- Node.js가 없으면 휴대용 zip을 임시 폴더에 풀어 PATH에 추가한 뒤 `npm install pptxgenjs react-icons react react-dom sharp`.
- 디스크 여유가 작으면 `node_modules`·렌더링 PNG는 임시 폴더에서 작업하고 최종 `.pptx`와 소스 스크립트만 프로젝트에 둔다.
- 한글이 든 코드는 셸 heredoc으로 넘기지 말고 파일로 저장해 실행 (글자 깨짐 방지).
- PowerPoint COM 내보내기:
  ```powershell
  $ppt = New-Object -ComObject PowerPoint.Application
  $pres = $ppt.Presentations.Open("<절대경로>\deck.pptx", $true, $false, $false)
  $pres.Export("<출력폴더>", "PNG", 1600, 900)
  $pres.Close(); $ppt.Quit()
  ```
- pptxgenjs 주의: hex는 `#` 없이 6자리, `shadow` 등 옵션 객체는 호출마다 새로 만든다, 곡선(웨이브)은 SVG→PNG(`sharp`)로 `addImage`.
