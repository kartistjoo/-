# 아파트 셀프 사전점검 체크리스트

- `아파트_셀프_사전점검_체크리스트.xlsx` — 현장용 엑셀 (점검표 118항목, 표준점검표 136항목, 하자기록, 자동 요약, 더샵 사전방문)
- `아파트_셀프_사전점검_체크리스트.md` — 같은 내용의 문서판 + YouTube 영상 요약

## 영상을 더 추가하고 싶을 때 (휴대폰에서도)

Claude 앱에서 이 저장소로 클라우드 세션을 열고 이렇게 말하면 됩니다.

- "사전점검 영상 더 찾아서 체크리스트에 넣어줘"
- "이 영상 반영해줘 https://youtu.be/…"

Claude는 `.claude/skills/youtube-checklist/SKILL.md` 순서대로
`tools/yt_tool.py`로 자막을 받고 → 체크리스트(md)를 고치고 → `tools/make_xlsx.py`로 엑셀을 다시 만듭니다.

필요 조건: 클라우드 환경 네트워크 설정에 `www.youtube.com` 허용, Claude GitHub App이 이 저장소에 설치되어 있을 것.
