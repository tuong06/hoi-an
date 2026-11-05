// Build: NODE_PATH=<dir with pptxgenjs, sharp, react-icons> node build_deck.js
// Output: Dongnae-Festival_Deck_v1.pptx + Dongnae-Festival_Script_v1.md (same data → 1:1 slide/script)
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const React = require("react");
const RDS = require("react-dom/server");
const fa = require("react-icons/fa6");

const SKILL = process.env.PPTX_SKILL || "";
const DIR = __dirname;
const PH = (n) => path.join(DIR, "photos", `p${String(n).padStart(3, "0")}.jpg`);
const OUT = path.join(DIR, "Dongnae-Festival_Deck_v1.pptx");
const MD = path.join(DIR, "Dongnae-Festival_Script_v1.md");

const THEME = {
  name: "Dongnae 1592",
  headFontFace: "Malgun Gothic",
  bodyFontFace: "Malgun Gothic",
  colors: {
    dk1: "1C2230", lt1: "FFFFFF", dk2: "2E6B5E", lt2: "F1F3F5",
    accent1: "B5382B", accent2: "2E6B5E", accent3: "D9A43A",
    accent4: "4A5568", accent5: "8C97A8", accent6: "F6E3E0",
    hlink: "2E6B5E", folHlink: "4A5568",
  },
};
const HEX = THEME.colors;
const W = 13.333, H = 7.5;

// ---------- image helpers ----------
const DPI = 110;
async function photo(n, w, h, opt = {}) {
  const pw = Math.round(w * DPI), ph = Math.round(h * DPI);
  const r = Math.round((opt.radius ?? 0.14) * DPI);
  let img = sharp(PH(n));
  if (opt.extract) img = img.extract(opt.extract);
  img = img.resize(pw, ph, { fit: "cover", position: opt.pos || "centre" });
  if (opt.darken) img = img.modulate({ brightness: opt.darken });
  const mask = Buffer.from(`<svg width="${pw}" height="${ph}"><rect width="${pw}" height="${ph}" rx="${r}" ry="${r}" fill="#fff"/></svg>`);
  const buf = r > 0
    ? await img.composite([{ input: mask, blend: "dest-in" }]).png({ compressionLevel: 9 }).toBuffer()
    : await img.jpeg({ quality: 80 }).toBuffer();
  return (r > 0 ? "image/png;base64," : "image/jpeg;base64,") + buf.toString("base64");
}
async function icon(Comp, color, px = 256) {
  const svg = RDS.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: px }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------- content (single source for deck + markdown) ----------
const S = [
  {
    kicker: "", title: "1592년 조선, 동래를 만나다",
    bullets: ["부제: 제32회 동래읍성역사축제 현장 답사 발표", "핵심 메시지: 역사 · 교육 · 체험이 하나 된 가족 축제", "답사일: 2026. 10. 17.(토) 13:00 ~ 야간 공연", "발표자: 광운대학교 정보융합학부 레깟뜨엉"],
    media: ["배경: 야간 뮤지컬 '외로운 성' 사진(어둡게 처리, 전체 화면)"],
    script: "안녕하세요. 저는 광운대학교 정보융합학부 레깟뜨엉이에요. 오늘은 제가 지난 10월 17일 토요일에 다녀온 '제32회 동래읍성역사축제'에 대해 발표해 볼게요. 이번 축제의 주제는 '1592년 조선, 동래를 만나다'예요. 저는 이 축제를 한마디로 '역사, 교육, 체험이 하나가 된 가족 축제'라고 소개하고 싶어요. 그럼 지금부터 저와 함께 1592년의 동래로 떠나 볼까요?",
  },
  {
    kicker: "", title: "오늘의 발표 순서",
    bullets: ["01 축제 소개 — 인증샷, 개요, 역사 배경, 주요 프로그램", "02 현장 답사 — 이동 동선, 첫인상", "03 좋았던 점 — 뮤지컬, 교육, 가족 친화, 장터와 먹거리", "04 아쉬운 점과 개선 방안", "05 안전관리 — 점검표, 서울세계불꽃축제와 비교", "06 맺음말 — AI 활용, 총평"],
    media: ["축제장 전경 사진(언덕 위에서 내려다본 천막 구역)"],
    script: "오늘 발표는 여섯 부분으로 나눠서 이야기할게요. 먼저 축제를 간단히 소개하고, 제가 실제로 걸었던 동선을 따라가 볼 거예요. 그다음에 좋았던 점, 아쉬웠던 점과 개선 방안을 말씀드리고, 안전관리를 꼼꼼하게 점검해 볼게요. 마지막으로 AI를 어떻게 활용했는지 소개하고 총평으로 마무리할게요.",
  },
  {
    kicker: "01 · 축제 소개", title: "저, 직접 다녀왔어요! — 현장 인증샷",
    bullets: ["입구 아치('제32회 동래읍성역사축제' 간판) 앞 인증샷", "함께 간 친구들과 단체 사진", "답사 일시: 2026. 10. 17.(토) 13:00 도착 → 야간 공연까지 관람", "장소: 동래읍성지 북문광장 일대(부산 동래구)"],
    media: ["필수: 얼굴과 축제 간판이 함께 보이는 인증샷 2장"],
    script: "먼저 제가 직접 다녀왔다는 증거부터 보여 드릴게요. 왼쪽 사진은 축제 입구 아치 앞에서 찍은 인증샷이에요. 뒤에 '제32회 동래읍성역사축제' 간판이 보이시죠? 오른쪽은 함께 간 친구들과 찍은 단체 사진이에요. 저희는 토요일 오후 1시에 도착해서 밤 공연이 끝날 때까지 축제를 즐겼어요. 낮과 밤의 분위기가 정말 달라서, 마치 두 개의 다른 축제에 온 것 같았어요.",
  },
  {
    kicker: "01 · 축제 소개", title: "축제 한눈에 보기",
    bullets: ["행사명: 제32회 동래읍성역사축제", "기간: 2026. 10. 16.(금) ~ 10. 18.(일), 3일간", "장소: 동래읍성지 북문광장 · 동래문화회관 · 명륜1번가 · 온천장 일원", "주제: '1592년 조선, 동래를 만나다'", "슬로건: '동래성 사람들, 역사와 미래를 잇다!'", "주요 관람층: 어린 자녀를 둔 가족, 청소년, 지역 주민, 외국인 관광객"],
    media: ["입구에서 받은 부채형 안내지도(앞면 포스터 + 뒷면 행사장 배치도)"],
    script: "동래읍성역사축제는 올해로 32번째를 맞은, 동래구의 대표 역사 축제예요. 올해는 10월 16일 금요일부터 18일 일요일까지 3일 동안 열렸고, 동래읍성지 북문광장을 중심으로 동래문화회관, 명륜1번가, 온천장 일대에서 진행됐어요. 입구에서 받은 이 부채 모양 안내지도에는 '동래성 사람들, 역사와 미래를 잇다'라는 문구가 적혀 있었어요. 뒷면에는 행사장 배치도가 있어서 부채로도 쓰고 지도로도 쓸 수 있었어요. 정말 실용적인 아이디어였어요.",
  },
  {
    kicker: "01 · 축제 소개", title: "왜 1592년 동래일까요?",
    bullets: ["1592. 4. 13.(음력) 왜군 부산포 상륙 — 임진왜란 시작", "4. 14. 부산진성 함락", "4. 15. 동래성 전투 — 동래부사 송상현과 군민이 끝까지 항전", "송상현의 답: 「戰死易 假道難」 싸워서 죽기는 쉬워도 길을 빌려주기는 어렵다", "축제의 의미: 그날의 용기와 희생을 기억하고 다음 세대에 전하는 자리"],
    media: ["포토존의 '帥 · 戰死' 깃발 벽화 사진", "북문언덕(전투재현장) 안내판 사진"],
    script: "그렇다면 왜 하필 1592년일까요? 1592년은 임진왜란이 시작된 해예요. 음력 4월 13일에 왜군이 부산에 상륙했고, 다음 날 부산진성이 무너졌어요. 그리고 4월 15일, 이곳 동래성에서 큰 전투가 벌어졌어요. 그때 동래부사 송상현은 '싸워서 죽기는 쉬워도, 길을 빌려주기는 어렵다'는 말을 남기고 끝까지 성을 지켰어요. 군인뿐만 아니라 백성들도 함께 싸웠다고 해요. 축제 포토존에도 이 '전사' 깃발이 그려져 있었는데요, 이 축제는 바로 그날의 용기를 기억하는 축제예요.",
  },
  {
    kicker: "01 · 축제 소개", title: "주요 프로그램 6가지",
    bullets: ["동래부사행차 길놀이 — 부사 행차를 재현한 거리 퍼레이드", "동래 세가닥 줄다리기 — 세 가닥 줄을 엮어 당기는 동래 전통 민속놀이", "동래성전투재현 뮤지컬 '외로운 성' — 축제의 핵심, 북문 언덕 야외무대", "관문체험 — 조선 시대 성문(관문)을 통과하는 참여형 체험", "동래장터체험 — 옛 장터 분위기의 먹거리 · 공예품 부스", "24가지 체험 — 꼬마 역사가, 장영실 과학체험, 한복 체험, 친환경 솜사탕 만들기 등"],
    media: ["프로그램별 대표 사진(전통 연희 공연, 장영실 과학체험 천막, 장터 부스)"],
    script: "축제의 주요 프로그램은 크게 여섯 가지예요. 동래부사 행차를 재현한 길놀이, 세 가닥 줄을 엮어서 당기는 동래 세가닥 줄다리기, 그리고 이번 축제의 핵심인 동래성 전투 재현 뮤지컬 '외로운 성'이 있어요. 그 밖에도 관문체험과 동래장터체험이 있고요. 아이들이 '꼬마 역사가'가 되어 보는 활동, 장영실 과학체험, 한복 체험, 친환경 솜사탕 만들기처럼 무려 24가지 체험 프로그램이 준비되어 있었어요.",
  },
  {
    kicker: "02 · 현장 답사", title: "제가 걸은 길 — 답사 동선",
    bullets: ["13:00 도착 — 언덕길 벽면의 역사 인물 · 동래성 그림", "입구 — 부채형 안내지도 수령", "체험존 — 장영실 과학체험, 소원지 터널", "장터 · 먹거리 — 4구역 공예품 부스, 동래읍성빵", "해 질 녘 — 북문 언덕 뮤지컬 '외로운 성' 관람", "밤 — 등불이 켜진 길을 따라 귀가"],
    media: ["언덕길 벽화, 장영실 체험 천막, 야간 공연, 등불길 사진 4장"],
    script: "이제 제가 실제로 움직인 동선을 따라가 볼게요. 오후 1시, 언덕길을 올라가는데 벽에 역사 인물과 동래성 그림이 있어서, 걸으면서 미리 공부하는 느낌이 들었어요. 입구에서는 부채 안내지도를 받았고요. 그다음 장영실 과학체험과 소원지 터널을 지나서 장터와 먹거리 구역을 구경했어요. 해가 진 뒤에는 북문 언덕에서 뮤지컬을 보고, 등불이 켜진 길을 따라 내려왔어요.",
  },
  {
    kicker: "02 · 현장 답사", title: "첫인상: 작은 배려가 만든 큰 감동",
    bullets: ["소리 — 아치형 터널에 걸린 아이들의 소원지가 바람에 사각사각", "배려 — 공연 관람객에게 방석 무료 배부", "청결 — 흙길을 걸은 뒤 신발 먼지를 터는 에어 먼지제거기"],
    media: ["소원지 터널, 방석 배부 현장, 에어 먼지제거기 사진", "가능하면 소원지 소리가 담긴 5초 영상"],
    script: "축제에 들어가자마자 작은 디테일들이 제 마음을 움직였어요. 첫 번째는 '소리'예요. 아치형 터널에 아이들이 쓴 소원지가 가득 걸려 있었는데, 바람이 불 때마다 종이가 사각사각 소리를 냈어요. 그 소리가 정말 오래 기억에 남았어요. 두 번째는 '배려'예요. 공연을 볼 때 바닥에 앉아야 하는데, 스태프분들이 방석을 나눠 주셨어요. 세 번째는 '청결'이에요. 흙길을 걷고 나면 신발에 먼지가 많이 묻는데, 에어 먼지제거기가 있어서 깨끗하게 털 수 있었어요.",
  },
  {
    kicker: "03 · 좋았던 점", title: "축제의 영혼, 뮤지컬 '외로운 성'",
    bullets: ["장소: 북문 언덕 야외 공연장 — 실제 '전투재현장'으로 지정된 곳", "내용: 1592년 송상현 부사와 백성들의 마지막 항전", "연출: 북소리 · 조명 · 대규모 출연진 → 강한 몰입감", "반응: 언덕 객석이 가득 참, 어린이도 끝까지 집중", "교육 효과: 교과서 속 한 줄이 '눈앞의 장면'이 됨"],
    media: ["야간 공연 사진 3장(전체 무대, 파란 조명 장면, 가득 찬 객석)", "추가 권장: 본인이 찍은 공연 영상 10~15초 또는 네이버TV · 유튜브 공식 영상 캡처"],
    script: "이제 이 축제의 영혼, 동래성 전투 재현 뮤지컬 '외로운 성'을 소개할게요. 이 공연은 북문 언덕의 야외 공연장에서 열렸어요. 이곳은 실제로 '전투재현장'으로 지정된 곳이라서, 그 자리에 앉아 있는 것만으로도 마음이 무거워졌어요. 송상현 부사와 백성들이 성을 지키는 장면에서는 북소리와 조명이 어우러져서 몰입감이 정말 컸어요. 제 옆에 있던 어린아이도 끝까지 집중해서 봤어요. 교과서 속 한 줄이 눈앞의 장면이 되니까, 아이들도 역사를 쉽게 이해할 수 있겠다고 느꼈어요.",
  },
  {
    kicker: "03 · 좋았던 점", title: "직접 못 본 프로그램은 네이버 자료로 보완",
    bullets: ["동래부사행차 길놀이 — 시간이 맞지 않음 → 네이버 '동래부사행차 길놀이' 사진 · 영상, 관람 후기", "동래 세가닥 줄다리기 — 위치를 찾지 못함 → 네이버 뉴스 기사 · 참가자 블로그 후기", "한복 체험 '청춘, 한복에 반하다' — 대기 줄이 너무 김 → 체험 후기(소요 시간, 만족도)", "후기에서 확인할 것: 관람객 반응, 대기 시간, 아이들 반응, 좋았던 점 · 불편한 점"],
    media: ["각 칸에 네이버 블로그 · 뉴스 캡처 1장씩(출처 URL과 날짜 표기)", "후기 인용은 한 줄로 짧게, 작성자 닉네임 · 날짜 함께 표기"],
    script: "아쉽게도 모든 프로그램에 참여하지는 못했어요. 동래부사행차 길놀이는 시간이 맞지 않았고, 세가닥 줄다리기는 위치를 찾지 못해서 놓쳤어요. 한복 체험은 줄이 너무 길었고요. 그래서 이 부분은 네이버 블로그 후기, 뉴스 기사, 영상을 찾아서 보완했어요. [네이버 후기에서 확인한 관람객 반응을 한두 문장으로 넣으세요.] 제가 직접 경험한 것과 다른 사람들의 후기를 함께 보니까, 축제를 더 넓게 이해할 수 있었어요.",
  },
  {
    kicker: "03 · 좋았던 점", title: "아이들이 '작은 역사가'가 되는 축제",
    bullets: ["장영실 과학체험전: 앙부일구(해시계) · 혼천의 모형 앞에서 해설사가 눈높이 설명", "과거와 현재의 연결: 수표 · 앙부일구 바로 옆에 오늘날 자동기상관측장비", "직접 해 보는 역사: 활동지 풀기, 소원지 쓰기, 만들기 체험", "학습 포인트: '보는' 역사 → '해 보는' 역사"],
    media: ["앙부일구 해설 사진, 혼천의 앞 아이들 사진, 어린이 활동지 체험 사진, 자동기상관측장비 사진"],
    script: "이 축제가 특히 좋았던 이유는 '교육'이에요. 장영실 과학체험전에서는 해시계인 앙부일구와 천체 관측 기구인 혼천의 모형 앞에서 해설사분이 아이들 눈높이에 맞게 설명해 주셨어요. 재미있는 점은 바로 옆에 오늘날의 자동기상관측장비가 있었다는 거예요. 약 600년 전 장영실의 과학이 지금까지 이어지고 있다는 걸 한눈에 볼 수 있었어요. 아이들은 활동지를 풀고 소원지를 쓰면서, '보는' 역사가 아니라 '직접 해 보는' 역사를 배우고 있었어요.",
  },
  {
    kicker: "03 · 좋았던 점", title: "가족을 위한 세심한 준비",
    bullets: ["유모차 보관 구역 — 입구 쪽 전용 공간, 포토존 앞에도 마련", "언덕 그늘 쉼터 — 돗자리를 펴고 공연을 기다리는 가족들", "분리수거함 · 지진 옥외대피장소 표지판이 눈에 잘 띔", "가족 단위 방문객이 많은 축제의 특성을 잘 반영"],
    media: ["유모차 보관소 사진 2장, 언덕 쉼터 사진, 지진 옥외대피장소 표지판 사진"],
    script: "이 축제에는 가족 단위 방문객, 특히 어린아이와 함께 온 가족이 정말 많았어요. 그래서인지 유모차 보관 구역이 따로 마련되어 있었어요. 포토존 바로 앞에도 유모차를 세워 둘 수 있었고요. 언덕 위 그늘에는 돗자리를 펴고 쉴 수 있는 공간이 있어서, 부모님들이 공연을 편하게 기다릴 수 있었어요. 또 분리수거함과 지진 옥외대피장소 표지판도 눈에 잘 띄었어요.",
  },
  {
    kicker: "03 · 좋았던 점", title: "장터의 정(情) — 냉장고 자석 이야기",
    bullets: ["4구역 장터: 직접 그림을 그린 냉장고 자석 부스", "처음 가격: 1개 2,000원 → 사장님: \"두 개 사면 3,000원이에요\"", "다시 찾아가 3개 선택 → 아주머니: \"다시 와 줘서 고마워요\" → 3개 3,000원", "느낀 점: 축제의 기억은 결국 '사람'이 만든다"],
    media: ["직접 산 냉장고 자석 3개 사진(손바닥 위)"],
    script: "이번 축제에서 제일 따뜻했던 순간을 이야기해 볼게요. 4구역 장터에서 직접 그림을 그린 냉장고 자석을 파는 가게를 발견했어요. 하나에 2,000원이었는데, 사장님이 웃으시면서 \"두 개 사면 3,000원이에요\"라고 하셨어요. 다른 곳을 구경하다가 자꾸 생각이 나서 다시 돌아갔어요. 제가 세 개를 고르니까, 아주머니께서 다시 와 줘서 고맙다고 하시면서 세 개를 3,000원에 주셨어요. 바로 이 자석들이에요. 저는 이때, 축제의 기억은 결국 '사람'이 만든다는 걸 느꼈어요.",
  },
  {
    kicker: "03 · 좋았던 점", title: "먹거리: 기대와 반전",
    bullets: ["먹거리 부스는 지역 봉사단체가 운영, 수익금은 봉사 활동에 사용", "닭꼬치 3,000원 — 줄이 길었지만 맛은 기대 이하 (△)", "동래읍성빵(십원빵 스타일) 팥 · 치즈 2,500원 — 쭉 늘어나는 치즈, 강력 추천 (◎)", "아쉬움: 인기 메뉴 대기 줄이 길어 번호표가 필요"],
    media: ["닭꼬치 부스 대기 줄 사진, 읍성제빵소 가격표 사진, 봉사단체 먹거리 부스 사진"],
    script: "먹거리도 빼놓을 수 없죠. 장터 먹거리 부스는 지역 봉사단체가 운영하고, 수익금은 봉사 활동에 쓰인다고 해서 더 의미가 있었어요. 닭꼬치는 줄을 오래 섰는데, 솔직히 맛은 기대만큼은 아니었어요. 하지만 동래읍성빵은 정말 맛있었어요! 동전 모양 빵 안에서 치즈가 쭉 늘어나서, 줄을 선 보람이 있었어요. 팥 맛과 치즈 맛이 있었고, 가격은 2,500원이었어요.",
  },
  {
    kicker: "04 · 아쉬운 점과 개선 방안", title: "현장에서 느낀 아쉬운 점 5가지",
    bullets: ["① 미끄러운 흙길 · 경사로 — 넘어지는 사람을 여러 번 목격", "② 줄다리기(2구역) 위치 안내 실패 — 자원봉사자 2명 모두 모름 → 참여 못 함", "③ 분실물 — 외투 분실 후 이름 · 연락처 등록, 아직 연락 없음", "④ 정수기 물 소진 — 한동안 교체하는 사람이 없음", "⑤ 한복 체험 — 대기 줄이 너무 길어 포기"],
    media: ["미끄러운 흙길 사진 2장, 빈 정수기 사진, 한복 체험 부스 사진"],
    script: "물론 아쉬운 점도 있었어요. 첫째, 흙길과 경사로가 미끄러워서 넘어지는 사람을 여러 번 봤어요. 둘째, 2구역에서 줄다리기에 참여하고 싶어서 자원봉사자 두 분께 여쭤봤는데, 두 분 다 위치를 모르셔서 결국 놓쳤어요. 셋째, 저는 외투를 잃어버려서 이름과 전화번호를 적고 분실물 신고를 했는데, 아직까지 연락이 없어요. 넷째, 정수기에 물이 떨어졌는데 한동안 아무도 갈아 주지 않았어요. 마지막으로, 한복 체험은 대기 줄이 너무 길어서 포기했어요.",
  },
  {
    kicker: "04 · 아쉬운 점과 개선 방안", title: "문제 → 원인 → 개선 방안",
    table: [
      ["미끄러운 흙길", "경사로 + 흙 · 우드칩 바닥, 야간 조명 부족", "미끄럼 방지 매트, 로프 손잡이, 야간 조명, 경사로 안전요원"],
      ["줄다리기 안내 실패", "자원봉사자 사전 교육 · 정보 공유 부족", "사전 교육, 시간표 · 위치 안내 카드 지급, 무전 · 단체방 공유"],
      ["분실물 연락 없음", "접수 후 처리 절차 불명확", "접수 번호 발급, 종료 후 일괄 문자, 온라인 분실물 게시판"],
      ["정수기 물 소진", "관리 담당자 부재", "담당자 지정, 1시간 단위 점검표, 예비 물통 비치"],
      ["한복 체험 대기", "수요에 비해 의상 · 인력 부족", "온라인 사전 예약, 번호표 · 회차제, 예상 대기 시간 안내판"],
    ],
    bullets: [],
    media: ["표 형식(문제 / 원인 / 개선 방안), 아이콘으로 각 행 구분"],
    script: "그래서 저는 문제마다 원인을 생각해 보고 개선 방안을 정리해 봤어요. 미끄러운 길에는 미끄럼 방지 매트와 손잡이 로프를 설치하고, 특히 밤에는 조명과 안전요원이 꼭 필요해요. 자원봉사자에게는 축제 전에 프로그램 시간과 위치를 교육하고, 한 장짜리 안내 카드를 나눠 주면 좋겠어요. 분실물은 접수 번호를 주고, 축제가 끝난 뒤에 문자로 결과를 꼭 알려 주는 시스템이 필요해요. 정수기는 담당자를 정해서 정해진 시간마다 확인하고, 한복 체험은 온라인 사전 예약이나 번호표로 운영하면 대기 시간을 줄일 수 있을 거예요.",
  },
  {
    kicker: "05 · 안전관리", title: "안전관리 점검표",
    checks: [
      ["동선 · 군중 관리", "good", "길이 잘 나뉘어 많은 인파에도 정체 거의 없음"],
      ["쓰레기 처리", "good", "분리수거함 곳곳 배치, 빨리 비워져 넘치지 않음"],
      ["화장실", "good", "이동식 화장실 청결 양호"],
      ["유모차 · 휴게 공간", "good", "전용 보관 구역, 언덕 그늘 쉼터"],
      ["급수 · 편의 시설", "mid", "정수기 물 소진 후 교체 지연"],
      ["바닥 · 보행 안전", "bad", "흙길 · 경사로 미끄러움, 넘어짐 다수"],
      ["분실물 관리", "bad", "접수 후 결과 안내 없음"],
    ],
    bullets: ["○ 잘됨: 동선 · 군중 관리, 쓰레기 처리, 화장실, 유모차 · 휴게 공간", "△ 보통: 급수 · 편의 시설", "× 개선 필요: 바닥 · 보행 안전, 분실물 관리"],
    media: ["근거 사진: 분리수거함, 이동식 화장실, 야간 귀가 인파"],
    script: "이번에는 안전관리를 항목별로 점검해 봤어요. 결과부터 말씀드리면, 동선과 군중 관리는 정말 잘 되어 있었어요. 사람이 많았지만 길이 잘 나뉘어 있어서 정체가 거의 없었어요. 쓰레기는 분리수거함이 금방 비워져서 넘치는 곳이 없었고, 화장실도 깨끗했어요. 유모차 보관 구역도 좋았고요. 반면에 정수기 같은 편의 시설은 '보통', 바닥 안전과 분실물 관리는 '개선 필요'로 평가했어요.",
  },
  {
    kicker: "05 · 안전관리", title: "서울세계불꽃축제에서 배우기",
    bullets: ["불꽃축제: 여의도 한강공원, 매년 약 100만 명 운집(주최 측 추산)", "주요 대책: 지하철 무정차 통과 · 도로 통제, 대규모 안전 인력 배치, 구역별 분산 귀가", "동래 적용 ① 대중교통 연계 — 동래역 · 명륜역 · 수안역 안내, 공연 후 분산 귀가", "동래 적용 ② 언덕 야간 안전 — 경사로 조명, 미끄럼 방지, 일방통행", "동래 적용 ③ 실시간 정보 — QR 지도로 혼잡도 · 대기 시간 · 분실물 공지"],
    media: ["네이버 뉴스: 서울세계불꽃축제 안전관리 기사 사진 1장(출처 표기)"],
    script: "수업 시간에 배운 서울세계불꽃축제 사례와도 비교해 봤어요. 불꽃축제에는 매년 100만 명 가까운 사람이 모이기 때문에, 지하철 무정차 통과와 도로 통제, 대규모 안전 인력 배치, 구역별로 나눠서 귀가하는 방법 등을 사용해요. 동래읍성역사축제는 규모는 작지만, 공연장이 언덕 위에 있어서 밤에 한꺼번에 내려올 때 위험할 수 있어요. 그래서 저는 세 가지를 제안하고 싶어요. 첫째, 동래역과 명륜역 같은 가까운 지하철역을 안내하고 공연 후에 나눠서 귀가하도록 유도하는 것. 둘째, 경사로에 조명과 미끄럼 방지 시설을 두고 일방통행으로 운영하는 것. 셋째, QR 지도로 혼잡도와 대기 시간을 실시간으로 알려 주는 것이에요.",
  },
  {
    kicker: "06 · 맺음말", title: "AI는 이렇게 '구체적으로' 활용했어요",
    prompts: [
      "토요일 13시에 동래읍성역사축제에 갔어요. 흙길에서 넘어지는 사람이 많았고 정수기 물이 비어 있었어요. 이 문제들을 '문제-원인-개선 방안' 표로 정리해 주세요.",
      "서울세계불꽃축제의 군중 관리 방법을 언덕 위 야외 공연장이 있는 지역 축제에 적용한다면, 현실적인 대책 3가지는 무엇인가요?",
      "외국인 학생의 발표 대본이에요. 해요체로 자연스럽게 고쳐 주고, 맞춤법이 틀린 곳을 표시해 주세요.",
    ],
    bullets: ["질문 1: 현장 문제를 '문제-원인-개선 방안' 표로 정리 요청", "질문 2: 서울세계불꽃축제 사례를 언덕 위 공연장에 적용하는 대책 질문", "질문 3: 대본 해요체 다듬기 + 맞춤법 확인", "원칙: 사진 · 경험은 직접, AI는 정리와 표현 점검에만, 사실 정보는 안내지도 · 공식 자료로 재확인"],
    media: ["실제 AI 대화 화면 캡처 1~2장(질문이 보이도록)"],
    script: "이번 발표를 준비하면서 AI도 활용했어요. 교수님께서 말씀하신 것처럼, 질문을 최대한 구체적으로 하려고 노력했어요. 예를 들어 미끄러운 흙길이나 비어 있는 정수기 같은 문제를 '문제, 원인, 개선 방안' 표로 정리해 달라고 했고, 불꽃축제 사례를 언덕 위 공연장에 적용하면 어떤 대책이 있는지도 물어봤어요. 대본의 맞춤법과 자연스러운 표현도 확인받았어요. 하지만 사진과 경험은 모두 제가 직접 겪은 것이고, 사실 정보는 안내지도와 공식 자료로 다시 확인했어요.",
  },
  {
    kicker: "06 · 맺음말", title: "총평: 역사를 '만나는' 축제",
    ratings: [["역사성", 5], ["교육성", 5], ["체험 · 재미", 4], ["편의 · 안전", 3]],
    bullets: ["역사성 5/5 — 실제 전투 현장에서 역사를 다시 만남", "교육성 5/5 — 장영실 과학체험, 꼬마 역사가 활동", "체험 · 재미 4/5 — 프로그램은 다양, 대기 줄 · 안내는 아쉬움", "편의 · 안전 3/5 — 기본 관리 우수, 바닥 안전 · 분실물 개선 필요", "한 줄 평: 역사를 '배우는' 곳이 아니라 '만나는' 곳"],
    media: ["5점 척도 막대그래프 + 한 줄 평"],
    script: "마지막으로 제 총평이에요. 역사성과 교육성은 5점 만점에 5점을 주고 싶어요. 실제 전투가 있었던 곳에서 그 역사를 다시 만나는 경험은 다른 축제에서는 하기 어려워요. 체험과 재미는 4점이에요. 프로그램은 다양했지만 대기 줄과 안내가 조금 아쉬웠어요. 편의와 안전은 3점이에요. 기본 관리는 좋았지만, 미끄러운 길과 분실물 처리는 꼭 개선되면 좋겠어요. 한 문장으로 정리하면, 이 축제는 역사를 '배우는' 곳이 아니라 '만나는' 곳이었어요.",
  },
  {
    kicker: "", title: "감사합니다",
    bullets: ["소원지 문구 인용: \"우리 가족 행복하고 항상 건강하기를\"", "메시지: 1592년 동래 사람들이 지키려던 것 = 평범한 가족의 행복", "권유: 내년 가을, 가족 · 친구와 함께 동래읍성역사축제에", "질의응답(Q&A)"],
    media: ["소원지 클로즈업 사진(배경, 어둡게 처리)"],
    script: "소원지 터널에서 이런 소원을 봤어요. \"우리 가족 행복하고 항상 건강하기를.\" 1592년 동래 사람들이 지키고 싶었던 것도 결국 이런 평범한 가족의 행복이 아니었을까 하는 생각이 들었어요. 여러분도 내년 가을에는 가족이나 친구와 함께 동래읍성역사축제에 꼭 가 보세요. 지금까지 들어 주셔서 감사합니다. 궁금한 점이 있으면 편하게 질문해 주세요.",
  },
];

// ---------- deck ----------
(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.author = "레깟뜨엉 (Le Cat Tuong)";
  pres.title = "1592년 조선, 동래를 만나다 — 제32회 동래읍성역사축제 현장 답사";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;

  pres.defineSlideMaster({
    title: "DARK",
    background: { color: HEX.dk1 },
    objects: [],
  });
  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: HEX.lt1 },
    margin: [0.5, 0.6, 0.6, 0.6],
    objects: [
      { placeholder: { options: { name: "kicker", type: "body", x: 0.6, y: 0.3, w: 9, h: 0.35, fontSize: 12, bold: true, color: C.accent1, margin: 0, valign: "middle" }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.62, w: 12.1, h: 0.8, fontSize: 30, bold: true, color: C.text1, margin: 0, valign: "middle", align: "left" }, text: "" } },
      { text: { text: "제32회 동래읍성역사축제 현장 답사  |  레깟뜨엉", options: { x: 0.6, y: 7.0, w: 8, h: 0.3, fontSize: 10, color: C.accent5, margin: 0 } } },
    ],
    slideNumber: { x: 12.2, y: 7.0, w: 0.55, h: 0.3, fontSize: 10, color: C.accent5, align: "right", margin: 0 },
  });

  const sections = ["표지", "01 축제 소개", "02 현장 답사", "03 좋았던 점", "04 아쉬운 점과 개선 방안", "05 안전관리", "06 맺음말"];
  const secOf = [0, 0, 1, 1, 1, 1, 2, 2, 3, 3, 3, 3, 3, 3, 4, 4, 5, 5, 6, 6, 6];
  let lastSec = -1;
  function newSlide(i, master) {
    const sec = secOf[i];
    if (sec !== lastSec) { pres.addSection({ title: sections[sec] }); lastSec = sec; }
    const s = pres.addSlide({ masterName: master, sectionTitle: sections[sec] });
    if (master === "CONTENT") {
      s.addText(S[i].kicker, { placeholder: "kicker" });
      s.addText(S[i].title, { placeholder: "title" });
    }
    s.addNotes(S[i].script);
    return s;
  }
  const T = (s, text, o) => s.addText(text, { isTextBox: true, margin: 0, color: C.text1, fontSize: 14, valign: "top", ...o });
  const card = (s, x, y, w, h, color = C.background2, name) =>
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color }, line: { type: "none" }, objectName: name });
  const badge = (s, x, y, label, color = C.accent1, d = 0.5, fs = 14) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color }, line: { type: "none" } });
    s.addText(String(label), { isTextBox: true, x, y, w: d, h: d, align: "center", valign: "middle", fontSize: fs, bold: true, color: C.background1, margin: 0 });
  };
  const pic = async (s, n, x, y, w, h, opt = {}) => s.addImage({ data: await photo(n, w, h, opt), x, y, w, h, altText: opt.alt || "현장 사진", objectName: "photo-" + n });
  const cap = (s, text, x, y, w, o = {}) => T(s, text, { x, y, w, h: 0.3, fontSize: 11, color: C.accent4, ...o });
  const bulletList = (items, o = {}) => items.map((t, k) => ({ text: t, options: { bullet: { indent: 18 }, breakLine: k < items.length - 1, paraSpaceAfter: o.gap ?? 8 } }));
  const dashed = (s, x, y, w, h, label) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: C.background1 }, line: { color: C.accent5, width: 1.25, dashType: "dash" } });
    T(s, label, { x: x + 0.2, y, w: w - 0.4, h, align: "center", valign: "middle", fontSize: 12, color: C.accent4 });
  };

  // 1 — cover
  {
    const s = newSlide(0, "DARK");
    await pic(s, 34, 0, 0, W, H, { radius: 0, darken: 0.45, alt: "야간 뮤지컬 공연 사진" });
    T(s, "제32회 동래읍성역사축제 · 현장 답사 발표", { x: 0.8, y: 1.9, w: 10, h: 0.45, fontSize: 18, color: C.accent3, bold: true });
    s.addText("1592년 조선,\n동래를 만나다", { isTextBox: true, x: 0.8, y: 2.45, w: 11, h: 2.2, fontSize: 54, bold: true, color: C.background1, margin: 0, fontFace: THEME.headFontFace, valign: "top", lineSpacingMultiple: 1.05 });
    T(s, "역사 · 교육 · 체험이 하나 된 가족 축제", { x: 0.8, y: 4.75, w: 10, h: 0.5, fontSize: 22, color: C.background1 });
    T(s, "답사일 2026. 10. 17.(토) 13:00 ~ 야간 공연   |   발표자 레깟뜨엉 · 광운대학교 정보융합학부", { x: 0.8, y: 6.3, w: 11.5, h: 0.4, fontSize: 14, color: C.background2 });
  }

  // 2 — agenda
  {
    const s = newSlide(1, "CONTENT");
    const items = [["01", "축제 소개", "인증샷 · 개요 · 1592년 · 프로그램"], ["02", "현장 답사", "이동 동선 · 첫인상"], ["03", "좋았던 점", "뮤지컬 · 교육 · 가족 · 장터"], ["04", "아쉬운 점과 개선 방안", "문제 → 원인 → 해결"], ["05", "안전관리", "점검표 · 불꽃축제 비교"], ["06", "맺음말", "AI 활용 · 총평"]];
    items.forEach(([n, t, d], k) => {
      const col = k % 2, row = Math.floor(k / 2);
      const x = 0.6 + col * 3.75, y = 1.75 + row * 1.68;
      card(s, x, y, 3.5, 1.45);
      badge(s, x + 0.25, y + 0.25, n, k === 2 ? C.accent1 : C.text2, 0.55, 13);
      T(s, t, { x: x + 0.95, y: y + 0.25, w: 2.4, h: 0.55, fontSize: 17, bold: true, valign: "middle" });
      T(s, d, { x: x + 0.25, y: y + 0.92, w: 3.1, h: 0.4, fontSize: 12, color: C.accent4 });
    });
    await pic(s, 39, 8.3, 1.75, 4.43, 4.81, { alt: "축제장 전경" });
    cap(s, "언덕 위에서 내려다본 축제장", 8.3, 6.62, 4.43);
  }

  // 3 — check-in
  {
    const s = newSlide(2, "CONTENT");
    await pic(s, 1, 0.6, 1.65, 3.6, 4.8, { alt: "입구 아치 앞 인증샷" });
    await pic(s, 2, 4.45, 1.65, 3.6, 4.8, { alt: "친구들과 단체 인증샷" });
    cap(s, "입구 아치 앞 인증샷", 0.6, 6.52, 3.6);
    cap(s, "함께 간 친구들과 단체 사진", 4.45, 6.52, 3.6);
    card(s, 8.4, 1.65, 4.33, 4.8);
    T(s, "답사 정보", { x: 8.7, y: 1.9, w: 3.8, h: 0.45, fontSize: 18, bold: true, color: C.accent1 });
    const rows = [["일시", "2026. 10. 17.(토)\n13:00 도착 ~ 야간 공연"], ["장소", "동래읍성지 북문광장 일대\n(부산 동래구)"], ["동행", "친구들과 함께"], ["증빙", "얼굴 + 축제 간판 인증샷"]];
    rows.forEach(([k, v], j) => {
      const y = 2.55 + j * 0.95;
      T(s, k, { x: 8.7, y, w: 0.8, h: 0.8, fontSize: 14, bold: true, color: C.text2 });
      T(s, v, { x: 9.5, y, w: 3.05, h: 0.8, fontSize: 14 });
    });
  }

  // 4 — overview
  {
    const s = newSlide(3, "CONTENT");
    const rows = [["행사명", "제32회 동래읍성역사축제"], ["기간", "2026. 10. 16.(금) ~ 10. 18.(일), 3일간"], ["장소", "동래읍성지 북문광장 · 동래문화회관 ·\n명륜1번가 · 온천장 일원"], ["주제", "1592년 조선, 동래를 만나다"], ["슬로건", "동래성 사람들, 역사와 미래를 잇다!"], ["관람층", "어린 자녀를 둔 가족 · 청소년 · 주민 · 외국인"]];
    rows.forEach(([k, v], j) => {
      const y = 1.7 + j * 0.85;
      card(s, 0.6, y, 6.9, 0.72, j === 3 ? C.accent6 : C.background2);
      T(s, k, { x: 0.85, y, w: 1.2, h: 0.72, valign: "middle", bold: true, fontSize: 14, color: j === 3 ? C.accent1 : C.text2 });
      T(s, v, { x: 2.1, y, w: 5.3, h: 0.72, valign: "middle", fontSize: 14, bold: j === 3 });
    });
    await pic(s, 31, 7.85, 1.7, 4.88, 3.66, { alt: "부채형 안내지도" });
    cap(s, "입구에서 받은 부채형 안내지도 — 앞면 포스터, 뒷면 행사장 배치도", 7.85, 5.45, 4.88, { h: 0.55 });
    T(s, "부채 + 지도 = 1석 2조", { x: 7.85, y: 6.05, w: 4.88, h: 0.4, fontSize: 16, bold: true, color: C.accent1 });
  }

  // 5 — history
  {
    const s = newSlide(4, "CONTENT");
    const tl = [["4. 13.", "왜군 부산포 상륙", "임진왜란 시작"], ["4. 14.", "부산진성 함락", "하루 만에 무너짐"], ["4. 15.", "동래성 전투", "송상현 부사와 군민의 항전"]];
    s.addShape(pres.shapes.LINE, { x: 0.85, y: 2.2, w: 6.6, h: 0, line: { color: C.accent5, width: 2 } });
    tl.forEach(([d, t, sub], k) => {
      const x = 0.6 + k * 2.3;
      badge(s, x + 0.0, 1.95, k + 1, k === 2 ? C.accent1 : C.text2, 0.5, 13);
      T(s, d + " (1592, 음력)", { x, y: 2.6, w: 2.2, h: 0.35, fontSize: 12, color: C.accent4 });
      T(s, t, { x, y: 2.95, w: 2.2, h: 0.4, fontSize: 16, bold: true, color: k === 2 ? C.accent1 : C.text1 });
      T(s, sub, { x, y: 3.35, w: 2.1, h: 0.6, fontSize: 12, color: C.accent4 });
    });
    card(s, 0.6, 4.25, 6.9, 2.35, C.text1);
    T(s, "戰死易 假道難", { x: 0.95, y: 4.45, w: 6.3, h: 0.7, fontSize: 32, bold: true, color: C.accent3 });
    T(s, "\"싸워서 죽기는 쉬워도, 길을 빌려주기는 어렵다\"", { x: 0.95, y: 5.2, w: 6.3, h: 0.45, fontSize: 16, color: C.background1 });
    T(s, "— 동래부사 송상현. 군인과 백성이 함께 끝까지 성을 지킴", { x: 0.95, y: 5.75, w: 6.3, h: 0.6, fontSize: 12, color: C.background2 });
    await pic(s, 45, 7.85, 1.7, 2.35, 4.9, { pos: "top", alt: "전사 깃발 포토존" });
    await pic(s, 29, 10.38, 1.7, 2.35, 4.9, { alt: "북문언덕 전투재현장 안내판" });
    cap(s, "포토존의 '帥 · 戰死' 깃발", 7.85, 6.65, 2.35);
    cap(s, "북문언덕(전투재현장)", 10.38, 6.65, 2.35);
  }

  // 6 — programs
  {
    const s = newSlide(5, "CONTENT");
    const P = [["동래부사행차 길놀이", "부사 행차를 재현한 거리 퍼레이드"], ["동래 세가닥 줄다리기", "세 가닥 줄을 엮어 당기는 동래 전통 민속놀이"], ["뮤지컬 '외로운 성'", "동래성 전투 재현 · 축제의 핵심 (북문 언덕)"], ["관문체험", "조선 시대 성문을 통과하는 참여형 체험"], ["동래장터체험", "옛 장터 분위기의 먹거리 · 공예품 부스"], ["24가지 체험", "꼬마 역사가 · 장영실 과학 · 한복 · 친환경 솜사탕"]];
    P.forEach(([t, d], k) => {
      const col = k % 3, row = Math.floor(k / 3);
      const x = 0.6 + col * 2.75, y = 1.7 + row * 2.5;
      const hi = k === 2;
      card(s, x, y, 2.55, 2.3, hi ? C.accent1 : C.background2);
      badge(s, x + 0.25, y + 0.25, k + 1, hi ? C.accent3 : C.text2, 0.5, 13);
      T(s, t, { x: x + 0.25, y: y + 0.9, w: 2.1, h: 0.65, fontSize: 15, bold: true, color: hi ? C.background1 : C.text1 });
      T(s, d, { x: x + 0.25, y: y + 1.5, w: 2.1, h: 0.75, fontSize: 12, color: hi ? C.background1 : C.accent4 });
    });
    await pic(s, 21, 8.95, 1.7, 3.78, 2.3, { alt: "전통 연희 공연" });
    cap(s, "북문광장 전통 연희 공연", 8.95, 4.05, 3.78);
    await pic(s, 9, 8.95, 4.45, 3.78, 2.05, { alt: "장영실 과학체험 천막" });
    cap(s, "장영실 과학체험전 천막", 8.95, 6.55, 3.78);
  }

  // 7 — route
  {
    const s = newSlide(6, "CONTENT");
    const steps = [["13:00", "언덕길 역사 벽화", 3], ["입구", "부채형 안내지도", 31], ["체험존", "장영실 · 소원지 터널", 16], ["장터", "4구역 · 먹거리", 46], ["해 질 녘", "뮤지컬 '외로운 성'", 33], ["밤", "등불길 따라 귀가", 32]];
    const cw = 1.9, gap = 0.142;
    for (let k = 0; k < steps.length; k++) {
      const [t, d, n] = steps[k];
      const x = 0.6 + k * (cw + gap);
      await pic(s, n, x, 1.75, cw, 2.9, { alt: d });
      badge(s, x + 0.15, 1.9, k + 1, k === 4 ? C.accent1 : C.text2, 0.45, 12);
      T(s, t, { x, y: 4.8, w: cw, h: 0.4, fontSize: 15, bold: true, color: k === 4 ? C.accent1 : C.text2 });
      T(s, d, { x, y: 5.2, w: cw, h: 0.6, fontSize: 13 });
    }
    card(s, 0.6, 6.0, 12.13, 0.7, C.background2);
    T(s, "낮에는 체험 · 장터, 밤에는 공연 · 등불 — 같은 장소가 시간대별로 전혀 다른 축제가 됨", { x: 0.85, y: 6.0, w: 11.6, h: 0.7, valign: "middle", fontSize: 14 });
  }

  // 8 — first impressions
  {
    const s = newSlide(7, "CONTENT");
    const C3 = [["소리", "소원지 터널", "바람이 불 때마다 아이들의 소원지가 사각사각", 16], ["배려", "방석 무료 배부", "바닥에 앉는 관람객을 위한 작은 준비", 25], ["청결", "에어 먼지제거기", "흙길을 걸은 뒤 신발 먼지를 깨끗하게", 19]];
    for (let k = 0; k < 3; k++) {
      const [tag, t, d, n] = C3[k];
      const x = 0.6 + k * 4.12;
      await pic(s, n, x, 1.7, 3.88, 3.2, { alt: t });
      badge(s, x + 0.2, 1.85, tag, C.accent1, 0.85, 13);
      T(s, t, { x, y: 5.05, w: 3.88, h: 0.45, fontSize: 18, bold: true });
      T(s, d, { x, y: 5.5, w: 3.88, h: 0.7, fontSize: 14, color: C.accent4 });
    }
  }

  // 9 — musical
  {
    const s = newSlide(8, "CONTENT");
    await pic(s, 34, 0.6, 1.7, 7.3, 3.5, { alt: "뮤지컬 외로운 성 장면" });
    await pic(s, 33, 0.6, 5.35, 3.55, 1.45, { pos: "centre", alt: "야간 공연 무대" });
    await pic(s, 36, 4.35, 5.35, 3.55, 1.45, { pos: "centre", alt: "가득 찬 객석" });
    const pts = [["장소", "북문 언덕 야외 공연장 — 실제 '전투재현장'"], ["내용", "송상현 부사와 백성들의 마지막 항전"], ["연출", "북소리 · 조명 · 대규모 출연진의 몰입감"], ["반응", "언덕 객석 만석, 어린이도 끝까지 집중"]];
    pts.forEach(([k, v], j) => {
      const y = 1.7 + j * 0.92;
      T(s, k, { x: 8.25, y, w: 0.85, h: 0.8, fontSize: 14, bold: true, color: C.accent1 });
      T(s, v, { x: 9.1, y, w: 3.63, h: 0.8, fontSize: 14 });
    });
    card(s, 8.25, 5.45, 4.48, 1.35, C.text1);
    T(s, "교과서 속 한 줄이\n'눈앞의 장면'이 되는 역사 교육", { x: 8.5, y: 5.45, w: 4.0, h: 1.35, valign: "middle", fontSize: 16, bold: true, color: C.background1 });
  }

  // 10 — Naver supplement
  {
    const s = newSlide(9, "CONTENT");
    const cols = [["동래부사행차 길놀이", "시간이 맞지 않음", "동래부사행차 길놀이 2026"], ["동래 세가닥 줄다리기", "위치를 찾지 못함", "동래 세가닥 줄다리기 동래읍성"], ["한복 체험", "대기 줄이 너무 김", "동래읍성역사축제 한복 체험 후기"]];
    cols.forEach(([t, why, kw], k) => {
      const x = 0.6 + k * 4.12;
      T(s, t, { x, y: 1.65, w: 3.88, h: 0.45, fontSize: 18, bold: true });
      T(s, "못 본 이유: " + why, { x, y: 2.1, w: 3.88, h: 0.35, fontSize: 12, color: C.accent1 });
      dashed(s, x, 2.55, 3.88, 2.55, "네이버 사진 · 영상 캡처 삽입\n(출처 URL · 날짜 표기)");
      card(s, x, 5.25, 3.88, 0.95);
      T(s, "검색어", { x: x + 0.2, y: 5.33, w: 3.5, h: 0.3, fontSize: 11, bold: true, color: C.text2 });
      T(s, kw, { x: x + 0.2, y: 5.63, w: 3.5, h: 0.5, fontSize: 13 });
    });
    T(s, "후기에서 확인할 것: 관람객 반응 · 대기 시간 · 아이들 반응 · 좋았던 점과 불편한 점", { x: 0.6, y: 6.35, w: 12.1, h: 0.4, fontSize: 14, color: C.accent4 });
  }

  // 11 — education
  {
    const s = newSlide(10, "CONTENT");
    await pic(s, 18, 0.6, 1.7, 2.9, 3.0, { alt: "앙부일구 해설" });
    await pic(s, 8, 3.65, 1.7, 2.9, 3.0, { alt: "혼천의 앞 어린이" });
    await pic(s, 17, 0.6, 4.85, 2.9, 1.95, { alt: "어린이 활동지 체험" });
    await pic(s, 7, 3.65, 4.85, 2.9, 1.95, { alt: "자동기상관측장비", pos: "bottom" });
    const pts = [["장영실 과학체험전", "앙부일구 · 혼천의 모형 앞에서 해설사의 눈높이 설명"], ["과거 ↔ 현재", "약 600년 전 과학 기구 바로 옆에 오늘날 자동기상관측장비"], ["직접 해 보는 역사", "활동지 풀기 · 소원지 쓰기 · 만들기 체험"]];
    pts.forEach(([t, d], j) => {
      const y = 1.7 + j * 1.3;
      badge(s, 6.95, y + 0.05, j + 1, C.text2, 0.5, 13);
      T(s, t, { x: 7.6, y, w: 5.1, h: 0.45, fontSize: 17, bold: true });
      T(s, d, { x: 7.6, y: y + 0.45, w: 5.1, h: 0.7, fontSize: 14, color: C.accent4 });
    });
    card(s, 6.95, 5.75, 5.78, 1.05, C.accent6);
    T(s, "'보는' 역사  →  '해 보는' 역사", { x: 7.2, y: 5.75, w: 5.3, h: 1.05, valign: "middle", fontSize: 20, bold: true, color: C.accent1 });
  }

  // 12 — family
  {
    const s = newSlide(11, "CONTENT");
    await pic(s, 27, 0.6, 1.7, 2.95, 4.0, { extract: { left: 80, top: 0, width: 405, height: 712 }, alt: "유모차 보관 구역" });
    await pic(s, 26, 3.7, 1.7, 2.95, 4.0, { alt: "포토존 앞 유모차" });
    await pic(s, 12, 6.8, 1.7, 2.95, 4.0, { alt: "언덕 그늘 쉼터" });
    await pic(s, 24, 9.9, 1.7, 2.83, 4.0, { alt: "지진 옥외대피장소와 분리수거함" });
    const caps = ["유모차 보관 구역", "포토존 앞 유모차 자리", "언덕 그늘 쉼터", "대피장소 표지 · 분리수거함"];
    caps.forEach((c, k) => cap(s, c, 0.6 + k * 3.1, 5.78, 2.95, { fontSize: 13, bold: true, color: C.text1 }));
    card(s, 0.6, 6.2, 12.13, 0.65, C.background2);
    T(s, "어린 자녀를 둔 가족이 많은 축제 특성을 정확히 읽은 준비", { x: 0.85, y: 6.2, w: 11.6, h: 0.65, valign: "middle", fontSize: 14 });
  }

  // 13 — magnets
  {
    const s = newSlide(12, "CONTENT");
    await pic(s, 14, 0.6, 1.7, 3.9, 5.1, { alt: "직접 산 냉장고 자석 3개" });
    const st = [["1", "1개 2,000원", "4구역 장터, 손으로 그린 냉장고 자석"], ["2", "\"두 개 사면 3,000원이에요\"", "웃으며 먼저 깎아 주신 사장님"], ["3", "다시 찾아가 3개 선택", "구경하다가 자꾸 생각나서 돌아감"], ["4", "\"다시 와 줘서 고마워요\"", "아주머니: 3개 모두 3,000원"]];
    st.forEach(([n, t, d], j) => {
      const y = 1.7 + j * 1.0;
      badge(s, 4.85, y + 0.05, n, j === 3 ? C.accent1 : C.text2, 0.5, 13);
      T(s, t, { x: 5.5, y, w: 3.9, h: 0.42, fontSize: 16, bold: true, color: j === 3 ? C.accent1 : C.text1 });
      T(s, d, { x: 5.5, y: y + 0.42, w: 3.9, h: 0.4, fontSize: 13, color: C.accent4 });
    });
    card(s, 9.75, 1.7, 2.98, 3.8, C.accent1);
    T(s, "3개", { x: 9.75, y: 2.0, w: 2.98, h: 1.0, align: "center", fontSize: 54, bold: true, color: C.background1 });
    T(s, "=", { x: 9.75, y: 3.0, w: 2.98, h: 0.5, align: "center", fontSize: 24, color: C.background1 });
    T(s, "3,000원", { x: 9.75, y: 3.5, w: 2.98, h: 0.9, align: "center", fontSize: 40, bold: true, color: C.accent3 });
    T(s, "원래 6,000원", { x: 9.75, y: 4.55, w: 2.98, h: 0.4, align: "center", fontSize: 12, color: C.background2 });
    card(s, 4.85, 5.85, 7.88, 0.95, C.background2);
    T(s, "축제의 기억은 결국 '사람'이 만든다 — 장터의 정(情)", { x: 5.1, y: 5.85, w: 7.4, h: 0.95, valign: "middle", fontSize: 17, bold: true });
  }

  // 14 — food
  {
    const s = newSlide(13, "CONTENT");
    const F = [[5, "닭꼬치 3,000원", "△", "긴 대기 줄, 맛은 기대 이하", C.accent4], [46, "동래읍성빵 2,500원", "◎", "십원빵 스타일, 팥 · 치즈 — 치즈가 쭉! 강력 추천", C.accent1], [4, "봉사단체 먹거리 부스", "♡", "수익금은 지역 봉사 활동에 사용", C.text2]];
    for (let k = 0; k < 3; k++) {
      const [n, t, mark, d, col] = F[k];
      const x = 0.6 + k * 4.12;
      await pic(s, n, x, 1.7, 3.88, 3.3, { alt: t });
      badge(s, x + 3.2, 1.85, mark, col, 0.55, 16);
      T(s, t, { x, y: 5.15, w: 3.88, h: 0.45, fontSize: 18, bold: true, color: k === 1 ? C.accent1 : C.text1 });
      T(s, d, { x, y: 5.6, w: 3.88, h: 0.7, fontSize: 14, color: C.accent4 });
    }
    T(s, "개선 제안: 인기 메뉴는 번호표 · 진동벨로 대기 줄 분산", { x: 0.6, y: 6.4, w: 12.1, h: 0.4, fontSize: 14, color: C.text2, bold: true });
  }

  // 15 — issues
  {
    const s = newSlide(14, "CONTENT");
    const I = [[fa.FaPersonFalling, "미끄러운 흙길 · 경사로", "넘어지는 사람을 여러 번 목격"], [fa.FaCircleQuestion, "줄다리기 위치 안내 실패", "2구역 자원봉사자 2명 모두 모름 → 참여 못 함"], [fa.FaShirt, "분실물 연락 없음", "외투 분실, 이름 · 연락처 등록 후 아직 소식 없음"], [fa.FaBottleWater, "정수기 물 소진", "한동안 교체하는 사람이 없음"], [fa.FaHourglassHalf, "한복 체험 대기", "줄이 너무 길어 포기"]];
    for (let j = 0; j < I.length; j++) {
      const [ic, t, d] = I[j];
      const y = 1.65 + j * 1.04;
      s.addShape(pres.shapes.OVAL, { x: 0.6, y, w: 0.75, h: 0.75, fill: { color: C.accent6 }, line: { type: "none" } });
      s.addImage({ data: await icon(ic, HEX.accent1), x: 0.78, y: y + 0.18, w: 0.39, h: 0.39, altText: t });
      T(s, `${j + 1}. ${t}`, { x: 1.6, y: y - 0.02, w: 5.6, h: 0.42, fontSize: 17, bold: true });
      T(s, d, { x: 1.6, y: y + 0.4, w: 5.6, h: 0.4, fontSize: 14, color: C.accent4 });
    }
    await pic(s, 37, 7.55, 1.7, 2.5, 3.3, { alt: "미끄러운 흙길" });
    await pic(s, 23, 10.23, 1.7, 2.5, 3.3, { alt: "물이 빈 정수기" });
    await pic(s, 30, 7.55, 5.15, 5.18, 1.65, { extract: { left: 0, top: 430, width: 720, height: 270 }, alt: "한복 체험 부스" });
  }

  // 16 — table
  {
    const s = newSlide(15, "CONTENT");
    const hdr = ["문제", "원인", "개선 방안"].map((t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 }, fontSize: 15 } }));
    const rows = [hdr, ...S[15].table.map((r, k) => r.map((c, j) => ({ text: c, options: { fontSize: 14, bold: j === 0, color: j === 2 ? C.text2 : C.text1, fill: { color: k % 2 ? C.background1 : C.background2 } } })))];
    s.addTable(rows, { x: 0.6, y: 1.7, w: 12.13, colW: [2.6, 4.2, 5.33], rowH: 0.8, valign: "middle", margin: [0.06, 0.15, 0.06, 0.15], border: { type: "solid", pt: 0.75, color: "D5D9DE" } });
  }

  // 17 — safety checklist
  {
    const s = newSlide(16, "CONTENT");
    const mark = { good: ["○", "잘됨", C.text2], mid: ["△", "보통", C.accent3], bad: ["×", "개선 필요", C.accent1] };
    S[16].checks.forEach(([k, st, ev], j) => {
      const y = 1.68 + j * 0.74;
      card(s, 0.6, y, 8.0, 0.64, st === "bad" ? C.accent6 : C.background2);
      const [m, label, col] = mark[st];
      badge(s, 0.75, y + 0.09, m, col, 0.46, 14);
      T(s, k, { x: 1.4, y, w: 2.3, h: 0.64, valign: "middle", fontSize: 14, bold: true });
      T(s, label, { x: 3.7, y, w: 1.1, h: 0.64, valign: "middle", fontSize: 13, bold: true, color: col });
      T(s, ev, { x: 4.8, y, w: 3.7, h: 0.64, valign: "middle", fontSize: 12, color: C.accent4 });
    });
    await pic(s, 28, 8.95, 1.68, 1.82, 2.4, { alt: "분리수거함" });
    await pic(s, 11, 10.91, 1.68, 1.82, 2.4, { alt: "이동식 화장실" });
    await pic(s, 32, 8.95, 4.25, 3.78, 2.6, { alt: "야간 귀가 인파" });
    cap(s, "분리수거함 · 이동식 화장실 · 야간 귀가 동선", 8.95, 6.88, 3.78, { fontSize: 10 });
  }

  // 18 — fireworks comparison
  {
    const s = newSlide(17, "CONTENT");
    card(s, 0.6, 1.7, 4.6, 5.1, C.text1);
    T(s, "서울세계불꽃축제", { x: 0.9, y: 1.95, w: 4.0, h: 0.45, fontSize: 19, bold: true, color: C.accent3 });
    T(s, "여의도 한강공원 · 매년 약 100만 명 운집(주최 측 추산)", { x: 0.9, y: 2.45, w: 4.0, h: 0.7, fontSize: 13, color: C.background2 });
    s.addText(bulletList(["지하철 무정차 통과 · 도로 통제", "대규모 안전 인력 · 경찰 배치", "구역별 분산 귀가 안내"], { gap: 10 }), { isTextBox: true, x: 0.9, y: 3.3, w: 4.0, h: 1.6, fontSize: 14, color: C.background1, margin: 0, valign: "top" });
    dashed(s, 0.9, 5.0, 4.0, 1.55, "네이버 뉴스 사진 삽입\n(불꽃축제 안전관리 기사)");
    T(s, "동래에 적용한다면", { x: 5.55, y: 1.7, w: 7.0, h: 0.45, fontSize: 19, bold: true, color: C.accent1 });
    const A = [[fa.FaTrainSubway, "대중교통 연계", "동래역 · 명륜역 · 수안역 안내, 공연 후 분산 귀가 유도"], [fa.FaLightbulb, "언덕 야간 안전", "경사로 조명 · 미끄럼 방지 시설, 공연 후 일방통행"], [fa.FaQrcode, "실시간 정보", "QR 지도로 혼잡도 · 대기 시간 · 분실물 공지"]];
    for (let j = 0; j < 3; j++) {
      const [ic, t, d] = A[j];
      const y = 2.3 + j * 1.5;
      card(s, 5.55, y, 7.18, 1.3);
      s.addShape(pres.shapes.OVAL, { x: 5.8, y: y + 0.27, w: 0.75, h: 0.75, fill: { color: C.text2 }, line: { type: "none" } });
      s.addImage({ data: await icon(ic, "FFFFFF"), x: 5.98, y: y + 0.45, w: 0.39, h: 0.39, altText: t });
      T(s, t, { x: 6.8, y: y + 0.18, w: 5.7, h: 0.42, fontSize: 17, bold: true });
      T(s, d, { x: 6.8, y: y + 0.62, w: 5.7, h: 0.55, fontSize: 14, color: C.accent4 });
    }
  }

  // 19 — AI
  {
    const s = newSlide(18, "CONTENT");
    S[18].prompts.forEach((p, j) => {
      const y = 1.7 + j * 1.38;
      card(s, 0.6, y, 7.9, 1.22);
      badge(s, 0.8, y + 0.36, "Q" + (j + 1), C.text2, 0.5, 12);
      T(s, "\"" + p + "\"", { x: 1.5, y, w: 6.85, h: 1.22, valign: "middle", fontSize: 13 });
    });
    card(s, 8.85, 1.7, 3.88, 3.98, C.text1);
    T(s, "활용 원칙", { x: 9.1, y: 1.9, w: 3.4, h: 0.45, fontSize: 18, bold: true, color: C.accent3 });
    s.addText(bulletList(["사진 · 경험은 100% 직접", "AI는 정리 · 표현 점검에만", "사실 정보는 안내지도 · 공식 자료로 재확인", "대본은 소리 내어 여러 번 연습"], { gap: 10 }), { isTextBox: true, x: 9.1, y: 2.45, w: 3.4, h: 3.1, fontSize: 14, color: C.background1, margin: 0, valign: "top" });
    T(s, "구체적인 질문 = 성의 있는 준비", { x: 0.6, y: 6.0, w: 12.1, h: 0.5, fontSize: 18, bold: true, color: C.accent1 });
  }

  // 20 — rating
  {
    const s = newSlide(19, "CONTENT");
    const R = S[19].ratings;
    s.addChart(pres.charts.BAR, [{ name: "점수", labels: R.map((r) => r[0]), values: R.map((r) => r[1]) }], {
      x: 0.6, y: 1.65, w: 7.2, h: 4.4, barDir: "bar", catAxisOrientation: "maxMin",
      valAxisMinVal: 0, valAxisMaxVal: 5, valAxisMajorUnit: 1, valAxisHidden: true,
      chartColors: [HEX.accent1], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0\"/5\"",
      dataLabelFontSize: 14, dataLabelFontBold: true, dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt",
      catAxisLabelFontSize: 15, catAxisLabelColor: HEX.dk1, catAxisLabelFontFace: "+mn-lt", catAxisLineShow: false,
      valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false, barGapWidthPct: 60,
      showTitle: true, title: "개인 평가 (5점 만점)", titleFontSize: 14, titleColor: HEX.accent4, titleFontFace: "+mn-lt",
    });
    T(s, "역사 · 교육: 실제 전투 현장 + 과학체험\n체험 · 재미: 다양하지만 대기 · 안내 아쉬움\n편의 · 안전: 바닥 안전 · 분실물 개선 필요", { x: 0.6, y: 6.1, w: 7.2, h: 0.85, fontSize: 12, color: C.accent4 });
    await pic(s, 44, 8.15, 1.7, 4.58, 3.4, { alt: "소원지", pos: "top" });
    card(s, 8.15, 5.3, 4.58, 1.55, C.text1);
    T(s, "역사를 '배우는' 곳이 아니라\n'만나는' 곳", { x: 8.4, y: 5.3, w: 4.1, h: 1.55, valign: "middle", fontSize: 19, bold: true, color: C.background1 });
  }

  // 21 — closing
  {
    const s = newSlide(20, "DARK");
    await pic(s, 44, 0, 0, W, H, { radius: 0, darken: 0.3, alt: "소원지 배경" });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: H, fill: { color: C.text1, transparency: 30 }, line: { type: "none" } });
    T(s, "\"우리 가족 행복하고 항상 건강하기를\"", { x: 0.8, y: 1.6, w: 11.5, h: 0.7, fontSize: 26, bold: true, color: C.accent3 });
    T(s, "1592년 동래 사람들이 지키고 싶었던 것도\n이런 평범한 가족의 행복이 아니었을까요?", { x: 0.8, y: 2.45, w: 11.5, h: 1.1, fontSize: 20, color: C.background1 });
    s.addText("감사합니다", { isTextBox: true, x: 0.8, y: 4.1, w: 11.5, h: 1.2, fontSize: 60, bold: true, color: C.background1, margin: 0, fontFace: THEME.headFontFace });
    T(s, "Q & A   |   레깟뜨엉 · 광운대학교 정보융합학부", { x: 0.8, y: 5.6, w: 11.5, h: 0.45, fontSize: 16, color: C.background2 });
  }

  await pres.writeFile({ fileName: OUT });
  if (SKILL) {
    const { applyTheme } = require(path.join(SKILL, "scripts/apply_theme.js"));
    await applyTheme(OUT, THEME);
  }

  // ---------- markdown (Part 1 + Part 2) ----------
  let md = "# 1592년 조선, 동래를 만나다 — 발표 자료\n\n";
  md += "제32회 동래읍성역사축제 현장 답사 · 발표자 레깟뜨엉 · 슬라이드 파일: `Dongnae-Festival_Deck_v1.pptx` (각 슬라이드 메모에 대본 포함)\n\n";
  md += "## PART 1. 슬라이드 구성\n\n";
  S.forEach((x, i) => {
    md += `### Slide ${i + 1}. ${x.title}\n\n`;
    if (x.kicker) md += `- 섹션: ${x.kicker}\n`;
    const b = x.table ? x.table.map((r) => `${r[0]} — 원인: ${r[1]} / 개선: ${r[2]}`) : x.checks ? x.checks.map((c) => `${c[0]}: ${{ good: "○ 잘됨", mid: "△ 보통", bad: "× 개선 필요" }[c[1]]} — ${c[2]}`) : x.bullets;
    md += "- 내용:\n" + b.map((t) => `  - ${t}`).join("\n") + "\n";
    if (x.prompts) md += "- AI 질문 예시:\n" + x.prompts.map((p) => `  - "${p}"`).join("\n") + "\n";
    md += "- 이미지 · 영상:\n" + x.media.map((t) => `  - ${t}`).join("\n") + "\n\n";
  });
  md += "## PART 2. 발표 대본 (해요체)\n\n";
  S.forEach((x, i) => { md += `### Slide ${i + 1}. ${x.title}\n\n${x.script}\n\n`; });
  fs.writeFileSync(MD, md);
  console.log("done", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
