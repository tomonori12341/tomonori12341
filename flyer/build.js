const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.defineLayout({ name: "A4", width: 8.27, height: 11.69 });
pres.layout = "A4";
pres.title = "映像講義モニター募集チラシ";
const F = "Meiryo UI";
const DARK = "2A2117", GOLD = "F2B21B", RED = "E60012", WHITE = "FFFFFF", INK = "2A2117", GRAY = "6B6257", LIGHT = "FBF3DC";
const s = pres.addSlide();
s.background = { color: DARK };
const T = (t, o) => s.addText(t, Object.assign({ fontFace: F, isTextBox: true, margin: 0, valign: "middle" }, o));

// header
T("日建学院  ｜  映像講義モニター", { x: 0.5, y: 0.35, w: 5.2, h: 0.5, fontSize: 18, bold: true, color: WHITE });
s.addImage({ path: "nikken_logo.png", x: 6.07, y: 0.35, w: 1.7, h: 0.473, objectName: "日建学院ロゴ" });

// headline
T("映像講義を観て、意見を聞かせて！", { x: 0.5, y: 1.15, w: 7.27, h: 0.5, fontSize: 22, bold: true, color: WHITE, align: "center" });
T("モニター募集", { x: 0.5, y: 1.7, w: 7.27, h: 1.25, fontSize: 70, bold: true, color: GOLD, align: "center" });
// badges
const badges = [["建築学生限定", RED], ["限定10名", GOLD]];
badges.forEach((b, i) => {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.75 + i * 2.5, y: 3.1, w: 2.3, h: 0.6, rectRadius: 0.3, fill: { color: b[1] }, line: { color: b[1] }, objectName: "バッジ" + (i + 1) });
  T(b[0], { x: 1.75 + i * 2.5, y: 3.1, w: 2.3, h: 0.6, fontSize: 20, bold: true, color: b[1] === GOLD ? INK : WHITE, align: "center" });
});

// reward card
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 3.95, w: 7.27, h: 1.75, rectRadius: 0.2, fill: { color: WHITE }, line: { color: GOLD, width: 4 }, objectName: "報酬カード" });
T("2時間で", { x: 0.8, y: 4.15, w: 2.4, h: 1.35, fontSize: 34, bold: true, color: INK, align: "center" });
T([{ text: "5,000", options: { fontSize: 80, bold: true, color: RED } }, { text: "円", options: { fontSize: 38, bold: true, color: RED } }], { x: 3.1, y: 4.1, w: 4.4, h: 1.45, align: "center" });

// flow
T("参加の流れ", { x: 0.5, y: 5.95, w: 3, h: 0.4, fontSize: 18, bold: true, color: GOLD });
const steps = [["応募", "QRコードから\nお申し込み"], ["映像講義を視聴", "約2時間の\n講義映像"], ["意見を記入", "視聴後に\nご意見を回答"], ["報酬お渡し", "後日\n手渡し"]];
steps.forEach((st, i) => {
  const x = 0.5 + i * 1.89;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 6.4, w: 1.6, h: 1.55, rectRadius: 0.12, fill: { color: LIGHT }, line: { color: LIGHT }, objectName: "ステップ" + (i + 1) });
  s.addShape(pres.shapes.OVAL, { x: x + 0.1, y: 6.5, w: 0.4, h: 0.4, fill: { color: RED }, line: { color: RED }, objectName: "番号" + (i + 1) });
  T(String(i + 1), { x: x + 0.1, y: 6.5, w: 0.4, h: 0.4, fontSize: 14, bold: true, color: WHITE, align: "center" });
  T(st[0], { x: x + 0.08, y: 6.95, w: 1.44, h: 0.35, fontSize: 13, bold: true, color: INK, align: "center" });
  T(st[1], { x: x + 0.08, y: 7.3, w: 1.44, h: 0.6, fontSize: 11, color: GRAY, align: "center", valign: "top" });
  if (i < 3) T("▶", { x: x + 1.6, y: 6.95, w: 0.29, h: 0.5, fontSize: 12, color: GOLD, align: "center" });
});

// details
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 8.2, w: 7.27, h: 1.85, rectRadius: 0.15, fill: { color: WHITE }, line: { color: WHITE }, objectName: "募集要項カード" });
const rows = [["対象", "建築を学ぶ大学生"], ["日時", "【　　月　　日（　）　　：　　～　　：　　】"], ["場所", "【　　　　　　大学　　　　　教室　】"], ["報酬", "5,000円（後日手渡し）"], ["締切", "【　　月　　日（　）】　※定員になり次第終了"]];
rows.forEach((r, i) => {
  const y = 8.3 + i * 0.35;
  T(r[0], { x: 0.75, y, w: 0.9, h: 0.3, fontSize: 13, bold: true, color: RED });
  T(r[1], { x: 1.7, y, w: 5.9, h: 0.3, fontSize: 13, color: INK });
});

// QR + contact
s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 10.2, w: 1.3, h: 1.3, fill: { color: WHITE }, line: { color: GOLD, width: 2, dashType: "dash" }, objectName: "QRコード貼付枠" });
T("応募はこちら", { x: 2.05, y: 10.25, w: 5.7, h: 0.5, fontSize: 22, bold: true, color: GOLD });
T("QRコードを読み取ってお申し込みください\n【お問い合わせ先：日建学院　　校　TEL 　　　　　　】", { x: 2.05, y: 10.85, w: 5.7, h: 0.65, fontSize: 12, color: WHITE, valign: "top" });
pres.writeFile({ fileName: "monitor_flyer.pptx" }).then(() => console.log("ok"));
