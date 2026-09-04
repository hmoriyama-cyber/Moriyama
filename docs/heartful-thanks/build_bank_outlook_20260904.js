const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  AlignmentType, HeadingLevel, BorderStyle, ShadingType, LevelFormat, PageNumber,
  Footer, Header, VerticalAlign, TabStopType, PageBreak,
} = require("docx");

const MINCHO = "游明朝";
const GOTHIC = "游ゴシック";
const NAVY = "1F3A5F";
const GRAY = "F2F2F2";
const LIGHT = "E8EEF5";

// ---------- helpers ----------
const run = (text, opts = {}) =>
  new TextRun({ text, font: { ascii: opts.gothic ? GOTHIC : MINCHO, eastAsia: opts.gothic ? GOTHIC : MINCHO, hAnsi: opts.gothic ? GOTHIC : MINCHO }, size: opts.size || 21, bold: opts.bold, color: opts.color, italics: opts.italics });

const p = (text, opts = {}) =>
  new Paragraph({
    alignment: opts.align || AlignmentType.BOTH,
    spacing: { before: opts.before ?? 0, after: opts.after ?? 120, line: opts.line || 320 },
    indent: opts.indent,
    children: Array.isArray(text) ? text : [run(text, opts)],
  });

const note = (text) => p(text, { size: 17, color: "555555", after: 60, line: 260 });

const h1 = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 2 } },
    children: [run(text, { gothic: true, size: 26, bold: true, color: NAVY })],
  });

const h2 = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 100 },
    children: [run(text, { gothic: true, size: 22, bold: true, color: NAVY })],
  });

const bullet = (text, level = 0) =>
  new Paragraph({
    numbering: { reference: "bul", level },
    spacing: { after: 80, line: 300 },
    children: Array.isArray(text) ? text : [run(text)],
  });

const numbered = (ref, text) =>
  new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { after: 80, line: 300 },
    children: Array.isArray(text) ? text : [run(text)],
  });

const border = { style: BorderStyle.SINGLE, size: 4, color: "999999" };
const borders = { top: border, bottom: border, left: border, right: border };

function cell(text, width, opts = {}) {
  const lines = String(text).split("\n");
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders,
    verticalAlign: VerticalAlign.CENTER,
    shading: opts.fill ? { type: ShadingType.CLEAR, color: opts.fill, fill: opts.fill } : undefined,
    margins: { top: 50, bottom: 50, left: 90, right: 90 },
    children: lines.map((l) =>
      new Paragraph({
        alignment: opts.align || AlignmentType.LEFT,
        spacing: { after: 0, line: 260 },
        children: [run(l, { size: opts.size || 18, bold: opts.bold, gothic: opts.head, color: opts.color })],
      })
    ),
  });
}

// rows: array of arrays; aligns: per column ("L"|"R"|"C"); widths in DXA
function table(widths, header, rows, opts = {}) {
  const aligns = opts.aligns || widths.map((_, i) => (i === 0 ? "L" : "R"));
  const al = (a) => (a === "R" ? AlignmentType.RIGHT : a === "C" ? AlignmentType.CENTER : AlignmentType.LEFT);
  const trs = [];
  if (header) {
    trs.push(new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, widths[i], { fill: LIGHT, bold: true, head: true, align: AlignmentType.CENTER })) }));
  }
  rows.forEach((r) => {
    const isTotal = r.__total;
    const vals = r.__total ? r.vals : r;
    trs.push(new TableRow({ children: vals.map((v, i) => cell(v, widths[i], { align: al(aligns[i]), bold: isTotal, fill: isTotal ? GRAY : undefined })) }));
  });
  return new Table({ columnWidths: widths, width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, rows: trs });
}
const total = (vals) => ({ __total: true, vals });
const spacer = (n = 120) => new Paragraph({ spacing: { after: n }, children: [] });

// ---------- content ----------
const W = 9640; // usable width (A4 210mm - 2*25mm margins ≈ 160mm = 9070 DXA; use 9640 w/ 22mm margins)

const children = [];

// header block
children.push(p("2026年9月4日", { align: AlignmentType.RIGHT, after: 240 }));
children.push(p("〇〇銀行　〇〇支店　御中", { size: 22, after: 360 }));
children.push(p("株式会社ハートフルサンク", { align: AlignmentType.RIGHT, after: 40 }));
children.push(p("代表取締役　森山　和哉", { align: AlignmentType.RIGHT, after: 40 }));
children.push(p("株式会社EEFULホールディングス（親会社）", { align: AlignmentType.RIGHT, after: 40 }));
children.push(p("代表取締役　森山　穂貴", { align: AlignmentType.RIGHT, after: 360 }));

children.push(new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { before: 120, after: 80 },
  children: [run("業績改善の進捗および今後の見通しについて（ご報告）", { gothic: true, size: 30, bold: true, color: NAVY })],
}));
children.push(p("― 2026年11月期 上半期実績と、年内に実行する収益改善計画 ―", { align: AlignmentType.CENTER, size: 20, color: "444444", after: 360 }));

children.push(p("拝啓　時下ますますご清栄のこととお慶び申し上げます。平素は格別のご高配を賜り、厚く御礼申し上げます。"));
children.push(p("さて、当社の経営再建につきましては、日頃より貴行から多大なご理解とご支援を頂戴しております。本書は、株主向けに報告いたしました2026年11月期上半期の実績、およびその後の月次推移を踏まえ、当社が現在実行している業績改善計画と今後の見通しについてご報告申し上げるものです。ご高覧のうえ、引き続きのご支援を賜りますようお願い申し上げます。", { after: 80 }));
children.push(p("敬具", { align: AlignmentType.RIGHT, after: 240 }));

// ---------- 要旨 ----------
children.push(h1("要旨"));
children.push(bullet([run("実態収益力は改善基調にあります。", { bold: true }), run("一時費用等を除いた調整後営業利益は、グループ入り時点（2025年9月）の月▲1.2百万円から、2026年11月期上半期（2025年12月〜2026年5月）は月平均＋5.0百万円、直近7か月（〜2026年6月）でも月平均＋4.9百万円へ回復しています。")]));
children.push(bullet([run("帳簿上の損益は一時要因により赤字です。", { bold: true }), run("上半期の帳簿上の営業損益は▲26.3百万円、当期純損益は▲67.0百万円ですが、その主因は合併に伴う一時損失（約49百万円）と賞与の一括計上等であり、いずれも2026年2月までに計上を終えています。")]));
children.push(bullet([run("課題は「返済負担に対する収益力の不足」です。", { bold: true }), run("月々の元金返済（約6百万円）を税引後利益で賄うために必要な営業利益は月9.0百万円であり、現状の月5.0百万円に対し約4.0百万円の不足があります。")]));
children.push(bullet([run("年内に月額約4.5百万円のコスト構造改善を実行します。", { bold: true }), run("人員体制の適正化、不採算事業所の整理、本社間接費の削減を柱とする施策を2026年9月〜12月に集中実行し、2027年初からは月次営業利益（調整後）を必要ライン（月9.0百万円）以上に引き上げる計画です。")]));
children.push(bullet([run("あわせて資産の整理により約36百万円の一時的な資金創出を見込みます。", { bold: true }), run("不採算事業所の譲渡および所有不動産の売却（または担保活用）を進め、有利子負債の圧縮または手元流動性の確保に充当する方針です。使途については貴行とご相談のうえ決定いたします。")]));

// ---------- 1 ----------
children.push(h1("1．本資料の位置づけ"));
children.push(p("本資料は、当社が株主に報告した2026年11月期上半期報告（調整後損益計算書・貸借対照表）、2026年6月までの月次管理会計（親会社EEFULホールディングスによる連結月次PL/BS）、および2026年9月に策定した収益改善KPI管理表を基礎として作成しています。"));
children.push(p("金額の単位は百万円（小数第2位以下四捨五入）、「▲」は負の値を示します。なお「調整後」とは、賞与の一括計上を対応月へ月割りし直し、合併に伴う一時損失、処遇改善給付の一括計上など一時的・期間帰属のずれによる要因を除いた、実態収益力を示す社内管理上の数値です。調整後営業利益では、親会社への経営指導料（月2.75百万円）を営業外に区分しています。"));
children.push(p("2026年11月期の数値は月次試算表に基づく未確定値であり、決算確定により変動する可能性があります。"));

// ---------- 2 ----------
children.push(h1("2．現状認識 ― 2026年11月期の実績推移"));
children.push(h2("2-1．月次損益の推移（2025年12月〜2026年6月）"));
children.push(p("グループ入り後の再編（2025年12月の事業引継ぎ、2026年2月の吸収合併）を経て、単体には過年度からのグループ累積損失と一時費用が集約されました。一方で、これらを除いた実態収益力は改善しています。"));

const mW = [1900, 860, 860, 860, 860, 860, 860, 860, 1720];
children.push(table(mW,
  ["（百万円）", "12月", "1月", "2月", "3月", "4月", "5月", "6月", "7か月累計"],
  [
    ["売上高", "92.1", "85.4", "82.0", "87.7", "84.0", "86.7", "89.2", "607.2"],
    ["営業利益（帳簿）", "▲15.4", "2.0", "▲2.4", "3.9", "▲5.3", "3.8", "0.6", "▲12.8"],
    ["営業利益（調整後）", "14.2", "6.5", "0.2", "3.9", "0.9", "3.9", "5.1", "34.4"],
    ["EBITDA（調整後）", "14.5", "6.8", "0.7", "4.4", "1.4", "4.3", "5.8", "37.5"],
    ["当期純利益（帳簿）", "▲7.4", "▲0.2", "▲54.1", "0.9", "▲7.4", "1.3", "▲2.5", "▲69.5"],
    ["当期純利益（調整後）", "22.3", "4.3", "▲2.6", "0.9", "▲1.1", "1.3", "1.9", "26.6"],
  ],
  { aligns: ["L", "R", "R", "R", "R", "R", "R", "R", "R"] }
));
children.push(note("※ 出典：連結月次PL/BS（HT単体）。2月の帳簿上の純損失▲54.1百万円には、合併関連の雑損失28.4百万円および抱合株式消滅差損20.5百万円（いずれも一時損失）を含みます。12月の調整後営業利益が高いのは、賞与一括計上（29.2百万円）を6か月へ月割りし直した結果であり、賞与按分後の各月には▲4.9百万円が計上されています。"));
children.push(note("※ 2026年7月の売上高は日次進捗管理ベースで91.7百万円（予算比98.7％）の速報値です。7月の損益は月次締め後に改めてご報告いたします。"));
children.push(spacer(60));

children.push(h2("2-2．上半期（2025年12月〜2026年5月）実績 ― 株主報告との整合"));
children.push(p("株主向け上半期報告では、帳簿数値と調整後数値を併記し、実態収益力と一時要因を区分してご報告しました。上半期の要約は次のとおりです。"));
children.push(table([3400, 2080, 2080, 2080],
  ["（百万円）", "帳簿のまま", "調整後", "調整幅"],
  [
    ["売上高", "517.9", "517.9", "―"],
    ["営業損益", "▲26.3", "＋29.7", "＋56.0"],
    ["月平均営業損益", "▲4.4", "＋5.0", "＋9.4"],
    ["当期純損益", "▲67.0", "＋38.0", "＋105.0"],
  ]
));
children.push(note("※ 主な調整事項：賞与の月割り再配分、合併に伴う一時損失（約49百万円）の除外、処遇改善給付の月割り配分、経営指導料の営業外区分。"));
children.push(spacer(60));

children.push(h2("2-3．財政状態（2026年6月末・単体）"));
children.push(table([3400, 1600, 4640],
  ["（百万円）", "6月末", "補足"],
  [
    ["現金及び預金", "84.6", "うち定期預金10.0は借入担保として拘束"],
    ["売掛金（介護報酬等）", "165.6", "約2か月分の報酬債権"],
    ["総資産", "392.9", ""],
    ["有利子負債（短期＋長期）", "337.5", "短期63.2（当座貸越等）、長期274.3"],
    ["純資産（帳簿）", "▲20.8", "一時損失の計上により2026年4月以降債務超過"],
  ],
  { aligns: ["L", "R", "L"] }
));
children.push(note("※ 株主報告では、事業引継ぎ済み・清算予定の関係法人向け貸付金51.1百万円を全額評価減した実態純資産（2026年5月末▲59.1百万円）を併せて開示しています。なお同法人から当社が引き受けた借入（約46百万円）については、2026年11月期決算までに当社の帳簿へ計上する予定であり、計上後の実態ベースの有利子負債は約380百万円となります。"));
children.push(spacer(60));

children.push(h2("2-4．課題の所在 ― 返済負担に対する収益力の不足"));
children.push(p("当社の課題は、収益力そのものよりも、増大した有利子負債の返済負担（財務キャッシュ・フロー）にあります。月々の元金返済を税引後利益で賄うために必要な営業利益の水準は次のとおりです。"));
children.push(table([4600, 1600, 3440],
  ["計算ステップ", "月額（百万円）", "説明"],
  [
    ["毎月の元金返済", "6.0", "経費ではなく税引後利益からのみ支払可能"],
    ["税負担を上乗せ（÷(1－30％)）", "8.6", "税引前で必要な利益"],
    ["支払利息を加算", "＋0.3", "借入17口の利息"],
    total(["必要な営業利益（持続性の最低ライン）", "≒9.0", "年換算108百万円。減価償却費（月約0.7）を足し戻せば約8.0"]),
    ["現在の水準（上半期平均・調整後）", "5.0", "直近7か月平均は4.9"],
    total(["不足額", "約4.0", "本計画で解消を図る金額"]),
  ],
  { aligns: ["L", "R", "L"] }
));
children.push(p("すなわち、調整後営業利益が月9.0百万円を下回り続ける限り、返済のたびに現預金が減少する構造にあります。逆に月9.0百万円を安定的に超えれば、借入を約定どおり返済しながら自力で事業を回し続けることができます。月5.0百万円まで回復した現在、残る約4.0百万円の解消が当面の最重要課題であり、本計画はこの解消を目的としています。", { before: 120 }));

// ---------- 3 ----------
children.push(h1("3．これまでの取り組み（2025年10月〜2026年8月）"));
children.push(p("親会社であるEEFULホールディングスが経営に参画して以降、構造改革の第一段階（グループ再編と管理基盤の整備）はおおむね完了し、運営フェーズへ移行しています。"));
children.push(bullet([run("グループ再編の完了：", { bold: true }), run("関係法人の事業引継ぎ（2025年12月）、関係2社の吸収合併（2026年2月）を実施し、分散していた損益・債務を単体に統合しました。合併に伴う一時損失は2月に計上を終え、以降の発生はありません。")]));
children.push(bullet([run("会計の正常化：", { bold: true }), run("賞与の期間按分、賞与引当金の整理、法定福利費の二重計上是正、処遇改善給付の按分を実施し、月次で実態収益力を把握できる体制を整えました。")]));
children.push(bullet([run("日次の売上予実管理：", { bold: true }), run("2026年7月より事業所別・日次の売上予実管理を導入し、稼働率・単価・営業日数の要因分解を毎日行っています。7月の全社売上は予算比98.7％で着地しました。")]));
children.push(bullet([run("収益基盤の拡充：", { bold: true }), run("2026年4月に大阪市内2拠点で訪問系事業所を新設し、売上計上を開始しています。東京都内での訪問看護・訪問介護の新規指定も進めています。")]));
children.push(bullet([run("直近の実行済み施策（8月〜9月初）：", { bold: true }), run("管理職1名の処遇条件の見直し（月0.4百万円）、不要設備の撤去（月0.1百万円）を完了し、収益改善KPI管理表による週次の進捗管理を開始しました。")]));

// ---------- 4 ----------
children.push(h1("4．業績改善計画（2026年9月〜12月）"));
children.push(h2("4-1．計画の目標（KPI）"));
children.push(p([run("「2026年内に、月額4.5百万円（年換算54百万円）のコスト削減を実行可能な状態にする」", { bold: true }), run("ことをKPIとして設定しています。これは前記2-4の不足額（約4.0百万円）を上回る水準であり、達成により月次の調整後営業利益は必要ライン（9.0百万円）を超える見込みです。")]));
children.push(p("施策は事業部・本社の各責任者が担当し、KPI管理表に実施可否・優先順位・数値インパクト・開始日・完了日を明記して週次で進捗を更新しています。以下、施策の柱ごとに概要をご説明します（個別の人事事項に関わるため、対象者・事業所の詳細は記載を控えております）。"));

children.push(h2("4-2．施策の柱と月額効果（概算）"));
const sW = [2500, 4340, 1200, 1600];
children.push(table(sW,
  ["施策の柱", "主な内容", "月額効果\n（百万円）", "実行時期"],
  [
    ["① 人員体制の適正化", "通所系事業所を中心に、人時売上（1人1時間あたり売上）4,500円を基準としてシフトと配置を再設計し、事業所ごとの総労働時間目標を設定。雇用契約の更新時期・定年到達に合わせた契約時間の見直し、相談支援職等への配置転換、管理体制の世代交代を段階的に実施。", "約2.6", "9月〜11月\n（面談開始→意向確認→契約変更）"],
    ["② 事業ポートフォリオの見直し", "収支が恒常的に赤字の一部サービス（夜間帯サービス等）の縮小・廃止、収支が見合わない事業所の移転による家賃削減、小規模拠点の統合・譲渡。利用者・ご家族への丁寧な説明と受入先の確保を前提に進める。", "約1.5", "9月〜12月\n（利用者説明→移転・譲渡）"],
    ["③ 本社・間接費の削減", "コーポレート部門の組織再編（5部体制への集約）による間接人員の適正化、車両保険の見直し、研修会場の自社施設化、不要設備の撤去等。", "約0.8", "9月〜11月\n（体制発表→新体制開始）"],
    total(["合計（施策の積み上げ）", "うち年内に実行可能と見込む金額：約4.5百万円／月（KPI）", "約4.9", "―"]),
  ],
  { aligns: ["L", "L", "R", "C"] }
));
children.push(note("※ 月額効果は各施策のランニングコスト削減額の概算（千円単位のKPI管理表を集計）。①には既に完了した管理職1名の処遇見直し（月0.4百万円）を含みます。複数施策で重複する人件費は一方にのみ計上しています。"));
children.push(spacer(60));

children.push(h2("4-3．資産整理による一時的な資金創出"));
children.push(table([3800, 1400, 4440],
  ["項目", "見込額\n（百万円）", "内容・時期"],
  [
    ["不採算事業所の事業譲渡（2拠点）", "約20", "就労支援・児童系の小規模2拠点を対象に譲渡先と条件（金額・時期・体制）を協議中。2026年10月〜12月の実行を目標。"],
    ["所有不動産の売却", "約16", "堺市内の所有物件（約240㎡）。売却後に賃借として継続利用する方法、または担保として活用する方法を含めて検討。2026年9月に流通登録、11月末までの契約を目標。"],
    total(["合計", "約36", "有利子負債の圧縮または手元流動性の確保に充当。使途は貴行とご相談のうえ決定。"]),
  ],
  { aligns: ["L", "R", "L"] }
));
children.push(spacer(60));

children.push(h2("4-4．実行スケジュール"));
children.push(table([1400, 8240],
  ["時期", "主な実行事項"],
  [
    ["2026年9月", "各施策の方針決定と対象者・関係者との面談開始／通所系事業所のシフト再設計に着手／所有不動産の流通登録／コーポレート再編の面談・確定／利用者・ご家族への説明開始"],
    ["2026年10月", "雇用契約・勤務条件の変更を順次確定・適用／新任管理体制への移行開始／組織再編の体制発表／事業譲渡先との条件確定"],
    ["2026年11月", "新体制の開始（本社）／契約変更の完了／一部サービスの縮小・廃止／事業所移転の解約通知／不動産売却契約／事業譲渡の実行"],
    ["2026年12月", "全施策の実行完了と月額効果の検証／2027年以降の月次計画への反映／通期決算（2026年11月期）の確定作業"],
  ],
  { aligns: ["C", "L"] }
));
children.push(spacer(60));

// ---------- 5 ----------
children.push(h1("5．業績の見通し"));
children.push(h2("5-1．月次営業利益（調整後）の改善ブリッジ"));
children.push(table([5400, 2120, 2120],
  ["（百万円／月）", "金額", "累計"],
  [
    ["現状：上半期平均の調整後営業利益", "5.0", "5.0"],
    ["＋ ① 人員体制の適正化", "＋2.6", "7.6"],
    ["＋ ② 事業ポートフォリオの見直し", "＋1.5", "9.1"],
    ["＋ ③ 本社・間接費の削減", "＋0.8", "9.9"],
    total(["施策完了後の水準（2027年初以降）", "", "約9.5〜9.9"]),
    ["（参考）必要ライン", "", "9.0"],
  ]
));
children.push(note("※ 売上面の施策（稼働率向上、加算取得、新設事業所の立ち上がり）は上振れ要因として本ブリッジには織り込んでいません。"));
children.push(spacer(60));

children.push(h2("5-2．シナリオ別の見通し（2027年11月期）"));
children.push(table([2200, 3300, 1500, 1500, 1140],
  ["シナリオ", "前提", "月次営業利益\n（調整後）", "年間営業利益\n（調整後）", "必要ライン比"],
  [
    ["保守的", "施策の実現率7割（月＋3.2）。売上は横ばい。", "約8.2", "約98", "▲10"],
    ["基本", "KPI（月＋4.5）を年内に達成。売上は横ばい。", "約9.5", "約114", "＋6"],
    ["上振れ", "施策全体（月＋4.9）を達成し、新設事業所の寄与と稼働率改善が加わる。", "10以上", "120以上", "＋12以上"],
  ],
  { aligns: ["C", "L", "R", "R", "R"] }
));
children.push(p("基本シナリオでは、2027年11月期の調整後営業利益は年間約110百万円と、元金返済（年約72百万円）と利息を税引後利益で賄える水準に達する見込みです。保守的シナリオでも、減価償却費を足し戻したキャッシュベースの必要ライン（月約8.0百万円）はおおむね確保できる見通しです。", { before: 120 }));

children.push(h2("5-3．2026年11月期（当期）の見通し"));
children.push(p("当期は上半期に計上した合併関連の一時損失および賞与の一括計上の影響により、帳簿上は通期で当期純損失を計上する見込みです。一方、調整後の営業利益は7か月累計で＋34.4百万円を確保しており、下半期に本計画の効果が段階的に発現することから、調整後ベースでは通期黒字を維持する見通しです。決算確定後、帳簿数値と調整後数値を対比したうえで改めてご報告いたします。"));

// ---------- 6 ----------
children.push(h1("6．資金繰りの見通しと貴行へのお願い"));
children.push(p("資金繰りは、月内に社会保険料・家賃（月初）、給与（16日・26日）の支出が先行し、22〜24日頃に介護報酬が入金され、月末に約定返済を行うリズムで推移しています。現状では月内の資金の谷が最大で約13百万円生じており、短期借入（当座貸越等63.2百万円）の継続が資金繰りの前提となっています。"));
children.push(p("本計画の実行により2027年初以降は月次の営業キャッシュ・フローが返済額を上回る見通しですが、施策の効果が本格的に発現するまでの期間（2026年12月〜2027年3月頃）は、冬季賞与の支給とも重なり、手元流動性が最も薄くなる局面と認識しています。つきましては、以下の点につきましてご理解とご支援を賜りたくお願い申し上げます。"));
children.push(numbered("req", [run("短期借入（当座貸越等）の期日更新の継続", { bold: true }), run("　本計画の効果発現までの間、現行の短期借入枠を維持いただきたくお願い申し上げます。")]));
children.push(numbered("req", [run("既存長期借入の約定返済の継続と、据置明け返済開始時期の事前協議", { bold: true }), run("　当社は約定どおりの返済を継続する方針です。2027年に据置期間が明ける借入については、返済開始時の資金負担を平準化できるよう、事前にご相談させていただきたく存じます。")]));
children.push(numbered("req", [run("資産整理により創出する資金の使途に関するご相談", { bold: true }), run("　不動産売却・事業譲渡による資金（約36百万円）について、有利子負債の圧縮と手元流動性の確保のバランスを貴行とご相談のうえ決定いたします。所有不動産については、売却ではなく担保としての活用も選択肢として検討しております。")]));
children.push(numbered("req", [run("月次でのご報告の継続", { bold: true }), run("　月次試算表、調整後損益、KPI管理表の進捗を毎月ご報告し、計画との乖離が生じた場合は速やかに追加施策をご説明いたします。")]));

// ---------- 7 ----------
children.push(h1("7．留意事項"));
children.push(bullet("本資料の見通しは、2026年9月時点で入手可能な情報および当社の判断に基づく前提により作成したものであり、介護報酬・障害福祉サービス報酬の改定、人材の採用・定着状況、利用者数の変動等により実際の結果は異なる可能性があります。"));
children.push(bullet("人員体制の適正化は、対象者との合意形成と法令・就業規則の遵守を前提に段階的に進めるため、効果の発現時期が前後する可能性があります。利用者・ご家族への影響を最小化することを最優先とし、サービスの質の維持を前提に実行します。"));
children.push(bullet("2026年11月期の数値は月次試算表に基づく未確定値です。関係法人から引き受けた借入の帳簿計上、資産の評価等により、決算確定時に純資産が変動する可能性があります。"));
children.push(bullet("本資料は貴行へのご説明を目的として作成したものであり、監査法人・公認会計士による監査または検証を受けたものではありません。"));

children.push(p("以上", { align: AlignmentType.RIGHT, before: 240 }));

// ---------- 別紙 ----------
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(h1("別紙　参考資料"));
children.push(h2("別紙1．業績の推移（グループ合算・年次）"));
children.push(table([2200, 1240, 1240, 1240, 1240, 1240, 1240],
  ["（百万円）", "2022年4月期", "2023年4月期", "2024年4月期", "2025年4月期", "上半期\n（帳簿）", "上半期\n年換算（調整後）"],
  [
    ["売上高", "530.1", "705.8", "912.2", "953.5", "517.9", "1,035.8"],
    ["営業利益", "▲22.3", "▲3.1", "▲5.6", "▲19.5", "▲26.3", "＋59.4"],
    ["当期純利益", "▲12.0", "＋4.5", "▲5.2", "▲18.3", "▲67.0", "＋76.0"],
  ]
));
children.push(note("※ 2023年4月期以降は、2026年に統合した関係法人・関係会社を含む合算ベース。2022年4月期は当時ほぼ1社であった単体。上半期は2025年12月〜2026年5月。年換算は上半期調整後×2の参考値。"));
children.push(spacer(60));

children.push(h2("別紙2．実態収益力の推移（月次営業利益・調整後）"));
children.push(table([3800, 1900, 3940],
  ["時点", "月額（百万円）", "備考"],
  [
    ["2025年9月（グループ入り時点）", "▲1.2", "買収時デュー・デリジェンスに基づく水準"],
    ["2026年11月期 上半期平均（12月〜5月）", "＋5.0", "株主向け上半期報告"],
    ["直近7か月平均（12月〜6月）", "＋4.9", "連結月次PL/BS（HT単体）"],
    ["2027年初以降（本計画完了後・基本シナリオ）", "＋9.5", "必要ライン9.0を上回る水準"],
  ],
  { aligns: ["L", "R", "L"] }
));
children.push(spacer(60));

children.push(h2("別紙3．有利子負債の概要（2026年6月末）"));
children.push(table([4400, 1600, 3640],
  ["区分", "残高（百万円）", "補足"],
  [
    ["長期借入金（銀行・公的金融機関・役員借入等 17口）", "274.3", "月額元金返済 約6百万円（据置中の口を含む）"],
    ["短期借入金（当座貸越等）", "63.2", "期日更新により維持"],
    total(["有利子負債合計（帳簿）", "337.5", "現預金84.6を差し引いたネット有利子負債 252.9"]),
    ["（参考）関係法人から引受済み・帳簿計上予定の借入", "約46", "公的金融機関2口。計上後の実態合計は約380"],
  ],
  { aligns: ["L", "R", "L"] }
));
children.push(note("※ 据置期間が明ける口（2027年3月〜8月に順次償還開始予定）があるため、月額元金返済は段階的に増加する見込みです。"));

// ---------- document ----------
const doc = new Document({
  creator: "株式会社ハートフルサンク",
  title: "業績改善の進捗および今後の見通しについて（ご報告）",
  styles: {
    default: { document: { run: { font: { ascii: MINCHO, eastAsia: MINCHO, hAnsi: MINCHO }, size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: GOTHIC, size: 26, bold: true, color: NAVY } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: GOTHIC, size: 22, bold: true, color: NAVY } },
    ],
  },
  numbering: {
    config: [
      { reference: "bul", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "●", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 480, hanging: 300 } }, run: { size: 14 } } },
        { level: 1, format: LevelFormat.BULLET, text: "－", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 900, hanging: 300 } } } },
      ] },
      { reference: "req", levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: "（%1）", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 640, hanging: 640 } } } },
      ] },
    ],
  },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1300, bottom: 1200, left: 1133, right: 1133 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [run("株式会社ハートフルサンク　業績改善の進捗および今後の見通しについて　　【貴行限り】", { size: 16, color: "666666" })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT, " / ", PageNumber.TOTAL_PAGES], font: MINCHO, size: 18, color: "666666" })] })] }) },
    children,
  }],
});

const out = process.argv[2] || "out.docx";
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("written", out, buf.length); });
