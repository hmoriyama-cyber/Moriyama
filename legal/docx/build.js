// Build Word documents for the EEFUL subsidiary approval package.
// Usage: node build.js <outDir>
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  AlignmentType, BorderStyle, VerticalAlign, HeightRule, Footer, PageNumber,
} = require('docx');

const OUT = process.argv[2] || './out';
fs.mkdirSync(OUT, { recursive: true });

// ---------- style primitives ----------
const FONT = { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: 'ＭＳ 明朝', cs: 'ＭＳ 明朝' };
const SZ = 21; // 10.5pt
const W = 9070; // A4 content width with 25mm margins (DXA)
const NONE = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NO_BORDERS = { top: NONE, bottom: NONE, left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE };
const CELL_NONE = { top: NONE, bottom: NONE, left: NONE, right: NONE };
const LINE = { style: BorderStyle.SINGLE, size: 6, color: '000000' };
const BOX = { top: LINE, bottom: LINE, left: LINE, right: LINE };

function seg(s, o = {}) {
  return new TextRun({
    text: s, font: FONT, size: o.size || SZ, bold: !!o.bold,
    highlight: o.hl ? 'yellow' : undefined,
  });
}
function runs(content, o = {}) {
  if (typeof content === 'string') return [seg(content, o)];
  return content.map(c => (typeof c === 'string' ? seg(c, o) : seg(c.t, { ...o, hl: c.hl, bold: c.bold })));
}
function P(content, o = {}) {
  return new Paragraph({
    children: runs(content, o),
    alignment: o.align,
    indent: o.indent,
    spacing: { before: o.before || 0, after: o.after || 0, line: o.line || 320 },
    keepNext: o.keepNext,
  });
}
const blank = (n = 1) => Array.from({ length: n }, () => P(''));
const title = t => P(t, { align: AlignmentType.CENTER, bold: true, size: 28, before: 240, after: 360 });
const right = (t, o = {}) => P(t, { ...o, align: AlignmentType.RIGHT });
const center = (t, o = {}) => P(t, { ...o, align: AlignmentType.CENTER });
const body = (t, o = {}) => P(t, { indent: { firstLine: 210 }, ...o });
const ki = () => center('記', { before: 240, after: 240 });
const ijou = () => right('以上', { before: 240 });
const h1 = t => P(t, { bold: true, before: 240, after: 80, indent: { left: 420, hanging: 420 }, keepNext: true });
const l1 = (t, o = {}) => P(t, { indent: { left: 420, firstLine: 210 }, ...o });
const l2 = (t, o = {}) => P(t, { indent: { left: 1050, hanging: 630 }, ...o });
const l2t = (t, o = {}) => P(t, { indent: { left: 1050 }, ...o });
const l3 = (t, o = {}) => P(t, { indent: { left: 1470, hanging: 420 }, ...o });
const l3t = (t, o = {}) => P(t, { indent: { left: 1470 }, ...o });
const item = (t, o = {}) => P(t, { indent: { left: 840, hanging: 630 }, after: 60, ...o });

function cell(children, width, o = {}) {
  return new TableCell({
    children,
    width: { size: width, type: WidthType.DXA },
    borders: o.borders || CELL_NONE,
    verticalAlign: o.valign,
    margins: o.margins || { top: 20, bottom: 20, left: 60, right: 60 },
  });
}

// label / value table
function kv(rows, o = {}) {
  const indent = o.indent ?? 630;
  const total = W - indent;
  const lw = o.labelWidth || 2300;
  const vw = total - lw;
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: [lw, vw],
    indent: { size: indent, type: WidthType.DXA },
    borders: NO_BORDERS,
    rows: rows.map(([label, value]) => new TableRow({
      cantSplit: true,
      children: [
        cell([P(label)], lw),
        cell((Array.isArray(value) ? value : [value]).map(v => (v instanceof Paragraph ? v : P(v))), vw),
      ],
    })),
  });
}

function sealBox() {
  const s = 1100;
  return new Table({
    width: { size: s, type: WidthType.DXA },
    columnWidths: [s],
    borders: NO_BORDERS,
    rows: [new TableRow({
      height: { value: s, rule: HeightRule.EXACT },
      children: [new TableCell({
        children: [P('')],
        width: { size: s, type: WidthType.DXA },
        borders: BOX,
        verticalAlign: VerticalAlign.CENTER,
      })],
    })],
  });
}

// signer block: lines indented to the right half, with an optional seal box
function signer(lines, o = {}) {
  const withSeal = o.seal !== false;
  const sealW = withSeal ? 1300 : 0;
  const textW = W - sealW;
  const indent = o.indent ?? 3400;
  const paras = lines.map(l => P(l, { indent: { left: indent } }));
  const cells = [cell(paras, textW, { valign: VerticalAlign.CENTER })];
  if (withSeal) {
    cells.push(new TableCell({
      children: [sealBox()],
      width: { size: sealW, type: WidthType.DXA },
      borders: CELL_NONE,
      verticalAlign: VerticalAlign.CENTER,
      margins: { top: 0, bottom: 0, left: 120, right: 0 },
    }));
  }
  return new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: withSeal ? [textW, sealW] : [textW],
    borders: NO_BORDERS,
    rows: [new TableRow({ cantSplit: true, children: cells })],
  });
}

function footer() {
  return new Footer({
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: ['- ', PageNumber.CURRENT, ' -'], font: FONT, size: 18 })],
    })],
  });
}

function makeDoc(children, docTitle) {
  return new Document({
    creator: '株式会社EEFULホールディングス',
    title: docTitle,
    styles: { default: { document: { run: { font: FONT, size: SZ }, paragraph: { spacing: { line: 320 } } } } },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1418, bottom: 1418, left: 1418, right: 1418, header: 708, footer: 708 },
        },
      },
      footers: { default: footer() },
      children,
    }],
  });
}

// docx-js emits <w:highlightCs>, which is not in the OOXML schema; strip it so the
// file validates cleanly (highlight itself is kept).
async function stripHighlightCs(buf) {
  const JSZip = require('jszip');
  const zip = await JSZip.loadAsync(buf);
  const p = 'word/document.xml';
  const xml = await zip.file(p).async('string');
  zip.file(p, xml.replace(/<w:highlightCs[^>]*\/>/g, ''));
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

async function save(children, docTitle, name) {
  const raw = await Packer.toBuffer(makeDoc(children, docTitle));
  const buf = await stripHighlightCs(raw);
  fs.writeFileSync(path.join(OUT, name), buf);
  console.log('wrote', name);
}

// ---------- shared facts ----------
const CO = '株式会社EEFULホールディングス';
const CO_ADDR = '東京都港区芝浦4-2-8 住友不動産三田ファーストビル12階';
const SUB = '株式会社イーフルケアパートナーズ';
const DATE = '2026年9月24日';
const MORI_ADDR = '東京都中央区晴海五丁目3番1号-824号';
const SHA = '2026年3月6日付株主間契約書';
const NOTICE_TITLE = '子会社の設立および代表取締役の兼務に関する通知書兼承認依頼書';
const PROPOSAL_TITLE = '株主総会の決議事項に関する提案書';
const HL_AMOUNT = { t: '金　　　　　　円', hl: true };

const PURPOSES = [
  '介護保険法に基づく居宅サービス事業、地域密着型サービス事業、居宅介護支援事業、介護予防サービス事業および施設サービス事業',
  '障害者の日常生活及び社会生活を総合的に支援するための法律に基づく障害福祉サービス事業',
  '有料老人ホーム、サービス付き高齢者向け住宅その他高齢者向け住宅・施設の企画、運営および管理',
  '訪問看護事業',
  '介護用品および福祉用具の販売、レンタルおよび輸出入',
  '介護事業者に対する経営コンサルティング、業務受託および支援業務',
  '介護・福祉に関する人材の教育、研修および資格取得支援',
  '有料職業紹介事業および労働者派遣事業',
  '介護・福祉に関する情報システムの開発、提供および運営',
  '前各号に附帯関連する一切の事業',
];
const purposeParas = () => PURPOSES.map((t, i) => P(`(${i + 1}) ${t}`, { indent: { left: 560, hanging: 560 } }));

function subsidiaryRows(o = {}) {
  const who = o.who || '発行会社';
  const rows = [
    ['商号', SUB],
    ['本店所在地', [CO_ADDR, `（${who}と同じ）`]],
    ['事業目的', purposeParas()],
    ['資本金', '金9,900,000円'],
    ['発行株式数', `1,000株（1株当たり金9,900円。全株を${who}が引き受け、出資比率100％）`],
    ['設立時取締役', '森山　穂貴'],
    ['設立時代表取締役', '森山　穂貴'],
    ['設立予定日', '2026年10月末日'],
  ];
  if (o.purpose) rows.push(['設立の目的', o.purpose]);
  return rows;
}

const companySigner = () => signer([CO_ADDR, CO, '代表取締役　森山　穂貴']);
const keieiSigner = () => signer([MORI_ADDR, '経営株主　森山　穂貴']);

const GLOBIS_VII_LINES = [
  'PO Box 2681, Cricket Square, Hutchins Drive, George Town, Grand Cayman, KY1-1111 Cayman Islands',
  'Globis Fund VII, L.P.',
  '無限責任組合員　グロービス7号ファンド有限責任事業組合',
  '組合員　グロービス・キャピタル・パートナーズ株式会社',
  '代表取締役　堀　義人',
];
const GLOBIS_7_LINES = [
  '東京都千代田区二番町5-1',
  'グロービス7号ファンド投資事業有限責任組合',
  '無限責任組合員　グロービス7号ファンド有限責任事業組合',
  '組合員　グロービス・キャピタル・パートナーズ株式会社',
  '代表取締役　堀　義人',
];
const DBC_LINES = [
  '東京都港区虎ノ門2-2-1 住友不動産虎ノ門タワー13階Room8',
  'DBC1号投資事業有限責任組合',
  '無限責任組合員　MV有限責任事業組合',
  '代表組合員　株式会社Dual Bridge Capital',
  '代表取締役　寺田　修輔',
];
const FIC_LINES = [
  '東京都渋谷区恵比寿1-25-7',
  'Future Identity Capital投資事業有限責任組合',
  '無限責任組合員　Future Identity Capital有限責任事業組合',
  '組合員　森山　博暢',
];

// ---------- 01 通知書兼承認依頼書 ----------
function docNotice() {
  return [
    P('グロービス・キャピタル・パートナーズ株式会社'),
    P('（Globis Fund VII, L.P. および グロービス7号ファンド投資事業有限責任組合）'),
    P('南　良平　様'),
    ...blank(),
    P('DBC1号投資事業有限責任組合'),
    P('無限責任組合員　MV有限責任事業組合'),
    P('代表組合員　株式会社Dual Bridge Capital'),
    P('代表取締役　寺田　修輔　様'),
    ...blank(),
    right(DATE),
    ...blank(),
    companySigner(),
    ...blank(),
    keieiSigner(),
    title(NOTICE_TITLE),
    body(`${SHA}（以下「本契約」といいます。）第2.3条第1項および第2.4条第2項に基づき、下記のとおり通知するとともに、多数投資者の書面による承認をお願いいたします。`),
    ki(),

    h1('1　株式会社イーフルケアパートナーズの設立（事前同意のお願い）'),
    l1('発行会社を発起人として、下記のとおり完全子会社を設立いたします。'),
    kv(subsidiaryRows({ purpose: '発行会社グループにおける介護事業全般の受け皿とするため' })),

    h1('2　経営株主 森山穂貴 の子会社代表取締役の兼務（本契約第2.4条第2項）'),
    l1('経営株主 森山穂貴 は、上記1により設立する株式会社イーフルケアパートナーズの代表取締役に就任し、発行会社の代表取締役と兼務いたします。本契約第2.4条第2項に基づき、多数投資者の事前の書面による承認をお願いいたします。'),
    kv([
      ['兼務先での地位', '代表取締役'],
      ['発行会社の業務への影響', '発行会社の業務執行体制に変更はなく、経営株主は引き続き発行会社の業務に専念いたします。子会社は発行会社グループの一部として運営します。'],
    ], { labelWidth: 2800 }),
    l1('あわせて、上記の就任は本契約別紙2.3-1(11)（発行会社の子会社における代表取締役の選任）に該当するため、第2.3条第1項に基づく承認もお願いいたします。', { before: 120 }),

    h1('3　発行会社と子会社との間の取引（本契約別紙2.3-1(17)）'),
    l1('発行会社と株式会社イーフルケアパートナーズとの間で、下記の取引を予定しております。関連当事者との取引に該当するため、第2.3条第1項に基づき、下記の範囲で包括的に承認をお願いいたします。取引条件は、独立した第三者間の取引において合理的に設定される条件より発行会社に不利なものとしません（第2.2条第1項）。'),
    l2('(1) 業務委託契約、経営指導契約およびこれらに付随する取引'),
    l2(['(2) 金銭の貸付、債務保証その他の資金支援（上限額 ', HL_AMOUNT, '）']),
    l2('(3) 事務所・設備の賃貸借および使用貸借'),
    l2('(4) 従業員の出向・兼務に関する取引'),
    l2('(5) その他上記に準ずる取引'),
    l2t('期間　森山穂貴 が同社の代表取締役に在任する期間', { before: 60 }),

    h1('4　今後のグループ会社役員等の兼務に関する包括承認のお願い（本契約第2.4条第2項および第2.3条第1項）'),
    l1('経営株主が今後、発行会社のグループ会社の役員等を兼務することについて、下記の範囲で、その都度の承認を要しないものとして包括的にご承認いただきたくお願いいたします。'),
    l2('(1) グループ会社の範囲'),
    l3('(ア) 発行会社が直接または間接に議決権の過半数を保有する会社'),
    l3('(イ) 発行会社または(ア)の会社が出資し、発行会社の事業と一体として運営される会社として発行会社が指定した会社'),
    l2('(2) 対象となる兼務'),
    l3t('グループ会社の取締役、代表取締役、監査役、執行役員、業務執行社員、理事その他これらに準ずる役員等への就任'),
    l2('(3) 条件'),
    l3('① 経営株主は引き続き発行会社の業務に専念し、兼務が発行会社の業務に支障を及ぼさないこと'),
    l3('② 兼務先の事業が発行会社グループの事業の範囲内であること'),
    l2('(4) 有効期間'),
    l3t('本承認の日から、本契約の終了または多数投資者による書面での撤回まで'),
    l2('(5) 本契約との関係'),
    l3t('本承認は、本契約第2.4条第2項および第2.3条第1項（別紙2.3-1(11)）に基づく承認を上記の範囲に限って包括的に与えるものであり、本契約のその他の条項に影響を及ぼしません。子会社の設立・取得そのもの、および別紙2.3-1・2.3-2に記載のその他の事項については、従来どおり個別に通知・承認をお願いいたします。'),

    h1('5　承認方法およびご回答期限'),
    l1('ご承認いただける場合は、別紙承認書への電子署名（クラウドサインによる合意締結を含みます。）または記名押印、もしくは本通知への電子メールでのご回答をお願いいたします。第2.3条第1項に基づく事項については、本通知の発信日から7日を経過する日までに承認しない旨のご通知がない場合、同項に基づき承認いただいたものとして取り扱わせていただきます。第2.4条第2項に基づく兼務の承認については、明示のご回答をお願いいたします。'),
    ijou(),
  ];
}

// ---------- 02 / 03 承認書 ----------
function docApproval(o) {
  const subj = o.plural ? '当組合ら' : '当組合';
  const role = o.majority ? '多数投資者として承認します。' : '承認します。';
  const signers = [];
  o.signerBlocks.forEach((lines, i) => {
    if (i > 0) signers.push(...blank());
    signers.push(signer(lines));
  });
  return [
    title('承　認　書'),
    right(DATE),
    ...blank(),
    P(`${CO}　御中`),
    P('森山　穂貴　殿'),
    ...blank(),
    ...signers,
    ...blank(),
    body(`${subj}は、${SHA}（以下「本契約」といいます。）第2.3条第1項および第2.4条第2項に基づき、貴社および経営株主 森山穂貴 から${DATE}付「${NOTICE_TITLE}」（以下「本通知書」といいます。）により通知を受けた下記の事項について、その内容を確認したうえで、${role}`),
    ki(),
    item('1　株式会社イーフルケアパートナーズの設立'),
    item('2　経営株主 森山穂貴 の株式会社イーフルケアパートナーズ代表取締役への就任および貴社代表取締役との兼務（本契約第2.4条第2項および別紙2.3-1(11)）'),
    item('3　貴社と株式会社イーフルケアパートナーズとの間の取引（本通知書3記載の範囲における包括承認。本契約別紙2.3-1(17)）'),
    item('4　経営株主による貴社グループ会社の役員等の兼務（本通知書4記載の範囲および条件における包括承認。本契約第2.4条第2項および別紙2.3-1(11)）'),
    body(`本承認は上記の事項に限るものであり、本契約のその他の条項に基づく${subj}の権利に影響を及ぼすものではありません。`, { before: 240 }),
    ijou(),
  ];
}

// ---------- 04 提案書 ----------
function docProposal() {
  return [
    P('株主各位'),
    ...blank(),
    right(DATE),
    ...blank(),
    companySigner(),
    title(PROPOSAL_TITLE),
    body('会社法第319条第1項の規定に基づき、下記の事項を株主総会の決議事項として提案いたします。ご同意いただける場合は、別紙同意書に電子署名（クラウドサインによる合意締結を含みます。）または記名押印のうえ、当社までご返送ください。'),
    ki(),

    h1('第1号議案　株式会社イーフルケアパートナーズ設立の件'),
    l1('当社を発起人として、下記のとおり完全子会社を設立する。設立に必要な定款の作成および認証、出資の履行、設立登記の申請その他一切の手続きは、代表取締役 森山穂貴 に一任する。'),
    kv(subsidiaryRows({ who: '当社' })),

    h1('第2号議案　代表取締役 森山穂貴 の競業取引および利益相反取引承認の件'),
    l2('(1) 競業取引の承認（会社法第356条第1項第1号）', { before: 60 }),
    l3t('代表取締役 森山穂貴 が、第1号議案により設立する株式会社イーフルケアパートナーズの代表取締役に就任し、同社のために当社の事業の部類に属する取引を行うことを、同社の代表取締役に在任する期間中の取引の全部について包括的に承認する。'),
    l3t('重要な事実の開示', { before: 60 }),
    kv([
      ['兼務先', [SUB, '（当社が議決権の100％を保有する子会社）']],
      ['兼務先での地位', '代表取締役'],
      ['兼務先の事業内容', '介護事業全般（第1号議案記載の事業目的のとおり）'],
      ['当社事業との関係', '当社グループの介護事業を同社が担う。同社は当社の完全子会社として当社グループの一部を構成するため、当社との利害の対立は実質的に生じない。'],
    ], { indent: 1470, labelWidth: 2300 }),
    l2('(2) 利益相反取引の包括承認（会社法第356条第1項第2号および第3号）', { before: 120 }),
    l3t('当社と株式会社イーフルケアパートナーズとの間で、代表取締役 森山穂貴 が双方を代表して行う下記の取引を、同人が同社の代表取締役に在任する期間について包括的に承認する。取引条件は、独立した第三者間の取引において合理的に設定される条件より当社に不利なものとしない。'),
    l3('① 業務委託契約、経営指導契約およびこれらに付随する取引'),
    l3(['② 金銭の貸付、債務保証その他の資金支援に関する取引（上限額 ', HL_AMOUNT, '）']),
    l3('③ 事務所・設備の賃貸借および使用貸借'),
    l3('④ 従業員の出向・兼務に関する取引'),
    l3('⑤ その他上記に準ずる取引'),

    h1('第3号議案　グループ会社の役員等の兼務に関する包括承認の件'),
    l1('当社の取締役が、今後、当社のグループ会社の役員等を兼務することについて、その都度の株主総会の承認を要しないものとし、下記のとおり包括的に承認する。'),
    l2('1　対象となる取締役', { before: 60 }),
    l3t('当社の取締役全員（現に在任する者および今後就任する者を含む）'),
    l2('2　グループ会社の範囲', { before: 60 }),
    l3('(ア) 当社が直接または間接に議決権の過半数を保有する会社'),
    l3('(イ) 当社または(ア)の会社が出資し、当社の事業と一体として運営される会社として当社が指定した会社'),
    l2('3　承認する兼務の内容', { before: 60 }),
    l3t('グループ会社の取締役、代表取締役、監査役、執行役員、業務執行社員、理事その他これらに準ずる役員等への就任'),
    l2('4　あわせて承認する取引', { before: 60 }),
    l3('(1) 兼務先のために当社の事業の部類に属する取引を行うこと（会社法第356条第1項第1号）'),
    l3('(2) 当社と兼務先との間で、第2号議案(2)①から⑤に掲げる種類の取引を行うこと（会社法第356条第1項第2号および第3号）'),
    l2('5　承認の期間', { before: 60 }),
    l3t('本決議の日から、株主総会が本決議を変更または撤回するまで'),
    ijou(),
  ];
}

// ---------- 05-09 同意書 ----------
function docConsent(o) {
  return [
    title('同　意　書'),
    right(DATE),
    ...blank(),
    P(`${CO}　御中`),
    ...blank(),
    signer([...o.lines, o.holding]),
    ...blank(),
    body(`${o.subj}は、${DATE}付「${PROPOSAL_TITLE}」により貴社代表取締役 森山穂貴 から提案のあった下記の事項について、会社法第319条第1項の規定に基づき、その内容を確認したうえで同意します。`),
    ki(),
    item('第1号議案　株式会社イーフルケアパートナーズ設立の件'),
    item('第2号議案　代表取締役 森山穂貴 の競業取引および利益相反取引承認の件'),
    item('第3号議案　グループ会社の役員等の兼務に関する包括承認の件'),
    ijou(),
  ];
}

// ---------- 10 議事録 ----------
function docMinutes() {
  const agenda = (t) => P(t, { indent: { left: 420 }, before: 120, keepNext: true });
  const agendaBody = (t) => P(t, { indent: { left: 840, firstLine: 210 } });
  return [
    title('臨時株主総会議事録'),
    body('会社法第319条第1項の規定に基づき、株主総会の決議があったものとみなされた事項は、次のとおりである。'),

    h1('1　株主総会の決議があったものとみなされた事項の内容'),
    agenda('第1号議案　株式会社イーフルケアパートナーズ設立の件'),
    agendaBody('当社を発起人として、下記のとおり完全子会社を設立し、設立に必要な一切の手続きを代表取締役 森山穂貴 に一任する。'),
    kv([
      ['商号', SUB],
      ['本店所在地', CO_ADDR],
      ['事業目的', `介護事業全般（${DATE}付提案書記載のとおり）`],
      ['資本金', '金9,900,000円'],
      ['発行株式数', '1,000株（全株を当社が引き受け、出資比率100％）'],
      ['設立時取締役', '森山　穂貴'],
      ['設立時代表取締役', '森山　穂貴'],
      ['設立予定日', '2026年10月末日'],
    ], { indent: 1050 }),
    agenda('第2号議案　代表取締役 森山穂貴 の競業取引および利益相反取引承認の件'),
    agendaBody(`代表取締役 森山穂貴 が株式会社イーフルケアパートナーズの代表取締役に就任し、同社のために当社の事業の部類に属する取引を行うこと（会社法第356条第1項第1号）、および当社と同社との間の取引（同項第2号・第3号）を、${DATE}付提案書に記載の範囲および期間において包括的に承認する。`),
    agenda('第3号議案　グループ会社の役員等の兼務に関する包括承認の件'),
    agendaBody(`当社の取締役が当社のグループ会社の役員等を兼務すること、およびこれに伴う競業取引・利益相反取引を、${DATE}付提案書に記載の範囲において包括的に承認し、その都度の株主総会の承認を要しないものとする。`),

    h1('2　前項の事項を提案した者の氏名'),
    P('代表取締役　森山　穂貴', { indent: { left: 630 } }),
    h1('3　株主総会の決議があったものとみなされた日'),
    P(`${DATE}（議決権を有する株主全員の同意書が当社に到達した日）`, { indent: { left: 630 } }),
    h1('4　議事録の作成に係る職務を行った取締役の氏名'),
    P('代表取締役　森山　穂貴', { indent: { left: 630 } }),

    body('上記のとおり、会社法第319条第1項に基づき株主総会の決議があったものとみなされたので、同法第318条第1項および会社法施行規則第72条第4項第1号に基づき本議事録を作成する。', { before: 360 }),
    ...blank(),
    right(DATE),
    ...blank(),
    signer([`${CO}　臨時株主総会`, '議事録作成者　代表取締役　森山　穂貴']),
  ];
}

// ---------- build all ----------
(async () => {
  await save(docNotice(), NOTICE_TITLE, '01_通知書兼承認依頼書_投資者宛.docx');
  await save(docApproval({ plural: true, majority: true, signerBlocks: [GLOBIS_VII_LINES, GLOBIS_7_LINES] }), '承認書（グロービス）', '02_承認書_グロービス.docx');
  await save(docApproval({ plural: false, majority: false, signerBlocks: [DBC_LINES] }), '承認書（DBC）', '03_承認書_DBC.docx');
  await save(docProposal(), PROPOSAL_TITLE, '04_株主総会提案書.docx');
  await save(docConsent({ subj: '私', lines: [MORI_ADDR, '森山　穂貴'], holding: ['（保有株式：普通株式 ', { t: '11,012株', hl: true }, '）'] }), '同意書（森山穂貴）', '05_同意書_森山穂貴.docx');
  await save(docConsent({ subj: '当組合', lines: FIC_LINES, holding: '（保有株式：普通株式 188株）' }), '同意書（Future Identity Capital投資事業有限責任組合）', '06_同意書_FutureIdentityCapital.docx');
  await save(docConsent({ subj: '当組合', lines: GLOBIS_VII_LINES, holding: '（保有株式：PA種優先株式 968株）' }), '同意書（Globis Fund VII, L.P.）', '07_同意書_GlobisFundVII.docx');
  await save(docConsent({ subj: '当組合', lines: GLOBIS_7_LINES, holding: '（保有株式：PA種優先株式 2,426株）' }), '同意書（グロービス7号ファンド投資事業有限責任組合）', '08_同意書_グロービス7号ファンド.docx');
  await save(docConsent({ subj: '当組合', lines: DBC_LINES, holding: ['（保有株式：PA種優先株式 1,131株', { t: '、S種優先株式 1,245株', hl: true }, '）'] }), '同意書（DBC1号投資事業有限責任組合）', '09_同意書_DBC1号.docx');
  await save(docMinutes(), '臨時株主総会議事録', '10_臨時株主総会議事録.docx');
})().catch(e => { console.error(e); process.exit(1); });
