// Nihai istatistik raporu (n = 15). Kullanım: node rapor_nihai.js cikti.docx
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
  AlignmentType, HeadingLevel, ImageRun, BorderStyle, LevelFormat, Footer, PageNumber, PageBreak,
} = require('docx');

const R = JSON.parse(fs.readFileSync(path.join(__dirname, 'sonuclar.json'), 'utf8'));
const FIG = (f) => path.join(__dirname, 'sekiller', f);
const FONT = 'Times New Roman';
const W = 9026;
const ACCENT = '1F3864';

// ---------- sayı biçimleri ----------
const f2 = (x) => Number(x).toFixed(2);
const f3 = (x) => Number(x).toFixed(3);
const pf = (p) => (p < 0.001 ? '< 0.001' : '= ' + Number(p).toFixed(3));
const pct = (x) => '%' + (100 * x).toFixed(1);
const D = R.desc, FR = R.friedman, IK = R.ikili, U = R.uyum, G = R.guc;

// ---------- yardımcılar ----------
function runs(text, base = {}) {
  const out = []; const re = /(\*\*[^*]+\*\*|~[^~]+~)/g; let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(new TextRun({ ...base, text: text.slice(last, m.index) }));
    const t = m[0];
    out.push(t.startsWith('**') ? new TextRun({ ...base, text: t.slice(2, -2), bold: true }) : new TextRun({ ...base, text: t.slice(1, -1), italics: true }));
    last = m.index + t.length;
  }
  if (last < text.length) out.push(new TextRun({ ...base, text: text.slice(last) }));
  return out;
}
const P = (t, o = {}) => new Paragraph({ children: runs(t, o.run || {}), alignment: o.align || AlignmentType.JUSTIFIED, spacing: { before: 60, after: 120, line: 300 }, keepNext: o.keepNext });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const BL = (t) => new Paragraph({ children: runs(t), numbering: { reference: 'bul', level: 0 }, alignment: AlignmentType.JUSTIFIED, spacing: { after: 70, line: 288 } });
const NL = (t, ref) => new Paragraph({ children: runs(t), numbering: { reference: ref, level: 0 }, alignment: AlignmentType.JUSTIFIED, spacing: { after: 70, line: 288 } });
const CAP = (t) => new Paragraph({ children: runs(t, { size: 21 }), spacing: { before: 200, after: 80 }, keepNext: true });
const NOTE = (t) => new Paragraph({ children: runs(t, { size: 18, color: '404040' }), spacing: { before: 40, after: 160 }, alignment: AlignmentType.JUSTIFIED });
const EQ = (t) => new Paragraph({ children: runs(t, { font: 'Cambria Math' }), alignment: AlignmentType.CENTER, spacing: { before: 60, after: 100 } });

let figNo = 0;
function figure(file, wIn, caption, comment) {
  const buf = fs.readFileSync(FIG(file));
  const w = buf.readUInt32BE(16), h = buf.readUInt32BE(20);
  const px = Math.round(wIn * 96);
  figNo += 1;
  const out = [
    new Paragraph({ children: [new ImageRun({ type: 'png', data: buf, transformation: { width: px, height: Math.round(px * h / w) } })], alignment: AlignmentType.CENTER, spacing: { before: 160, after: 40 }, keepNext: true }),
    new Paragraph({ children: runs(`**Şekil ${figNo}.** ${caption}`, { size: 19 }), alignment: AlignmentType.CENTER, spacing: { after: comment ? 60 : 200 }, keepNext: !!comment }),
  ];
  if (comment) out.push(commentBox(comment));
  return out;
}
function commentBox(text) {
  const b = { style: BorderStyle.SINGLE, size: 12, color: '2F6DB5' }, n = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    rows: [new TableRow({ children: [new TableCell({
      width: { size: W, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, fill: 'F2F6FB', color: 'auto' },
      borders: { left: b, top: n, bottom: n, right: n }, margins: { top: 80, bottom: 80, left: 160, right: 120 },
      children: [new Paragraph({ children: runs('**Yorum: **' + text, { size: 20 }), alignment: AlignmentType.JUSTIFIED, spacing: { after: 0, line: 276 } })],
    })] })],
  });
}
const SP = () => new Paragraph({ children: [], spacing: { after: 100 } });

const thin = { style: BorderStyle.SINGLE, size: 4, color: '9AA5B1' };
function table(headers, rows, widths, opt = {}) {
  const s = widths.reduce((a, b) => a + b, 0); const ws = widths.map(x => Math.round(x * W / s)); ws[ws.length - 1] += W - ws.reduce((a, b) => a + b, 0);
  const cell = (t, i, hdr, last) => new TableCell({
    width: { size: ws[i], type: WidthType.DXA },
    shading: hdr ? { type: ShadingType.CLEAR, fill: 'DCE6F0', color: 'auto' } : (last && opt.boldLast ? { type: ShadingType.CLEAR, fill: 'F3F3F3', color: 'auto' } : undefined),
    margins: { top: 50, bottom: 50, left: 90, right: 90 }, borders: { top: thin, bottom: thin, left: thin, right: thin },
    children: String(t).split('\n').map(l => new Paragraph({ children: runs(l, { size: 19, bold: (hdr || (last && opt.boldLast)) || undefined }), alignment: i === 0 || opt.left ? AlignmentType.LEFT : AlignmentType.CENTER, spacing: { after: 0, line: 252 }, keepNext: true })),
  });
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: ws, rows: [
    new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, i, true)) }),
    ...rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, i, false, ri === rows.length - 1)) })),
  ] });
}
function keyBox(title, items) {
  const b = { style: BorderStyle.SINGLE, size: 8, color: '5B8C3A' };
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ children: [new TableCell({
    width: { size: W, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, fill: 'EEF5EA', color: 'auto' },
    borders: { top: b, bottom: b, left: b, right: b }, margins: { top: 120, bottom: 120, left: 180, right: 180 },
    children: [new Paragraph({ children: [new TextRun({ text: title, bold: true, size: 24, color: '3A6B22' })], spacing: { after: 100 } }),
      ...items.map(t => new Paragraph({ children: runs(t, { size: 21 }), numbering: { reference: 'bul', level: 0 }, spacing: { after: 70, line: 276 }, alignment: AlignmentType.JUSTIFIED }))],
  })] })] });
}

// ---------- içerik ----------
const TITLE = 'Üst Çenesi Tam Dişsiz Hastalarda Dijital İş Akışı Kapsamında Farklı Tekniklerle Elde Edilen Üç Boyutlu Sanal Modellerin Doğruluklarının İn Vivo Olarak Değerlendirilmesi';
const [K, PS, FU] = D; const [KP, KF, PF] = IK;
const fs15 = G.fried_sim['15'], fs3 = G.fried_sim['3'];

const C = [];
C.push(
  new Paragraph({ children: [new TextRun({ text: 'NİHAİ İSTATİSTİKSEL ANALİZ RAPORU', bold: true, size: 34, color: ACCENT })], alignment: AlignmentType.CENTER, spacing: { before: 200, after: 160 } }),
  new Paragraph({ children: [new TextRun({ text: TITLE, italics: true, size: 23 })], alignment: AlignmentType.CENTER, spacing: { after: 300 } }),
  table(['Başlık', 'Bilgi'], [
    ['Araştırma tasarımı', 'Tekrarlı ölçüm (bağımlı örneklem) tasarımı: her katılımcıda üç yöntemle elde edilen sanal model aynı referansa göre değerlendirilmiştir.'],
    ['Katılımcı sayısı', 'n = 15 üst çenesi tam dişsiz birey (15 × 3 = 45 ölçüm; eksik veri yok)'],
    ['Karşılaştırılan yöntemler', 'KIBT – Referans tarama  |  Primescan (ağız içi tarayıcı) – Referans tarama  |  Füzyon model – Referans veri'],
    ['Sonuç değişkeni', 'Referans modele göre yüzey sapması (mm); düşük değer = yüksek doğruluk (trueness)'],
    ['Anlamlılık düzeyi', 'α = 0.05 (iki yönlü); ikili karşılaştırmalarda Bonferroni düzeltmesi (α′ = 0.0167)'],
  ], [26, 74], { left: true }),
  SP(),
  keyBox('Ana bulgular', [
    `Üç yöntemin doğruluğu arasında anlamlı fark vardır: **Friedman χ²(2) = ${f2(FR.chi2)}; p < 0.001; Kendall’s W = ${f3(FR.W)} (büyük etki).**`,
    `**KIBT** modelleri en yüksek sapmayı göstermiştir (${f3(K.ort)} ± ${f3(K.ss)} mm). KIBT; hem Primescan’den (${f3(KP.ort)} mm) hem Füzyon modelden (${f3(KF.ort)} mm) daha fazla sapmaktadır (her ikisi p < 0.001). 15 katılımcının **tamamında** KIBT en kötü sonucu vermiştir.`,
    `**Primescan** (${f3(PS.ort)} ± ${f3(PS.ss)} mm) ve **Füzyon model** (${f3(FU.ort)} ± ${f3(FU.ss)} mm) arasında fark yoktur (fark ${f3(PF.ort)} mm; p = ${f3(PF.p_bonf)}). İki yöntem ±0.10 mm sınırında **istatistiksel olarak eşdeğerdir** (TOST p < 0.001).`,
    `Füzyon işlemi, ağız içi taramanın doğruluğunu korurken KIBT’nin tek başına sağladığı doğruluğu yaklaşık **5 kat** iyileştirmiştir.`,
    `**n = 15 yeterlidir:** tasarıma uygun a priori hesaplama 15 katılımcı vermektedir; gözlenen verilerde genel karşılaştırmanın gücü **≥ %99**, eşdeğerlik testinin gücü **${pct(G.tost15['0.1'])}**’dir.`,
  ]),

  // 1
  new Paragraph({ children: [new PageBreak()] }),
  H1('1. İstatistiksel Yöntem'),
  P('Analiz, diş hekimliğinde dijital model doğruluğu (trueness) çalışmalarında yaygın olarak kullanılan ve küçük örneklemli tekrarlı ölçüm tasarımlarına uygun yaklaşımlarla yapılmıştır. Uygulanan adımlar ve seçim gerekçeleri aşağıda özetlenmiştir:'),
  table(['Amaç', 'Yöntem', 'Gerekçe'], [
    ['Verileri özetlemek', 'Ortalama ± SS, %95 GA, ortanca (ÇAA), min–maks', 'Hem parametrik hem parametrik olmayan özetler birlikte sunulmuştur.'],
    ['Dağılımı incelemek', 'Shapiro–Wilk testi, Q–Q grafikleri', 'n < 50 için en güçlü normallik testidir.'],
    ['Üç yöntemi birlikte karşılaştırmak', 'Friedman testi; etki büyüklüğü Kendall’s W', 'Aynı katılımcıda üç bağımlı ölçüm vardır ve bir değişken normal dağılmamaktadır.'],
    ['Hangi yöntemlerin farklı olduğunu bulmak', 'Wilcoxon işaretli sıralar testi + Bonferroni düzeltmesi; etki büyüklüğü r = Z/√n', 'Üç ikili karşılaştırmada I. tip hata kontrol altına alınmıştır.'],
    ['Primescan ve Füzyon’un benzerliğini göstermek', 'Eşdeğerlik testi (TOST, ±0.10 mm); Bland–Altman analizi; ICC(A,1)', '“Fark bulunamadı” sonucu benzerliği kanıtlamaz; eşdeğerlik ayrıca test edilmelidir.'],
    ['Örneklem yeterliliğini doğrulamak', 'A priori güç (G*Power, eşleştirilmiş test); post-hoc güç (merkezi olmayan χ² ve simülasyon)', 'Tasarıma uygun hesaplama ve gözlenen verilerle doğrulama birlikte yapılmıştır.'],
  ], [24, 38, 38], { left: true }),
  NOTE('SS: standart sapma; GA: güven aralığı; ÇAA: çeyrekler arası aralık; TOST: iki tek yönlü test; ICC(A,1): iki yönlü rastgele etkiler modeli, mutlak uyum, tek ölçüm sınıf içi korelasyon katsayısı. Analizler Python 3 (SciPy, statsmodels, pingouin) ile yapılmış olup SPSS 25 ile birebir yeniden üretilebilir.'),

  // 2
  H1('2. Tanımlayıcı Bulgular'),
  CAP('**Tablo 1.** Yöntemlere göre yüzey sapması değerleri (n = 15, mm)'),
  table(['Yöntem', 'Ortalama ± SS', '%95 GA', 'Ortanca (ÇAA)', 'Min – Maks'], D.map((d, i) => [
    ['KIBT – Referans tarama', 'Primescan – Referans tarama', 'Füzyon model – Referans veri'][i],
    `${f3(d.ort)} ± ${f3(d.ss)}`, `${f3(d.alt)} – ${f3(d.ust)}`, `${f3(d.med)} (${f3(d.q1)} – ${f3(d.q3)})`, `${f3(d.min)} – ${f3(d.max)}`]),
    [30, 17, 17, 21, 15]),
  SP(),
  ...figure('s01_ortalama.png', 5.6, 'Yöntemlere göre ortalama sapma ve %95 güven aralıkları.',
    `KIBT modellerinin ortalama sapması (${f3(K.ort)} mm), Primescan (${f3(PS.ort)} mm) ve Füzyon (${f3(FU.ort)} mm) modellerinin yaklaşık 5.3 katıdır. Primescan ve Füzyon’un ortalamaları ve güven aralıkları neredeyse tamamen örtüşmektedir.`),
  ...figure('s02_kutu.png', 5.6, 'Sapma değerlerinin dağılımı: kutu grafiği (ortanca ve çeyrekler) ve her katılımcının değeri.',
    'KIBT değerleri diğer iki yöntemle hiç çakışmamaktadır: en düşük KIBT değeri (1.187 mm) bile en yüksek Primescan (0.628 mm) ve Füzyon (0.614 mm) değerlerinin yaklaşık iki katıdır. KIBT’de iki katılımcının (K8, K14) değerleri grubun geri kalanından düşüktür.'),
  ...figure('s03_katilimci.png', 6.0, 'Her katılımcı için üç yöntemin sapma değerleri.',
    'Bulgu her katılımcıda tekrarlanmaktadır: 15 katılımcının tamamında KIBT çubuğu belirgin biçimde en yüksektir. Primescan ve Füzyon çubukları ise her katılımcıda birbirine yakındır.'),
  ...figure('s04_profil.png', 6.0, 'Katılımcı profilleri: (a) üç yöntem, (b) Primescan ve Füzyon’un yakından görünümü. Siyah çizgi ortalamayı gösterir.',
    `(a) KIBT’den diğer yöntemlere tüm çizgiler aşağı inmektedir. (b) Primescan’den Füzyon’a geçişte çizgilerin yaklaşık yarısı yukarı (${15 - U.ust_say}), yarısı aşağı (${U.ust_say}) yönelmektedir; ortalama çizgisi yataydır. Bu, iki yöntem arasında sistematik bir fark olmadığını gösterir.`),

  // 3
  H1('3. Normal Dağılım Değerlendirmesi'),
  CAP('**Tablo 2.** Shapiro–Wilk normallik testi'),
  table(['Değişken', 'W', 'p', 'Sonuç'], D.map((d, i) => [['KIBT – Referans', 'Primescan – Referans', 'Füzyon – Referans'][i], f3(d.sw_w), f3(d.sw_p), d.sw_p < 0.05 ? 'Normal dağılmıyor' : 'Normal dağılıma uygun']), [34, 16, 16, 34]),
  ...figure('s05_qq.png', 6.2, 'Normal Q–Q grafikleri. Noktaların doğru üzerinde olması normal dağılıma uygunluğu gösterir.',
    'Primescan verileri normal dağılmamaktadır (p = 0.021); noktalar doğrudan belirgin şekilde sapmaktadır. Bu nedenle ve örneklemin küçük olması nedeniyle tüm karşılaştırmalarda parametrik olmayan testler (Friedman ve Wilcoxon) kullanılmıştır.'),

  // 4
  H1('4. Üç Yöntemin Genel Karşılaştırılması (Friedman Testi)'),
  P('Friedman testinde her katılımcının üç değeri kendi içinde sıralanır (1 = en düşük sapma, 3 = en yüksek sapma) ve yöntemlerin sıra toplamları karşılaştırılır.'),
  EQ(`χ² = [12 / (n·k·(k+1))] · ΣR² − 3·n·(k+1) = [12 / 180] · (45² + 23² + 22²) − 180 = ${f2(FR.chi2)}`),
  EQ(`Kendall’s W = χ² / [n·(k − 1)] = ${f2(FR.chi2)} / 30 = ${f3(FR.W)}`),
  ...figure('s06_friedman.png', 5.6, 'Friedman testinde yöntemlerin ortalama sıraları.',
    `KIBT’nin ortalama sırası 3.00’tür; yani her katılımcıda en yüksek sapma KIBT’dedir. Primescan (${f2(FR.ort_sira[1])}) ve Füzyon (${f2(FR.ort_sira[2])}) neredeyse eşit sıralara sahiptir. Üç yöntem arasında anlamlı fark vardır (**χ²(2) = ${f2(FR.chi2)}; p < 0.001**). Kendall’s W = ${f3(FR.W)} büyük bir etkiyi gösterir (0.1 küçük, 0.3 orta, ≥ 0.5 büyük).`),

  // 5
  H1('5. İkili Karşılaştırmalar'),
  CAP('**Tablo 3.** Yöntemlerin ikili karşılaştırılması (Wilcoxon işaretli sıralar testi, Bonferroni düzeltmeli)'),
  table(['Karşılaştırma', 'Ortalama fark\n(%95 GA), mm', 'Ortanca fark,\nmm', 'Wilcoxon T', 'Düzeltilmiş p', 'Etki (r)'], IK.map(p => [
    p.ad, `${f3(p.ort)}\n(${f3(p.alt)} – ${f3(p.ust)})`, f3(p.med), String(p.T), pf(p.p_bonf).replace('= ', ''), f2(p.r)]),
    [24, 20, 14, 13, 15, 14]),
  NOTE('Düzeltilmiş p = ham p × 3. Etki büyüklüğü r: 0.1 küçük, 0.3 orta, ≥ 0.5 büyük. T = 0, tüm katılımcılarda farkın aynı yönde olduğunu gösterir.'),
  ...figure('s07_ikili.png', 5.8, 'İkili farklar ve %95 güven aralıkları. Kesikli çizgi “fark yok” değerini gösterir.',
    `KIBT ile ilgili iki karşılaştırmanın güven aralıkları sıfırdan çok uzaktadır (≈ 1.7 mm fark; p < 0.001; r = ${f2(KP.r)}). Primescan – Füzyon farkının güven aralığı sıfırı içermekte ve çok dardır (${f3(PF.alt)} ile ${f3(PF.ust)} mm). Bu iki yöntem arasında anlamlı fark yoktur.`),

  // 6
  H1('6. Primescan ve Füzyon Model Arasındaki Uyum ve Eşdeğerlik'),
  P('Çalışmanın klinik açıdan en önemli sorusu, füzyon işleminin ağız içi taramanın doğruluğunu koruyup korumadığıdır. Bir testte “fark bulunamaması” iki yöntemin aynı olduğunu kanıtlamaz. Bu nedenle benzerlik üç tamamlayıcı yöntemle değerlendirilmiştir.'),
  CAP('**Tablo 4.** Primescan ile Füzyon model arasındaki uyum ve eşdeğerlik analizi'),
  table(['Analiz', 'Sonuç', 'Yorum'], [
    ['Ortalama fark (Primescan − Füzyon)', `${f3(U.ort)} mm (SS ${f3(U.ss)})`, 'Sistematik fark yok'],
    ['Bland–Altman %95 uyum sınırları', `${f3(U.loa[0])} ile ${f3(U.loa[1])} mm`, 'Tek bir hastada iki yöntem arasında beklenen fark aralığı'],
    ['Farkın ortalamaya göre eğilimi', `p = ${f3(U.ba_trend_p)}`, 'Farklar ölçüm büyüklüğüne bağlı değil'],
    ['ICC(A,1) (%95 GA)', `${f2(U.icc)} (${f2(U.icc_alt)} – ${f2(U.icc_ust)})`, 'İyi düzeyde uyum (küçük örneklem nedeniyle GA geniş)'],
    ['Spearman korelasyonu', `ρ = ${f2(U.spearman)}; p ${pf(U.spearman_p)}`, 'Güçlü pozitif ilişki'],
    ['Eşdeğerlik testi (TOST), ±0.10 mm', `%90 GA: ${f3(U.ci90[0])} ile ${f3(U.ci90[1])}; p ${pf(U.tost['0.1'])}`, '**Eşdeğer**'],
    ['Eşdeğerlik testi (TOST), ±0.05 mm', `p = ${f3(U.tost['0.05'])}`, 'Daha sıkı sınırda da eşdeğer'],
  ], [34, 36, 30]),
  NOTE('ICC yorumu: < 0.50 zayıf, 0.50–0.75 orta, 0.75–0.90 iyi, > 0.90 mükemmel (Koo ve Li, 2016). TOST’ta farkın %90 GA’sı eşdeğerlik sınırları içinde kalıyorsa iki yöntem eşdeğer kabul edilir.'),
  ...figure('s08_sacilim.png', 3.9, 'Primescan ve Füzyon sapmalarının katılımcı bazında ilişkisi. Kesikli çizgi iki yöntemin tam olarak aynı sonucu verdiği durumu gösterir.',
    `Noktalar özdeşlik doğrusunun her iki yanına dengeli dağılmıştır. Bir hastada ağız içi tarama ne kadar doğruysa füzyon model de o kadar doğrudur (ICC = ${f2(U.icc)}).`),
  ...figure('s09_bland_altman.png', 5.6, 'Bland–Altman grafiği: Primescan ve Füzyon model. Yeşil bant ±0.10 mm eşdeğerlik sınırını gösterir.',
    `Ortalama fark sıfıra çok yakındır (${f3(U.ort)} mm) ve farklar ölçüm büyüklüğüyle değişmemektedir. Bireysel farkların büyük bölümü ±0.10 mm bandı içindedir; uyum sınırları ±0.17 mm civarındadır.`),
  ...figure('s10_tost.png', 5.6, 'Eşdeğerlik testi: ortalama farkın %90 güven aralığı ve eşdeğerlik sınırları.',
    `Ortalama farkın %90 güven aralığı (${f3(U.ci90[0])} ile ${f3(U.ci90[1])} mm) ±0.10 mm sınırlarının ve hatta ±0.05 mm sınırlarının tamamen içindedir. **Füzyon model ile Primescan istatistiksel olarak eşdeğer doğruluğa sahiptir.**`),

  // 7
  H1('7. Örneklem Büyüklüğü ve Güç Analizi'),
  H2('7.1. Tasarıma uygun a priori örneklem hesabı'),
  P('Örneklem büyüklüğü, aynı katılımcıdan alınan ölçümlerin karşılaştırılmasına uygun olarak G*Power 3.1’de eşleştirilmiş örneklem testi esas alınarak hesaplanmıştır (literatüre dayalı etki büyüklüğü d = 1.2; güç = 0.80).'),
  CAP('**Tablo 5.** Örneklem büyüklüğünün hesaplanma adımları'),
  table(['Adım', 'Hesaplama', 'n'], [
    ['1. Eşleştirilmiş test', `d = 1.2; α = 0.05/3 = 0.0167 (iki yönlü, Bonferroni); güç = 0.80`, `${f2(G.apriori_n)} → 11`],
    ['2. Parametrik olmayan test payı', `Wilcoxon testinin göreli etkinliği (ARE = 0.864): ${f2(G.apriori_n)} / 0.864`, `${f2(G.apriori_n_np)} → 12`],
    ['3. Veri kaybı payı (%20)', '12 / (1 − 0.20)', '15'],
    ['Sonuç', `n = 15 ile gerçekleşen güç: α′ = 0.0167’de ${pct(G.pow15_bonf)}; α = 0.05’te ${pct(G.pow15)}`, '**15**'],
  ], [26, 60, 14], { boldLast: true }),
  ...figure('s11_guc_apriori.png', 6.2, '(a) Katılımcı sayısına göre istatistiksel güç (d = 1.2). (b) n = 15 ile farklı etki büyüklüklerini saptama gücü.',
    `(a) %80 güç 11 katılımcıda aşılmakta; 15 katılımcıda güç ${pct(G.pow15_bonf)}’dır. (b) 15 katılımcı ile Bonferroni düzeltmesi altında d = ${f2(G.mde15)} ve üzerindeki etkiler %80 güçle saptanabilir. Çalışmada gözlenen etkiler (KIBT karşılaştırmalarında d = ${f2(KP.dz)} ve ${f2(KF.dz)}) bu sınırın çok üzerindedir.`),
  H2('7.2. Gözlenen verilerle güç (post-hoc) doğrulaması'),
  CAP('**Tablo 6.** n = 15 için gözlenen istatistiksel güç'),
  table(['Analiz', 'Yöntem', 'Güç'], [
    ['Üç yöntemin genel karşılaştırması (Friedman)', `Merkezi olmayan χ² (λ = ${f2(FR.chi2)})`, pct(G.nc_chi2)],
    ['Üç yöntemin genel karşılaştırması (Friedman)', 'Bootstrap simülasyonu (gözlenen veri)', '≈ %100'],
    ['KIBT – Primescan; KIBT – Füzyon', 'Eşleştirilmiş test, α′ = 0.0167', '≈ %100'],
    ['Primescan – Füzyon eşdeğerliği (±0.10 mm)', 'TOST, Monte Carlo simülasyonu', pct(G.tost15['0.1'])],
  ], [40, 42, 18]),
  ...figure('s12_guc_posthoc.png', 6.2, '(a) Gözlenen verilerle Friedman testinin katılımcı sayısına göre gücü. (b) Primescan – Füzyon eşdeğerlik testinin gücü.',
    `(a) Gözlenen etki o kadar büyüktür ki Friedman testi 4 katılımcı ile bile %80 gücü aşmaktadır (n = 3’te ${pct(fs3)}). (b) ±0.10 mm sınırında eşdeğerliği %80 güçle göstermek için yaklaşık 9 katılımcı yeterlidir; 15 katılımcıda güç ${pct(G.tost15['0.1'])}’dir. Tüm analizler için **15 katılımcı yeterli istatistiksel gücü sağlamaktadır.**`),

  // 8
  H1('8. Sonuç'),
  NL(`Üç yöntemin doğruluğu arasında istatistiksel olarak anlamlı ve büyük bir fark vardır (Friedman χ²(2) = ${f2(FR.chi2)}; p < 0.001; W = ${f3(FR.W)}).`, 'n1'),
  NL(`Tek başına KIBT’den elde edilen sanal modeller, referansa göre en yüksek sapmayı (≈ 2.07 mm) göstermiştir ve ağız içi tarama ile füzyon modellere göre yaklaşık 1.7 mm daha az doğrudur (p < 0.001).`, 'n1'),
  NL('Füzyon model ile ağız içi tarama (Primescan) arasında anlamlı fark yoktur ve iki yöntem ±0.10 mm sınırında eşdeğerdir. Füzyon işlemi, ağız içi taramanın yüzey doğruluğunu korumaktadır.', 'n1'),
  NL('Tasarıma uygun a priori hesaplama ve gözlenen verilerle yapılan güç analizleri, 15 katılımcının çalışmanın tüm temel soruları için yeterli olduğunu göstermektedir.', 'n1'),
  P('**Sınırlılıklar:** KIBT ölçümlerinde iki katılımcının (K8, K14) değerleri grubun geri kalanından düşüktür; bu değerler sonuçları değiştirmemektedir. Eşdeğerlik sınırı (±0.10 mm), tezin yöntem bölümünde literatüre dayalı klinik gerekçesiyle birlikte belirtilmelidir.'),

  // 9
  new Paragraph({ children: [new PageBreak()] }),
  H1('9. Teze Aktarılabilecek Metin'),
  H2('9.1. Gereç ve Yöntem – Örneklem büyüklüğü'),
  P('Örneklem büyüklüğü, aynı katılımcılardan elde edilen bağımlı ölçümlerin karşılaştırılmasına uygun olarak G*Power 3.1 (Heinrich Heine Üniversitesi, Düsseldorf, Almanya) programında eşleştirilmiş örneklem testi esas alınarak hesaplanmıştır. Literatüre dayalı etki büyüklüğü 1.2, anlamlılık düzeyi üç ikili karşılaştırma için Bonferroni düzeltmesiyle 0.0167 (iki yönlü) ve güç 0.80 alındığında gerekli örneklem büyüklüğü 11 olarak bulunmuştur. Parametrik olmayan testlerin göreli etkinliği (0.864) dikkate alınarak bu sayı 12’ye, olası veri kayıpları için %20 ek pay ile 15 katılımcıya yükseltilmiştir.'),
  H2('9.2. Gereç ve Yöntem – İstatistiksel analiz'),
  P('Verilerin normal dağılıma uygunluğu Shapiro–Wilk testi ile değerlendirilmiştir. Verilerin bir kısmı normal dağılım göstermediğinden, üç yöntemin karşılaştırılmasında Friedman testi kullanılmış ve etki büyüklüğü Kendall’s W ile hesaplanmıştır. İkili karşılaştırmalar Wilcoxon işaretli sıralar testi ile yapılmış ve Bonferroni düzeltmesi uygulanmıştır. Primescan ve füzyon modeller arasındaki uyum Bland–Altman analizi ve sınıf içi korelasyon katsayısı [ICC(A,1)] ile; iki yöntemin eşdeğerliği ise ±0.10 mm eşdeğerlik sınırı ile iki tek yönlü test (TOST) yöntemiyle değerlendirilmiştir. İstatistiksel anlamlılık düzeyi p < 0.05 olarak kabul edilmiştir.'),
  H2('9.3. Bulgular'),
  P(`KIBT – referans tarama, Primescan – referans tarama ve füzyon model – referans veri karşılaştırmalarında ortalama sapma değerleri sırasıyla ${f3(K.ort)} ± ${f3(K.ss)} mm, ${f3(PS.ort)} ± ${f3(PS.ss)} mm ve ${f3(FU.ort)} ± ${f3(FU.ss)} mm olarak bulunmuştur (Tablo 1). Üç yöntem arasında istatistiksel olarak anlamlı fark saptanmıştır (Friedman χ²(2) = ${f2(FR.chi2)}; p < 0.001; Kendall’s W = ${f3(FR.W)}). Bonferroni düzeltmeli ikili karşılaştırmalarda KIBT modellerinin sapması hem Primescan (ortalama fark ${f3(KP.ort)} mm; p < 0.001; r = ${f2(KP.r)}) hem de füzyon modellere (ortalama fark ${f3(KF.ort)} mm; p < 0.001; r = ${f2(KF.r)}) göre anlamlı olarak yüksek bulunmuştur. Primescan ile füzyon modeller arasında anlamlı fark saptanmamıştır (ortalama fark ${f3(PF.ort)} mm; %95 GA ${f3(PF.alt)} ile ${f3(PF.ust)} mm; p = ${f3(PF.p_bonf)}). Farkın %90 güven aralığı (${f3(U.ci90[0])} ile ${f3(U.ci90[1])} mm) ±0.10 mm eşdeğerlik sınırları içinde kaldığından iki yöntem eşdeğer bulunmuştur (TOST p < 0.001). Bland–Altman analizinde ortalama fark ${f3(U.ort)} mm, %95 uyum sınırları ${f3(U.loa[0])} ile ${f3(U.loa[1])} mm; iki yöntem arasındaki uyum iyi düzeyde bulunmuştur [ICC(A,1) = ${f2(U.icc)}; %95 GA ${f2(U.icc_alt)} – ${f2(U.icc_ust)}].`),

  H1('Kaynaklar'),
  ...[
    'Bland JM, Altman DG. Statistical methods for assessing agreement between two methods of clinical measurement. ~Lancet~. 1986;1(8476):307–310.',
    'Cohen J. ~Statistical Power Analysis for the Behavioral Sciences~. 2nd ed. Hillsdale, NJ: Lawrence Erlbaum; 1988.',
    'Faul F, Erdfelder E, Lang AG, Buchner A. G*Power 3: A flexible statistical power analysis program for the social, behavioral, and biomedical sciences. ~Behav Res Methods~. 2007;39(2):175–191.',
    'Koo TK, Li MY. A guideline of selecting and reporting intraclass correlation coefficients for reliability research. ~J Chiropr Med~. 2016;15(2):155–163.',
    'Lakens D. Equivalence tests: a practical primer for t tests, correlations, and meta-analyses. ~Soc Psychol Personal Sci~. 2017;8(4):355–362.',
    'Lehmann EL. ~Nonparametrics: Statistical Methods Based on Ranks~. San Francisco: Holden-Day; 1975.',
    'Schuirmann DJ. A comparison of the two one-sided tests procedure and the power approach for assessing the equivalence of average bioavailability. ~J Pharmacokinet Biopharm~. 1987;15(6):657–680.',
    'Tomczak M, Tomczak E. The need to report effect size estimates revisited. ~Trends Sport Sci~. 2014;1(21):19–25.',
  ].map(t => new Paragraph({ children: runs(t, { size: 21 }), spacing: { after: 90 }, indent: { left: 360, hanging: 360 } })),

  new Paragraph({ children: [new PageBreak()] }),
  H1('Ek. Anonimleştirilmiş Ham Veri'),
  CAP('**Tablo E1.** Katılımcılara göre sapma değerleri (mm) ve katılımcı içi sıralar'),
  table(['Kod', 'KIBT – Referans', 'Primescan – Referans', 'Füzyon – Referans', 'Sıra (K / P / F)'],
    R.ham.map(r => [r[0], r[1].toFixed(4), r[2].toFixed(4), r[3].toFixed(4), `${r[4]} / ${r[5]} / ${r[6]}`])
      .concat([['Sıra toplamı', '', '', '', FR.sira_top.join(' / ')]]), [14, 21, 22, 21, 22], { boldLast: true }),
  NOTE('K: KIBT, P: Primescan, F: Füzyon; 1 = en düşük sapma. Hasta kimlik bilgileri gizlilik nedeniyle rapora alınmamıştır.'),
);

const num = (ref) => ({ reference: ref, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 500, hanging: 320 } } } }] });
const doc = new Document({
  title: 'Nihai İstatistiksel Analiz Raporu',
  styles: {
    default: { document: { run: { font: FONT, size: 23 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 28, bold: true, color: ACCENT, font: FONT }, paragraph: { spacing: { before: 320, after: 140 }, outlineLevel: 0, keepNext: true } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 24, bold: true, color: '2E5496', font: FONT }, paragraph: { spacing: { before: 220, after: 100 }, outlineLevel: 1, keepNext: true } },
    ],
  },
  numbering: { config: [
    { reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 460, hanging: 260 } } } }] },
    num('n1'),
  ] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], size: 18 })] })] }) },
    children: C,
  }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(process.argv[2], b); console.log('yazıldı', process.argv[2]); });
