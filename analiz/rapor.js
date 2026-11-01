const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
  AlignmentType, HeadingLevel, ImageRun, BorderStyle, LevelFormat, Footer, PageNumber,
  TableOfContents, PageBreak
} = require('docx');

const FONT = 'Times New Roman';
const W = 9026; // A4 metin genişliği (2.54 cm kenar boşluğu), DXA

// ---------- yardımcılar ----------
function runs(text, base = {}) {
  // **kalın** ve _italik_ işaretlerini çözümle
  const out = [];
  const re = /(\*\*[^*]+\*\*|~[^~]+~)/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), ...base }));
    const t = m[0];
    if (t.startsWith('**')) out.push(new TextRun({ ...base, text: t.slice(2, -2), bold: true }));
    else out.push(new TextRun({ ...base, text: t.slice(1, -1), italics: true }));
    last = m.index + t.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), ...base }));
  return out;
}
const P = (text, opt = {}) => new Paragraph({
  children: runs(text, opt.run || {}), alignment: opt.align || AlignmentType.JUSTIFIED,
  spacing: { before: 80, after: 120, line: 300, ...(opt.spacing || {}) }, indent: opt.indent, keepNext: opt.keepNext,
});
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)], spacing: { before: 300, after: 140 } });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)], spacing: { before: 220, after: 100 } });
const B = (text, lvl = 0) => new Paragraph({ children: runs(text), numbering: { reference: 'bul', level: lvl }, alignment: AlignmentType.JUSTIFIED, spacing: { after: 80, line: 288 } });
const N = (text, ref = 'num') => new Paragraph({ children: runs(text), numbering: { reference: ref, level: 0 }, alignment: AlignmentType.JUSTIFIED, spacing: { after: 80, line: 288 } });
const EQ = (text) => new Paragraph({ children: runs(text, { font: 'Cambria Math' }), alignment: AlignmentType.CENTER, spacing: { before: 60, after: 120 } });
const CAP = (text) => new Paragraph({ children: runs(text, { size: 22 }), spacing: { before: 200, after: 80 }, keepNext: true });
const NOTE = (text) => new Paragraph({ children: runs(text, { size: 18 }), spacing: { before: 40, after: 200 }, alignment: AlignmentType.JUSTIFIED });

const thin = { style: BorderStyle.SINGLE, size: 4, color: '808080' };
const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
function table(headers, rows, widths, opts = {}) {
  const sum = widths.reduce((a, b) => a + b, 0);
  const ws = widths.map(w => Math.round(w * W / sum));
  ws[ws.length - 1] += W - ws.reduce((a, b) => a + b, 0);
  const mk = (txt, i, hdr, shade) => new TableCell({
    width: { size: ws[i], type: WidthType.DXA },
    shading: hdr ? { type: ShadingType.CLEAR, fill: 'DCE6F0', color: 'auto' } : (shade ? { type: ShadingType.CLEAR, fill: shade, color: 'auto' } : undefined),
    margins: { top: 50, bottom: 50, left: 90, right: 90 },
    borders: { top: thin, bottom: thin, left: thin, right: thin },
    children: String(txt).split('\n').map(line => new Paragraph({
      children: runs(line, { size: 19, bold: hdr || undefined }),
      alignment: (i === 0 || opts.leftAll) ? AlignmentType.LEFT : AlignmentType.CENTER,
      spacing: { after: 0, line: 252 },
    })),
  });
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: ws,
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((h, i) => mk(h, i, true)) }),
      ...rows.map(r => new TableRow({ cantSplit: true, children: r.map((c, i) => mk(c, i, false, opts.shade && opts.shade(r))) })),
    ],
  });
}
function box(title, lines, fill = 'EEF4EA', border = '5B8C3A') {
  const b = { style: BorderStyle.SINGLE, size: 8, color: border };
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    rows: [new TableRow({ children: [new TableCell({
      width: { size: W, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, fill, color: 'auto' },
      borders: { top: b, bottom: b, left: b, right: b }, margins: { top: 120, bottom: 120, left: 160, right: 160 },
      children: [new Paragraph({ children: [new TextRun({ text: title, bold: true })], spacing: { after: 80 } }),
        ...lines.map(l => new Paragraph({ children: runs(l, { size: 21 }), numbering: { reference: 'bul', level: 0 }, spacing: { after: 60, line: 276 }, alignment: AlignmentType.JUSTIFIED }))],
    })] })],
  });
}
const SP = () => new Paragraph({ children: [], spacing: { after: 60 } });

// ---------- veri (anonim, K1–K15) ----------
const raw = [
  [1.9306, 0.3172, 0.3088], [2.1430, 0.4692, 0.4333], [2.4561, 0.2762, 0.2342], [2.0178, 0.3088, 0.2703],
  [2.1252, 0.2706, 0.1888], [2.2082, 0.2560, 0.2815], [2.5767, 0.5895, 0.6136], [1.1868, 0.6281, 0.4312],
  [1.9876, 0.4636, 0.5533], [2.5263, 0.3028, 0.3899], [2.2788, 0.3656, 0.4503], [2.0214, 0.3635, 0.3762],
  [1.9875, 0.3620, 0.5184], [1.4485, 0.5909, 0.4832], [2.1270, 0.2601, 0.2550],
];

const TITLE = 'Üst Çenesi Tam Dişsiz Hastalarda Dijital İş Akışı Kapsamında Farklı Tekniklerle Elde Edilen Üç Boyutlu Sanal Modellerin Doğruluklarının İn Vivo Olarak Değerlendirilmesi';

const img = fs.readFileSync(path.join(__dirname, 'sekil1.png'));

const children = [
  new Paragraph({ children: [new TextRun({ text: 'ARA İSTATİSTİKSEL DEĞERLENDİRME VE GÜÇ ANALİZİ RAPORU', bold: true, size: 30 })], alignment: AlignmentType.CENTER, spacing: { after: 200 } }),
  new Paragraph({ children: [new TextRun({ text: 'Örneklem büyüklüğünün n = 15 olarak revize edilmesine ilişkin doğrulama analizi', italics: true, size: 24 })], alignment: AlignmentType.CENTER, spacing: { after: 300 } }),
  table(['Başlık', 'Bilgi'], [
    ['Çalışmanın adı', TITLE],
    ['Araştırma niteliği', 'Uzmanlık tezi – Ankara Üniversitesi Diş Hekimliği Fakültesi'],
    ['Araştırma tasarımı', 'Aynı katılımcılarda üç farklı yöntemle elde edilen sanal modellerin referans veri ile karşılaştırıldığı tekrarlı ölçüm (bağımlı örneklem, within-subject) tasarımı'],
    ['Bağımlı değişken', 'Test modeli ile referans model arasındaki yüzey sapması (mm)'],
    ['Karşılaştırılan yöntemler', '(1) KIBT – Referans tarama  (2) Primescan (ağız içi tarayıcı) – Referans tarama  (3) Füzyon model – Referans veri'],
    ['Değerlendirilen katılımcı sayısı', 'n = 15 (eksik veri yok; 15 × 3 = 45 ölçüm)'],
    ['Anlamlılık düzeyi', 'α = 0.05 (iki yönlü); ikili karşılaştırmalarda Bonferroni düzeltmesi (α′ = 0.05/3 = 0.0167)'],
    ['Analiz yazılımı', 'Python 3 (SciPy, statsmodels); sonuçlar SPSS 25 ile birebir yeniden üretilebilir'],
  ], [28, 72], { leftAll: true }),
  SP(),
  box('Özet bulgular', [
    'Ara raporda verilen tanımlayıcı istatistikler, Friedman χ²(2) = 22.53 (p < 0.001) ve Kendall’s W = 0.751 değerleri ham veriden **birebir doğrulanmıştır**.',
    'Üç yöntem arasındaki genel farklılık için gözlenen güç, kullanılan yönteme göre **%99.3 (yaklaşık yöntem) ile ≈ %100 (simülasyon)** arasındadır; ara rapordaki “≈ %99” ifadesi doğrudur.',
    'Genel farklılığın tamamı KIBT modelinden kaynaklanmaktadır: KIBT, hem Primescan hem Füzyon modelden ortalama **≈ 1.68 mm** daha fazla sapma göstermiştir (her iki karşılaştırmada Bonferroni düzeltmeli p < 0.001, güç ≈ %100).',
    'Primescan ile Füzyon model arasında fark **yoktur** (ortalama fark 0.002 mm; p = 0.92). Bu karşılaştırma için “fark” testinin gücü yalnızca **%5**’tir; bu nedenle iki yöntemin benzerliği **eşdeğerlik testi (TOST)** ile gösterilmiştir: ±0.10 mm eşdeğerlik sınırında p < 0.001 (güç %98.7).',
    'Tekrarlı ölçüm tasarımına uygun (eşleştirilmiş) yeniden yapılan a priori hesaplama, ilk analizdeki etki büyüklüğü (d = 1.2) ve %80 güç hedefiyle **15 katılımcının yeterli olduğunu** bağımsız olarak doğrulamaktadır (11 → parametrik olmayan düzeltme ile 12 → %20 kayıp payı ile 15).',
    'Önceki güç analizinde, test seçiminden (bağımsız iki grup t testi) kaynaklanan önemli bir yöntemsel hata ve belgeler arasında çeşitli tutarsızlıklar saptanmıştır (Bölüm 9).',
  ]),

  // ---------------- 1 ----------------
  H1('1. Amaç ve Veri Seti'),
  P('Bu raporun amacı; (i) mevcut 15 katılımcıdan elde edilen verilerin istatistiksel analizini yapmak, (ii) ara raporda sunulan Friedman testi, Kendall’s W ve post-hoc güç değerlerini ham veriden bağımsız olarak doğrulamak, (iii) planlanan örneklem büyüklüğünün 15 katılımcıya indirilmesinin istatistiksel olarak savunulabilir olup olmadığını değerlendirmek ve (iv) etik kurul başvurusunda ve ilk güç analizi raporunda yer alan istatistiksel kararları gözden geçirerek varsa hataları belirtmektir.'),
  P('Veri setinde her katılımcı için üç ölçüm bulunmaktadır: KIBT’den elde edilen modelin referans taramaya göre sapması, Primescan ağız içi tarayıcı ile elde edilen modelin referans taramaya göre sapması ve KIBT ile ağız içi tarama verilerinin birleştirilmesiyle oluşturulan füzyon modelin referans veriye göre sapması. Ölçümler mm cinsindendir; değer büyüdükçe referanstan sapma artmakta, yani doğruluk azalmaktadır. Hasta kimlik bilgileri analiz ve rapordan çıkarılmış, katılımcılar K1–K15 olarak kodlanmıştır (Ek-1). Eksik veri bulunmamaktadır.'),

  // ---------------- 2 ----------------
  H1('2. İstatistiksel Yöntem'),
  B('Tanımlayıcı istatistikler ortalama ± standart sapma (SS), ortalamanın %95 güven aralığı (GA), ortanca (çeyrekler arası aralık, ÇAA) ve minimum–maksimum olarak verilmiştir.'),
  B('Normal dağılıma uygunluk, küçük örneklemlerde daha güçlü olan **Shapiro–Wilk** testi ile hem her yöntem için hem de ikili farklar için değerlendirilmiştir.'),
  B('Üç bağımlı ölçümün genel karşılaştırması **Friedman testi** ile yapılmış, etki büyüklüğü **Kendall’s W** ile verilmiştir. Duyarlılık analizi olarak tekrarlı ölçümlerde varyans analizi (RM-ANOVA; Mauchly küresellik testi ve Greenhouse–Geisser düzeltmesi ile) uygulanmıştır.'),
  B('İkili karşılaştırmalar **Wilcoxon işaretli sıralar testi** (kesin p) ile yapılmış ve Bonferroni düzeltmesi uygulanmıştır (düzeltilmiş p = ham p × 3). SPSS’in Friedman ikili karşılaştırma çıktısı (Dunn–Bonferroni) ve eşleştirilmiş t testi destekleyici olarak raporlanmıştır. Etki büyüklüğü olarak r = Z/√n ve Cohen’s d_z = ortalama fark / farkların SS’si kullanılmıştır.'),
  B('Primescan ve Füzyon model arasındaki benzerlik, “fark yok” hipotezinin anlamsızlığı ile değil, **iki tek yönlü test (TOST) eşdeğerlik yaklaşımı** ve **Bland–Altman uyum analizi** ile değerlendirilmiştir.'),
  B('Post-hoc güç; (a) merkezi olmayan χ² yaklaşımı, (b) 20 000 tekrarlı bootstrap ve çok değişkenli normal Monte Carlo simülasyonu, (c) RM-ANOVA için merkezi olmayan F dağılımı ile hesaplanmıştır. Örneklem büyüklüğü yeniden hesaplaması eşleştirilmiş t testi gücü üzerinden yapılmıştır (G*Power: “Means: Difference between two dependent means (matched pairs)”).'),

  // ---------------- 3 ----------------
  H1('3. Tanımlayıcı Bulgular'),
  CAP('**Tablo 1.** Yöntemlere göre yüzey sapması değerlerinin tanımlayıcı istatistikleri (n = 15, mm)'),
  table(['Yöntem', 'Ortalama ± SS', '%95 GA', 'Ortanca (ÇAA)', 'Min – Maks'], [
    ['KIBT – Referans tarama', '2.068 ± 0.368', '1.865 – 2.272', '2.125 (1.988 – 2.244)', '1.187 – 2.577'],
    ['Primescan – Referans tarama', '0.388 ± 0.129', '0.317 – 0.460', '0.362 (0.290 – 0.466)', '0.256 – 0.628'],
    ['Füzyon model – Referans veri', '0.386 ± 0.127', '0.316 – 0.456', '0.390 (0.276 – 0.467)', '0.189 – 0.614'],
  ], [30, 17, 17, 21, 15]),
  NOTE('SS: standart sapma; GA: güven aralığı; ÇAA: çeyrekler arası aralık (25.–75. yüzdelik).'),
  P('Ara raporda verilen ortalama ve standart sapma değerleri (2.068 ± 0.368; 0.388 ± 0.129; 0.386 ± 0.127) ham veriden **doğru** olarak hesaplanmıştır. KIBT modelinin ortalama sapması, ağız içi tarama ve füzyon modelin yaklaşık 5.3 katıdır. Primescan ve Füzyon model ortalamaları birbirine neredeyse eşittir (fark 0.002 mm = 2 µm). KIBT ölçümlerinde iki katılımcının değeri (K8: 1.187 mm; K14: 1.449 mm) kutu grafiğinde alt uç değer olarak görünmektedir; bu değerlerin ölçüm/segmentasyon hatasından kaynaklanıp kaynaklanmadığı kaynak verilerden kontrol edilmelidir. Bu iki değer, KIBT’nin diğer yöntemlerden daha kötü olduğu sonucunu değiştirmemektedir (her iki katılımcıda da KIBT sapması diğer iki yöntemden belirgin olarak yüksektir).'),
  new Paragraph({ children: [new ImageRun({ type: 'png', data: img, transformation: { width: 600, height: 252 } })], alignment: AlignmentType.CENTER, spacing: { before: 120, after: 60 } }),
  NOTE('**Şekil 1.** (a) Yöntemlere göre sapma değerleri; gri çizgiler aynı katılımcının üç ölçümünü birleştirmektedir. (b) Primescan ile Füzyon model arasındaki uyumun Bland–Altman grafiği (ortalama fark ve %95 uyum sınırları).'),

  // ---------------- 4 ----------------
  H1('4. Normal Dağılım Değerlendirmesi'),
  CAP('**Tablo 2.** Shapiro–Wilk normallik testi sonuçları'),
  table(['Değişken', 'W', 'p', 'Yorum'], [
    ['KIBT – Referans', '0.898', '0.088', 'Normal dağılım varsayımı reddedilmedi (sınırda)'],
    ['Primescan – Referans', '0.855', '0.021', 'Normal dağılmıyor'],
    ['Füzyon – Referans', '0.967', '0.817', 'Normal dağılım varsayımı reddedilmedi'],
    ['Fark: KIBT − Primescan', '—', '0.015', 'Normal dağılmıyor'],
    ['Fark: KIBT − Füzyon', '—', '0.126', 'Normal dağılım varsayımı reddedilmedi'],
    ['Fark: Primescan − Füzyon', '—', '0.931', 'Normal dağılım varsayımı reddedilmedi'],
  ], [34, 12, 12, 42]),
  P('Primescan ölçümleri ve KIBT − Primescan farkları normal dağılmamaktadır (p < 0.05). Ayrıca RM-ANOVA’nın küresellik varsayımı ciddi biçimde ihlal edilmiştir (Mauchly W = 0.090; χ²(2) = 31.34; p < 0.001; Greenhouse–Geisser ε = 0.52). Örneklemin küçük olması (n = 15) da dikkate alındığında, ara raporda olduğu gibi **parametrik olmayan Friedman testinin seçilmesi doğrudur**.'),

  // ---------------- 5 ----------------
  H1('5. Üç Yöntemin Genel Karşılaştırılması'),
  H2('5.1. Friedman testi'),
  P('Her katılımcıda üç yöntem kendi içinde sıralanmıştır (1 = en düşük sapma). Ortalama sıralar: KIBT 3.00, Primescan 1.53, Füzyon 1.47. KIBT, 15 katılımcının tamamında en yüksek sapmayı göstermiştir.'),
  EQ('χ²_F = [12 / (n·k·(k+1))] · ΣR_j² − 3·n·(k+1)'),
  EQ('χ²_F = [12 / (15·3·4)] · (45² + 23² + 22²) − 3·15·4 = (12/180) · 3038 − 180 = 22.53'),
  P('**Friedman χ²(2) = 22.53; p = 1.28 × 10⁻⁵ (p < 0.001).** Üç yöntem arasında istatistiksel olarak anlamlı genel farklılık vardır.'),
  H2('5.2. Etki büyüklüğü: Kendall’s W'),
  EQ('W = χ²_F / [n · (k − 1)] = 22.53 / (15 · 2) = 0.751'),
  P('Kendall’s W için yaygın kabul gören eşiklere göre (0.1 küçük, 0.3 orta, ≥ 0.5 büyük) W = 0.751 **büyük** bir etkiyi göstermektedir. Ara rapordaki hesaplama ve yorum doğrudur. Ancak W’nin 1’e ulaşmamasının nedeni, Primescan ve Füzyon’un katılımcılar arasında tutarlı bir sıra oluşturmamasıdır (yaklaşık yarısında Primescan, yarısında Füzyon daha düşük sapma göstermiştir); yani “güçlü genel etki” tümüyle KIBT’nin diğer iki yöntemden ayrışmasından kaynaklanmaktadır. Bu durum tez metninde açıkça belirtilmelidir.'),
  H2('5.3. Duyarlılık analizi: Tekrarlı ölçümlerde ANOVA'),
  P('Parametrik alternatif olarak uygulanan RM-ANOVA aynı sonucu vermiştir: F(2, 28) = 230.09; Greenhouse–Geisser düzeltmesi ile F(1.05, 14.66) = 230.09, p < 0.001; kısmi η² = 0.943 (Cohen’s f = 4.05). Sonuç, test seçiminden bağımsızdır.'),

  // ---------------- 6 ----------------
  H1('6. İkili (Post-hoc) Karşılaştırmalar'),
  P('Ara raporda Friedman testinin anlamlı çıkmasından sonra hangi yöntemlerin birbirinden farklı olduğu gösterilmemiştir. Omnibus testin anlamlı olması tek başına yeterli değildir; ikili karşılaştırmalar aşağıda verilmiştir.'),
  CAP('**Tablo 3.** Yöntemlerin ikili karşılaştırılması (n = 15)'),
  table(['Karşılaştırma', 'Ort. fark (mm)\n[%95 GA]', 'Wilcoxon\nT; p (kesin)', 'Bonferroni\np', 'Etki\nr | d_z', 'Dunn–Bonf.\np (SPSS)'], [
    ['KIBT vs Primescan', '1.680\n[1.432 – 1.928]', 'T = 0; p < 0.001', '< 0.001', '0.88 | 3.76', '< 0.001'],
    ['KIBT vs Füzyon', '1.682\n[1.460 – 1.904]', 'T = 0; p < 0.001', '< 0.001', '0.88 | 4.20', '< 0.001'],
    ['Primescan vs Füzyon', '0.002\n[−0.047 – 0.052]', 'T = 59; p = 0.978', '1.000', '0.01 | 0.03', '1.000'],
  ], [22, 18, 18, 13, 14, 15]),
  NOTE('Wilcoxon kesin p (KIBT karşılaştırmaları) = 6.1 × 10⁻⁵; Bonferroni ile 1.8 × 10⁻⁴. T = 0: 15 katılımcının tamamında KIBT sapması diğer yöntemden büyüktür. Eşleştirilmiş t testi: KIBT–Primescan t(14) = 14.55; KIBT–Füzyon t(14) = 16.26 (her ikisi p < 0.001); Primescan–Füzyon t(14) = 0.11, p = 0.918.'),
  P('**Yorum:** KIBT’den elde edilen modeller, hem ağız içi tarama (Primescan) hem de füzyon modele göre referanstan anlamlı ve klinik olarak büyük ölçüde (≈ 1.7 mm) daha fazla sapmaktadır. Primescan ve füzyon model arasında anlamlı bir fark bulunmamıştır.'),

  // ---------------- 7 ----------------
  H1('7. Primescan ile Füzyon Modelin Eşdeğerliği'),
  P('“Anlamlı fark bulunamadı” (p = 0.92) ifadesi, iki yöntemin “aynı” olduğunu kanıtlamaz; yalnızca farkın gösterilemediğini ifade eder. Bu karşılaştırma tezin klinik açıdan en önemli sorusu olduğundan (füzyon modelin ağız içi taramanın doğruluğunu koruyup korumadığı), benzerlik **eşdeğerlik testi** ile gösterilmelidir. TOST yaklaşımında, ortalama farkın %90 güven aralığı önceden belirlenen eşdeğerlik sınırları (±Δ) içinde kalıyorsa iki yöntem eşdeğer kabul edilir.'),
  P('Primescan − Füzyon ortalama farkı = 0.0024 mm; farkların SS’si = 0.0887 mm; **%90 GA: −0.038 ile +0.043 mm.**'),
  CAP('**Tablo 4.** Farklı eşdeğerlik sınırları için TOST sonuçları'),
  table(['Eşdeğerlik sınırı (±Δ)', 'TOST p', 'Sonuç', 'n = 15 için güç', '%80 güç için gereken n'], [
    ['±0.05 mm (50 µm)', '0.028', 'Eşdeğer', '%34', '29'],
    ['±0.10 mm (100 µm)', '< 0.001', 'Eşdeğer', '%98.7', '9'],
    ['±0.30 mm (300 µm)*', '< 0.001', 'Eşdeğer', '≈ %100', '3'],
  ], [26, 14, 16, 20, 24]),
  NOTE('* Etik kurul başvurusunda hareketli protezler için klinik tolerans olarak bildirilen 300 µm doku deplasmanı değeri. Güç: gerçek farkın 0 ve farkların SS’sinin 0.0887 mm olduğu varsayımıyla 50 000 tekrarlı Monte Carlo simülasyonu.'),
  P('Klinik olarak savunulabilir ve oldukça sıkı bir sınır olan ±0.10 mm alındığında iki yöntemin eşdeğer olduğu yüksek güçle (%98.7) gösterilmiştir. Bland–Altman analizinde ortalama fark 0.002 mm, %95 uyum sınırları −0.172 ile +0.176 mm olup sistematik bir sapma (bias) gözlenmemiştir (Şekil 1b). **Önemli not:** Eşdeğerlik sınırı veriye bakıldıktan sonra değil, klinik gerekçeyle belirlenmelidir; tezde kullanılacak sınır (öneri: ±0.10 mm) literatürle gerekçelendirilmeli ve yöntem bölümünde önceden tanımlanmış olarak yazılmalıdır.'),

  // ---------------- 8 ----------------
  H1('8. Güç Analizi ve Örneklem Büyüklüğünün Doğrulanması'),
  H2('8.1. Üç yöntemin genel karşılaştırması için post-hoc güç'),
  P('Ara raporda belirtilen “≈ %99” değeri, hesaplama yöntemi belirtilmeden verilmiştir. Bu değer, Friedman istatistiğinin H₁ altında yaklaşık olarak merkezi olmayan χ² dağılımına uyduğu ve merkezi olmayanlık parametresinin gözlenen χ² değerine eşit alındığı (λ ≈ 22.53) yaklaşımla birebir yeniden üretilmiştir. Daha doğrudan bir yöntem olan simülasyon da aynı sonucu desteklemektedir.'),
  CAP('**Tablo 5.** Genel karşılaştırma (3 yöntem) için post-hoc güç (n = 15, α = 0.05)'),
  table(['Yöntem', 'Girdi', 'Güç'], [
    ['Merkezi olmayan χ² yaklaşımı (ara rapordaki yöntem)', 'sd = 2; λ = 22.53; kritik χ² = 5.99', '%99.3'],
    ['Bootstrap simülasyonu (Friedman, 20 000 tekrar)', 'Katılımcılar yerine konarak yeniden örnekleme', '≈ %100'],
    ['Çok değişkenli normal simülasyon (Friedman)', 'Gözlenen ortalamalar ve kovaryans matrisi', '≈ %100'],
    ['RM-ANOVA (G*Power “within factors”)', 'f = 4.05; ε = 0.52 (GG)', '≈ %100'],
  ], [42, 40, 18]),
  P('Sonuç olarak, mevcut 15 katılımcı, üç yöntemin genel karşılaştırması için **en az %99** istatistiksel güç sağlamaktadır. Simülasyonlar, gözlenen etki büyüklüğünde Friedman testinin %80 güce yalnızca **4 katılımcı** ile ulaştığını göstermektedir (n = 3’te %26; n = 4 ve üzerinde ≈ %100).'),

  H2('8.2. İkili karşılaştırmalar için güç'),
  CAP('**Tablo 6.** İkili karşılaştırmalar için gözlenen güç (n = 15, Bonferroni α′ = 0.0167, iki yönlü)'),
  table(['Karşılaştırma', 'Gözlenen d_z', 'Güç (n = 15)', 'Yorum'], [
    ['KIBT vs Primescan', '3.76', '≈ %100', 'Yeterli'],
    ['KIBT vs Füzyon', '4.20', '≈ %100', 'Yeterli'],
    ['Primescan vs Füzyon (fark testi)', '0.03', '%1.7 (α = 0.05’te %5.1)', 'Fark testi anlamsız; bu farkı yakalamak için ≈ 10 600 katılımcı gerekir'],
    ['Primescan vs Füzyon (TOST, ±0.10 mm)', '—', '%98.7', 'Eşdeğerlik için yeterli'],
  ], [30, 14, 22, 34]),
  P('Bu tablo, ara rapordaki “%99 güç” ifadesinin **yalnızca genel (omnibus) karşılaştırma ve KIBT ile ilgili karşılaştırmalar** için geçerli olduğunu göstermektedir. Primescan ile Füzyon model arasındaki 2 µm’lik fark klinik olarak önemsizdir ve hiçbir makul örneklem büyüklüğüyle “anlamlı fark” olarak gösterilemez; bu karşılaştırma için doğru soru “fark var mı?” değil “eşdeğer mi?” sorusudur ve bu soruya 15 katılımcı ile yeterli güçle yanıt verilmektedir (Bölüm 7).'),

  H2('8.3. Örneklem büyüklüğünün 15’e indirilmesinin a priori gerekçesi'),
  P('Post-hoc güç, gözlenen p değerinin matematiksel bir dönüşümüdür ve örneklem revizyonunu tek başına gerekçelendirmek için yeterli kabul edilmez (Hoenig ve Heisey, 2001). Bu nedenle, örneklem büyüklüğü ilk güç analizindeki **aynı varsayımlar** (etki büyüklüğü 1.2; α = 0.05; güç = %80) kullanılarak, ancak bu kez çalışma tasarımına uygun **eşleştirilmiş (bağımlı) örneklem** testi için yeniden hesaplanmıştır:'),
  CAP('**Tablo 7.** Tasarıma uygun a priori örneklem büyüklüğü hesaplaması'),
  table(['Adım', 'Gerekçe', 'n'], [
    ['1. Eşleştirilmiş t testi; d_z = 1.2; α = 0.05/3 = 0.0167 (iki yönlü, Bonferroni); güç = 0.80', 'G*Power: Means – Difference between two dependent means (matched pairs). Üç ikili karşılaştırma için Bonferroni düzeltmesi uygulandı.', '10.3 → 11'],
    ['2. Parametrik olmayan test düzeltmesi', 'Wilcoxon testinin t testine göre en kötü durum asimptotik göreli etkinliği (ARE = 0.864): 10.27 / 0.864', '11.9 → 12'],
    ['3. Çalışma dışı kalma / kullanılamaz veri payı (%20)', '12 / (1 − 0.20) = 15 (ilk raporla tutarlı olarak %10 + %10 ek pay)', '**15**'],
  ], [36, 50, 14]),
  P('n = 15 katılımcı ile d_z = 1.2 için gerçekleşen güç Bonferroni düzeltmeli α′ = 0.0167’de **%96.0**, α = 0.05’te **%99.1**’dir. Başka bir deyişle, 15 katılımcı ile Bonferroni düzeltmesi altında %80 güçle saptanabilecek en küçük etki büyüklüğü d_z = 0.93’tür. Gözlenen etkiler (KIBT karşılaştırmalarında d_z = 3.76 ve 4.20) bu değerin çok üzerindedir.'),
  P('Eşleştirilmiş tasarımda d_z = d / √(2(1 − ρ)) ilişkisi geçerlidir. Burada d_z = d = 1.2 alınması, ölçümler arası korelasyonun ρ = 0.5 olduğu varsayımına karşılık gelir ve literatürdeki bağımsız grup etki büyüklüğünün doğrudan kullanılmasına göre muhafazakâr (temkinli) bir yaklaşımdır.'),

  H2('8.4. Örneklem büyüklüğüne ilişkin genel değerlendirme'),
  P('Üç bağımsız gerekçe aynı sonuca ulaşmaktadır: (1) gözlenen verilerle genel karşılaştırma için post-hoc güç ≥ %99; (2) gözlenen etki için %80 güce ulaşmak için gereken en küçük örneklem 4 katılımcıdır; (3) tasarıma uygun yeniden yapılan a priori hesaplama, ilk analizdeki etki büyüklüğü ile 15 katılımcıyı vermektedir. **Bu nedenle n = 15, çalışmanın birincil sorusu (üç yöntemin doğruluk açısından karşılaştırılması) için istatistiksel olarak yeterli ve savunulabilir bir örneklem büyüklüğüdür.**'),
  P('Sınırlılık: Ara analiz sonrasında örneklem büyüklüğünün değiştirilmesi, yöntem bölümünde ve etik kurula bildirilecek protokol değişikliğinde şeffaf biçimde belirtilmelidir. Bu çalışmada p değerleri anlamlılık sınırının çok altında (p < 0.001) olduğundan ara analizin I. tip hata oranına etkisi pratikte önemsizdir. Ayrıca bölgesel analizler (damak, alveolar kret, posterior palatal alan) planlanmışsa, bunlar için ayrıca güç değerlendirmesi yapılmalıdır; mevcut veri setinde yalnızca genel sapma değerleri bulunmaktadır.'),

  // ---------------- 9 ----------------
  H1('9. Önceki Analizlerin Değerlendirilmesi: Tespit Edilen Hatalar ve Tutarsızlıklar'),
  P('Ara rapor, ilk güç analizi belgesi (“İstatistik Değerlendirme Çalışması”, G*Power çıktısı) ve etik kurul başvuru formu ayrıntılı olarak incelenmiştir. Ara rapordaki **sayısal değerlerin tamamı doğrudur**; hatalar ağırlıklı olarak ilk güç analizinin tasarım seçiminde ve belgeler arasındaki tutarlılıktadır.'),
  CAP('**Tablo 8.** Tespit edilen hatalar, önem düzeyleri ve önerilen düzeltmeler'),
  table(['#', 'Belge', 'Tespit', 'Neden sorunlu / Önerilen düzeltme', 'Önem'], [
    ['1', 'İlk güç analizi (G*Power)', 'Test olarak “Means: Difference between two **independent** means (two groups)” seçilmiş; “her grupta n = 10” denmiştir.', 'Çalışmada iki ayrı hasta grubu yoktur; aynı hastadan üç ölçüm alınmaktadır (tekrarlı ölçüm). Doğru test “dependent means (matched pairs)” veya “ANOVA: repeated measures, within factors”dır. Bağımsız grup testi, aynı kişide ölçüm yapmanın sağladığı istatistiksel avantajı yok sayar. Bölüm 8.3’teki hesaplama kullanılmalıdır.', 'Yüksek'],
    ['2', 'İlk güç analizi', '“β = 0,80” yazılmıştır.', 'β, II. tip hata olasılığıdır. Doğrusu: “güç (1 − β) = 0.80”, yani β = 0.20.', 'Orta'],
    ['3', 'İlk güç analizi', 'Tek yönlü (“Tail(s): One”) test kullanılmıştır.', 'Yöntemlerden hangisinin daha doğru olacağı önceden kesin olmadığından iki yönlü test kullanılmalıdır. İki yönlü test ile aynı hatalı tasarımda bile 12 + 12 = 24 gerekirdi; tek yönlü seçim gerekçelendirilmemiştir.', 'Orta'],
    ['4', 'İlk güç analizi', 'Analiz planında “gruplar t testi ile karşılaştırılacaktır” denmiştir.', 'Üç yöntem karşılaştırıldığından iki örneklem t testi uygun değildir; Friedman (veya RM-ANOVA) + çoklu karşılaştırma düzeltmeli ikili testler gereklidir. Ara raporda Friedman doğru seçilmiştir.', 'Yüksek'],
    ['5', 'İlk güç analizi', 'Normallik için Kolmogorov–Smirnov ve Shapiro–Wilk birlikte önerilmiştir.', 'n < 50 için Shapiro–Wilk tercih edilmelidir; K–S küçük örneklemde gücü düşüktür.', 'Düşük'],
    ['6', 'Etik kurul formu', '“Güç analizi %95 güven aralığı ve %80 güç hedeflenerek yapılmıştır.”', 'Güç analizinde girdi güven aralığı değil anlamlılık düzeyidir: “α = 0.05 (iki yönlü) ve %80 güç” şeklinde yazılmalıdır.', 'Düşük'],
    ['7', 'Etik kurul formu – ilk güç analizi', 'Etik kurulda 24 hasta; G*Power çıktısında 20 hasta.', '24 = 20 × 1.2 (%10 + %10 ek pay) olduğu belgelerde açıkça yazılmalıdır. Revizyonla n = 15 olduğundan etik kurula protokol değişikliği bildirilmelidir.', 'Orta'],
    ['8', 'Etik kurul formu – veri dosyası', 'Etik formda ağız içi tarayıcı “Trios 5”; veri dosyasında “Primescan”.', 'Kullanılan cihaz tezde ve etik kurul kayıtlarında tutarlı olmalıdır. Cihaz değiştiyse protokol değişikliği olarak bildirilmelidir.', 'Yüksek'],
    ['9', 'Ara rapor', 'Post-hoc güç (≈ %99) hesaplama yöntemi belirtilmemiştir.', 'Yöntem (merkezi olmayan χ², λ = 22.53) ve yazılım belirtilmelidir. Değer doğrudur.', 'Düşük'],
    ['10', 'Ara rapor', 'Friedman sonrasında ikili karşılaştırma yapılmamıştır.', 'Hangi yöntemlerin farklı olduğu gösterilmelidir (Tablo 3).', 'Yüksek'],
    ['11', 'Ara rapor', '%99 gücün tüm araştırma sorusu için geçerli olduğu ima edilmektedir.', 'Güç yalnızca genel fark ve KIBT karşılaştırmaları için yüksektir. Primescan–Füzyon fark testi için güç %5’tir; bu karşılaştırma eşdeğerlik testiyle raporlanmalıdır (Bölüm 7).', 'Yüksek'],
    ['12', 'Ara rapor', 'Post-hoc güç örneklem revizyonunun temel gerekçesi olarak sunulmuştur.', 'Gözlenen güç p değerinin bir fonksiyonudur ve tek başına yetersiz kabul edilir. Bölüm 8.3’teki a priori gerekçe birlikte sunulmalıdır.', 'Orta'],
    ['13', 'Ara rapor', 'Terminoloji: “Referans tarama” ve “Referans veri” birlikte kullanılmış; Shapiro–Wilk sonuçları yok.', 'Referansın her üç karşılaştırmada aynı model olup olmadığı açıkça yazılmalıdır; normallik sonuçları Friedman seçiminin gerekçesi olarak eklenmelidir.', 'Düşük'],
    ['14', 'Etik kurul formu', 'Araştırma türü “retrospektif”; ancak yöntemde aynı hastalardan yeni aljinat ölçü ve ağız içi tarama alınacağı belirtilmiştir.', 'KIBT verisi retrospektif, diğer ölçümler prospektiftir. Tasarım “retrospektif KIBT verisi kullanılan prospektif, kesitsel in vivo çalışma” olarak tanımlanmalıdır.', 'Orta'],
    ['15', 'Etik kurul formu', 'Dahil edilmeme kriteri 6: “Damak bölgesinde kazanılmış deformitesi **olmayan**”.', 'Yazım hatası; “olan” olmalıdır (aksi hâlde uygun hastalar dışlanmış olur).', 'Düşük'],
    ['16', 'Veri dosyası', 'Excel dosyasında hasta ad-soyadları açık olarak yer almaktadır.', 'Etik formda “veriler anonimleştirilecektir” denmiştir. Analiz dosyalarında kod (K1–K15) kullanılmalı, isim–kod listesi ayrı ve güvenli saklanmalıdır.', 'Orta'],
  ], [5, 14, 26, 45, 10], { leftAll: true, shade: r => r[4] === 'Yüksek' ? 'FBEAEA' : undefined }),

  // ---------------- 10 ----------------
  H1('10. Sonuç'),
  N('Ara rapordaki tüm sayısal bulgular (tanımlayıcı istatistikler, Friedman χ²(2) = 22.53, p < 0.001, Kendall’s W = 0.751, post-hoc güç ≈ %99) ham veriden doğrulanmıştır.', 'num2'),
  N('KIBT’den elde edilen sanal modeller, referansa göre ağız içi tarama ve füzyon modele kıyasla yaklaşık 1.7 mm daha fazla sapma göstermektedir (p < 0.001; büyük etki).', 'num2'),
  N('Füzyon model, ağız içi tarama (Primescan) ile eşdeğer doğruluk sağlamaktadır (fark 0.002 mm; ±0.10 mm sınırında TOST p < 0.001).', 'num2'),
  N('n = 15; genel karşılaştırma ve KIBT ile ilgili ikili karşılaştırmalar için ≥ %99, eşdeğerlik analizi için %98.7 güç sağlamaktadır. Tasarıma uygun yeniden yapılan a priori hesaplama da 15 katılımcıyı desteklemektedir. **Örneklem büyüklüğünün 15’e indirilmesi istatistiksel olarak gerekçelendirilebilir.**', 'num2'),
  N('İlk güç analizindeki test seçimi hatası (bağımsız iki grup t testi), tek yönlü test ve “β = 0.80” ifadesi ile belgeler arası tutarsızlıklar (cihaz adı, örneklem sayısı, araştırma türü) tezde ve etik kurul bildiriminde düzeltilmelidir.', 'num2'),

  // ---------------- 11 ----------------
  new Paragraph({ children: [new PageBreak()] }),
  H1('11. Teze Aktarılabilecek Önerilen Metin'),
  H2('11.1. Gereç ve Yöntem – Örneklem büyüklüğü'),
  P('Örneklem büyüklüğü, aynı katılımcılardan elde edilen bağımlı ölçümlerin karşılaştırılmasına uygun olarak G*Power 3.1 programında eşleştirilmiş örneklem t testi (“Means: Difference between two dependent means”) esas alınarak hesaplanmıştır. Literatüre dayalı etki büyüklüğü 1.2, anlamlılık düzeyi üç ikili karşılaştırma için Bonferroni düzeltmesiyle α = 0.0167 (iki yönlü) ve güç (1 − β) = 0.80 alındığında gerekli örneklem büyüklüğü 11 olarak bulunmuştur. Parametrik olmayan testlerin göreli etkinliği (ARE = 0.864) dikkate alınarak bu sayı 12’ye, olası veri kayıpları için %20 ek pay ile 15 katılımcıya yükseltilmiştir. Ara değerlendirmede 15 katılımcıdan elde edilen verilerle üç yöntemin genel karşılaştırması için gözlenen istatistiksel gücün %99’un üzerinde olduğu doğrulanmıştır.'),
  H2('11.2. Gereç ve Yöntem – İstatistiksel analiz'),
  P('Verilerin normal dağılıma uygunluğu Shapiro–Wilk testi ile değerlendirilmiştir. Normal dağılım ve küresellik varsayımları sağlanmadığından, üç yöntemin genel karşılaştırmasında Friedman testi kullanılmış ve etki büyüklüğü Kendall’s W ile hesaplanmıştır. İkili karşılaştırmalar Wilcoxon işaretli sıralar testi ile yapılmış ve Bonferroni düzeltmesi uygulanmıştır. Primescan ve füzyon modellerin eşdeğerliği, önceden belirlenen ±0.10 mm eşdeğerlik sınırıyla iki tek yönlü test (TOST) yöntemi ve Bland–Altman analizi ile değerlendirilmiştir. Post-hoc güç, merkezi olmayan χ² yaklaşımı ve Monte Carlo simülasyonu ile hesaplanmıştır. İstatistiksel anlamlılık düzeyi p < 0.05 olarak kabul edilmiştir.'),
  H2('11.3. Bulgular'),
  P('KIBT – referans tarama, Primescan – referans tarama ve füzyon model – referans veri karşılaştırmalarında ortalama sapma değerleri sırasıyla 2.068 ± 0.368 mm, 0.388 ± 0.129 mm ve 0.386 ± 0.127 mm olarak bulunmuştur (Tablo 1). Friedman testi, üç yöntem arasında istatistiksel olarak anlamlı fark olduğunu göstermiştir (χ²(2) = 22.53; p < 0.001; Kendall’s W = 0.751). Bonferroni düzeltmeli ikili karşılaştırmalarda KIBT modellerinin sapması hem Primescan (ortalama fark 1.680 mm; p < 0.001; r = 0.88) hem de füzyon modellere (ortalama fark 1.682 mm; p < 0.001; r = 0.88) göre anlamlı olarak yüksek bulunmuştur. Primescan ile füzyon modeller arasında anlamlı fark saptanmamıştır (ortalama fark 0.002 mm; %95 GA −0.047 ile 0.052 mm; p = 0.978). Farkın %90 güven aralığı (−0.038 ile 0.043 mm) ±0.10 mm eşdeğerlik sınırları içinde kaldığından iki yöntem eşdeğer bulunmuştur (TOST p < 0.001). Bland–Altman analizinde ortalama fark 0.002 mm, %95 uyum sınırları −0.172 ile 0.176 mm olarak hesaplanmıştır. Mevcut 15 katılımcı ile üç yöntemin genel karşılaştırması için gözlenen istatistiksel güç %99’un üzerindedir.'),

  // ---------------- Kaynaklar ----------------
  H1('Kaynaklar'),
  P('Bland JM, Altman DG. Statistical methods for assessing agreement between two methods of clinical measurement. ~Lancet~. 1986;1(8476):307–310.'),
  P('Cohen J. ~Statistical Power Analysis for the Behavioral Sciences~. 2nd ed. Hillsdale, NJ: Lawrence Erlbaum; 1988.'),
  P('Faul F, Erdfelder E, Lang AG, Buchner A. G*Power 3: A flexible statistical power analysis program for the social, behavioral, and biomedical sciences. ~Behav Res Methods~. 2007;39(2):175–191.'),
  P('Hoenig JM, Heisey DM. The abuse of power: the pervasive fallacy of power calculations for data analysis. ~Am Stat~. 2001;55(1):19–24.'),
  P('Lakens D. Equivalence tests: a practical primer for t tests, correlations, and meta-analyses. ~Soc Psychol Personal Sci~. 2017;8(4):355–362.'),
  P('Lehmann EL. ~Nonparametrics: Statistical Methods Based on Ranks~. San Francisco: Holden-Day; 1975.'),
  P('Tomczak M, Tomczak E. The need to report effect size estimates revisited. ~Trends Sport Sci~. 2014;1(21):19–25.'),

  // ---------------- Ekler ----------------
  new Paragraph({ children: [new PageBreak()] }),
  H1('Ek-1. Anonimleştirilmiş Ham Veri'),
  CAP('**Tablo E1.** Katılımcılara göre yüzey sapması değerleri (mm) ve katılımcı içi sıralar'),
  table(['Kod', 'KIBT – Referans', 'Primescan – Referans', 'Füzyon – Referans', 'Sıralar (K / P / F)'],
    raw.map((r, i) => {
      const order = [...r].map((v, j) => [v, j]).sort((a, b) => a[0] - b[0]);
      const rk = [0, 0, 0]; order.forEach((o, idx) => { rk[o[1]] = idx + 1; });
      return ['K' + (i + 1), r[0].toFixed(4), r[1].toFixed(4), r[2].toFixed(4), rk.join(' / ')];
    }).concat([['Sıra toplamı', '', '', '', '45 / 23 / 22']]), [12, 20, 22, 22, 24]),
  NOTE('K: KIBT, P: Primescan, F: Füzyon (1 = en düşük sapma). Hasta ad-soyad bilgileri gizlilik nedeniyle rapora alınmamıştır.'),
  H1('Ek-2. Yeniden Üretilebilirlik'),
  P('Tüm hesaplamalar Python 3 (SciPy 1.x, statsmodels) ile yapılmış olup analiz betikleri (analiz.py, guc.py, sekil.py) rapor ile birlikte saklanmaktadır. Simülasyonlarda rastgele sayı tohumu 20261101 olarak sabitlenmiştir. Sonuçlar SPSS 25’te şu menülerle doğrulanabilir: Analyze → Nonparametric Tests → Related Samples (Friedman, ikili karşılaştırmalar ile); Analyze → General Linear Model → Repeated Measures; Analyze → Compare Means → Paired-Samples T Test. G*Power’da a priori hesaplama için: t tests → Means: Difference between two dependent means (matched pairs) → d_z = 1.2, α = 0.0167, power = 0.80, two tails.'),
];

const doc = new Document({
  creator: 'Ara istatistiksel değerlendirme', title: 'Ara İstatistiksel Değerlendirme ve Güç Analizi Raporu',
  styles: {
    default: { document: { run: { font: FONT, size: 23 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 27, bold: true, color: '1F3864', font: FONT }, paragraph: { spacing: { before: 300, after: 140 }, outlineLevel: 0, keepNext: true } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 24, bold: true, color: '2E5496', font: FONT }, paragraph: { spacing: { before: 220, after: 100 }, outlineLevel: 1, keepNext: true } },
    ],
  },
  numbering: { config: [
    { reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 500, hanging: 280 } } } }] },
    { reference: 'num', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 500, hanging: 320 } } } }] },
    { reference: 'num2', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 500, hanging: 320 } } } }] },
  ] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], size: 18 })] })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then(buf => { fs.writeFileSync(process.argv[2], buf); console.log('yazıldı', process.argv[2]); });
