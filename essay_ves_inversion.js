const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, LevelFormat, BorderStyle, WidthType,
  ShadingType, PageNumber, Footer, TabStopType, TabStopPosition,
  TableOfContents, PageBreak, UnderlineType
} = require('docx');
const fs = require('fs');

// ─── helpers ─────────────────────────────────────────────────────────────────
const font = "Times New Roman";
const sz   = 24; // 12 pt
const szSm = 20; // 10 pt

function p(children, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 360, after: 160 },  // 1.5 line, 8pt after
    ...opts,
    children: Array.isArray(children) ? children : [children],
  });
}

function t(text, opts = {}) {
  return new TextRun({ text, font, size: sz, ...opts });
}

function bold(text) { return t(text, { bold: true }); }
function italic(text) { return t(text, { italics: true }); }

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 160 },
    children: [new TextRun({ text, font, size: 28, bold: true })],
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, font, size: 26, bold: true })],
  });
}

function heading3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text, font, size: 24, bold: true, italics: true })],
  });
}

// Equation paragraph (centered, indented formula label)
function eq(formulaText, label) {
  const children = [new TextRun({ text: formulaText, font: "Courier New", size: sz, italics: true })];
  if (label) {
    children.push(new TextRun({ text: `          ${label}`, font, size: sz }));
  }
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 280, before: 120, after: 120 },
    children,
  });
}

// Caption paragraph (center-aligned, smaller)
function caption(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 240, before: 80, after: 160 },
    children: [new TextRun({ text, font, size: szSm, italics: true })],
  });
}

// Reference item
function ref(text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 280, after: 120 },
    indent: { left: 720, hanging: 720 },
    children: [new TextRun({ text, font, size: sz })],
  });
}

// Bullet item
function bullet(text, numbered = false) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 320, after: 80 },
    numbering: { reference: numbered ? "numbers" : "bullets", level: 0 },
    children: [t(text)],
  });
}

// Page break
function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// ─── TABLE: rubrik / notation ─────────────────────────────────────────────────
const border = { style: BorderStyle.SINGLE, size: 1, color: "999999" };
const borders = { top: border, bottom: border, left: border, right: border };

function makeTable(headers, rows, colWidths) {
  const totalWidth = colWidths.reduce((a, b) => a + b, 0);
  const headerRow = new TableRow({
    children: headers.map((h, i) => new TableCell({
      borders,
      width: { size: colWidths[i], type: WidthType.DXA },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      shading: { fill: "D9E1F2", type: ShadingType.CLEAR },
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: h, font, size: sz, bold: true })],
      })],
    })),
  });

  const dataRows = rows.map(row => new TableRow({
    children: row.map((cell, i) => new TableCell({
      borders,
      width: { size: colWidths[i], type: WidthType.DXA },
      margins: { top: 60, bottom: 60, left: 120, right: 120 },
      children: [new Paragraph({
        alignment: AlignmentType.LEFT,
        children: [new TextRun({ text: cell, font, size: sz })],
      })],
    })),
  }));

  return new Table({
    width: { size: totalWidth, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [headerRow, ...dataRows],
  });
}

// ─── DOCUMENT ─────────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
      {
        reference: "numbers",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
    ],
  },
  styles: {
    default: {
      document: { run: { font, size: sz } },
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal",
        run: { font, size: 28, bold: true },
        paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 0 },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal",
        run: { font, size: 26, bold: true },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 },
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal",
        run: { font, size: 24, bold: true, italics: true },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 },
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 },
      },
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: [PageNumber.CURRENT], font, size: sz })],
        })],
      }),
    },
    children: [

      // ─── COVER / TITLE ────────────────────────────────────────────────────
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 480, after: 80 },
        children: [new TextRun({ text: "INVERSI GLOBAL METODE GEOLISTRIK VERTICAL ELECTRICAL SOUNDING (VES) MENGGUNAKAN GRID SEARCH DAN PARTICLE SWARM OPTIMIZATION (PSO)", font, size: 28, bold: true })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 40, after: 40 },
        children: [new TextRun({ text: "Esai Ilmiah", font, size: sz, italics: true })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 40, after: 40 },
        children: [new TextRun({ text: "Mata Kuliah Pemodelan dalam Geofisika", font, size: sz })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 40, after: 40 },
        children: [new TextRun({ text: "Teknik Geofisika - Institut Teknologi Sepuluh Nopember", font, size: sz })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 400 },
        children: [new TextRun({ text: "2025", font, size: sz })],
      }),

      // divider line via paragraph border
      new Paragraph({
        spacing: { before: 0, after: 320 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E4057", space: 1 } },
        children: [],
      }),

      // ─── ABSTRAK ─────────────────────────────────────────────────────────
      heading1("ABSTRAK"),
      p([
        t("Metode Vertical Electrical Sounding (VES) merupakan salah satu teknik geolistrik yang banyak digunakan dalam eksplorasi air tanah, pemetaan stratigrafi, dan investigasi lingkungan. Dalam konteks inversi, VES menghadapi tantangan khas berupa non-linearitas hubungan antara resistivitas semu yang terukur dan distribusi resistivitas bawah permukaan yang sesungguhnya. Esai ini membahas penerapan dua metode inversi global, yaitu "),
        italic("Grid Search"),
        t(" dan "),
        italic("Particle Swarm Optimization"),
        t(" (PSO), untuk menyelesaikan problem inversi VES pada konfigurasi Schlumberger. Pembahasan mencakup persamaan fundamental forward model menggunakan transformasi Hankel, formulasi parameter model dalam ruang multi-dimensi, analisis noise dan ketidakpastian data, serta implementasi komputasional kedua algoritma inversi tersebut. Perbandingan antara Grid Search dan PSO menunjukkan bahwa meskipun Grid Search memberikan jaminan eksplorasi ruang parameter yang sistematis, kompleksitas komputasinya meningkat secara eksponensial seiring bertambahnya jumlah lapisan. PSO, sebagai metode meta-heuristik berbasis swarm intelligence, menawarkan eksplorasi yang jauh lebih efisien dan skalabel untuk problem dengan dimensi parameter yang tinggi."),
      ]),
      new Paragraph({
        spacing: { after: 80, before: 80 },
        children: [bold("Kata kunci: "), t("Geolistrik, VES, Schlumberger, Inversi, Grid Search, Particle Swarm Optimization, Resistivitas")],
      }),

      pageBreak(),

      // ─── 1. PENDAHULUAN ───────────────────────────────────────────────────
      heading1("1. PENDAHULUAN"),
      p([
        t("Eksplorasi geofisika pada dasarnya adalah upaya untuk memahami kondisi bawah permukaan bumi dari pengukuran-pengukuran yang dilakukan di permukaan. Di antara berbagai metode yang tersedia, metode resistivitas geolistrik, khususnya Vertical Electrical Sounding (VES), telah menjadi salah satu alat yang paling andal dan banyak digunakan selama beberapa dekade terakhir. Metode ini memanfaatkan prinsip sederhana namun powerful: batuan dan material bawah permukaan memiliki kemampuan yang berbeda-beda dalam menghantarkan arus listrik, dan perbedaan ini dapat diukur dari permukaan."),
      ]),
      p([
        t("Namun, interpretasi data VES bukanlah pekerjaan trivial. Hubungan antara resistivitas semu yang terukur di permukaan ("),
        italic("apparent resistivity"),
        t(") dan resistivitas sesungguhnya dari tiap lapisan batuan di bawah permukaan bersifat non-linear dan tidak unik. Permasalahan ini masuk dalam kategori "),
        italic("ill-posed inverse problem"),
        t(": satu set data observasi dapat dijelaskan oleh lebih dari satu model bumi yang berbeda. Di sinilah peran metode inversi menjadi sangat krusial."),
      ]),
      p([
        t("Secara historis, inversi VES dilakukan secara manual menggunakan kurva-kurva standar ("),
        italic("type curves"),
        t("). Pendekatan ini tidak hanya memakan waktu tetapi juga sangat bergantung pada keahlian interpreter. Perkembangan komputasi modern membuka jalan bagi pendekatan inversi otomatis, baik berbasis gradien ("),
        italic("local optimization"),
        t(") maupun berbasis populasi ("),
        italic("global optimization"),
        t("). Metode berbasis gradien seperti Gauss-Newton atau Occam's Inversion sangat efisien secara komputasi, tetapi rentan terjebak pada minimum lokal, terutama untuk data dengan tingkat noise yang tinggi."),
      ]),
      p([
        t("Metode inversi global, di sisi lain, dirancang untuk mengeksplorasi ruang parameter secara lebih menyeluruh dan tidak bergantung pada solusi awal yang baik. Dua metode yang akan dibahas dalam esai ini adalah Grid Search dan Particle Swarm Optimization (PSO). Meskipun keduanya bersifat global, filosofi dan mekanisme kerjanya sangat berbeda. Grid Search menggunakan pendekatan deterministik dengan melakukan evaluasi fungsi misfit pada semua titik dalam kisi parameter yang telah ditentukan, sementara PSO menggunakan mekanisme evolusi berbasis koloni yang lebih adaptif dan stokastik."),
      ]),
      p([
        t("Esai ini bertujuan untuk memberikan pemaparan mendalam tentang bagaimana kedua metode inversi tersebut diterapkan pada problem VES, mulai dari persamaan fisika yang mendasarinya, formulasi matematis, hingga implementasi komputasional dan interpretasi output yang dihasilkan. Metode yang digunakan sebagai studi kasus adalah konfigurasi Schlumberger untuk pengukuran VES 1D berlapis."),
      ]),

      // ─── 2. PERSAMAAN FUNDAMENTAL FORWARD MODEL ───────────────────────────
      heading1("2. PERSAMAAN FUNDAMENTAL FORWARD MODEL"),
      p([
        t("Secara umum, hubungan antara model bumi dan data observasi dalam geofisika dapat ditulis dalam bentuk operator forward:"),
      ]),
      eq("d = G(m)", "(1)"),
      p([
        t("di mana "),
        bold("d"),
        t(" adalah vektor data observasi (resistivitas semu) berukuran N, "),
        bold("G"),
        t(" adalah operator forward modeling yang umumnya bersifat non-linear, dan "),
        bold("m"),
        t(" adalah vektor parameter model berukuran M."),
      ]),

      heading2("2.1 Hukum Fisika yang Mendasari"),
      p([
        t("Metode geolistrik didasarkan pada Hukum Ohm dalam bentuk diferensialnya. Ketika arus listrik disuntikkan ke dalam medium bumi yang heterogen dan isotropik, distribusi potensial listrik V di dalam medium tersebut mengikuti Persamaan Poisson:"),
      ]),
      eq("∇ · (σ(r) ∇V(r)) = -I δ(r - rs)", "(2)"),
      p([
        t("di mana σ(r) adalah konduktivitas listrik [S/m] sebagai fungsi posisi r, V(r) adalah potensial listrik [V], I adalah besar arus yang diinjeksikan [A], dan δ(r - rs) adalah fungsi delta Dirac yang merepresentasikan sumber arus di posisi rs."),
      ]),
      p([
        t("Untuk model bumi berlapis horizontal 1D ("),
        italic("horizontally layered earth"),
        t("), di mana setiap lapisan memiliki resistivitas ρ"),
        t("i"),
        t(" dan ketebalan h"),
        t("i"),
        t(" yang konstan dalam arah horizontal, Persamaan (2) dapat diselesaikan secara analitik menggunakan transformasi Hankel. Resistivitas semu pada jarak elektroda AB/2 = r dapat diekspresikan sebagai:"),
      ]),
      eq("ρa(r) = 2π r² ∫₀^∞ T(λ) J₁(λr) λ dλ", "(3)"),
      p([
        t("di mana J₁ adalah fungsi Bessel orde pertama, λ adalah bilangan gelombang (variabel transformasi), dan T(λ) adalah fungsi kernel yang merangkum seluruh informasi tentang struktur lapisan bawah permukaan. Fungsi kernel T(λ) dihitung secara rekursif dari lapisan terbawah ke permukaan menggunakan algoritma Pekeris:"),
      ]),
      eq("T_n(λ) = ρ_n (T_{n+1}(λ) + ρ_n tanh(λh_n)) / (ρ_n + T_{n+1}(λ) tanh(λh_n))", "(4)"),
      p([
        t("di mana T_{N+1}(λ) = ρ_N adalah resistivitas lapisan terbawah (setengah ruang tak terhingga atau "),
        italic("halfspace"),
        t("). Ini adalah persamaan rekursif yang dievaluasi dari n = N ke n = 1 untuk mendapatkan T_1(λ) yang kemudian dimasukkan ke Persamaan (3)."),
      ]),

      heading2("2.2 Asumsi dan Penyederhanaan"),
      p([
        t("Forward model VES 1D di atas dibangun di atas beberapa asumsi fundamental:"),
      ]),
      bullet("Bumi diasumsikan terdiri dari lapisan-lapisan horizontal yang seragam (horizontally stratified), sehingga resistivitas hanya merupakan fungsi dari kedalaman: ρ = ρ(z)."),
      bullet("Medium bumi diasumsikan isotropik: konduktivitas listrik tidak bergantung pada arah."),
      bullet("Sumber arus diasumsikan sebagai sumber titik (point source) di permukaan bumi."),
      bullet("Permukaan bumi diasumsikan datar (no topography correction)."),
      bullet("Tidak ada efek elektromagnetik (frekuensi pengukuran sangat rendah, kondisi quasi-statik)."),
      new Paragraph({ spacing: { after: 80 }, children: [] }),
      p([
        t("Dalam praktik di lapangan, asumsi-asumsi ini tentu saja tidak selalu terpenuhi secara sempurna. Namun, untuk banyak aplikasi survei geolistrik di lapangan yang relatif datar dengan perubahan lateral yang tidak terlalu drastis, model 1D berlapis memberikan pendekatan yang sangat memadai."),
      ]),

      heading2("2.3 Hubungan Parameter Model dan Data Observasi"),
      p([
        t("Persamaan (3) dan (4) secara eksplisit menunjukkan bahwa data yang terukur di lapangan, yaitu resistivitas semu ρ"),
        italic("a"),
        t("(r), merupakan fungsi non-linear dari parameter-parameter lapisan (ρ"),
        t("i"),
        t(", h"),
        t("i"),
        t("). Informasi tentang lapisan-lapisan yang lebih dalam tercermin dalam pengukuran pada spasi elektroda yang lebih besar. Semakin besar spasi AB/2, semakin dalam penetrasi arus listrik ke dalam bumi, dan semakin besar pengaruh lapisan-lapisan yang lebih dalam terhadap nilai ρ"),
        italic("a"),
        t(" yang terukur. Inilah prinsip dasar dari metode sounding (pendugaan kedalaman) dalam geolistrik."),
      ]),

      // ─── 3. PARAMETER MODEL ───────────────────────────────────────────────
      heading1("3. PARAMETER MODEL UNTUK INVERSI"),
      p([
        t("Target inversi dalam VES 1D berlapis adalah parameter-parameter yang mendefinisikan distribusi resistivitas bawah permukaan. Untuk model dengan N lapisan (termasuk lapisan setengah ruang terbawah), vektor model secara eksplisit didefinisikan sebagai:"),
      ]),
      eq("m = [ρ₁, h₁, ρ₂, h₂, ..., ρ_{N-1}, h_{N-1}, ρ_N]ᵀ", "(5)"),
      p([
        t("di mana ρ"),
        t("i"),
        t(" adalah resistivitas lapisan ke-i [Ω·m] dan h"),
        t("i"),
        t(" adalah ketebalan lapisan ke-i [m]. Lapisan terbawah (N) merupakan halfspace yang tidak memiliki ketebalan terbatas, sehingga vektor model memiliki dimensi M = 2N - 1."),
      ]),
      p([
        t("Sebagai contoh konkret, untuk model empat lapisan (N = 4), dimensi ruang model adalah M = 2(4) - 1 = 7, dengan vektor:"),
      ]),
      eq("m = [ρ₁, h₁, ρ₂, h₂, ρ₃, h₃, ρ₄]ᵀ ∈ ℝ⁷", "(6)"),

      heading2("3.1 Batasan Fisik Parameter"),
      p([
        t("Dalam proses inversi, setiap parameter harus dibatasi pada rentang fisik yang masuk akal. Tabel 1 merangkum batasan yang umum digunakan dalam inversi VES untuk studi air tanah dan investigasi geoteknik."),
      ]),
      new Paragraph({ spacing: { after: 160 }, children: [] }),
      makeTable(
        ["Parameter", "Simbol", "Rentang Minimum", "Rentang Maksimum", "Satuan"],
        [
          ["Resistivitas lapisan", "ρᵢ", "1", "10.000", "Ω·m"],
          ["Ketebalan lapisan", "hᵢ", "0,5", "200", "m"],
          ["Kedalaman total", "Σhᵢ", "—", "500", "m"],
        ],
        [2500, 1500, 1500, 1800, 1200]
      ),
      caption("Tabel 1. Batasan fisik parameter model inversi VES untuk aplikasi hidrologi dan geoteknik."),

      p([
        t("Batasan ini bersifat fleksibel dan dapat disesuaikan berdasarkan pengetahuan geologi daerah setempat ("),
        italic("a priori"),
        t(" information). Dalam praktik, informasi dari sumur bor yang ada (borehole data) sangat berguna untuk mempersempit rentang parameter, sehingga meningkatkan keunikan solusi inversi."),
      ]),

      heading2("3.2 Transformasi Parameter untuk Inversi"),
      p([
        t("Karena rentang resistivitas dapat mencakup beberapa orde besaran (dari 1 hingga ribuan Ω·m), sangat umum untuk melakukan inversi dalam ruang logaritmik. Dengan mendefinisikan parameter transformasi:"),
      ]),
      eq("m'ᵢ = log₁₀(mᵢ)", "(7)"),
      p([
        t("transformasi ini membuat distribusi parameter lebih seragam dan meningkatkan stabilitas numerik proses inversi. Dalam esai ini, semua operasi inversi diasumsikan bekerja dalam ruang logaritmik untuk parameter resistivitas, sementara ketebalan tetap dalam ruang linear."),
      ]),

      // ─── 4. DATA LAPANGAN ─────────────────────────────────────────────────
      heading1("4. DATA YANG DIUKUR DI LAPANGAN"),

      heading2("4.1 Konfigurasi Pengukuran Schlumberger"),
      p([
        t("Konfigurasi Schlumberger adalah konfigurasi elektroda yang paling umum digunakan dalam pengukuran VES. Pada konfigurasi ini, dua elektroda arus (A dan B) dan dua elektroda potensial (M dan N) ditempatkan sepanjang garis lurus dengan elektroda potensial di tengah. Konfigurasi ini ditunjukkan pada Gambar 1 (skema)."),
      ]),
      p([
        t("Syarat geometri Schlumberger adalah AB/2 >> MN/2. Elektroda arus A dan B diperlebar secara simetris terhadap titik pusat pengukuran, sementara elektroda potensial M dan N dijaga tetap atau diperlebar secara bertahap ketika pembacaan tegangan menjadi terlalu kecil. Faktor geometri konfigurasi Schlumberger adalah:"),
      ]),
      eq("K = π [(AB/2)² - (MN/2)²] / MN", "(8)"),
      p([
        t("dan resistivitas semu dihitung dari pengukuran tegangan (ΔV) dan arus (I) sebagai:"),
      ]),
      eq("ρa = K · ΔV / I", "(9)"),

      heading2("4.2 Besaran yang Diukur dan Vektor Data Observasi"),
      p([
        t("Besaran fisis yang diukur secara langsung di lapangan adalah beda potensial ΔV [mV] antara elektroda M dan N, serta arus I [mA] yang diinjeksikan melalui elektroda A-B. Dari kedua besaran ini, resistivitas semu ρ"),
        italic("a"),
        t(" [Ω·m] dihitung menggunakan Persamaan (8) dan (9) untuk setiap spasi elektroda r = AB/2 yang digunakan."),
      ]),
      p([
        t("Dalam satu titik VES, pengukuran dilakukan pada rentang spasi elektroda dari yang terkecil (beberapa desimeter) hingga yang terbesar (ratusan meter), tergantung pada kedalaman investigasi yang diinginkan. Vektor data observasi ditulis sebagai:"),
      ]),
      eq("d_obs = [ρa(r₁), ρa(r₂), ..., ρa(rN)]ᵀ ∈ ℝᴺ", "(10)"),
      p([
        t("di mana r₁ < r₂ < ... < r_N adalah spasi elektroda AB/2 yang digunakan dalam pengukuran. Jumlah titik data N umumnya berkisar antara 20 hingga 50 titik per sounding. Semua nilai ρ"),
        italic("a"),
        t(" dalam vektor ini memiliki satuan Ohm-meter [Ω·m]."),
      ]),

      new Paragraph({ spacing: { after: 160 }, children: [] }),
      makeTable(
        ["Spasi AB/2 (m)", "Nilai ρa Tipikal (Ω·m)", "Lapisan yang Terpenetrasi"],
        [
          ["1 – 5", "Bervariasi tinggi", "Lapisan permukaan / top soil"],
          ["10 – 30", "Tergantung litologi", "Lapisan 1-2"],
          ["50 – 100", "Tergantung akuifer", "Lapisan 2-3"],
          ["100 – 500", "Menurun / naik", "Lapisan dalam / halfspace"],
        ],
        [2500, 2500, 3500]
      ),
      caption("Tabel 2. Hubungan antara spasi elektroda, rentang ρa tipikal, dan lapisan yang terpenetrasi dalam pengukuran VES."),

      // ─── 5. FORMULASI MATRIKS ─────────────────────────────────────────────
      heading1("5. FORMULASI MATRIKS BERDASARKAN DATA LAPANGAN"),
      p([
        t("Problem VES adalah problem non-linear. Operator forward G tidak dapat ditulis dalam bentuk matriks linear A yang sederhana. Namun, dalam konteks linearisasi lokal yang digunakan dalam beberapa metode inversi iteratif, Jacobian (atau matriks sensitivitas) memainkan peran yang sangat penting."),
      ]),

      heading2("5.1 Representasi Non-Linear"),
      p([
        t("Secara umum, hubungan forward ditulis sebagai:"),
      ]),
      eq("d = G(m) = [ρa(r₁|m), ρa(r₂|m), ..., ρa(rN|m)]ᵀ", "(11)"),
      p([
        t("di mana setiap elemen d"),
        t("i"),
        t(" = ρ"),
        italic("a"),
        t("(r"),
        t("i"),
        t("|m) dievaluasi menggunakan Persamaan (3) dan (4) untuk vektor model m yang diberikan. Karena hubungan ini melibatkan integrasi numerik fungsi Bessel, evaluasi G(m) memerlukan algoritma komputasi khusus."),
      ]),

      heading2("5.2 Matriks Jacobian (Sensitivitas)"),
      p([
        t("Walaupun G non-linear, pada integrasi dalam metode berbasis gradient kita perlu matriks Jacobian J, didefinisikan sebagai:"),
      ]),
      eq("Jᵢⱼ = ∂ρa(rᵢ) / ∂mⱼ", "(12)"),
      p([
        t("Matriks J berukuran N × M, di mana elemen J"),
        t("ij"),
        t(" menyatakan sensitivitas data ke-i terhadap parameter ke-j. Secara fisik, J"),
        t("ij"),
        t(" yang besar menunjukkan bahwa pengukuran pada spasi r"),
        t("i"),
        t(" sangat sensitif terhadap perubahan parameter m"),
        t("j"),
        t(" (misalnya resistivitas lapisan ke-j). Untuk inversi global seperti Grid Search dan PSO yang akan dibahas, matriks Jacobian tidak diperlukan secara eksplisit karena tidak ada proses linearisasi — forward model G(m) dievaluasi langsung."),
      ]),

      heading2("5.3 Dimensi Masalah"),
      p([
        t("Untuk kasus tipikal dengan N = 30 titik data dan M = 7 parameter (model 4 lapisan), dimensi masalah adalah:"),
      ]),
      eq("d ∈ ℝ³⁰,  m ∈ ℝ⁷,  J ∈ ℝ³⁰ˣ⁷", "(13)"),
      p([
        t("Kondisi N > M merupakan kondisi "),
        italic("overdetermined"),
        t(", yang berarti data lebih banyak dari parameter yang dicari. Meskipun demikian, karena non-linearitas dan adanya noise, solusi tidak bisa diperoleh secara langsung dan harus dicari melalui minimisasi fungsi misfit."),
      ]),

      // ─── 6. ANALISIS NOISE ────────────────────────────────────────────────
      heading1("6. ANALISIS NOISE DAN KETIDAKPASTIAN DATA"),
      p([
        t("Dalam praktik lapangan, data yang terukur selalu mengandung noise dari berbagai sumber. Model data yang realistis dapat dituliskan sebagai:"),
      ]),
      eq("d_obs = G(m_true) + ε", "(14)"),
      p([
        t("di mana m_true adalah model bumi yang sesungguhnya (yang hendak kita cari) dan ε adalah vektor noise. Pemahaman terhadap karakteristik noise sangat penting karena akan mempengaruhi pemilihan fungsi misfit dan interpretasi solusi inversi."),
      ]),

      heading2("6.1 Sumber-Sumber Noise"),
      heading3("6.1.1 Noise Instrumen"),
      p([
        t("Resistivitimeter modern umumnya memiliki resolusi pengukuran tegangan pada orde mikrovolt. Namun, pada spasi elektroda yang sangat besar, tegangan yang terukur menjadi sangat kecil sehingga rasio sinyal terhadap noise (SNR) menurun drastis. Noise instrumen mencakup: resolusi ADC (Analog-to-Digital Converter), drift termal pada rangkaian elektronik, dan kesalahan kalibrasi elektroda. Noise ini biasanya bersifat acak dengan distribusi mendekati Gaussian."),
      ]),
      heading3("6.1.2 Noise Lingkungan"),
      p([
        t("Sumber gangguan eksternal yang paling signifikan dalam pengukuran VES meliputi: arus stray dari jaringan listrik (50 Hz atau 60 Hz), efek polarisasi terinduksi (IP) pada elektroda yang tidak sempurna, variasi alami potensial spontan (SP) dalam tanah, dan aktivitas petir atau badai elektromagnetik. Penggunaan peralatan dengan kemampuan stack (penumpukan pembacaan) dapat mereduksi noise acak secara signifikan."),
      ]),
      heading3("6.1.3 Noise Geologis"),
      p([
        t("Noise geologis atau "),
        italic("geological noise"),
        t(" timbul dari struktur geologi yang tidak menjadi target inversi tetapi mempengaruhi pengukuran. Dalam konteks VES, asumsi 1D adalah simplifikasi yang sering tidak sepenuhnya valid. Perubahan lateral resistivitas (misalnya akibat adanya sesar, lensa pasir, atau variasi muka air tanah) akan menyebabkan kurva ρ"),
        italic("a"),
        t(" yang terukur menyimpang dari prediksi model 1D. Ini merupakan bentuk "),
        italic("model error"),
        t(" yang tidak bisa direduksi dengan pengulangan pengukuran."),
      ]),
      heading3("6.1.4 Ketidakpastian Geometri"),
      p([
        t("Kesalahan dalam penempatan elektroda di lapangan (positioning error) secara langsung mempengaruhi perhitungan faktor geometri K pada Persamaan (8). Kesalahan posisi sebesar 1-2% dapat mengakibatkan kesalahan ρ"),
        italic("a"),
        t(" yang signifikan. Topografi yang tidak datar juga harus dikoreksi sebelum inversi."),
      ]),

      heading2("6.2 Model Statistik Noise"),
      p([
        t("Asumsi yang paling umum digunakan adalah bahwa noise berdistribusi Gaussian dengan rata-rata nol:"),
      ]),
      eq("ε ~ N(0, C_d)", "(15)"),
      p([
        t("di mana C_d adalah matriks kovariansi data berukuran N × N. Untuk kasus sederhana di mana kesalahan antar titik data tidak berkorelasi ("),
        italic("uncorrelated errors"),
        t("), C_d adalah matriks diagonal:"),
      ]),
      eq("C_d = diag(σ₁², σ₂², ..., σN²)", "(16)"),
      p([
        t("di mana σ"),
        t("i"),
        t(" adalah standar deviasi pengukuran pada titik ke-i. Dalam banyak aplikasi praktis, σ"),
        t("i"),
        t(" diasumsikan sebagai persentase dari nilai data itu sendiri (misalnya 5%), memberikan:"),
      ]),
      eq("σᵢ = ε_rel · |d_obs,i|", "(17)"),
      p([
        t("dengan ε_rel adalah level noise relatif (biasanya 0,02–0,10 atau 2–10%). Pendekatan ini disebut "),
        italic("proportional noise model"),
        t(" dan sangat sesuai untuk data VES karena rentang nilai ρ"),
        italic("a"),
        t(" yang dapat mencakup beberapa orde besaran."),
      ]),

      // ─── 7. IMPLEMENTASI ──────────────────────────────────────────────────
      heading1("7. IMPLEMENTASI GRID SEARCH DAN PSO"),
      p([
        t("Fungsi misfit atau "),
        italic("objective function"),
        t(" yang digunakan dalam kedua metode inversi ini adalah L2-norm (chi-square misfit) yang dinormalisasi dengan ketidakpastian data:"),
      ]),
      eq("E(m) = Σᵢ₌₁ᴺ [(d_obs,i - d_calc,i(m)) / σᵢ]²", "(18)"),
      p([
        t("di mana d_calc,i(m) = ρ"),
        italic("a"),
        t("(r_i|m) adalah respon forward dari model m untuk titik ke-i. Nilai E(m) mendekati N untuk solusi yang baik (sesuai dengan kriteria misfit yang ternormalisasi). Dalam praktik, sering digunakan RMS error:"),
      ]),
      eq("RMS = √[E(m) / N]", "(19)"),
      p([
        t("Nilai RMS mendekati 1 menunjukkan bahwa residual data berada dalam batas ketidakpastian pengukuran."),
      ]),

      heading2("7.1 Grid Search"),

      heading3("7.1.1 Definisi Search Space"),
      p([
        t("Langkah pertama Grid Search adalah mendefinisikan ruang pencarian untuk setiap parameter. Untuk parameter ke-j, rentang yang dieksplorasi adalah:"),
      ]),
      eq("mⱼ ∈ {m_min,j, m_min,j + Δmⱼ, m_min,j + 2Δmⱼ, ..., m_max,j}", "(20)"),
      p([
        t("di mana Δm_j = (m_max,j - m_min,j) / (K_j - 1) adalah langkah diskritisasi dan K_j adalah jumlah titik grid untuk parameter j. Grid total dibentuk sebagai produk Kartesian dari semua parameter:"),
      ]),
      eq("Grid = {m_min,1, ..., m_max,1} × {m_min,2, ..., m_max,2} × ... × {m_min,M, ..., m_max,M}", "(21)"),
      p([
        t("Jumlah total kombinasi yang harus dievaluasi adalah:"),
      ]),
      eq("N_total = Π_{j=1}^{M} Kⱼ", "(22)"),
      p([
        t("Ini adalah inti dari "),
        italic("curse of dimensionality"),
        t(": untuk model 4 lapisan dengan M = 7 parameter, jika masing-masing parameter didiskritisasi menjadi K = 20 titik, total evaluasi forward yang diperlukan adalah 20⁷ = 1,28 × 10⁹. Dengan asumsi satu evaluasi forward membutuhkan 1 ms, waktu komputasi yang dibutuhkan adalah sekitar 15 hari! Ini jelas tidak praktis."),
      ]),

      heading3("7.1.2 Prosedur Grid Search"),
      p([t("Algoritma Grid Search dapat diuraikan sebagai berikut:")]),
      bullet("Inisialisasi: Tentukan rentang [m_min,j, m_max,j] dan jumlah titik K_j untuk setiap parameter j = 1, ..., M.", true),
      bullet("Pembentukan Grid: Buat semua kombinasi parameter menggunakan produk Kartesian (Persamaan 21).", true),
      bullet("Evaluasi Forward: Untuk setiap titik grid m_k (k = 1, ..., N_total), hitung d_calc = G(m_k) menggunakan Persamaan (3) dan (4).", true),
      bullet("Hitung Misfit: Evaluasi E(m_k) menggunakan Persamaan (18) untuk setiap titik grid.", true),
      bullet("Pilih Model Terbaik: m_best = argmin_k E(m_k).", true),
      new Paragraph({ spacing: { after: 80 }, children: [] }),
      p([
        t("Seluruh informasi tentang topografi misfit surface tersimpan dalam evaluasi ini, sehingga Grid Search memungkinkan analisis yang sangat kaya tentang non-uniqueness dan trade-off antar parameter."),
      ]),

      heading3("7.1.3 Visualisasi Misfit Surface"),
      p([
        t("Salah satu keunggulan utama Grid Search adalah kemampuannya menghasilkan "),
        italic("misfit surface"),
        t(" atau "),
        italic("misfit landscape"),
        t(" secara lengkap. Untuk model dua parameter, misfit surface dapat divisualisasikan sebagai kontur plot E(m_1, m_2) dalam ruang parameter. Untuk model multi-parameter, visualisasi dilakukan melalui proyeksi marginal: misfit diminimisasi terhadap semua parameter kecuali dua yang ditampilkan. Misfit surface ini sangat informatif untuk:"),
      ]),
      bullet("Mengidentifikasi keunikan solusi: minimum tunggal atau multiple minima."),
      bullet("Menganalisis trade-off antar parameter: parameter mana yang coupled atau decoupled."),
      bullet("Menentukan resolusi inversi: seberapa tajam minimum di ruang parameter."),
      new Paragraph({ spacing: { after: 80 }, children: [] }),

      heading3("7.1.4 Keterbatasan Grid Search"),
      p([
        t("Meskipun Grid Search menjamin bahwa solusi global tidak terlewatkan (asalkan grid cukup halus dan rentang parameter tepat), metode ini memiliki keterbatasan fundamental:"),
      ]),
      bullet("Curse of Dimensionality: Kompleksitas komputasi meningkat secara eksponensial dengan M. Untuk M > 5, metode ini menjadi sangat mahal secara komputasi."),
      bullet("Trade-off Resolusi vs. Kecepatan: Grid yang kasar mungkin melewatkan solusi optimal; grid yang halus membutuhkan waktu komputasi yang sangat lama."),
      bullet("Tidak Adaptif: Evaluasi dilakukan secara merata di seluruh ruang parameter, termasuk di wilayah yang jelas tidak menjanjikan (misfit sangat tinggi)."),
      new Paragraph({ spacing: { after: 80 }, children: [] }),

      // ─── PSO ──────────────────────────────────────────────────────────────
      heading2("7.2 Particle Swarm Optimization (PSO)"),
      p([
        t("PSO dikembangkan oleh Kennedy dan Eberhart (1995) terinspirasi dari perilaku kolektif kawanan burung ("),
        italic("flocking"),
        t(") dan gerombolan ikan ("),
        italic("schooling"),
        t("). Setiap individu dalam kawanan disebut "),
        italic("partikel"),
        t(", dan setiap partikel merepresentasikan satu solusi calon (satu set parameter model) dalam ruang parameter."),
      ]),

      heading3("7.2.1 Inisialisasi Swarm"),
      p([
        t("Swarm terdiri dari K partikel. Setiap partikel k (k = 1, ..., K) diinisialisasi dengan posisi acak dalam ruang parameter:"),
      ]),
      eq("m_k^(0) ~ Uniform(m_min, m_max)", "(23)"),
      p([
        t("dan kecepatan awal yang kecil atau nol:"),
      ]),
      eq("v_k^(0) ~ Uniform(-|m_max - m_min|/2, |m_max - m_min|/2)", "(24)"),
      p([
        t("Matriks posisi swarm pada iterasi t dapat dituliskan sebagai:"),
      ]),
      eq("M_swarm^(t) = [m_1^(t), m_2^(t), ..., m_K^(t)]ᵀ ∈ ℝ^{K×M}", "(25)"),
      p([
        t("Jumlah partikel K biasanya dipilih antara 20 hingga 100, bergantung pada kompleksitas ruang parameter. Dalam praktik, K = 30–50 partikel sudah cukup untuk problem VES dengan M ≤ 10."),
      ]),

      heading3("7.2.2 Persamaan Update PSO"),
      p([
        t("Pada setiap iterasi t, kecepatan dan posisi setiap partikel diperbarui menggunakan:"),
      ]),
      eq("v_k^(t+1) = w·v_k^(t) + c₁r₁(p_best,k - m_k^(t)) + c₂r₂(g_best - m_k^(t))", "(26)"),
      eq("m_k^(t+1) = m_k^(t) + v_k^(t+1)", "(27)"),
      p([
        t("Makna fisik dari setiap komponen dalam Persamaan (26) adalah:"),
      ]),
      bullet("w (inertia weight): Mengontrol seberapa besar kecepatan sebelumnya dipertahankan. Nilai w besar (mendekati 1) mendorong eksplorasi global; nilai w kecil mendorong eksploitasi lokal. Umumnya w menurun secara linear dari 0,9 menjadi 0,4 sepanjang iterasi."),
      bullet("c₁ (cognitive coefficient): Mengontrol kecenderungan partikel untuk kembali ke posisi terbaiknya sendiri (p_best,k). Merepresentasikan 'memori personal' partikel."),
      bullet("c₂ (social coefficient): Mengontrol kecenderungan partikel untuk bergerak menuju posisi terbaik global (g_best). Merepresentasikan 'pengetahuan kolektif' kawanan."),
      bullet("r₁, r₂ (bilangan acak): Bilangan acak terdistribusi uniform dalam [0,1], dihasilkan baru pada setiap iterasi untuk setiap partikel. Memberikan sifat stokastik pada PSO."),
      bullet("p_best,k (personal best): Posisi terbaik yang pernah dikunjungi partikel k sepanjang riwayat pencarian, berdasarkan nilai misfit terkecil."),
      bullet("g_best (global best): Posisi terbaik di antara seluruh partikel dalam swarm, yaitu argmin_k [min_t E(m_k^(t))]."),
      new Paragraph({ spacing: { after: 80 }, children: [] }),
      p([
        t("Nilai khas yang digunakan dalam literatur adalah c₁ = c₂ = 2,0 (Kennedy & Eberhart, 1995). Nilai ini memastikan bahwa rata-rata pergerakan partikel mengarah ke titik tengah antara p_best dan g_best, dengan jangkauan yang cukup untuk eksplorasi."),
      ]),

      heading3("7.2.3 Boundary Handling"),
      p([
        t("Ketika partikel bergerak keluar dari batas fisik yang telah ditentukan (Tabel 1), perlu ada mekanisme untuk mempertahankan partikel dalam ruang yang valid. Tiga pendekatan yang umum adalah: "),
        italic("reflection"),
        t(" (memantulkan kecepatan ketika mengenai batas), "),
        italic("clamping"),
        t(" (memotong nilai pada batas dan mengatur ulang kecepatan), dan "),
        italic("periodic boundary"),
        t(". Dalam inversi geofisika, clamping paling sering digunakan karena sederhana dan menjaga fisikalitas parameter."),
      ]),

      heading3("7.2.4 Kriteria Konvergensi"),
      p([
        t("PSO dihentikan ketika salah satu dari kondisi berikut terpenuhi:"),
      ]),
      bullet("Jumlah iterasi maksimum tercapai: t = t_max (biasanya 500–2000 iterasi)."),
      bullet("Nilai misfit g_best sudah sangat kecil: E(g_best) < E_tol (misalnya E_tol = N, yang berarti RMS ≈ 1)."),
      bullet("Variasi misfit g_best dalam sejumlah iterasi terakhir sangat kecil: |E(g_best^(t)) - E(g_best^(t-100))| < δ."),
      new Paragraph({ spacing: { after: 80 }, children: [] }),

      heading3("7.2.5 Pseudokode PSO untuk Inversi VES"),
      p([
        t("Algoritma PSO untuk inversi VES dapat dirangkum sebagai berikut dalam pseudokode:"),
      ]),
      new Paragraph({
        spacing: { before: 120, after: 120, line: 280 },
        indent: { left: 720 },
        children: [new TextRun({ text: "INISIALISASI:", font: "Courier New", size: szSm, bold: true })],
      }),
      new Paragraph({
        spacing: { after: 40, line: 260 },
        indent: { left: 720 },
        children: [new TextRun({ text: "  Untuk k = 1 sampai K:", font: "Courier New", size: szSm })],
      }),
      new Paragraph({
        spacing: { after: 40, line: 260 },
        indent: { left: 720 },
        children: [new TextRun({ text: "    m_k ~ Uniform(m_min, m_max)", font: "Courier New", size: szSm })],
      }),
      new Paragraph({
        spacing: { after: 40, line: 260 },
        indent: { left: 720 },
        children: [new TextRun({ text: "    v_k = 0;  p_best,k = m_k", font: "Courier New", size: szSm })],
      }),
      new Paragraph({
        spacing: { after: 40, line: 260 },
        indent: { left: 720 },
        children: [new TextRun({ text: "  g_best = argmin_k [E(m_k)]", font: "Courier New", size: szSm })],
      }),
      new Paragraph({
        spacing: { before: 120, after: 40, line: 260 },
        indent: { left: 720 },
        children: [new TextRun({ text: "ITERASI (t = 1, 2, ..., t_max):", font: "Courier New", size: szSm, bold: true })],
      }),
      new Paragraph({
        spacing: { after: 40, line: 260 },
        indent: { left: 720 },
        children: [new TextRun({ text: "  Untuk k = 1 sampai K:", font: "Courier New", size: szSm })],
      }),
      new Paragraph({
        spacing: { after: 40, line: 260 },
        indent: { left: 720 },
        children: [new TextRun({ text: "    Hitung v_k(t+1) dengan Persamaan (26)", font: "Courier New", size: szSm })],
      }),
      new Paragraph({
        spacing: { after: 40, line: 260 },
        indent: { left: 720 },
        children: [new TextRun({ text: "    Perbarui m_k(t+1) dengan Persamaan (27)", font: "Courier New", size: szSm })],
      }),
      new Paragraph({
        spacing: { after: 40, line: 260 },
        indent: { left: 720 },
        children: [new TextRun({ text: "    Terapkan boundary handling", font: "Courier New", size: szSm })],
      }),
      new Paragraph({
        spacing: { after: 40, line: 260 },
        indent: { left: 720 },
        children: [new TextRun({ text: "    Hitung E(m_k(t+1)) menggunakan Persamaan (18)", font: "Courier New", size: szSm })],
      }),
      new Paragraph({
        spacing: { after: 40, line: 260 },
        indent: { left: 720 },
        children: [new TextRun({ text: "    Jika E(m_k(t+1)) < E(p_best,k): p_best,k = m_k(t+1)", font: "Courier New", size: szSm })],
      }),
      new Paragraph({
        spacing: { after: 40, line: 260 },
        indent: { left: 720 },
        children: [new TextRun({ text: "    Jika E(p_best,k) < E(g_best): g_best = p_best,k", font: "Courier New", size: szSm })],
      }),
      new Paragraph({
        spacing: { after: 40, line: 260 },
        indent: { left: 720 },
        children: [new TextRun({ text: "  Kurangi w secara linear", font: "Courier New", size: szSm })],
      }),
      new Paragraph({
        spacing: { after: 40, line: 260 },
        indent: { left: 720 },
        children: [new TextRun({ text: "  Cek kriteria konvergensi; jika terpenuhi, STOP", font: "Courier New", size: szSm })],
      }),
      new Paragraph({
        spacing: { before: 80, after: 40, line: 260 },
        indent: { left: 720 },
        children: [new TextRun({ text: "OUTPUT: m_best = g_best, E_final = E(g_best)", font: "Courier New", size: szSm, bold: true })],
      }),
      new Paragraph({ spacing: { after: 160 }, children: [] }),

      heading2("7.3 Perbandingan Grid Search dan PSO"),
      p([
        t("Tabel 3 merangkum perbandingan menyeluruh antara kedua metode dalam konteks inversi VES."),
      ]),
      new Paragraph({ spacing: { after: 160 }, children: [] }),
      makeTable(
        ["Aspek", "Grid Search", "PSO"],
        [
          ["Sifat pencarian", "Deterministik, exhaustive", "Stokastik, adaptif"],
          ["Jumlah evaluasi forward", "K^M (eksponensial)", "K × t_max (linear dalam M)"],
          ["Jaminan solusi global", "Ya (jika grid halus)", "Probabilistik (hampir pasti untuk K, t besar)"],
          ["Skalabilitas (M besar)", "Sangat buruk (curse of dimensionality)", "Baik, skalabel"],
          ["Informasi misfit surface", "Lengkap, dapat divisualisasikan", "Terbatas (hanya sepanjang trajektori partikel)"],
          ["Kebutuhan initial guess", "Tidak perlu", "Tidak perlu"],
          ["Sensitivitas terhadap noise", "Sedang", "Rendah (karena populasi besar)"],
          ["Implementasi", "Sangat sederhana", "Relatif sederhana"],
          ["Parallelisasi", "Mudah sekali", "Mudah (evaluasi partikel independen)"],
        ],
        [2200, 3300, 3300]
      ),
      caption("Tabel 3. Perbandingan Grid Search dan PSO untuk inversi VES."),

      p([
        t("Dari perbandingan ini terlihat jelas bahwa Grid Search dan PSO memiliki niche yang berbeda. Grid Search sangat cocok untuk problem dengan sedikit parameter (M ≤ 3) di mana visualisasi misfit surface lengkap sangat berharga untuk memahami non-uniqueness. Sebaliknya, PSO lebih cocok untuk problem multi-parameter di mana eksplorasi efisien lebih diutamakan daripada pemetaan lengkap ruang parameter."),
      ]),
      p([
        t("Mekanisme yang membuat PSO mengatasi curse of dimensionality adalah sifat adaptifnya: partikel tidak mengeksplorasi ruang secara seragam, melainkan bergerak secara kolektif menuju area yang menjanjikan (nilai misfit rendah). Komponen kognitif (c₁) memastikan setiap partikel 'mengingat' sejarah pencarian terbaiknya, sementara komponen sosial (c₂) memastikan transfer informasi antar partikel. Kombinasi ini menciptakan eksplorasi yang jauh lebih efisien dibandingkan grid regular."),
      ]),

      // ─── 8. OUTPUT ────────────────────────────────────────────────────────
      heading1("8. OUTPUT YANG DIHARAPKAN DARI INVERSI"),

      heading2("8.1 Model Optimal"),
      p([
        t("Output utama dari proses inversi adalah model optimal:"),
      ]),
      eq("m_best = [ρ₁*, h₁*, ρ₂*, h₂*, ρ₃*, h₃*, ρ₄*]ᵀ", "(28)"),
      p([
        t("yaitu vektor parameter yang menghasilkan nilai misfit terkecil E(m_best). Nilai ini kemudian harus diinterpretasi dalam konteks geologi setempat."),
      ]),

      heading2("8.2 Data Kalkulasi dan Kurva Fitting"),
      p([
        t("Dari model optimal, data kalkulasi diperoleh melalui forward modeling:"),
      ]),
      eq("d_calc = G(m_best) = [ρa_calc(r₁), ..., ρa_calc(rN)]ᵀ", "(29)"),
      p([
        t("Kurva fitting, yaitu plot d_obs vs d_calc sebagai fungsi spasi elektroda AB/2 pada skala log-log, merupakan cara standar untuk mengevaluasi kualitas inversi. Kurva yang fitted dengan baik menunjukkan bahwa model m_best mampu mereproduksi data lapangan dalam batas ketidakpastian pengukuran."),
      ]),

      heading2("8.3 Nilai Misfit dan RMS Error"),
      p([
        t("Nilai misfit akhir E(m_best) dan RMS error (Persamaan 19) memberikan ukuran kuantitatif kualitas solusi. Kriteria umum yang diterima adalah RMS < 1,5, yang menunjukkan bahwa rata-rata residual tidak melebihi 1,5 kali standar deviasi noise."),
      ]),

      heading2("8.4 Misfit Surface (Grid Search)"),
      p([
        t("Untuk Grid Search, seluruh landscape E(m) di ruang parameter tersimpan dan dapat divisualisasikan. Contoh visualisasi yang informatif:"),
      ]),
      bullet("Kontur plot E(ρ₁, ρ₂) dengan parameter lain dipegang pada nilai optimal: menunjukkan trade-off antara resistivitas lapisan 1 dan 2."),
      bullet("Heatmap E(ρᵢ, hᵢ) untuk setiap lapisan i: menunjukkan resolusi masing-masing parameter."),
      bullet("1D marginal plot E(mⱼ) setelah minimisasi terhadap semua parameter lain: menunjukkan ketidakpastian parameter individual."),
      new Paragraph({ spacing: { after: 80 }, children: [] }),

      heading2("8.5 Kurva Konvergensi (PSO)"),
      p([
        t("Untuk PSO, kurva konvergensi yang memplot nilai misfit g_best sebagai fungsi nomor iterasi merupakan diagnostik penting. Kurva yang turun monoton menuju plateau menunjukkan konvergensi yang sehat. Penurunan yang terlalu cepat dan prematur mungkin mengindikasikan premature convergence, di mana swarm terperangkap pada minimum lokal."),
      ]),

      heading2("8.6 Interpretasi Geologis"),
      p([
        t("Interpretasi geologis model yang diperoleh adalah tujuan akhir dari seluruh proses inversi. Dalam konteks VES untuk survei air tanah, misalnya:"),
      ]),
      bullet("Lapisan dengan ρ < 10 Ω·m: umumnya mengindikasikan tanah lempung jenuh air atau air asin."),
      bullet("Lapisan dengan ρ = 10–100 Ω·m: tipikal pasir berkerikil jenuh air (zona akuifer potensial)."),
      bullet("Lapisan dengan ρ > 1000 Ω·m: biasanya batuan keras (basement) atau material kering."),
      new Paragraph({ spacing: { after: 80 }, children: [] }),
      p([
        t("Korelasi antara model resistivitas 1D dari berbagai titik VES sepanjang sebuah lintasan memungkinkan konstruksi penampang geologi 2D yang sangat berguna untuk perencanaan pengeboran atau penilaian risiko geoteknik."),
      ]),

      // ─── 9. KESIMPULAN ────────────────────────────────────────────────────
      heading1("9. KESIMPULAN"),
      p([
        t("Esai ini telah memaparkan secara komprehensif penerapan inversi global metode Grid Search dan PSO pada data VES konfigurasi Schlumberger. Beberapa poin utama dapat dirangkum sebagai berikut:"),
      ]),
      bullet("Forward model VES 1D berlapis didasarkan pada solusi analitik Persamaan Laplace menggunakan transformasi Hankel dan algoritma rekursi Pekeris. Hubungan antara data (resistivitas semu) dan model (resistivitas dan ketebalan lapisan) bersifat non-linear, menjadikan problem inversi VES sebagai tantangan komputasi yang sesungguhnya.", true),
      bullet("Parameter inversi VES untuk model N lapisan adalah vektor berdimensi M = 2N-1, yang terdiri dari pasangan (ρᵢ, hᵢ) untuk i = 1 sampai N-1 dan ρ_N untuk halfspace terbawah. Batasan fisik yang realistis harus diterapkan untuk semua parameter.", true),
      bullet("Data lapangan berupa resistivitas semu ρa yang diukur pada berbagai spasi elektroda AB/2, dengan konfigurasi Schlumberger yang memberikan faktor geometri K yang didefinisikan oleh Persamaan (8). Noise dalam data muncul dari sumber instrumen, lingkungan, dan geologis, dan umumnya dimodelkan sebagai Gaussian proporsional.", true),
      bullet("Grid Search memberikan eksplorasi ruang parameter yang lengkap dan deterministik, tetapi mengalami curse of dimensionality yang membuatnya tidak praktis untuk M > 4-5. Kelebihannya adalah kemampuan menghasilkan misfit surface yang lengkap untuk analisis non-uniqueness.", true),
      bullet("PSO mengatasi curse of dimensionality melalui mekanisme pencarian adaptif berbasis populasi. Dengan hanya K × t_max evaluasi forward (yang hanya tumbuh secara linear terhadap M), PSO mampu menemukan solusi optimal yang kompetitif untuk problem multi-parameter dalam fraksi waktu komputasi Grid Search.", true),
      bullet("Kombinasi kedua metode dapat dilakukan secara praktis: Grid Search kasar digunakan untuk mendapatkan gambaran awal misfit landscape dan menentukan rentang parameter yang lebih sempit, kemudian PSO dijalankan dalam ruang yang dipersempit untuk konvergensi yang lebih cepat dan akurat.", true),
      new Paragraph({ spacing: { after: 80 }, children: [] }),
      p([
        t("Ke depannya, pengembangan metode inversi VES yang semakin canggih, termasuk integrasi dengan informasi geologi a priori dalam kerangka inversi Bayesian dan penggunaan komputasi paralel untuk Grid Search dan PSO, akan terus meningkatkan kemampuan interpretasi data geolistrik dalam berbagai aplikasi eksplorasi dan investigasi lingkungan."),
      ]),

      // ─── DAFTAR PUSTAKA ───────────────────────────────────────────────────
      pageBreak(),
      heading1("DAFTAR PUSTAKA"),
      ref("[1] Menke, W. (2018). Geophysical Data Analysis: Discrete Inverse Theory. 4th ed. Academic Press, London. 352 hlm."),
      ref("[2] Aster, R. C., Borchers, B., & Thurber, C. H. (2019). Parameter Estimation and Inverse Problems. 3rd ed. Elsevier, Amsterdam. 404 hlm."),
      ref("[3] Tarantola, A. (2005). Inverse Problem Theory and Methods for Model Parameter Estimation. SIAM, Philadelphia. 342 hlm."),
      ref("[4] Kennedy, J. & Eberhart, R. (1995). Particle swarm optimization. Proceedings of the IEEE International Conference on Neural Networks, 4, 1942–1948."),
      ref("[5] Grandis, H. (2009). Pengantar Pemodelan Inversi Geofisika. HAGI, Jakarta."),
      ref("[6] Sen, M. K. & Stoffa, P. L. (2013). Global Optimization Methods in Geophysical Inversion. 2nd ed. Cambridge University Press, Cambridge. 289 hlm."),
      ref("[7] Koefoed, O. (1979). Geosounding Principles, Vol. 1: Resistivity Sounding Measurements. Elsevier Scientific Publishing Co., Amsterdam. 276 hlm."),
      ref("[8] Loke, M. H. (2021). Tutorial: 2-D and 3-D Electrical Imaging Surveys. Universiti Sains Malaysia. Tersedia online di www.geotomosoft.com."),
      ref("[9] Pekeris, C. L. (1940). Direct method of interpretation in resistivity prospecting. Geophysics, 5(1), 31–42."),
      ref("[10] Shi, Y. & Eberhart, R. C. (1998). A modified particle swarm optimizer. Proceedings of the IEEE World Congress on Computational Intelligence, 69–73."),

    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("/mnt/user-data/outputs/Esai_Inversi_VES_GridSearch_PSO.docx", buf);
  console.log("Done! File written.");
}).catch(err => {
  console.error("Error:", err);
  process.exit(1);
});
