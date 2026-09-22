/**
 * ============================================================================
 * PintarSaham - Data Ingestion: Sheet Ringkasan Profil Emiten (69 Saham)
 * ============================================================================
 * Menambahkan sheet "Ringkasan" ke spreadsheet Data Dashboard dan mengisinya
 * dengan profil bisnis, sektor, dan highlight keunggulan untuk 69 emiten.
 *
 * CARA PAKAI:
 *   1. Buka Apps Script Editor di Spreadsheet Anda.
 *   2. Buat file script baru bernama "populate_ringkasan" (atau paste di bawah kode.gs).
 *   3. Pilih fungsi "populateRingkasanData" di toolbar atas, lalu klik "Run".
 *   4. Sheet "Ringkasan" akan otomatis terbuat dan terisi rapi!
 * ============================================================================
 */

function populateRingkasanData() {
  const SPREADSHEET_ID = '1tbkh-ulOyzm-SQm2H_cLEpZZDzQY7A0uB1v3fv-PAjE';
  let ss;
  try {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  if (!ss) {
    throw new Error('Spreadsheet tidak ditemukan. Pastikan SPREADSHEET_ID valid.');
  }

  const SHEET_NAME = 'Ringkasan';
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  const header = [
    'Kode',
    'Nama Perusahaan',
    'Sektor',
    'Subsektor',
    'Ringkasan Bisnis',
    'Highlight / Keunggulan'
  ];

  const data = [
    [
      'AADI',
      'PT Adaro Andalan Indonesia Tbk',
      'Energi',
      'Pertambangan Batubara Termal',
      'Perusahaan pertambangan batubara termal hasil spin-off dari Adaro Energy yang mengelola tambang batubara berkualitas tinggi di Kalimantan Selatan dan Tengah.',
      '• Salah satu produsen batubara termal berbiaya rendah di Indonesia.\n• Cadangan batubara masif dengan infrastruktur rantai pasok sungai terintegrasi.\n• Kebijakan pembagian dividen yang menarik bagi pemegang saham.'
    ],
    [
      'ACES',
      'PT Aspirasi Hidup Indonesia Tbk',
      'Konsumer Siklikal',
      'Ritel Perbaikan Rumah & Gaya Hidup',
      'Ritel perlengkapan rumah tangga dan gaya hidup terkemuka di Indonesia (sebelumnya pemegang lisensi Ace Hardware) dengan ratusan gerai di berbagai kota besar.',
      '• Jaringan toko luas dan loyalitas pelanggan yang solid (member base kuat).\n• Portofolio produk private label eksklusif ber-margin tinggi.\n• Pertumbuhan gerai konsisten di kota tier-1 dan tier-2.'
    ],
    [
      'ADMR',
      'PT Adaro Minerals Indonesia Tbk',
      'Energi',
      'Pertambangan Batubara Metalurgi',
      'Produsen batubara metalurgi kokas keras (coking coal) pertama dan terbesar di Indonesia, serta berekspansi ke proyek smelter aluminium hijau di Kalimantan Utara.',
      '• Memonopoli pasokan coking coal premium domestik untuk industri peleburan baja.\n• Marjin laba operasional tinggi didukung biaya penambangan yang kompetitif.\n• Prospek jangka panjang strategis dari proyek hilirisasi smelter aluminium.'
    ],
    [
      'ADRO',
      'PT Alamtri Resources Indonesia Tbk',
      'Energi',
      'Investasi & Hilirisasi Energi Hijau',
      'Grup energi terdiversifikasi yang bertransformasi dari tambang energi fosil menuju pembangkit EBT, pengolahan mineral, dan industri hijau berkelanjutan.',
      '• Neraca keuangan sangat sehat dengan kas melimpah (net cash position).\n• Diversifikasi agresif ke smelter aluminium hijau dan energi terbarukan.\n• Rekam jejak tata kelola perusahaan (GCG) dan pembagian dividen konsisten.'
    ],
    [
      'AKRA',
      'PT AKR Corporindo Tbk',
      'Energi',
      'Distribusi BBM, Logistik & Kawasan Industri',
      'Penyedia rantai pasok logistik dan distributor BBM/kimia dasar terintegrasi, serta pengembang kawasan industri terpadu JIIPE Gresik dengan pelabuhan laut dalam.',
      '• Skema marjin fee distribusi BBM dan kimia yang stabil dan minim risiko harga komoditas.\n• Monetisasi penjualan lahan kawasan industri JIIPE (Java Integrated Industrial and Ports Estate).\n• Kemitraan strategis ritel SPBU bersama BP dan ekosistem smelter Freeport di JIIPE.'
    ],
    [
      'AMMN',
      'PT Amman Mineral Internasional Tbk',
      'Bahan Baku',
      'Pertambangan Tembaga & Emas',
      'Perusahaan tambang tembaga dan emas terbesar kedua di Indonesia yang mengoperasikan tambang Batu Hijau dan mengembangkan proyek prospektif Elang di Sumbawa, NTB.',
      '• Cadangan tembaga dan emas masif dengan umur tambang produktif jangka panjang.\n• Penyelesaian fasilitas smelter tembaga domestik meningkatkan nilai tambah produk.\n• Tingginya permintaan tembaga global didorong tren elektrifikasi dan transisi energi.'
    ],
    [
      'AMRT',
      'PT Sumber Alfaria Trijaya Tbk',
      'Konsumer Non-Siklikal',
      'Ritel Minimarket (Alfamart)',
      'Pengelola jaringan ritel minimarket Alfamart dan Alfamidi terbesar di Indonesia dengan lebih dari 20.000 gerai yang tersebar hingga ke pelosok pemukiman.',
      '• Dominasi jaringan distribusi ritel modern terdekat dari pemukiman konsumen (proximity store).\n• Pertumbuhan fee-based income dari layanan transaksi digital dan logistik pickup point.\n• Perputaran persediaan (inventory turnover) sangat efisien dan ekspansi luar Jawa.'
    ],
    [
      'ANTM',
      'PT Aneka Tambang Tbk',
      'Bahan Baku',
      'Pertambangan Emas, Nikel & Bauksit',
      'BUMN tambang terintegrasi vertikal yang memproduksi emas batangan Logam Mulia, feronikel, bijih nikel, dan bauksit, serta memegang peran kunci ekosistem baterai EV.',
      '• Merek emas batangan Logam Mulia memiliki likuiditas dan kepercayaan tertinggi di Indonesia.\n• Posisi strategis dalam rantai pasok ekosistem baterai kendaraan listrik (EV).\n• Penjualan emas ritel menjadi instrumen lindung nilai (safe haven) saat ketidakpastian makro.'
    ],
    [
      'ASII',
      'PT Astra International Tbk',
      'Perindustrian',
      'Konglomerasi Otomotif, Alat Berat & Keuangan',
      'Konglomerasi terbesar di Indonesia dengan portofolio bisnis terdiversifikasi di sektor otomotif (pemimpin pasar mobil/motor), alat berat & pertambangan (UNTR), dan jasa keuangan.',
      '• Penguasa pangsa pasar mobil (>50%) dan sepeda motor (>70%) di Indonesia.\n• Ekosistem bisnis terintegrasi yang saling menopang (pembiayaan, asuransi, perbengkelan).\n• Arus kas operasional yang masif dan kebijakan dividen yang sangat royal.'
    ],
    [
      'AVIA',
      'PT Avia Avian Tbk',
      'Bahan Baku',
      'Manufaktur Cat & Bahan Bangunan',
      'Pemimpin pasar cat dekoratif dan arsitektural di Indonesia dengan merek legendaris Avian dan No Drop yang didukung jaringan distribusi mandiri ke puluhan ribu toko.',
      '• Pangsa pasar cat dekoratif nomor 1 di Indonesia (>20%).\n• Sistem distribusi langsung mandiri (in-house logistics) ke puluhan ribu toko bahan bangunan.\n• Marjin laba kotor dan marjin operasional tinggi berkat efisiensi formulasi bahan baku.'
    ],
    [
      'BBCA',
      'PT Bank Central Asia Tbk',
      'Keuangan',
      'Bank Konvensional',
      'Bank swasta terbesar di Indonesia yang terkenal dengan keunggulan transaksi digital, basis nasabah setia, serta likuiditas dana murah (CASA) yang sangat mendominasi.',
      '• Rasio dana murah (CASA) konsisten di atas 80%, memberikan biaya dana (Cost of Funds) terendah di industri.\n• Kualitas aset superior dengan NPL sangat rendah dan pencadangan tebal.\n• Return on Equity (ROE) konsisten di atas 20% didukung ekosistem perbankan digital masif.'
    ],
    [
      'BBNI',
      'PT Bank Negara Indonesia (Persero) Tbk',
      'Keuangan',
      'Bank Konvensional',
      'Salah satu bank BUMN terbesar di Indonesia dengan mandat fokus pada pembiayaan korporasi berskala besar, sindikasi, dan bisnis transaksional internasional/global banking.',
      '• Transformasi digital dan pembersihan kualitas aset yang sukses menekan credit cost.\n• Penguasaan pembiayaan segmen korporasi tier-1 dan nasabah diaspora global.\n• Pertumbuhan laba solid didukung perbaikan marjin bunga bersih (NIM) dan efisiensi operasional.'
    ],
    [
      'BBRI',
      'PT Bank Rakyat Indonesia (Persero) Tbk',
      'Keuangan',
      'Bank Konvensional & Finansial Mikro',
      'Bank dengan laba terbesar di Indonesia dan penguasa pasar kredit mikro serta ultra-mikro (Holding Ultra Mikro bersama Pegadaian dan PNM).',
      '• Jaringan branchless banking (AgenBRILink) mencapai ratusan ribu agen di pelosok desa.\n• Marjin bunga bersih (NIM) tebal berkat dominasi segmen kredit mikro berimbal hasil tinggi.\n• Dividen payout ratio historis tinggi (>80%) ditopang rasio kecukupan modal (CAR) yang kokoh.'
    ],
    [
      'BBTN',
      'PT Bank Tabungan Negara (Persero) Tbk',
      'Keuangan',
      'Bank Pembiayaan Perumahan (KPR)',
      'Bank BUMN dengan spesialisasi utama pada pembiayaan perumahan rakyat (KPR subsidi dan non-subsidi), menguasai pangsa pasar terbesar KPR di Indonesia.',
      '• Penguasa pasar KPR nasional dengan mandat strategis program perumahan rakyat pemerintah.\n• Permintaan kredit KPR yang defensif seiring tingginya angka backlog perumahan di Indonesia.\n• Upaya terukur menurunkan biaya dana melalui perbaikan rasio CASA dan platform digital perumahan.'
    ],
    [
      'BMRI',
      'PT Bank Mandiri (Persero) Tbk',
      'Keuangan',
      'Bank Konvensional',
      'Bank dengan total aset terbesar di Indonesia yang unggul dalam pembiayaan korporasi, ekosistem wholesale, serta aplikasi digital retail Livin\' dan wholesale Kopra.',
      '• Penguasa segmen kredit korporasi dan komersial dengan ekosistem supply chain terintegrasi.\n• Aplikasi Livin\' by Mandiri menjadi salah satu super-app finansial paling aktif di Indonesia.\n• Pertumbuhan CASA dan fee-based income masif dengan rasio NPL terendah dalam sejarah bank.'
    ],
    [
      'BRPT',
      'PT Barito Pacific Tbk',
      'Bahan Baku',
      'Petrokimia & Energi Terbarukan',
      'Holding konglomerasi industri milik Prajogo Pangestu yang menaungi raksasa petrokimia Chandra Asri Pacific (TPIA) dan produsen panas bumi Star Energy Geothermal (BREN).',
      '• Portofolio aset strategis di sektor petrokimia terintegrasi terbesar di Indonesia.\n• Produsen energi geotermal dengan kontrak jangka panjang berdenominasi USD.\n• Sinergi ekspansi infrastruktur energi, pelabuhan, dan logistik industri.'
    ],
    [
      'BSDE',
      'PT Bumi Serpong Damai Tbk',
      'Properti & Real Estat',
      'Pengembangan Kota Mandiri',
      'Pengembang properti terkemuka di bawah grup Sinar Mas Land yang mengelola kawasan kota mandiri BSD City serta berbagai proyek komersial dan residensial strategis.',
      '• Cadangan lahan (landbank) sangat luas di area penyangga Jakarta dengan nilai perolehan historis murah.\n• Marketing sales kuat ditopang segmen residensial menengah-atas dan insentif PPN DTP.\n• Posisi kas sehat dan rasio utang (gearing) yang relatif konservatif dibanding rata-rata industri.'
    ],
    [
      'BTPS',
      'PT Bank BTPN Syariah Tbk',
      'Keuangan',
      'Bank Syariah Finansial Mikro',
      'Bank syariah yang fokus melayani segmen prasejahtera produktif perempuan di pedesaan melalui pembiayaan kelompok tanpa agunan (pola sentra).',
      '• Marjin imbal hasil pembiayaan tinggi sejalan dengan karakteristik segmen ultra-mikro.\n• Model pendampingan lapangan berbasis komunitas yang menciptakan kedekatan nasabah.\n• Rasio permodalan (CAR) sangat tebal dan komitmen dividen yang menarik.'
    ],
    [
      'BUKA',
      'PT Bukalapak.com Tbk',
      'Teknologi',
      'E-Commerce & Layanan O2O (Mitra)',
      'Perusahaan teknologi yang berfokus pada digitalisasi warung tradisional melalui platform Mitra Bukalapak, gaming marketplace, dan manajemen investasi.',
      '• Posisi kas dan setara kas sangat melimpah tanpa beban utang berbunga.\n• Pendapatan bunga dan hasil investasi instrumen likuid menopang profitabilitas bottom-line.\n• Efisiensi operasional dengan fokus pada lini bisnis ber-marjin positif (Mitra & Gaming).'
    ],
    [
      'CMRY',
      'PT Cisarua Mountain Dairy Tbk',
      'Konsumer Non-Siklikal',
      'Olahan Susu & Daging Premium',
      'Produsen produk susu fermentasi (yogurt), susu cair UHT bernutrisi, dan olahan daging beku (Kanzler) premium dengan pertumbuhan penetrasi pasar yang sangat pesat.',
      '• Pemimpin pasar nomor 1 di kategori minuman yogurt (Cimory) dan sosis siap saji premium (Kanzler).\n• Saluran distribusi unik melalui jaringan Miss Cimory melengkapi jaringan ritel modern.\n• Inovasi produk berkelanjutan dengan kekuatan branding premium ber-marjin tebal.'
    ],
    [
      'CPIN',
      'PT Charoen Pokphand Indonesia Tbk',
      'Konsumer Non-Siklikal',
      'Pakan Ternak & Perunggasan Terintegrasi',
      'Produsen pakan ternak, Day-Old-Chicks (DOC), dan olahan ayam (Fiesta) terintegrasi terbesar di Indonesia di bawah naungan Charoen Pokphand Group.',
      '• Penguasa pangsa pasar pakan ternak nasional dengan efisiensi pabrik skala besar.\n• Integrasi vertikal dari hulu (pakan, pembibitan) hingga hilir (makanan olahan Fiesta & Champ).\n• Jaringan gerai ritel Prima Freshmart yang memperkuat penetrasi langsung ke konsumen.'
    ],
    [
      'CTRA',
      'PT Ciputra Development Tbk',
      'Properti & Real Estat',
      'Pengembangan Residensial & Kota Mandiri',
      'Pengembang properti paling terdiversifikasi secara geografis di Indonesia dengan puluhan proyek kota mandiri dan perumahan di lebih dari 30 kota.',
      '• Merek Ciputra memiliki reputasi brand equity tertinggi dalam segmen residensial keluarga.\n• Diversifikasi geografis luas mengurangi risiko konsentrasi di satu wilayah saja.\n• Peningkatan recurring income dari portofolio rumah sakit (Ciputra Hospital), mal, dan hotel.'
    ],
    [
      'CUAN',
      'PT Petrindo Jaya Kreasi Tbk',
      'Energi',
      'Holding Tambang Mineral & Batubara',
      'Holding pertambangan milik Prajogo Pangestu yang membawahi tambang batubara termal, batubara metalurgi, serta diversifikasi ke mineral nikel dan emas.',
      '• Pertumbuhan agresif melalui akuisisi strategis aset-aset tambang berproduksi.\n• Sinergi rantai pasok energi di bawah ekosistem grup Barito.\n• Efisiensi logistik pengangkutan tambang yang terhubung langsung dengan infrastruktur pelabuhan.'
    ],
    [
      'DSNG',
      'PT Dharma Satya Nusantara Tbk',
      'Konsumer Non-Siklikal',
      'Perkebunan Kelapa Sawit & Kayu Olahan',
      'Produsen minyak kelapa sawit mentah (CPO) bereputasi tinggi dengan produktivitas kebun terdepan, serta manufaktur produk kayu olahan berkelanjutan.',
      '• Profil usia kebun prima dengan produktivitas TBS dan oil extraction rate (OER) di atas rata-rata.\n• Penerapan prinsip ESG ketat dengan pemanfaatan limbah pabrik kelapa sawit menjadi biomassa.\n• Biaya produksi tunai (cash cost) terkendali memberikan marjin solid saat harga CPO tinggi.'
    ],
    [
      'DSSA',
      'PT Dian Swastatika Sentosa Tbk',
      'Energi',
      'Pertambangan Batubara, Pembangkit & Teknologi',
      'Holding investasi energi dan infrastruktur grup Sinar Mas yang mengendalikan produsen batubara Golden Energy Mines (GEMS) serta bisnis telekomunikasi dan teknologi.',
      '• Kontribusi laba masif dari ekspor batubara kalori menengah-tinggi melalui anak usaha.\n• Diversifikasi portofolio ke bisnis penyediaan internet fiber optic, data center, dan energi bersih.\n• Likuiditas neraca dan kapasitas dividen yang sangat tebal.'
    ],
    [
      'ELSA',
      'PT Elnusa Tbk',
      'Energi',
      'Jasa Penunjang Hulu Migas & Distribusi BBM',
      'Anak usaha Pertamina Hulu Energi yang bergerak di bidang jasa hulu minyak dan gas (seismik, drilling, well services), logistik distribusi BBM, dan jasa energi maritim.',
      '• Kontrak jangka panjang captive dari Pertamina Group menjamin visibilitas pendapatan.\n• Kinerja divisi distribusi dan logistik energi yang stabil dan konsisten bertumbuh.\n• Peningkatan aktivitas eksplorasi dan lifting migas domestik menopang permintaan jasa hulu.'
    ],
    [
      'EMTK',
      'PT Elang Mahkota Teknologi Tbk',
      'Teknologi',
      'Media, Telekomunikasi & Investasi Digital',
      'Konglomerasi teknologi dan media yang mengendalikan stasiun TV nasional (SCTV, Indosiar melalui SCMA), platform streaming Vidio, serta investasi di ekosistem perbankan digital.',
      '• Vidio menjadi platform OTT streaming lokal nomor 1 di Indonesia dengan hak siar konten olahraga bergengsi.\n• Penguasa pangsa pasar penonton (audience share) TV free-to-air melalui SCTV dan Indosiar.\n• Neraca keuangan sehat dengan likuiditas kas tinggi untuk ekspansi investasi digital.'
    ],
    [
      'ENRG',
      'PT Energi Mega Persada Tbk',
      'Energi',
      'Eksplorasi & Produksi Minyak dan Gas Bumi',
      'Perusahaan migas hulu terafiliasi grup Bakrie yang mengoperasikan berbagai blok konsesi minyak dan gas bumi di Sumatra, Jawa Timur, dan luar negeri.',
      '• Porsi cadangan dan produksi gas bumi yang dominan dengan kontrak penjualan jangka panjang fixed price.\n• Akuisisi blok-blok migas produktif baru yang meningkatkan kapasitas lifting harian.\n• Penurunan rasio utang secara signifikan meningkatkan fleksibilitas keuangan perusahaan.'
    ],
    [
      'ERAA',
      'PT Erajaya Swasembada Tbk',
      'Konsumer Siklikal',
      'Distributor & Ritel Gadget / Elektronik',
      'Distributor dan peritel perangkat telekomunikasi (iPhone, Samsung, Xiaomi), gadget IoT, farmasi, dan gaya hidup terbesar di Indonesia dengan ribuan gerai Erafone dan iBox.',
      '• Mitra distribusi resmi utama produk Apple (iBox) dan merek smartphone global terdepan.\n• Diversifikasi sukses ke segmen ritel gaya hidup (JD Sports), kecantikan, dan makanan.\n• Keberhasilan aturan IMEI pemerintah yang melindungi pasar ritel resmi dari barang black market.'
    ],
    [
      'ESSA',
      'PT ESSA Industries Indonesia Tbk',
      'Bahan Baku',
      'Pengolahan Amonia & Kilang LPG',
      'Produsen amonia biru/bersih dan pengolah gas alam cair (LPG) swasta terkemuka di Indonesia yang mengoperasikan fasilitas pabrik amonia modern di Luwuk, Sulawesi Tengah.',
      '• Salah satu produsen amonia dengan efisiensi energi tertinggi di dunia.\n• Pelopor proyek Blue Ammonia di Asia Tenggara untuk memasok kebutuhan energi bersih masa depan.\n• Keuntungan dari struktur kontrak pasokan gas alam domestik berbiaya kompetitif.'
    ],
    [
      'HEAL',
      'PT Medikaloka Hermina Tbk',
      'Kesehatan',
      'Layanan Rumah Sakit (Hermina)',
      'Jaringan rumah sakit swasta terbesar di Indonesia yang berfokus pada layanan ibu dan anak serta bertransformasi menjadi rumah sakit umum tipe B dan C melayani BPJS dan swasta.',
      '• Operator rumah sakit dengan efisiensi operasional terbaik dalam melayani pasien BPJS Kesehatan.\n• Rencana ekspansi rumah sakit baru yang terukur di kota tier-2 dan proyek IKN Nusantara.\n• Kemitraan strategis dengan investor global (Astra) yang memperkuat struktur permodalan.'
    ],
    [
      'HMSP',
      'PT H.M. Sampoerna Tbk',
      'Konsumer Non-Siklikal',
      'Manufaktur Rokok & Tembakau',
      'Produsen rokok terbesar di Indonesia (afiliasi Philip Morris International) dengan merek legendaris Dji Sam Soe, Sampoerna A, serta pelopor produk tembakau bebas asap IQOS.',
      '• Penguasa pangsa pasar rokok nasional dengan kekuatan merek dan jaringan distribusi terluas.\n• Pertumbuhan segmen Sigaret Kretek Tangan (SKT) yang menikmati tarif cukai lebih rendah.\n• Neraca keuangan sangat sehat tanpa utang bank dan dividen payout ratio mendekati 100%.'
    ],
    [
      'HRTA',
      'PT Hartadinata Abadi Tbk',
      'Konsumer Siklikal',
      'Manufaktur & Ritel Perhiasan Emas',
      'Produsen dan peritel perhiasan emas serta emas batangan terintegrasi di Indonesia dengan jaringan pabrik, toko ritel, dan kemitraan ekspor emas perhiasan ke pasar global.',
      '• Integrasi rantai bisnis dari pabrik pengolahan emas hingga gerai ritel mandiri.\n• Lonjakan pendapatan ekspor perhiasan emas ke negara-negara Timur Tengah dan Asia Selatan.\n• Permintaan emas domestik yang defensif sebagai instrumen tabungan masyarakat.'
    ],
    [
      'HRUM',
      'PT Harum Energy Tbk',
      'Energi',
      'Tambang Batubara & Hilirisasi Nikel',
      'Perusahaan pertambangan batubara milik keluarga Barki yang bertransformasi agresif menjadi pemain terintegrasi di ekosistem pengolahan nikel dan bahan baku baterai.',
      '• Biaya penambangan batubara kalori tinggi yang sangat efisien dan berorientasi ekspor.\n• Sukses mengakuisisi smelter nikel RKEF dan proyek High Pressure Acid Leach (HPAL).\n• Struktur neraca keuangan yang sehat dengan rasio utang terkendali.'
    ],
    [
      'ICBP',
      'PT Indofood CBP Sukses Makmur Tbk',
      'Konsumer Non-Siklikal',
      'Produsen Makanan Olahan (Indomie)',
      'Produsen mie instan terbesar di dunia dengan merek ikonik Indomie, serta mengoperasikan divisi susu (Indomilk), makanan ringan, bumbu, biskuit, dan nutrisi khusus.',
      '• Merek Indomie memiliki loyalitas konsumen (pricing power) yang tak tertandingi di pasar global.\n• Kehadiran internasional masif di Timur Tengah dan Afrika melalui Pinehill Company Limited.\n• Marjin laba kotor yang tebal dan defensif terhadap fluktuasi ekonomi makro.'
    ],
    [
      'INCO',
      'PT Vale Indonesia Tbk',
      'Bahan Baku',
      'Pertambangan & Pengolahan Nikel Matte',
      'Produsen nikel dalam matte berbiaya rendah terbesar di Indonesia yang mengoperasikan tambang dan fasilitas pengolahan pirometalurgi berbasis energi terbarukan (PLTA).',
      '• Salah satu produsen nikel dengan jejak karbon terendah di dunia berkat 3 PLTA mandiri.\n• Kemitraan strategis dengan BUMN Mind ID serta produsen otomotif global.\n• Proyek ekspansi HPAL di Pomalaa dan Morowali yang menjamin pertumbuhan jangka panjang.'
    ],
    [
      'INDF',
      'PT Indofood Sukses Makmur Tbk',
      'Konsumer Non-Siklikal',
      'Holding Pangan & Agribisnis Terintegrasi',
      'Induk usaha konglomerasi pangan grup Salim yang mengendalikan ICBP, Bogasari (produsen tepung terigu terbesar), perkebunan kelapa sawit, dan jaringan distribusi ritel.',
      '• Integrasi hulu-hilir yang memberikan ketahanan pasokan bahan baku pangan nasional.\n• Valuasi saham atraktif dengan diskon holding terhadap nilai pasar anak usahanya (ICBP).\n• Cash flow operasional yang stabil dan dividen reguler yang menopang imbal hasil investor.'
    ],
    [
      'INKP',
      'PT Indah Kiat Pulp & Paper Tbk',
      'Bahan Baku',
      'Manufaktur Bubur Kertas & Kertas Industri',
      'Salah satu produsen pulp, kertas budaya, dan kemasan industri terbesar di dunia di bawah naungan grup Asia Pulp & Paper (APP) Sinar Mas dengan pasar ekspor global.',
      '• Skala ekonomi raksasa dengan integrasi pasokan serat kayu bersertifikasi lestari.\n• Permintaan kemasan karton industri didorong oleh pertumbuhan logistik dan e-commerce global.\n• Efisiensi biaya produksi pulp per ton yang menempatkan perusahaan di kuartil terendah kurva biaya global.'
    ],
    [
      'INTP',
      'PT Indocement Tunggal Prakarsa Tbk',
      'Bahan Baku',
      'Manufaktur Semen (Tiga Roda)',
      'Produsen semen terbesar kedua di Indonesia dengan merek terpercaya Semen Tiga Roda, mengoperasikan kompleks pabrik modern di Citeureup, Cirebon, dan Tarjun.',
      '• Penguasaan pangsa pasar semen yang sangat dominan di wilayah Jabodetabek dan Jawa Barat.\n• Akuisisi Semen Grobogan memperkuat penetrasi logistik di Jawa Tengah dan Jawa Timur.\n• Neraca keuangan sangat kokoh tanpa utang berbunga (zero debt) dan kas melimpah.'
    ],
    [
      'ISAT',
      'PT Indosat Tbk',
      'Komunikasi',
      'Operator Telekomunikasi Seluler',
      'Operator seluler terbesar kedua di Indonesia (Indosat Ooredoo Hutchison) hasil merger sukses yang melayani puluhan juta pelanggan dengan merek IM3 dan Tri.',
      '• Realisasi sinergi merger yang melampaui target, menghasilkan efisiensi capex dan opex.\n• Peningkatan Average Revenue Per User (ARPU) didorong oleh monetisasi data dan paket digital.\n• Ekspansi jaringan ke luar Jawa dan adopsi teknologi AI untuk pengoptimalan infrastruktur jaringan.'
    ],
    [
      'ITMG',
      'PT Indo Tambangraya Megah Tbk',
      'Energi',
      'Pertambangan Batubara & Perdagangan Energi',
      'Produsen batubara kalori menengah-tinggi yang terkenal dengan efisiensi operasional tinggi dan kebijakan pembagian dividen bernilai besar (high dividend yield).',
      '• Rata-rata harga jual batubara (ASP) premium berkat kualitas batubara kalori tinggi.\n• Komitmen dividen payout ratio yang sangat royal kepada pemegang saham publik.\n• Diversifikasi ke proyek energi surya fotovoltaik dan perdagangan bahan bakar industri.'
    ],
    [
      'JPFA',
      'PT Japfa Comfeed Indonesia Tbk',
      'Konsumer Non-Siklikal',
      'Agri-Food & Peternakan Terintegrasi',
      'Perusahaan agri-food terintegrasi vertikal terbesar kedua di Indonesia yang bergerak di bidang pakan ternak, pembibitan ayam, budidaya perikanan, dan produk daging olahan.',
      '• Efisiensi pabrik pakan ternak modern dengan sebaran fasilitas di seluruh kepulauan Indonesia.\n• Penguatan hilirisasi produk protein hewani bernilai tambah tinggi (So Good dan Best Meat).\n• Perbaikan tata kelola pasokan DOC nasional oleh regulator yang menjaga stabilitas marjin industri.'
    ],
    [
      'JSMR',
      'PT Jasa Marga (Persero) Tbk',
      'Infrastruktur',
      'Operator Jalan Tol',
      'BUMN pengelola dan operator jalan tol terbesar di Indonesia yang menguasai lebih dari 50% panjang jalan tol beroperasi di seluruh Indonesia, termasuk Jalan Tol Trans Jawa.',
      '• Arus kas operasional (toll revenue) yang sangat stabil dan terus meningkat seiring volume mobilitas.\n• Monetisasi aset melalui kemitraan strategis ekuitas di anak usaha jalan tol.\n• Kenaikan tarif tol berkala setiap dua tahun yang terjamin regulasi sebagai lindung nilai inflasi.'
    ],
    [
      'KIJA',
      'PT Kawasan Industri Jababeka Tbk',
      'Properti & Real Estat',
      'Pengembangan Kawasan Industri Terpadu',
      'Pengembang kawasan industri kota mandiri terintegrasi di Cikarang (Kota Jababeka) dan Kendal (KIK bersama Sembcorp), lengkap dengan pelabuhan kering dan pembangkit listrik.',
      '• Ekosistem industri terpadu dengan tenant multinasional terkemuka dari Jepang dan Eropa.\n• Pertumbuhan pesat penjualan lahan di Kawasan Industri Kendal berstatus KEK.\n• Recurring revenue kuat dari penyediaan air bersih, pengolahan limbah, dan listrik industri.'
    ],
    [
      'KLBF',
      'PT Kalbe Farma Tbk',
      'Kesehatan',
      'Farmasi, Nutrisi & Layanan Kesehatan',
      'Perusahaan farmasi dan produk kesehatan terbesar di Asia Tenggara yang memproduksi obat resep, obat bebas (Promag, Mixagrip), produk nutrisi, dan jaringan distribusi Enseval.',
      '• Portofolio produk farmasi dan kesehatan konsumen paling komprehensif di Indonesia.\n• Rantai pasok distribusi farmasi (Enseval) terluas yang menjangkau seluruh apotek dan klinik.\n• Investasi strategis dalam pengembangan obat biologis (biopharmaceuticals) dan vaksin inovatif.'
    ],
    [
      'KPIG',
      'PT MNC Land Tbk',
      'Properti & Real Estat',
      'Pengembangan Properti & Hospitality',
      'Perusahaan properti grup MNC yang mengembangkan kawasan ekonomi khusus (KEK) pariwisata terintegrasi MNC Lido City, hotel bintang lima, dan gedung perkantoran.',
      '• Status Kawasan Ekonomi Khusus (KEK) Lido dengan fasilitas insentif perpajakan dan kepabeanan.\n• Fasilitas pariwisata terpadu termasuk lapangan golf kelas dunia dan theme park.\n• Portofolio aset hospitality komersial di lokasi strategis Jakarta, Surabaya, dan Bali.'
    ],
    [
      'MAPA',
      'PT MAP Aktif Adiperkasa Tbk',
      'Konsumer Siklikal',
      'Ritel Olahraga, Sepatu & Anak (Planet Sports)',
      'Anak usaha MAPI yang memegang lisensi dan mengoperasikan gerai ritel produk olahraga dan gaya hidup terkemuka (Planet Sports, Sports Station, Foot Locker, Kidz Station).',
      '• Hak distribusi eksklusif merek olahraga global nomor 1 dunia (Nike, Adidas, Puma, Skechers, New Balance).\n• Tren gaya hidup sehat dan olahraga perkotaan (running, gym, padel) mendorong pertumbuhan ritel.\n• Ekspansi sukses ke pasar regional Asia Tenggara (Filipina, Vietnam, Thailand, Malaysia).'
    ],
    [
      'MAPI',
      'PT Mitra Adiperkasa Tbk',
      'Konsumer Siklikal',
      'Ritel Gaya Hidup, Departemen Store & F&B',
      'Ritel gaya hidup nomor 1 di Indonesia dengan portofolio lebih dari 150 merek global di segmen busana (Zara, Mango), F&B (Starbucks, Subway), dan department store (Sogo, Seibu).',
      '• Portofolio merek paling bergengsi menyasar konsumen segmen menengah-atas dengan daya beli tangguh.\n• Jaringan gerai ritel luas di mal-mal premium di seluruh kota utama Indonesia.\n• Strategi omnichannel digital sukses melalui platform e-commerce dan program loyalitas MAPCLUB.'
    ],
    [
      'MEDC',
      'PT Medco Energi Internasional Tbk',
      'Energi',
      'Eksplorasi Migas, Tembaga & Pembangkit Listrik',
      'Perusahaan energi independen terkemuka di Asia Tenggara yang bergerak di bidang eksplorasi migas hulu, pertambangan tembaga (kepemilikan signifikan di AMMN), dan pembangkit listrik.',
      '• Rekam jejak akuisisi aset migas berproduksi tinggi yang sukses (Blok Corridor dan Natuna).\n• Kontribusi laba dan dividen signifikan dari kepemilikan saham di Amman Mineral (AMMN).\n• Biaya lifting migas kompetitif serta bauran penjualan gas jangka panjang yang stabil.'
    ],
    [
      'MIKA',
      'PT Mitra Keluarga Karyasehat Tbk',
      'Kesehatan',
      'Layanan Rumah Sakit Swasta',
      'Jaringan rumah sakit swasta premium terkemuka di Indonesia yang melayani pasien segmen menengah dan menengah-atas dengan fokus pada keunggulan klinis dan kepuasan pasien.',
      '• Marjin profitabilitas (EBITDA margin) tertinggi di antara operator rumah sakit publik di Indonesia.\n• Manajemen biaya operasional yang disiplin dan efisiensi permodalan yang prima.\n• Neraca keuangan sangat sehat tanpa beban utang bank dan posisi kas operasional tebal.'
    ],
    [
      'MTEL',
      'PT Dayamitra Telekomunikasi Tbk',
      'Infrastruktur',
      'Penyedia Menara Telekomunikasi (Mitratel)',
      'Anak usaha Telkom Indonesia yang merupakan pemilik dan operator menara telekomunikasi terbesar di Asia Tenggara dengan lebih dari 38.000 menara di seluruh Indonesia.',
      '• Kepemilikan menara terbesar dengan keunggulan footprint menara di luar pulau Jawa.\n• Rasio penyewaan (tenancy ratio) terus bertumbuh didorong ekspansi jaringan 4G/5G operator seluler.\n• Ekspansi agresif ke bisnis fiber-to-the-tower (FTTT) sebagai tulang punggung data modern.'
    ],
    [
      'MYOR',
      'PT Mayora Indah Tbk',
      'Konsumer Non-Siklikal',
      'Makanan & Minuman Olahan (Kopiko, Torabika)',
      'Produsen makanan dan minuman olahan terkemuka dengan merek global seperti biskuit Roma, permen Kopiko, kopi Torabika, sereal Energen, dan air mineral Le Minerale.',
      '• Kekuatan inovasi produk makanan/minuman dengan penetrasi pasar global ke lebih dari 100 negara.\n• Porsi penjualan ekspor tinggi (>40%) memberikan lindung nilai alami terhadap valuta asing.\n• Volume penjualan sangat defensif terhadap siklus ekonomi berkat harga produk terjangkau.'
    ],
    [
      'PGAS',
      'PT Perusahaan Gas Negara Tbk',
      'Energi',
      'Transmisi & Distribusi Gas Bumi',
      'Subholding gas Pertamina yang menguasai lebih dari 90% infrastruktur pipa transmisi dan distribusi gas bumi nasional untuk memasok sektor industri dan pembangkit listrik.',
      '• Monopoli alami pada jaringan pipa gas bumi nasional sepanjang ribuan kilometer.\n• Peran krusial dalam program transisi energi nasional sebagai penyedia energi gas ramah lingkungan.\n• Arus kas operasional stabil dari bisnis transmisi dan niaga gas berorientasi industri.'
    ],
    [
      'PGEO',
      'PT Pertamina Geothermal Energy Tbk',
      'Energi',
      'Pembangkit Listrik Panas Bumi (Geotermal)',
      'Subholding energi baru terbarukan Pertamina yang mengelola wilayah kerja panas bumi terbesar di Indonesia dengan kapasitas terpasang ribuan megawatt.',
      '• Pemain murni energi hijau terbarukan dengan kontrak penjualan listrik jangka panjang berdenominasi USD ke PLN.\n• Potensi cadangan panas bumi terbesar di dunia yang terletak di sepanjang ring-of-fire Indonesia.\n• Marjin EBITDA sangat tebal (>75%) dan dukungan penuh pemerintah dalam agenda dekarbonisasi.'
    ],
    [
      'PNLF',
      'PT Panin Financial Tbk',
      'Keuangan',
      'Jasa Asuransi Jiwa & Investasi Keuangan',
      'Perusahaan holding jasa keuangan grup Panin yang bergerak di bidang asuransi jiwa dan merupakan pemegang saham pengendali di PT Bank Pan Indonesia Tbk (PNBN).',
      '• Kepemilikan saham signifikan di Bank Panin (PNBN) menyumbang kontribusi nilai buku dan dividen solid.\n• Tingkat solvabilitas (RBC) asuransi sangat kuat jauh melampaui ketentuan regulasi minimum.\n• Valuasi harga saham kerap diperdagangkan pada diskon yang dalam terhadap nilai aktiva bersih (NAV).'
    ],
    [
      'POWR',
      'PT Cikarang Listrindo Tbk',
      'Utilitas',
      'Penyedia Listrik Kawasan Industri (IPP)',
      'Penyedia tenaga listrik swasta terlama di Indonesia yang melayani ribuan pelanggan industri di lima kawasan industri besar Cikarang serta memasok listrik ke PLN.',
      '• Kontrak penyediaan listrik eksklusif dengan pelanggan kawasan industri tanpa risiko gagal bayar berarti.\n• Rekam jejak keandalan operasional pabrik yang prima dengan ketersediaan listrik 99%.\n• Kebijakan dividen tunai berimbal hasil tinggi (high dividend yield) secara berkala.'
    ],
    [
      'PTBA',
      'PT Bukit Asam Tbk',
      'Energi',
      'Pertambangan Batubara BUMN',
      'Anggota holding tambang BUMN MIND ID yang memproduksi batubara termal dengan cadangan terbesar di Indonesia di Tanjung Enim, Sumatra Selatan.',
      '• Cadangan batubara melimpah dengan jalur pengangkutan kereta api terdedikasi bersama PT KAI.\n• Biaya penambangan tunai rendah dan porsi pasokan jangka panjang untuk ketahanan listrik PLN (DMO).\n• Rekam jejak pembagian dividen bernilai jumbo dengan payout ratio historis hingga 100%.'
    ],
    [
      'PWON',
      'PT Pakuwon Jati Tbk',
      'Properti & Real Estat',
      'Pengembangan Properti Superblok & Mal',
      'Raja mal dan superblok di Indonesia yang mengelola portofolio pusat perbelanjaan tersukses di Jakarta dan Surabaya (Gandaria City, Kota Kasablanka, Tunjungan Plaza).',
      '• Porsi pendapatan berulang (recurring income) tertinggi di industri properti Indonesia (>70%).\n• Tingkat okupansi mal konsisten di atas 90% dengan antrean tenant ritel terkemuka.\n• Neraca keuangan sangat sehat dengan posisi kas kuat dan rasio utang konservatif.'
    ],
    [
      'RAJA',
      'PT Rukun Raharja Tbk',
      'Energi',
      'Infrastruktur & Perdagangan Gas Bumi',
      'Penyedia infrastruktur energi terintegrasi yang bergerak di bidang transmisi pipa gas bumi, perdagangan gas, pengolahan gas, serta participating interest di Blok Rokan.',
      '• Kepemilikan participating interest di blok minyak terbesar Indonesia (Blok Rokan) menyumbang arus kas stabil.\n• Jaringan pipa transmisi gas bumi yang melayani pembangkit listrik dan kawasan industri utama.\n• Diversifikasi ke fasilitas pengolahan air bersih dan kompresi gas untuk industri.'
    ],
    [
      'SCMA',
      'PT Surya Citra Media Tbk',
      'Komunikasi',
      'Media Penyiaran & Platform Streaming',
      'Anak usaha EMTK yang menaungi stasiun televisi terpopuler SCTV dan Indosiar, rumah produksi SinemArt, serta platform streaming nomor satu di Indonesia, Vidio.',
      '• Penguasa pangsa pemirsa (audience share) dan kue iklan televisi nasional melalui SCTV dan Indosiar.\n• Vidio memimpin pasar streaming OTT domestik dengan subscriber berbayar tertinggi berkat konten olahraga eksklusif.\n• Integrasi produksi konten in-house yang kuat menghasilkan efisiensi biaya pemrograman media.'
    ],
    [
      'SIDO',
      'PT Industri Jamu dan Farmasi Sido Muncul Tbk',
      'Kesehatan',
      'Jamu Herbal & Suplemen Kesehatan',
      'Produsen obat herbal dan suplemen terkemuka di Indonesia dengan produk legendaris Tolak Angin, Kuku Bima, serta berbagai produk herbal inovatif berstandar farmasi.',
      '• Kekuatan brand Tolak Angin memiliki loyalitas konsumen yang sangat kuat (market leader tak tergoyahkan).\n• Efisiensi pabrik ekstraksi modern menghasilkan marjin laba kotor (>50%) dan ROE superior (>30%).\n• Neraca keuangan tanpa utang bank (zero debt), posisi kas berlimpah, dan dividen payout ratio konsisten di atas 85%.'
    ],
    [
      'SMRA',
      'PT Summarecon Agung Tbk',
      'Properti & Real Estat',
      'Pengembangan Kota Mandiri & Mal Ritel',
      'Pengembang kota mandiri terkemuka yang sukses membangun kawasan Summarecon Kelapa Gading, Serpong, Bekasi, Bandung, Karawang, Bogor, dan Makassar.',
      '• Kemampuan eksekusi terbukti dalam menciptakan ekosistem kota mandiri terpadu yang bernilai tinggi.\n• Pertumbuhan pendapatan berulang yang solid dari jaringan Summarecon Mall yang ramai pengunjung.\n• Permintaan residensial yang tetap kuat berkat reputasi kualitas bangunan dan fasilitas kawasan yang lengkap.'
    ],
    [
      'SSMS',
      'PT Sawit Sumbermas Sarana Tbk',
      'Konsumer Non-Siklikal',
      'Perkebunan & Pengolahan Kelapa Sawit',
      'Produsen minyak kelapa sawit yang beroperasi di Kalimantan Tengah dengan produktivitas kebun berstandar tinggi serta integrasi kilang refinery hilir melalui Citra Borneo Utama.',
      '• Profil perkebunan muda dengan hasil panen tandan buah segar (TBS) per hektar di atas rata-rata industri.\n• Seluruh areal perkebunan telah mengantongi sertifikasi keberlanjutan sawit lestari (RSPO).\n• Integrasi dengan pabrik pengolahan hilir (refinery) untuk menangkap marjin minyak goreng dan produk turunan sawit.'
    ],
    [
      'TAPG',
      'PT Triputra Agro Persada Tbk',
      'Konsumer Non-Siklikal',
      'Perkebunan Kelapa Sawit & Karet',
      'Produsen CPO terkemuka di bawah grup Triputra yang memiliki salah satu profil usia tanaman kebun paling prima dan efisiensi mekanisasi panen modern.',
      '• Produktivitas kebun sawit (yield CPO/ha) tertinggi di antara emiten perkebunan terbuka di BEI.\n• Usia rata-rata pohon sawit berada pada masa produksi puncak (prime production age).\n• Pengendalian biaya produksi tunai yang ketat menjaga margin keuntungan tetap tebal di berbagai siklus harga sawit.'
    ],
    [
      'TLKM',
      'PT Telkom Indonesia (Persero) Tbk',
      'Komunikasi',
      'Telekomunikasi Digital & Jaringan Terintegrasi',
      'Badan Usaha Milik Negara operator telekomunikasi terbesar di Indonesia yang memimpin pangsa pasar seluler (Telkomsel) dan internet pita lebar rumah (IndiHome).',
      '• Infrastruktur jaringan kabel serat optik dan menara telekomunikasi terluas dan terlengkap di seluruh Nusantara.\n• Strategi Five Bold Moves yang memisahkan infrastruktur (InfraCo) dan data center untuk membuka nilai tersembunyi.\n• Arus kas bebas yang masif dan komitmen pembagian dividen bernilai besar kepada pemegang saham.'
    ],
    [
      'TOWR',
      'PT Sarana Menara Nusantara Tbk',
      'Infrastruktur',
      'Menara Telekomunikasi & Jaringan Fiber Optik',
      'Perusahaan holding infrastruktur telekomunikasi di bawah grup Djarum yang memiliki puluhan ribu menara telekomunikasi dan jaringan kabel serat optik terluas di Indonesia.',
      '• Jaringan menara dan kabel fiber optik (FTTT) terintegrasi dengan kontrak sewa jangka panjang dari operator telekomunikasi papan atas.\n• Manajemen utang berbiaya rendah dan diversifikasi ke layanan konektivitas enterprise.\n• Arus kas berulang yang sangat prediktibel dan defensif terhadap ketidakpastian makroekonomi.'
    ],
    [
      'UNTR',
      'PT United Tractors Tbk',
      'Perindustrian',
      'Alat Berat, Kontraktor Tambang & Mineral',
      'Anak usaha Astra International yang merajai distribusi alat berat (Komatsu), kontraktor penambangan terbesar (PAMA), serta diversifikasi agresif ke tambang emas dan nikel.',
      '• Penguasa mutlak pasar alat berat di Indonesia (>50% market share) dan kontraktor tambang dengan volume overburden terbesar.\n• Arus kas melimpah yang diinvestasikan ke aset masa depan seperti tambang emas Martabe dan smelter nikel.\n• Neraca keuangan sangat kuat dengan sejarah pembagian dividen bernilai sangat besar (jumbo dividend).'
    ],
    [
      'UNVR',
      'PT Unilever Indonesia Tbk',
      'Konsumer Non-Siklikal',
      'Barang Konsumen Cepat Habis (FMCG)',
      'Perusahaan FMCG terkemuka di Indonesia yang memproduksi puluhan merek legendaris di kategori perawatan tubuh (Lifebuoy, Pepsodent, Sunsilk) dan makanan (Bango, Royco).',
      '• Penetrasi produk ke hampir 100% rumah tangga di Indonesia dengan saluran distribusi grosir dan ritel terdalam.\n• Portofolio merek-merek unggulan yang memimpin pangsa pasar di berbagai kategori produk kebutuhan harian.\n• Transformasi efisiensi saluran distribusi dan inovasi produk premium untuk mendorong kembali margin pertumbuhan.'
    ],
    [
      'WIFI',
      'PT Solusi Sinergi Digital Tbk',
      'Teknologi',
      'Konektivitas Internet & Ekosistem Digital (Surge)',
      'Penyedia infrastruktur jaringan serat optik terintegrasi di sepanjang jalur rel kereta api Pulau Jawa yang menghadirkan layanan internet terjangkau dan periklanan digital.',
      '• Jalur kabel optik eksklusif di sepanjang rel kereta api Pulau Jawa dengan latensi sangat rendah dan keamanan tinggi.\n• Peluang monetisasi dari penyediaan internet broadband murah (Affordable Internet) untuk jutaan rumah tangga di kota lapis 2 dan 3.\n• Pertumbuhan pesat segmen sewa bandwidth (core leasing) untuk operator telekomunikasi dan data center.'
    ]
  ];

  // Tulis Header & Styling
  sheet.clearContents();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(1, 1, 1, header.length)
       .setFontWeight('bold')
       .setBackground('#0f2240')
       .setFontColor('#ffffff')
       .setHorizontalAlignment('center');

  // Tulis Data 69 Emiten
  sheet.getRange(2, 1, data.length, header.length).setValues(data);
  sheet.getRange(2, 1, data.length, header.length).setVerticalAlignment('top');

  // Format Kolom
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 80);   // Kode
  sheet.setColumnWidth(2, 260);  // Nama
  sheet.setColumnWidth(3, 160);  // Sektor
  sheet.setColumnWidth(4, 220);  // Subsektor
  sheet.setColumnWidth(5, 500);  // Ringkasan Bisnis
  sheet.setColumnWidth(6, 450);  // Highlight / Keunggulan
  sheet.getRange(2, 5, data.length, 2).setWrap(true);

  SpreadsheetApp.flush();
  Logger.log('Berhasil menambahkan data ringkasan untuk ' + data.length + ' emiten.');
}
