import React, { useState, useEffect, useMemo, useCallback } from "react";
import {
  LayoutDashboard, Users, Flame, CalendarClock, MessageSquareText,
  RotateCcw, BarChart3, GitBranch, History, Settings as SettingsIcon,
  Search, Plus, Pencil, Trash2, X, Copy, Check, Upload, Download,
  ChevronDown, ChevronRight, AlertTriangle, Phone, MapPin, Car,
  TrendingUp, TrendingDown, Minus, Filter, ArrowUpDown, Sparkles,
  ClipboardList, Clock, CheckCircle2, XCircle, FileSpreadsheet, Menu,
} from "lucide-react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid,
  Tooltip, Cell, PieChart, Pie, Legend,
} from "recharts";

/* ============================================================
   CONSTANTS
   ============================================================ */

const SALES_LIST = ["Andi Pratama", "Siti Nurhaliza", "Budi Santoso", "Rina Wulandari", "Dedi Kurniawan"];

const SUMBER_LIST = [
  "Walk In", "Facebook", "Instagram", "TikTok", "WhatsApp", "Website",
  "Event", "Pameran", "Referral", "Database Lama", "Trade In", "Customer Existing",
];

const UNIT_LIST = [
  "Avanza", "Veloz", "Innova Zenix", "Rush", "Raize", "Agya", "Calya",
  "Yaris Cross", "Fortuner", "Hilux", "Rangga",
];

const KOTA_LIST = ["Baubau", "Wolio", "Betoambari", "Kendari", "Muna", "Buton Selatan", "Wangi-Wangi", "Bombana", "Raha"];

const STATUS_LIST = [
  "NEW LEAD", "BELUM DIHUBUNGI", "SUDAH DIHUBUNGI", "RESPON", "PRESENTASI",
  "TEST DRIVE", "PENAWARAN", "NEGOSIASI", "PENGAJUAN KREDIT", "APPROVED",
  "SPK", "DELIVERY", "FOLLOW UP LAGI", "LOST",
];

const FUNNEL_STAGES = ["LEAD", "CONTACTED", "RESPON", "PRESENTASI", "TEST DRIVE", "NEGOSIASI", "SPK", "DELIVERY"];

const STATUS_TO_FUNNEL = {
  "NEW LEAD": 0, "BELUM DIHUBUNGI": 0, "SUDAH DIHUBUNGI": 1, "RESPON": 2,
  "PRESENTASI": 3, "PENAWARAN": 3, "TEST DRIVE": 4, "NEGOSIASI": 5,
  "PENGAJUAN KREDIT": 5, "APPROVED": 5, "SPK": 6, "DELIVERY": 7,
  "FOLLOW UP LAGI": 2, "LOST": -1,
};

const ALASAN_LIST = [
  "Masih pikir-pikir", "Angsuran terlalu besar", "Masih bandingkan merek lain",
  "Menunggu gaji", "Menunggu bonus", "Menunggu persetujuan pasangan",
  "Menunda pembelian", "Belum ada budget DP", "Tidak merespon", "Lainnya",
];

const TUJUAN_LIST = [
  { id: "harga", label: "Follow Up Harga" },
  { id: "simulasi", label: "Follow Up Simulasi Kredit" },
  { id: "tradein", label: "Follow Up Trade In" },
  { id: "testdrive", label: "Follow Up Test Drive" },
  { id: "pameran", label: "Follow Up Setelah Pameran" },
  { id: "kirimharga", label: "Follow Up Setelah Kirim Harga" },
  { id: "noresponse", label: "Customer Tidak Respon" },
  { id: "mahal", label: "Customer Bilang Mahal" },
  { id: "bandingkan", label: "Customer Masih Bandingkan" },
  { id: "gaji", label: "Customer Menunggu Gaji" },
  { id: "bonus", label: "Customer Menunggu Bonus" },
  { id: "pasangan", label: "Customer Menunggu Persetujuan Pasangan" },
  { id: "tunda", label: "Customer Menunda Pembelian" },
  { id: "lama", label: "Follow Up Customer Lama" },
  { id: "closing", label: "Closing Follow Up" },
  { id: "spk", label: "Follow Up SPK" },
  { id: "kreditapproved", label: "Follow Up Setelah Kredit Approved" },
];

const GAYA_BAHASA = ["Hangat & Personal", "Sopan Netral", "Santai"];

/* ============================================================
   HELPERS
   ============================================================ */

const todayISO = () => new Date().toISOString().slice(0, 10);

function daysBetween(dateStr) {
  if (!dateStr) return 999;
  const d = new Date(dateStr);
  const today = new Date();
  d.setHours(0, 0, 0, 0);
  today.setHours(0, 0, 0, 0);
  return Math.round((today - d) / (1000 * 60 * 60 * 24));
}

function formatDate(dateStr) {
  if (!dateStr) return "-";
  const d = new Date(dateStr);
  return d.toLocaleDateString("id-ID", { day: "2-digit", month: "short", year: "numeric" });
}

function calculateScore(p) {
  let score = 0;
  if (p.sudahMerespon) score += 20;
  if (p.mintaHarga) score += 15;
  if (p.mintaSimulasiKredit) score += 15;
  if (p.mintaTestDrive) score += 15;
  if (p.sudahKirimDokumen) score += 10;
  if (p.statusTradeIn === "Ya") score += 10;
  if (p.unit) score += 10;
  if (p.sebutkanTarget) score += 5;

  const days = daysBetween(p.lastFollowUp);
  if (days > 7) score -= 15;
  else if (days > 3) score -= 10;

  if (p.belumBerminat) score -= 20;
  if (p.tundaLama) score -= 20;

  return Math.max(0, Math.min(100, score));
}

function getSuhu(score) {
  if (score >= 70) return "HOT";
  if (score >= 40) return "WARM";
  return "COLD";
}

function suhuColor(suhu) {
  if (suhu === "HOT") return { bg: "#FEE2E2", text: "#B91C3C", dot: "#DC2626" };
  if (suhu === "WARM") return { bg: "#FEF3C7", text: "#B45309", dot: "#D97706" };
  return { bg: "#E2E8F0", text: "#475569", dot: "#64748B" };
}

function getFollowUpFlag(p) {
  const next = p.nextFollowUp;
  if (!next) return "NO_FOLLOW_UP";
  const days = daysBetween(next);
  if (days > 0) return "OVERDUE";
  if (days === 0) return "DUE_TODAY";
  return "UPCOMING";
}

function getPriority(p) {
  const score = calculateScore(p);
  const suhu = getSuhu(score);
  const flag = getFollowUpFlag(p);
  const daysSinceFU = daysBetween(p.lastFollowUp);

  if (suhu === "HOT" && daysSinceFU >= 3) return "SEKARANG";
  if ((flag === "OVERDUE" || flag === "DUE_TODAY") && (suhu === "HOT" || suhu === "WARM")) return "HARI_INI";
  if (flag === "UPCOMING") return "TERJADWAL";
  if (suhu === "COLD") return "LOW";
  return "TERJADWAL";
}

const PRIORITY_META = {
  SEKARANG: { label: "Follow Up Sekarang", icon: "🔴", color: "#DC2626" },
  HARI_INI: { label: "Follow Up Hari Ini", icon: "🟠", color: "#D97706" },
  TERJADWAL: { label: "Follow Up Terjadwal", icon: "🟡", color: "#CA8A04" },
  LOW: { label: "Low Priority", icon: "⚪", color: "#94A3B8" },
};

const FLAG_META = {
  OVERDUE: { label: "Overdue", color: "#DC2626", bg: "#FEE2E2" },
  DUE_TODAY: { label: "Due Today", color: "#D97706", bg: "#FEF3C7" },
  UPCOMING: { label: "Upcoming", color: "#2563EB", bg: "#DBEAFE" },
  NO_FOLLOW_UP: { label: "No Follow Up", color: "#64748B", bg: "#E2E8F0" },
};

function isActiveFunnel(p) {
  return p.statusProspek !== "LOST";
}

/* ---------- Smart follow up analysis (section 9) ---------- */
const SMART_ANALYSIS = {
  "Masih pikir-pikir": {
    analisis: "Customer belum menolak. Masih ada kemungkinan membeli, hanya butuh waktu dan keyakinan tambahan.",
    strategi: "Jangan langsung menekan customer untuk SPK. Beri ruang, tapi tetap hadir dengan informasi yang membantu.",
    nextAction: "Follow up dengan memberikan pilihan atau solusi tambahan, bukan sekadar menanyakan kepastian.",
  },
  "Angsuran terlalu besar": {
    analisis: "Hambatan utama ada di budget cicilan, bukan di minat terhadap unit.",
    strategi: "Jangan hanya memberikan diskon. Cari alternatif unit, DP, tenor, atau skema trade in yang lebih ringan.",
    nextAction: "Siapkan 2-3 alternatif simulasi kredit dengan variasi DP dan tenor.",
  },
  "Masih bandingkan merek lain": {
    analisis: "Customer sedang melakukan perbandingan (comparison shopping), wajar dan bukan tanda tidak berminat.",
    strategi: "Jangan menjelekkan kompetitor. Tonjolkan value Toyota yang relevan dengan kebutuhan customer.",
    nextAction: "Kirim perbandingan value (bukan harga saja) yang sesuai kebutuhan customer.",
  },
  "Menunggu gaji": {
    analisis: "Minat sudah ada, hanya menunggu ketersediaan dana pada waktu tertentu.",
    strategi: "Jaga komunikasi tetap hangat tanpa terkesan menagih. Sesuaikan waktu follow up dengan estimasi gajian.",
    nextAction: "Jadwalkan follow up mendekati tanggal gajian yang disebutkan customer.",
  },
  "Menunggu bonus": {
    analisis: "Sama seperti menunggu gaji, keputusan tertunda karena faktor waktu ketersediaan dana.",
    strategi: "Follow up ringan secara berkala, hindari kesan terlalu sering menagih kepastian.",
    nextAction: "Follow up kembali mendekati periode bonus yang disebutkan, sertakan info promo bila relevan.",
  },
  "Menunggu persetujuan pasangan": {
    analisis: "Keputusan bersifat dua pihak. Customer sendiri sudah cenderung setuju.",
    strategi: "Bantu customer meyakinkan pasangannya dengan materi yang mudah dibagikan (simulasi, foto unit, promo).",
    nextAction: "Tawarkan bantuan berupa ringkasan simulasi yang bisa diteruskan ke pasangan customer.",
  },
  "Menunda pembelian": {
    analisis: "Keputusan pembelian ditunda untuk jangka waktu tertentu, kemungkinan closing masih ada di masa depan.",
    strategi: "Jangan terlalu sering follow up dalam waktu dekat. Gunakan pendekatan reactivation berkala.",
    nextAction: "Jadwalkan follow up reactivation dan berikan alasan baru untuk kembali berkomunikasi.",
  },
  "Belum ada budget DP": {
    analisis: "Hambatan ada di kesiapan dana awal (DP), bukan minat terhadap produk.",
    strategi: "Tawarkan skema DP ringan, program trade in, atau kerja sama leasing dengan DP lebih rendah.",
    nextAction: "Siapkan simulasi dengan opsi DP minimum dan opsi trade in bila customer punya mobil lama.",
  },
  "Tidak merespon": {
    analisis: "Customer tidak memberikan respon dalam beberapa waktu terakhir, kemungkinan prioritas sedang rendah.",
    strategi: "Gunakan pesan reactivation yang ringan, jangan langsung menawarkan promo atau menekan kepastian.",
    nextAction: "Kirim pesan singkat dengan sudut pandang baru, bukan mengulang pesan follow up sebelumnya.",
  },
  Lainnya: {
    analisis: "Kendala customer bersifat spesifik dan perlu digali lebih lanjut melalui percakapan langsung.",
    strategi: "Dengarkan kebutuhan customer secara aktif sebelum menawarkan solusi apa pun.",
    nextAction: "Follow up dengan pertanyaan terbuka untuk memahami kendala sebenarnya.",
  },
};

function getNextBestAction(p) {
  const score = calculateScore(p);
  const suhu = getSuhu(score);
  const daysSinceFU = daysBetween(p.lastFollowUp);
  const flag = getFollowUpFlag(p);

  if (p.statusProspek === "TEST DRIVE") return "Customer sudah test drive. Prioritaskan follow up ke arah closing/SPK hari ini.";
  if (p.statusProspek === "NEGOSIASI") return "Customer sudah negosiasi namun belum closing. Tawarkan solusi konkret untuk menutup transaksi.";
  if (p.statusProspek === "APPROVED") return "Kredit sudah approved. Segera follow up untuk penandatanganan SPK.";
  if (daysSinceFU >= 7) return "Customer tidak merespon selama lebih dari 7 hari. Gunakan pesan reactivation, jangan langsung menawarkan promo.";
  if (suhu === "HOT" && flag === "OVERDUE") return "Prospek HOT sudah overdue. Hubungi customer hari ini dan tawarkan simulasi/alternatif yang relevan.";
  if (p.mintaSimulasiKredit) return "Customer meminta simulasi kredit. Hubungi hari ini dan tawarkan simulasi alternatif dengan DP lebih rendah.";
  if (suhu === "WARM") return "Jaga momentum dengan memberi informasi atau value tambahan agar naik ke level HOT.";
  return "Follow up ringan untuk menjaga komunikasi tetap terbuka tanpa terkesan menekan.";
}

/* ---------- WhatsApp message generator (section 7-8, 23) ---------- */

const OPENERS = {
  "Hangat & Personal": (n) => `${n ? "Pak/Bu " + n.split(" ")[0] : "Bapak/Ibu"}, izin follow up sedikit ya.`,
  "Sopan Netral": (n) => `Selamat siang ${n ? "Bapak/Ibu " + n.split(" ")[0] : "Bapak/Ibu"}, mohon izin menghubungi kembali.`,
  Santai: (n) => `Halo ${n ? "Pak/Bu " + n.split(" ")[0] : "Pak/Bu"}, boleh follow up sebentar ya.`,
};

function goalContext(tujuanId, ctx) {
  const unit = ctx.unit || "unit yang diminati";
  const map = {
    harga: `Kemarin ${ctx.nama || "Bapak/Ibu"} sempat menanyakan harga ${unit}.`,
    simulasi: `Kemarin sempat dibahas simulasi kredit untuk ${unit}.`,
    tradein: `Sebelumnya ${ctx.nama || "Bapak/Ibu"} menyebutkan ada rencana trade in mobil lama untuk ${unit}.`,
    testdrive: `Terima kasih sudah menyempatkan waktu test drive ${unit} kemarin.`,
    pameran: `Senang bisa berkenalan saat pameran kemarin, khususnya soal ${unit}.`,
    kirimharga: `Sudah saya kirimkan info harga ${unit} beberapa waktu lalu.`,
    noresponse: `Sudah beberapa waktu kita belum ngobrol lagi soal ${unit}.`,
    mahal: `Saya mengerti soal angsuran ${unit} yang kemarin dirasa cukup berat.`,
    bandingkan: `Saya paham ${ctx.nama || "Bapak/Ibu"} masih mempertimbangkan beberapa pilihan selain ${unit}.`,
    gaji: `Kemarin ${ctx.nama || "Bapak/Ibu"} sempat menyebutkan akan mempertimbangkan ${unit} setelah gajian.`,
    bonus: `Kemarin sempat dibahas rencana ${unit} setelah bonus cair.`,
    pasangan: `Kemarin ${ctx.nama || "Bapak/Ibu"} menyebutkan perlu diskusi dulu bersama pasangan soal ${unit}.`,
    tunda: `Saya mengerti rencana ${unit} sempat ditunda beberapa waktu.`,
    lama: `Sudah cukup lama kita tidak mengobrol lagi sejak membahas ${unit}.`,
    closing: `Melanjutkan pembahasan kita soal ${unit} kemarin.`,
    spk: `Semua sudah hampir siap untuk proses SPK ${unit}.`,
    kreditapproved: `Kabar baik, pengajuan kredit untuk ${unit} sudah approved.`,
  };
  return map[tujuanId] || `Melanjutkan pembahasan kita soal ${unit}.`;
}

function goalValue(tujuanId, ctx) {
  const map = {
    harga: "Supaya ada gambaran yang lebih jelas, saya bisa bantu breakdown harga on the road-nya.",
    simulasi: "Saya bisa bantu hitungkan beberapa opsi tenor supaya cicilannya lebih nyaman di kantong.",
    tradein: "Saya bisa bantu cek estimasi harga mobil lama supaya bisa mengurangi beban DP.",
    testdrive: "Kalau ada hal yang masih ingin ditanyakan soal performa atau fitur, saya bantu jelaskan.",
    kirimharga: "Kalau ada bagian yang masih perlu didiskusikan dari penawaran kemarin, saya siap bantu.",
    noresponse: "Kalau ada kendala atau pertanyaan yang tertunda, saya bantu carikan solusinya.",
    mahal: "Ada beberapa opsi DP dan tenor lain yang mungkin lebih sesuai dengan budget bulanan.",
    bandingkan: "Saya bisa bantu bandingkan fitur dan value yang paling relevan dengan kebutuhan sehari-hari.",
    gaji: "Kalau nanti waktunya sudah pas, saya siapkan dulu simulasinya supaya prosesnya lebih cepat.",
    bonus: "Saya siapkan dulu simulasinya supaya begitu waktunya tiba, prosesnya bisa langsung jalan.",
    pasangan: "Saya bisa siapkan ringkasan simulasi yang mudah dibagikan untuk didiskusikan berdua.",
    tunda: "Kalau nanti sudah siap kembali, saya bantu update info promo atau harga terbaru.",
    lama: "Ada beberapa update model dan promo terbaru yang mungkin relevan untuk kebutuhan sekarang.",
    closing: "Semua dokumen dan simulasi sudah bisa saya siapkan supaya prosesnya lebih cepat.",
    spk: "Saya bantu siapkan seluruh berkas supaya proses SPK bisa berjalan lancar.",
    kreditapproved: "Tinggal satu langkah lagi menuju unit barunya siap dipakai.",
  };
  return map[tujuanId] || "Saya siap bantu prosesnya supaya lebih mudah.";
}

const CTA_PERSONAL = [
  "Kalau Bapak/Ibu berkenan, saya bantu hitungkan simulasinya.",
  "Kalau berkenan, saya kirimkan infonya ya Pak/Bu.",
  "Kalau ada waktu, boleh saya telepon sebentar untuk jelaskan lebih detail?",
];
const CTA_VALUE = [
  "Saya bantu cekkan dulu ya Pak/Bu, supaya ada gambaran yang lebih jelas.",
  "Saya siapkan dulu opsinya, nanti tinggal Bapak/Ibu pilih yang paling nyaman.",
  "Boleh saya kirimkan perbandingannya supaya lebih mudah dipertimbangkan?",
];
const CTA_CLOSING = [
  "Kalau semua sudah sesuai, saya bantu siapkan proses selanjutnya ya Pak/Bu.",
  "Kalau berkenan, kita jadwalkan waktu untuk lanjut ke proses berikutnya.",
  "Saya siap bantu prosesnya kapan pun Bapak/Ibu sudah siap.",
];

function pick(arr, seed) {
  return arr[seed % arr.length];
}

function generateMessages(ctx) {
  const opener = (OPENERS[ctx.gaya] || OPENERS["Hangat & Personal"])(ctx.nama);
  const context = goalContext(ctx.tujuan, ctx);
  const value = goalValue(ctx.tujuan, ctx);
  const seed = (ctx.nama ? ctx.nama.length : 0) + (ctx.unit ? ctx.unit.length : 0);

  const A = `${opener} ${context} ${pick(CTA_PERSONAL, seed)}`;
  const B = `${opener} ${context} ${value} ${pick(CTA_VALUE, seed + 1)}`;
  const C = `${opener} ${context} ${value} ${pick(CTA_CLOSING, seed + 2)}`;

  return {
    A: { label: "Personal & Ramah", text: A },
    B: { label: "Value Selling", text: B },
    C: { label: "Closing", text: C },
  };
}

/* ============================================================
   DEMO DATA
   ============================================================ */

const NAMES = [
  "Muh. Ridwan", "Sartika Amalia", "La Ode Arman", "Nur Aisyah", "Hasanuddin",
  "Wa Ode Fitri", "Irfan Maulana", "Suryani Dewi", "La Ode Fajar", "Endang Susanti",
  "Muhammad Yusuf", "Andi Nurul", "La Ode Bahtiar", "Rahmawati", "Fajar Nugroho",
  "Wa Ode Sartika", "Bayu Segara", "Nining Kurnia", "La Ode Ilham", "Citra Lestari",
];

function makeDemoData() {
  const rows = [];
  for (let i = 0; i < 20; i++) {
    const nama = NAMES[i];
    const sales = SALES_LIST[i % SALES_LIST.length];
    const unit = UNIT_LIST[i % UNIT_LIST.length];
    const sumber = SUMBER_LIST[i % SUMBER_LIST.length];
    const kota = KOTA_LIST[i % KOTA_LIST.length];
    const statusIdx = [0, 1, 2, 3, 4, 5, 3, 6, 7, 8, 9, 10, 11, 2, 1, 6, 13, 0, 5, 12][i];
    const status = STATUS_LIST[statusIdx];
    const lastFU = new Date();
    lastFU.setDate(lastFU.getDate() - [1, 2, 8, 0, 4, 10, 3, 1, 6, 15, 0, 20, 2, 9, 5, 1, 30, 3, 0, 12][i]);
    const nextFU = new Date();
    nextFU.setDate(nextFU.getDate() + [-2, 0, -5, 1, 2, -8, 0, 3, -1, -10, 1, -15, 4, -3, 2, 0, -20, 1, 5, -6][i]);
    const masuk = new Date();
    masuk.setDate(masuk.getDate() - (10 + i * 3));

    const p = {
      id: `PSP-${String(1000 + i)}`,
      tanggalMasuk: masuk.toISOString().slice(0, 10),
      nama,
      whatsapp: `08${(1000000000 + i * 7654321).toString().slice(0, 10)}`,
      kota,
      sales,
      sumber,
      unit,
      tahunMobilLama: i % 3 === 0 ? String(2015 + (i % 6)) : "",
      merekMobilLama: i % 3 === 0 ? ["Toyota Avanza", "Honda Jazz", "Daihatsu Xenia", "Suzuki Ertiga"][i % 4] : "",
      estimasiHargaMobilLama: i % 3 === 0 ? String(80000000 + i * 3500000) : "",
      statusTradeIn: i % 3 === 0 ? "Ya" : "Tidak",
      budgetDP: String(20000000 + (i % 6) * 8000000),
      targetAngsuran: String(2500000 + (i % 5) * 700000),
      statusKredit: ["Belum Diajukan", "Diajukan", "Approved", "Ditolak"][i % 4],
      statusProspek: status,
      lastFollowUp: lastFU.toISOString().slice(0, 10),
      nextFollowUp: nextFU.toISOString().slice(0, 10),
      jumlahFollowUp: 1 + (i % 6),
      responTerakhir: [
        "Masih dipertimbangkan", "Minta dikirimkan simulasi", "Tertarik tapi budget terbatas",
        "Menunggu diskusi dengan pasangan", "Bandingkan dengan merek lain", "Belum ada respon",
      ][i % 6],
      alasanBelumClosing: ALASAN_LIST[i % ALASAN_LIST.length],
      potensiSPK: ["Tinggi", "Sedang", "Rendah"][i % 3],
      catatanSales: "Customer kooperatif, perlu follow up rutin.",
      sudahMerespon: i % 5 !== 0,
      mintaHarga: i % 2 === 0,
      mintaSimulasiKredit: i % 3 === 0,
      mintaTestDrive: statusIdx >= 5,
      sudahKirimDokumen: i % 4 === 0,
      sebutkanTarget: i % 3 !== 0,
      belumBerminat: i === 17,
      tundaLama: i === 12,
      history: [
        {
          tanggal: lastFU.toISOString().slice(0, 10), jam: "10:00", media: "WhatsApp",
          isi: "Follow up rutin terkait kebutuhan unit.", respon: "Customer merespon positif.",
          nextAction: "Kirimkan simulasi kredit.", nextDate: nextFU.toISOString().slice(0, 10),
        },
      ],
    };
    rows.push(p);
  }
  return rows;
}

/* ============================================================
   SMALL UI PRIMITIVES
   ============================================================ */

function Badge({ children, bg, color, style }) {
  return (
    <span
      style={{ background: bg, color, ...style }}
      className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md text-xs font-semibold whitespace-nowrap"
    >
      {children}
    </span>
  );
}

function ScoreRing({ score }) {
  const suhu = getSuhu(score);
  const c = suhuColor(suhu);
  const r = 16;
  const circ = 2 * Math.PI * r;
  const offset = circ - (score / 100) * circ;
  return (
    <div className="relative w-11 h-11 shrink-0">
      <svg width="44" height="44" viewBox="0 0 44 44">
        <circle cx="22" cy="22" r={r} fill="none" stroke="#EEF1F5" strokeWidth="4" />
        <circle
          cx="22" cy="22" r={r} fill="none" stroke={c.dot} strokeWidth="4"
          strokeDasharray={circ} strokeDashoffset={offset} strokeLinecap="round"
          transform="rotate(-90 22 22)"
        />
      </svg>
      <div className="absolute inset-0 flex items-center justify-center text-[11px] font-bold" style={{ color: c.text }}>
        {score}
      </div>
    </div>
  );
}

function KPICard({ label, value, sub, accent }) {
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-4 flex flex-col gap-1 min-w-0">
      <div className="text-[11px] font-medium text-slate-500 truncate">{label}</div>
      <div className="text-2xl font-bold tracking-tight" style={{ color: accent || "#12233F", fontFamily: "'Sora', sans-serif" }}>
        {value}
      </div>
      {sub ? <div className="text-[11px] text-slate-400">{sub}</div> : null}
    </div>
  );
}

function EmptyState({ title, desc }) {
  return (
    <div className="flex flex-col items-center justify-center py-16 text-center gap-2">
      <div className="w-12 h-12 rounded-full bg-slate-100 flex items-center justify-center">
        <Search size={20} className="text-slate-400" />
      </div>
      <div className="font-semibold text-slate-700">{title}</div>
      <div className="text-sm text-slate-400 max-w-xs">{desc}</div>
    </div>
  );
}

/* ============================================================
   NAV
   ============================================================ */

const NAV_ITEMS = [
  { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { id: "prospek", label: "Prospek", icon: Users },
  { id: "hot", label: "Prospek HOT", icon: Flame },
  { id: "followup", label: "Follow Up", icon: CalendarClock },
  { id: "generator", label: "Follow Up Generator", icon: MessageSquareText },
  { id: "reactivation", label: "Reactivation", icon: RotateCcw },
  { id: "performance", label: "Performance Sales", icon: BarChart3 },
  { id: "funnel", label: "Funnel", icon: GitBranch },
  { id: "history", label: "History", icon: History },
  { id: "settings", label: "Settings", icon: SettingsIcon },
];

function Sidebar({ active, setActive, mobileOpen, setMobileOpen, alertCount }) {
  return (
    <>
      {mobileOpen && (
        <div className="fixed inset-0 bg-black/40 z-30 lg:hidden" onClick={() => setMobileOpen(false)} />
      )}
      <aside
        className={`fixed lg:static z-40 top-0 left-0 h-full w-64 shrink-0 flex flex-col transition-transform duration-200 ${
          mobileOpen ? "translate-x-0" : "-translate-x-full lg:translate-x-0"
        }`}
        style={{ background: "#12233F" }}
      >
        <div className="px-5 pt-6 pb-5 border-b border-white/10">
          <div className="flex items-center gap-2">
            <div className="w-9 h-9 rounded-lg flex items-center justify-center font-bold text-white" style={{ background: "#B91C3C" }}>
              T
            </div>
            <div>
              <div className="text-white font-bold leading-tight text-[15px]" style={{ fontFamily: "'Sora', sans-serif" }}>
                Prospect Monitoring
              </div>
              <div className="text-[11px] text-white/50">Kalla Toyota Baubau</div>
            </div>
          </div>
        </div>
        <nav className="flex-1 overflow-y-auto py-3 px-2.5 space-y-0.5">
          {NAV_ITEMS.map((item) => {
            const Icon = item.icon;
            const isActive = active === item.id;
            return (
              <button
                key={item.id}
                onClick={() => { setActive(item.id); setMobileOpen(false); }}
                className="w-full flex items-center gap-2.5 px-3 py-2.5 rounded-lg text-[13.5px] font-medium transition-colors relative"
                style={{
                  background: isActive ? "rgba(255,255,255,0.08)" : "transparent",
                  color: isActive ? "#FFFFFF" : "rgba(255,255,255,0.6)",
                }}
              >
                {isActive && <span className="absolute left-0 top-1.5 bottom-1.5 w-[3px] rounded-full" style={{ background: "#B91C3C" }} />}
                <Icon size={17} strokeWidth={2} />
                <span className="truncate">{item.label}</span>
                {item.id === "hot" && alertCount > 0 && (
                  <span className="ml-auto text-[10px] font-bold px-1.5 py-0.5 rounded-full" style={{ background: "#B91C3C", color: "white" }}>
                    {alertCount}
                  </span>
                )}
              </button>
            );
          })}
        </nav>
        <div className="p-4 border-t border-white/10 text-[11px] text-white/40">
          Data tersimpan otomatis di perangkat ini.
        </div>
      </aside>
    </>
  );
}

/* ============================================================
   PROSPECT FORM MODAL
   ============================================================ */

function Field({ label, children }) {
  return (
    <label className="flex flex-col gap-1 text-sm">
      <span className="text-xs font-medium text-slate-500">{label}</span>
      {children}
    </label>
  );
}

const inputCls = "border border-slate-200 rounded-lg px-3 py-2 text-sm outline-none focus:border-[#12233F] focus:ring-1 focus:ring-[#12233F] bg-white";

function ProspectModal({ initial, onClose, onSave }) {
  const [f, setF] = useState(
    initial || {
      id: `PSP-${Math.floor(1000 + Math.random() * 8999)}`,
      tanggalMasuk: todayISO(), nama: "", whatsapp: "", kota: KOTA_LIST[0],
      sales: SALES_LIST[0], sumber: SUMBER_LIST[0], unit: UNIT_LIST[0],
      tahunMobilLama: "", merekMobilLama: "", estimasiHargaMobilLama: "",
      statusTradeIn: "Tidak", budgetDP: "", targetAngsuran: "",
      statusKredit: "Belum Diajukan", statusProspek: "NEW LEAD",
      lastFollowUp: todayISO(), nextFollowUp: "", jumlahFollowUp: 0,
      responTerakhir: "", alasanBelumClosing: ALASAN_LIST[0], potensiSPK: "Sedang",
      catatanSales: "", sudahMerespon: false, mintaHarga: false, mintaSimulasiKredit: false,
      mintaTestDrive: false, sudahKirimDokumen: false, sebutkanTarget: false,
      belumBerminat: false, tundaLama: false, history: [],
    }
  );

  const set = (k, v) => setF((prev) => ({ ...prev, [k]: v }));
  const score = calculateScore(f);
  const suhu = getSuhu(score);
  const c = suhuColor(suhu);

  return (
    <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center p-3">
      <div className="bg-white rounded-2xl w-full max-w-3xl max-h-[90vh] flex flex-col overflow-hidden">
        <div className="flex items-center justify-between px-6 py-4 border-b border-slate-100">
          <div>
            <div className="font-bold text-lg" style={{ fontFamily: "'Sora', sans-serif" }}>{initial ? "Edit Prospek" : "Tambah Prospek"}</div>
            <div className="text-xs text-slate-400">{f.id}</div>
          </div>
          <div className="flex items-center gap-3">
            <div className="flex items-center gap-2">
              <ScoreRing score={score} />
              <Badge bg={c.bg} color={c.text}>{suhu}</Badge>
            </div>
            <button onClick={onClose} className="w-8 h-8 rounded-full hover:bg-slate-100 flex items-center justify-center">
              <X size={18} />
            </button>
          </div>
        </div>

        <div className="overflow-y-auto px-6 py-4 space-y-5">
          <div>
            <div className="text-xs font-bold uppercase tracking-wide text-slate-400 mb-2">Data Customer</div>
            <div className="grid grid-cols-2 gap-3">
              <Field label="Nama Customer"><input className={inputCls} value={f.nama} onChange={(e) => set("nama", e.target.value)} /></Field>
              <Field label="No. WhatsApp"><input className={inputCls} value={f.whatsapp} onChange={(e) => set("whatsapp", e.target.value)} /></Field>
              <Field label="Kota/Kabupaten">
                <select className={inputCls} value={f.kota} onChange={(e) => set("kota", e.target.value)}>
                  {KOTA_LIST.map((k) => <option key={k}>{k}</option>)}
                </select>
              </Field>
              <Field label="Sales">
                <select className={inputCls} value={f.sales} onChange={(e) => set("sales", e.target.value)}>
                  {SALES_LIST.map((k) => <option key={k}>{k}</option>)}
                </select>
              </Field>
              <Field label="Sumber Prospek">
                <select className={inputCls} value={f.sumber} onChange={(e) => set("sumber", e.target.value)}>
                  {SUMBER_LIST.map((k) => <option key={k}>{k}</option>)}
                </select>
              </Field>
              <Field label="Unit yang Diminati">
                <select className={inputCls} value={f.unit} onChange={(e) => set("unit", e.target.value)}>
                  {UNIT_LIST.map((k) => <option key={k}>{k}</option>)}
                </select>
              </Field>
              <Field label="Tanggal Masuk"><input type="date" className={inputCls} value={f.tanggalMasuk} onChange={(e) => set("tanggalMasuk", e.target.value)} /></Field>
            </div>
          </div>

          <div>
            <div className="text-xs font-bold uppercase tracking-wide text-slate-400 mb-2">Trade In & Kredit</div>
            <div className="grid grid-cols-2 gap-3">
              <Field label="Status Trade In">
                <select className={inputCls} value={f.statusTradeIn} onChange={(e) => set("statusTradeIn", e.target.value)}>
                  <option>Ya</option><option>Tidak</option>
                </select>
              </Field>
              <Field label="Status Kredit">
                <select className={inputCls} value={f.statusKredit} onChange={(e) => set("statusKredit", e.target.value)}>
                  <option>Belum Diajukan</option><option>Diajukan</option><option>Approved</option><option>Ditolak</option>
                </select>
              </Field>
              {f.statusTradeIn === "Ya" && (
                <>
                  <Field label="Merek Mobil Lama"><input className={inputCls} value={f.merekMobilLama} onChange={(e) => set("merekMobilLama", e.target.value)} /></Field>
                  <Field label="Tahun Mobil Lama"><input className={inputCls} value={f.tahunMobilLama} onChange={(e) => set("tahunMobilLama", e.target.value)} /></Field>
                  <Field label="Estimasi Harga Mobil Lama"><input className={inputCls} value={f.estimasiHargaMobilLama} onChange={(e) => set("estimasiHargaMobilLama", e.target.value)} /></Field>
                </>
              )}
              <Field label="Budget DP"><input className={inputCls} value={f.budgetDP} onChange={(e) => set("budgetDP", e.target.value)} /></Field>
              <Field label="Target Angsuran"><input className={inputCls} value={f.targetAngsuran} onChange={(e) => set("targetAngsuran", e.target.value)} /></Field>
            </div>
          </div>

          <div>
            <div className="text-xs font-bold uppercase tracking-wide text-slate-400 mb-2">Status & Follow Up</div>
            <div className="grid grid-cols-2 gap-3">
              <Field label="Status Prospek">
                <select className={inputCls} value={f.statusProspek} onChange={(e) => set("statusProspek", e.target.value)}>
                  {STATUS_LIST.map((k) => <option key={k}>{k}</option>)}
                </select>
              </Field>
              <Field label="Potensi SPK">
                <select className={inputCls} value={f.potensiSPK} onChange={(e) => set("potensiSPK", e.target.value)}>
                  <option>Tinggi</option><option>Sedang</option><option>Rendah</option>
                </select>
              </Field>
              <Field label="Last Follow Up"><input type="date" className={inputCls} value={f.lastFollowUp} onChange={(e) => set("lastFollowUp", e.target.value)} /></Field>
              <Field label="Next Follow Up"><input type="date" className={inputCls} value={f.nextFollowUp} onChange={(e) => set("nextFollowUp", e.target.value)} /></Field>
              <Field label="Jumlah Follow Up"><input type="number" className={inputCls} value={f.jumlahFollowUp} onChange={(e) => set("jumlahFollowUp", Number(e.target.value))} /></Field>
              <Field label="Alasan Belum Closing">
                <select className={inputCls} value={f.alasanBelumClosing} onChange={(e) => set("alasanBelumClosing", e.target.value)}>
                  {ALASAN_LIST.map((k) => <option key={k}>{k}</option>)}
                </select>
              </Field>
              <Field label="Respon Terakhir"><input className={inputCls} value={f.responTerakhir} onChange={(e) => set("responTerakhir", e.target.value)} /></Field>
            </div>
            <div className="mt-3"><Field label="Catatan Sales"><textarea rows={2} className={inputCls} value={f.catatanSales} onChange={(e) => set("catatanSales", e.target.value)} /></Field></div>
          </div>

          <div>
            <div className="text-xs font-bold uppercase tracking-wide text-slate-400 mb-2">Faktor Scoring (mempengaruhi suhu prospek)</div>
            <div className="grid grid-cols-2 gap-2 text-sm">
              {[
                ["sudahMerespon", "Sudah memberikan respon (+20)"],
                ["mintaHarga", "Meminta harga (+15)"],
                ["mintaSimulasiKredit", "Meminta simulasi kredit (+15)"],
                ["mintaTestDrive", "Meminta test drive (+15)"],
                ["sudahKirimDokumen", "Sudah kirim dokumen (+10)"],
                ["sebutkanTarget", "Menyebutkan target pembelian (+5)"],
                ["belumBerminat", "Menyatakan belum berminat (-20)"],
                ["tundaLama", "Menunda pembelian > 3 bulan (-20)"],
              ].map(([key, label]) => (
                <label key={key} className="flex items-center gap-2 px-2 py-1.5 rounded-lg hover:bg-slate-50 cursor-pointer">
                  <input type="checkbox" checked={!!f[key]} onChange={(e) => set(key, e.target.checked)} />
                  <span>{label}</span>
                </label>
              ))}
            </div>
          </div>
        </div>

        <div className="flex items-center justify-end gap-2 px-6 py-4 border-t border-slate-100">
          <button onClick={onClose} className="px-4 py-2 rounded-lg text-sm font-medium text-slate-600 hover:bg-slate-100">Batal</button>
          <button
            onClick={() => onSave(f)}
            className="px-4 py-2 rounded-lg text-sm font-semibold text-white"
            style={{ background: "#12233F" }}
            disabled={!f.nama}
          >
            Simpan Prospek
          </button>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   PROSPECT DETAIL DRAWER (history + smart follow up)
   ============================================================ */

function ProspectDetail({ p, onClose, onAddHistory }) {
  const score = calculateScore(p);
  const suhu = getSuhu(score);
  const c = suhuColor(suhu);
  const analysis = SMART_ANALYSIS[p.alasanBelumClosing] || SMART_ANALYSIS.Lainnya;
  const [note, setNote] = useState({ media: "WhatsApp", isi: "", respon: "", nextAction: "", nextDate: "" });

  return (
    <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-end">
      <div className="bg-white w-full max-w-lg h-full overflow-y-auto shadow-2xl">
        <div className="sticky top-0 bg-white px-5 py-4 border-b border-slate-100 flex items-start justify-between z-10">
          <div>
            <div className="font-bold text-lg" style={{ fontFamily: "'Sora', sans-serif" }}>{p.nama}</div>
            <div className="text-xs text-slate-400 flex items-center gap-1"><Phone size={12} />{p.whatsapp} · <MapPin size={12} />{p.kota}</div>
          </div>
          <button onClick={onClose} className="w-8 h-8 rounded-full hover:bg-slate-100 flex items-center justify-center"><X size={18} /></button>
        </div>

        <div className="p-5 space-y-5">
          <div className="flex items-center gap-3">
            <ScoreRing score={score} />
            <div>
              <Badge bg={c.bg} color={c.text}>{suhu} · {score}</Badge>
              <div className="text-xs text-slate-400 mt-1">{p.statusProspek} · <Car size={11} className="inline" /> {p.unit}</div>
            </div>
          </div>

          <div className="grid grid-cols-2 gap-3 text-sm">
            <div><div className="text-xs text-slate-400">Sales</div><div className="font-medium">{p.sales}</div></div>
            <div><div className="text-xs text-slate-400">Sumber</div><div className="font-medium">{p.sumber}</div></div>
            <div><div className="text-xs text-slate-400">Last Follow Up</div><div className="font-medium">{formatDate(p.lastFollowUp)}</div></div>
            <div><div className="text-xs text-slate-400">Next Follow Up</div><div className="font-medium">{formatDate(p.nextFollowUp)}</div></div>
            <div><div className="text-xs text-slate-400">Potensi SPK</div><div className="font-medium">{p.potensiSPK}</div></div>
            <div><div className="text-xs text-slate-400">Jumlah Follow Up</div><div className="font-medium">{p.jumlahFollowUp}x</div></div>
          </div>

          <div className="rounded-xl p-4" style={{ background: "#F8FAFC" }}>
            <div className="flex items-center gap-1.5 text-xs font-bold uppercase tracking-wide text-slate-500 mb-2">
              <Sparkles size={13} /> Smart Follow Up · {p.alasanBelumClosing}
            </div>
            <div className="text-sm space-y-2">
              <div><span className="font-semibold">Analisis: </span>{analysis.analisis}</div>
              <div><span className="font-semibold">Strategi: </span>{analysis.strategi}</div>
              <div><span className="font-semibold">Next Action: </span>{analysis.nextAction}</div>
            </div>
          </div>

          <div className="rounded-xl p-4 border" style={{ borderColor: "#FDE68A", background: "#FFFBEB" }}>
            <div className="flex items-center gap-1.5 text-xs font-bold uppercase tracking-wide text-amber-700 mb-1">
              <TrendingUp size={13} /> Next Best Action
            </div>
            <div className="text-sm text-amber-900">{getNextBestAction(p)}</div>
          </div>

          <div>
            <div className="text-xs font-bold uppercase tracking-wide text-slate-400 mb-2">History Follow Up</div>
            <div className="space-y-2 max-h-48 overflow-y-auto pr-1">
              {(p.history || []).slice().reverse().map((h, i) => (
                <div key={i} className="border border-slate-100 rounded-lg p-2.5 text-sm">
                  <div className="flex justify-between text-xs text-slate-400"><span>{formatDate(h.tanggal)} · {h.jam}</span><span>{h.media}</span></div>
                  <div className="mt-1">{h.isi}</div>
                  <div className="text-xs text-slate-500 mt-1">Respon: {h.respon}</div>
                  {h.nextAction && <div className="text-xs text-slate-500">Next: {h.nextAction} {h.nextDate ? `(${formatDate(h.nextDate)})` : ""}</div>}
                </div>
              ))}
              {(!p.history || p.history.length === 0) && <div className="text-sm text-slate-400">Belum ada history.</div>}
            </div>
          </div>

          <div className="border-t border-slate-100 pt-4">
            <div className="text-xs font-bold uppercase tracking-wide text-slate-400 mb-2">Catat Follow Up Baru</div>
            <div className="grid grid-cols-2 gap-2 mb-2">
              <select className={inputCls} value={note.media} onChange={(e) => setNote({ ...note, media: e.target.value })}>
                <option>WhatsApp</option><option>Telepon</option><option>Datang Langsung</option><option>Email</option>
              </select>
              <input type="date" className={inputCls} value={note.nextDate} onChange={(e) => setNote({ ...note, nextDate: e.target.value })} />
            </div>
            <textarea rows={2} placeholder="Isi percakapan" className={inputCls + " w-full mb-2"} value={note.isi} onChange={(e) => setNote({ ...note, isi: e.target.value })} />
            <textarea rows={1} placeholder="Respon customer" className={inputCls + " w-full mb-2"} value={note.respon} onChange={(e) => setNote({ ...note, respon: e.target.value })} />
            <input placeholder="Next action" className={inputCls + " w-full mb-3"} value={note.nextAction} onChange={(e) => setNote({ ...note, nextAction: e.target.value })} />
            <button
              className="w-full py-2 rounded-lg text-sm font-semibold text-white"
              style={{ background: "#12233F" }}
              onClick={() => {
                if (!note.isi) return;
                onAddHistory(p.id, note, note.nextDate);
                setNote({ media: "WhatsApp", isi: "", respon: "", nextAction: "", nextDate: "" });
              }}
            >
              Simpan Catatan
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   PROSPECT TABLE ROW / LIST
   ============================================================ */

function ProspectRow({ p, onEdit, onDelete, onOpen }) {
  const score = calculateScore(p);
  const suhu = getSuhu(score);
  const c = suhuColor(suhu);
  const flag = getFollowUpFlag(p);
  const fm = FLAG_META[flag];

  return (
    <tr className="border-b border-slate-50 hover:bg-slate-50/70 cursor-pointer" onClick={() => onOpen(p)}>
      <td className="py-2.5 px-3"><ScoreRing score={score} /></td>
      <td className="py-2.5 px-3">
        <div className="font-semibold text-slate-800 text-sm">{p.nama}</div>
        <div className="text-xs text-slate-400">{p.kota} · {p.sales}</div>
      </td>
      <td className="py-2.5 px-3 text-sm text-slate-600">{p.unit}</td>
      <td className="py-2.5 px-3"><Badge bg={c.bg} color={c.text}>{suhu}</Badge></td>
      <td className="py-2.5 px-3 text-xs text-slate-600">{p.statusProspek}</td>
      <td className="py-2.5 px-3"><Badge bg={fm.bg} color={fm.color}>{fm.label}</Badge></td>
      <td className="py-2.5 px-3 text-xs text-slate-500">{formatDate(p.nextFollowUp)}</td>
      <td className="py-2.5 px-3">
        <div className="flex items-center gap-1" onClick={(e) => e.stopPropagation()}>
          <button onClick={() => onEdit(p)} className="w-7 h-7 rounded-md hover:bg-slate-100 flex items-center justify-center text-slate-500"><Pencil size={14} /></button>
          <button onClick={() => onDelete(p.id)} className="w-7 h-7 rounded-md hover:bg-red-50 flex items-center justify-center text-red-500"><Trash2 size={14} /></button>
        </div>
      </td>
    </tr>
  );
}

function FilterBar({ filters, setFilters, onAdd, onImport, onExportCsv, onExportXlsx }) {
  const fileRef = React.useRef();
  return (
    <div className="flex flex-wrap items-center gap-2 mb-4">
      <div className="flex items-center gap-2 bg-white border border-slate-200 rounded-lg px-3 py-2 flex-1 min-w-[180px]">
        <Search size={15} className="text-slate-400" />
        <input
          placeholder="Cari nama, WhatsApp, atau unit..."
          className="text-sm outline-none flex-1"
          value={filters.q}
          onChange={(e) => setFilters({ ...filters, q: e.target.value })}
        />
      </div>
      <select className={inputCls} value={filters.sales} onChange={(e) => setFilters({ ...filters, sales: e.target.value })}>
        <option value="">Semua Sales</option>
        {SALES_LIST.map((s) => <option key={s}>{s}</option>)}
      </select>
      <select className={inputCls} value={filters.suhu} onChange={(e) => setFilters({ ...filters, suhu: e.target.value })}>
        <option value="">Semua Suhu</option>
        <option>HOT</option><option>WARM</option><option>COLD</option>
      </select>
      <select className={inputCls} value={filters.status} onChange={(e) => setFilters({ ...filters, status: e.target.value })}>
        <option value="">Semua Status</option>
        {STATUS_LIST.map((s) => <option key={s}>{s}</option>)}
      </select>
      <select className={inputCls} value={filters.sumber} onChange={(e) => setFilters({ ...filters, sumber: e.target.value })}>
        <option value="">Semua Sumber</option>
        {SUMBER_LIST.map((s) => <option key={s}>{s}</option>)}
      </select>
      <button onClick={() => fileRef.current.click()} className="flex items-center gap-1.5 px-3 py-2 rounded-lg text-sm font-medium border border-slate-200 hover:bg-slate-50">
        <Upload size={14} /> Import
      </button>
      <input ref={fileRef} type="file" accept=".csv" className="hidden" onChange={(e) => { if (e.target.files[0]) onImport(e.target.files[0]); e.target.value = ""; }} />
      <button onClick={onExportCsv} className="flex items-center gap-1.5 px-3 py-2 rounded-lg text-sm font-medium border border-slate-200 hover:bg-slate-50">
        <Download size={14} /> CSV
      </button>
      <button onClick={onExportXlsx} className="flex items-center gap-1.5 px-3 py-2 rounded-lg text-sm font-medium border border-slate-200 hover:bg-slate-50">
        <FileSpreadsheet size={14} /> Excel
      </button>
      <button onClick={onAdd} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-semibold text-white" style={{ background: "#B91C3C" }}>
        <Plus size={15} /> Tambah Prospek
      </button>
    </div>
  );
}

function filterProspects(list, filters) {
  return list.filter((p) => {
    if (filters.q) {
      const q = filters.q.toLowerCase();
      if (!(p.nama.toLowerCase().includes(q) || p.whatsapp.includes(q) || p.unit.toLowerCase().includes(q))) return false;
    }
    if (filters.sales && p.sales !== filters.sales) return false;
    if (filters.status && p.statusProspek !== filters.status) return false;
    if (filters.sumber && p.sumber !== filters.sumber) return false;
    if (filters.suhu && getSuhu(calculateScore(p)) !== filters.suhu) return false;
    return true;
  });
}

function ProspectsPage({ prospects, setProspects, openDetail }) {
  const [filters, setFilters] = useState({ q: "", sales: "", suhu: "", status: "", sumber: "" });
  const [modal, setModal] = useState(null); // null | 'new' | prospect object

  const filtered = useMemo(() => {
    return filterProspects(prospects, filters).sort((a, b) => calculateScore(b) - calculateScore(a));
  }, [prospects, filters]);

  const handleSave = (data) => {
    setProspects((prev) => {
      const exists = prev.some((p) => p.id === data.id);
      if (exists) return prev.map((p) => (p.id === data.id ? data : p));
      return [data, ...prev];
    });
    setModal(null);
  };

  const handleDelete = (id) => {
    if (window.confirm("Hapus prospek ini?")) setProspects((prev) => prev.filter((p) => p.id !== id));
  };

  const handleImport = (file) => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: (res) => {
        const rows = res.data.map((r, i) => ({
          id: `PSP-IMP-${Date.now()}-${i}`,
          tanggalMasuk: r.Tanggal || todayISO(),
          nama: r.Nama || "Tanpa Nama",
          whatsapp: r.WhatsApp || r.Whatsapp || "",
          kota: KOTA_LIST[0],
          sales: r.Sales || SALES_LIST[0],
          sumber: r.Sumber || SUMBER_LIST[0],
          unit: r.Unit || UNIT_LIST[0],
          tahunMobilLama: "", merekMobilLama: "", estimasiHargaMobilLama: "",
          statusTradeIn: "Tidak", budgetDP: "", targetAngsuran: "",
          statusKredit: "Belum Diajukan",
          statusProspek: STATUS_LIST.includes(r.Status) ? r.Status : "NEW LEAD",
          lastFollowUp: todayISO(), nextFollowUp: "", jumlahFollowUp: 0,
          responTerakhir: "", alasanBelumClosing: ALASAN_LIST[0], potensiSPK: "Sedang",
          catatanSales: r.Catatan || "",
          sudahMerespon: false, mintaHarga: false, mintaSimulasiKredit: false,
          mintaTestDrive: false, sudahKirimDokumen: false, sebutkanTarget: false,
          belumBerminat: false, tundaLama: false, history: [],
        }));
        setProspects((prev) => [...rows, ...prev]);
        alert(`${rows.length} prospek berhasil diimpor.`);
      },
    });
  };

  const exportCsv = () => {
    const csv = Papa.unparse(prospects.map((p) => ({
      ID: p.id, Nama: p.nama, WhatsApp: p.whatsapp, Kota: p.kota, Sales: p.sales,
      Unit: p.unit, Sumber: p.sumber, Status: p.statusProspek, Suhu: getSuhu(calculateScore(p)),
      Score: calculateScore(p), LastFollowUp: p.lastFollowUp, NextFollowUp: p.nextFollowUp,
    })));
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href = url; a.download = "prospek_toyota.csv"; a.click();
  };

  const exportXlsx = () => {
    const ws = XLSX.utils.json_to_sheet(prospects.map((p) => ({
      ID: p.id, Nama: p.nama, WhatsApp: p.whatsapp, Kota: p.kota, Sales: p.sales,
      Unit: p.unit, Sumber: p.sumber, Status: p.statusProspek, Suhu: getSuhu(calculateScore(p)),
      Score: calculateScore(p), LastFollowUp: p.lastFollowUp, NextFollowUp: p.nextFollowUp,
    })));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Prospek");
    XLSX.writeFile(wb, "prospek_toyota.xlsx");
  };

  return (
    <div>
      <FilterBar filters={filters} setFilters={setFilters} onAdd={() => setModal("new")} onImport={handleImport} onExportCsv={exportCsv} onExportXlsx={exportXlsx} />
      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left">
            <thead>
              <tr className="text-[11px] uppercase tracking-wide text-slate-400 border-b border-slate-100">
                <th className="py-2.5 px-3 font-semibold">Score</th>
                <th className="py-2.5 px-3 font-semibold">Customer</th>
                <th className="py-2.5 px-3 font-semibold">Unit</th>
                <th className="py-2.5 px-3 font-semibold">Suhu</th>
                <th className="py-2.5 px-3 font-semibold">Status</th>
                <th className="py-2.5 px-3 font-semibold">Follow Up</th>
                <th className="py-2.5 px-3 font-semibold">Next FU</th>
                <th className="py-2.5 px-3 font-semibold"></th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((p) => (
                <ProspectRow key={p.id} p={p} onEdit={setModal} onDelete={handleDelete} onOpen={openDetail} />
              ))}
            </tbody>
          </table>
        </div>
        {filtered.length === 0 && <EmptyState title="Tidak ada prospek" desc="Coba ubah filter atau tambahkan prospek baru." />}
      </div>
      <div className="text-xs text-slate-400 mt-2">{filtered.length} dari {prospects.length} prospek ditampilkan</div>

      {modal && (
        <ProspectModal
          initial={modal === "new" ? null : modal}
          onClose={() => setModal(null)}
          onSave={handleSave}
        />
      )}
    </div>
  );
}

/* ============================================================
   DASHBOARD
   ============================================================ */

function Dashboard({ prospects, setActiveTab }) {
  const stats = useMemo(() => {
    const total = prospects.length;
    const today = todayISO();
    const baru = prospects.filter((p) => p.tanggalMasuk === today).length;
    const scored = prospects.map((p) => ({ p, score: calculateScore(p) }));
    const hot = scored.filter((x) => getSuhu(x.score) === "HOT").length;
    const warm = scored.filter((x) => getSuhu(x.score) === "WARM").length;
    const cold = scored.filter((x) => getSuhu(x.score) === "COLD").length;
    const belumFU = prospects.filter((p) => p.jumlahFollowUp === 0).length;
    const fuHariIni = prospects.filter((p) => getFollowUpFlag(p) === "DUE_TODAY").length;
    const fuTerlambat = prospects.filter((p) => getFollowUpFlag(p) === "OVERDUE").length;
    const testDrive = prospects.filter((p) => p.statusProspek === "TEST DRIVE").length;
    const negosiasi = prospects.filter((p) => p.statusProspek === "NEGOSIASI").length;
    const spk = prospects.filter((p) => p.statusProspek === "SPK" || p.statusProspek === "DELIVERY").length;
    const lost = prospects.filter((p) => p.statusProspek === "LOST").length;

    const contacted = prospects.filter((p) => p.statusProspek !== "NEW LEAD" && p.statusProspek !== "BELUM DIHUBUNGI").length;
    const conversionRate = total ? Math.round((spk / total) * 100) : 0;
    const followUpRate = total ? Math.round(((total - belumFU) / total) * 100) : 0;
    const spkRate = total ? Math.round((spk / total) * 100) : 0;

    return { total, baru, hot, warm, cold, belumFU, fuHariIni, fuTerlambat, testDrive, negosiasi, spk, lost, conversionRate, followUpRate, spkRate };
  }, [prospects]);

  const funnelData = useMemo(() => {
    return FUNNEL_STAGES.map((stage, idx) => ({
      stage,
      count: prospects.filter((p) => isActiveFunnel(p) && STATUS_TO_FUNNEL[p.statusProspek] >= idx).length,
    }));
  }, [prospects]);

  const perSales = useMemo(() => {
    return SALES_LIST.map((s) => ({
      sales: s.split(" ")[0],
      prospek: prospects.filter((p) => p.sales === s).length,
      spk: prospects.filter((p) => p.sales === s && (p.statusProspek === "SPK" || p.statusProspek === "DELIVERY")).length,
    }));
  }, [prospects]);

  const alerts = useMemo(() => {
    const list = [];
    const hotBelumFU = prospects.filter((p) => getSuhu(calculateScore(p)) === "HOT" && getPriority(p) === "SEKARANG").length;
    const overdue = prospects.filter((p) => getFollowUpFlag(p) === "OVERDUE").length;
    const mintaSimulasi = prospects.filter((p) => p.mintaSimulasiKredit && daysBetween(p.lastFollowUp) >= 2).length;
    const testDriveBelumFU = prospects.filter((p) => p.statusProspek === "TEST DRIVE" && daysBetween(p.lastFollowUp) >= 1).length;
    const negoBelumClosing = prospects.filter((p) => p.statusProspek === "NEGOSIASI").length;
    if (hotBelumFU) list.push(`🚨 ${hotBelumFU} Prospek HOT belum di-follow up.`);
    if (overdue) list.push(`🚨 ${overdue} Prospek OVERDUE.`);
    if (mintaSimulasi) list.push(`🚨 ${mintaSimulasi} Customer meminta simulasi tetapi belum ditindaklanjuti.`);
    if (testDriveBelumFU) list.push(`🚨 ${testDriveBelumFU} Customer sudah test drive tetapi belum di-follow up.`);
    if (negoBelumClosing) list.push(`🚨 ${negoBelumClosing} Prospek sudah negosiasi tetapi belum closing.`);
    return list;
  }, [prospects]);

  return (
    <div className="space-y-5">
      {alerts.length > 0 && (
        <div className="bg-white rounded-xl border border-red-100 p-4">
          <div className="flex items-center gap-2 text-sm font-bold text-red-600 mb-2"><AlertTriangle size={16} /> Alert Otomatis</div>
          <div className="grid sm:grid-cols-2 gap-2">
            {alerts.map((a, i) => (
              <div key={i} className="text-sm text-slate-700 bg-red-50/60 rounded-lg px-3 py-2">{a}</div>
            ))}
          </div>
        </div>
      )}

      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
        <KPICard label="Total Prospek" value={stats.total} />
        <KPICard label="Baru Hari Ini" value={stats.baru} accent="#2563EB" />
        <KPICard label="Prospek HOT" value={stats.hot} accent="#DC2626" />
        <KPICard label="Prospek WARM" value={stats.warm} accent="#D97706" />
        <KPICard label="Prospek COLD" value={stats.cold} accent="#64748B" />
        <KPICard label="Belum Follow Up" value={stats.belumFU} accent="#DC2626" />
        <KPICard label="Follow Up Hari Ini" value={stats.fuHariIni} accent="#D97706" />
        <KPICard label="Follow Up Terlambat" value={stats.fuTerlambat} accent="#DC2626" />
        <KPICard label="Test Drive" value={stats.testDrive} accent="#2563EB" />
        <KPICard label="Negosiasi" value={stats.negosiasi} accent="#7C3AED" />
        <KPICard label="SPK" value={stats.spk} accent="#15803D" />
        <KPICard label="Lost" value={stats.lost} accent="#94A3B8" />
      </div>

      <div className="grid lg:grid-cols-3 gap-4">
        <div className="lg:col-span-2 bg-white rounded-xl border border-slate-200 p-5">
          <div className="font-bold mb-4 flex items-center gap-2" style={{ fontFamily: "'Sora', sans-serif" }}><GitBranch size={16} /> Funnel Prospek</div>
          <div className="space-y-1.5">
            {funnelData.map((f, i) => {
              const max = funnelData[0].count || 1;
              const pct = Math.max(8, Math.round((f.count / max) * 100));
              return (
                <div key={f.stage} className="flex items-center gap-3">
                  <div className="w-24 text-xs text-slate-500 shrink-0">{f.stage}</div>
                  <div className="flex-1 h-8 bg-slate-50 rounded-md overflow-hidden">
                    <div
                      className="h-full rounded-md flex items-center justify-end px-2 text-white text-xs font-semibold"
                      style={{ width: `${pct}%`, background: `linear-gradient(90deg, #12233F, #B91C3C)` }}
                    >
                      {f.count}
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>

        <div className="bg-white rounded-xl border border-slate-200 p-5">
          <div className="font-bold mb-4" style={{ fontFamily: "'Sora', sans-serif" }}>Rate Kunci</div>
          <div className="space-y-4">
            {[
              { label: "Conversion Rate", value: stats.conversionRate, color: "#15803D" },
              { label: "Follow Up Rate", value: stats.followUpRate, color: "#2563EB" },
              { label: "SPK Rate", value: stats.spkRate, color: "#B91C3C" },
            ].map((r) => (
              <div key={r.label}>
                <div className="flex justify-between text-sm mb-1"><span className="text-slate-600">{r.label}</span><span className="font-bold">{r.value}%</span></div>
                <div className="h-2 bg-slate-100 rounded-full overflow-hidden"><div className="h-full rounded-full" style={{ width: `${r.value}%`, background: r.color }} /></div>
              </div>
            ))}
          </div>
        </div>
      </div>

      <div className="bg-white rounded-xl border border-slate-200 p-5">
        <div className="font-bold mb-4" style={{ fontFamily: "'Sora', sans-serif" }}>Prospek & SPK per Sales</div>
        <div style={{ width: "100%", height: 260 }}>
          <ResponsiveContainer>
            <BarChart data={perSales} barGap={4}>
              <CartesianGrid vertical={false} stroke="#F1F5F9" />
              <XAxis dataKey="sales" tick={{ fontSize: 12, fill: "#64748B" }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fontSize: 12, fill: "#64748B" }} axisLine={false} tickLine={false} />
              <Tooltip contentStyle={{ borderRadius: 10, border: "1px solid #E2E8F0", fontSize: 12 }} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Bar dataKey="prospek" name="Prospek" fill="#12233F" radius={[4, 4, 0, 0]} />
              <Bar dataKey="spk" name="SPK" fill="#B91C3C" radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   HOT PROSPECTS PAGE
   ============================================================ */

function HotPage({ prospects, openDetail }) {
  const hot = useMemo(() => {
    return prospects
      .map((p) => ({ p, score: calculateScore(p) }))
      .filter((x) => getSuhu(x.score) === "HOT")
      .sort((a, b) => b.score - a.score);
  }, [prospects]);

  return (
    <div>
      <div className="mb-4 text-sm text-slate-500">{hot.length} prospek dengan skor tertinggi & potensi paling besar untuk closing.</div>
      <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">
        {hot.map(({ p, score }) => {
          const priority = getPriority(p);
          const pm = PRIORITY_META[priority];
          return (
            <div key={p.id} onClick={() => openDetail(p)} className="bg-white rounded-xl border border-red-100 p-4 cursor-pointer hover:shadow-md transition-shadow">
              <div className="flex items-center justify-between mb-2">
                <ScoreRing score={score} />
                <Badge bg="#FEE2E2" color="#B91C3C">{pm.icon} {pm.label}</Badge>
              </div>
              <div className="font-semibold text-slate-800">{p.nama}</div>
              <div className="text-xs text-slate-400 mb-2">{p.unit} · {p.sales}</div>
              <div className="text-xs text-slate-500">{p.statusProspek}</div>
              <div className="text-xs text-slate-400 mt-1">Next FU: {formatDate(p.nextFollowUp)}</div>
            </div>
          );
        })}
      </div>
      {hot.length === 0 && <EmptyState title="Belum ada prospek HOT" desc="Prospek HOT akan muncul di sini setelah score mencapai 70 ke atas." />}
    </div>
  );
}

/* ============================================================
   FOLLOW UP PAGE (grouped by priority)
   ============================================================ */

function FollowUpPage({ prospects, openDetail }) {
  const groups = useMemo(() => {
    const g = { SEKARANG: [], HARI_INI: [], TERJADWAL: [], LOW: [] };
    prospects.forEach((p) => g[getPriority(p)].push(p));
    Object.keys(g).forEach((k) => g[k].sort((a, b) => calculateScore(b) - calculateScore(a)));
    return g;
  }, [prospects]);

  return (
    <div className="space-y-6">
      {Object.entries(groups).map(([key, list]) => {
        const meta = PRIORITY_META[key];
        if (list.length === 0) return null;
        return (
          <div key={key}>
            <div className="flex items-center gap-2 mb-2">
              <span>{meta.icon}</span>
              <span className="font-bold" style={{ fontFamily: "'Sora', sans-serif" }}>{meta.label}</span>
              <Badge bg="#F1F5F9" color="#475569">{list.length}</Badge>
            </div>
            <div className="bg-white rounded-xl border border-slate-200 divide-y divide-slate-50">
              {list.map((p) => {
                const flag = getFollowUpFlag(p);
                const fm = FLAG_META[flag];
                const score = calculateScore(p);
                const c = suhuColor(getSuhu(score));
                return (
                  <div key={p.id} onClick={() => openDetail(p)} className="flex items-center gap-3 px-4 py-3 hover:bg-slate-50 cursor-pointer">
                    <ScoreRing score={score} />
                    <div className="flex-1 min-w-0">
                      <div className="font-medium text-sm text-slate-800 truncate">{p.nama} <span className="text-slate-400 font-normal">· {p.unit}</span></div>
                      <div className="text-xs text-slate-400">{p.sales} · {p.statusProspek}</div>
                    </div>
                    <Badge bg={c.bg} color={c.text}>{getSuhu(score)}</Badge>
                    <Badge bg={fm.bg} color={fm.color}>{fm.label}</Badge>
                    <div className="text-xs text-slate-400 w-20 text-right">{formatDate(p.nextFollowUp)}</div>
                  </div>
                );
              })}
            </div>
          </div>
        );
      })}
    </div>
  );
}

/* ============================================================
   FOLLOW UP GENERATOR PAGE
   ============================================================ */

function MessageCard({ opt, onCopy, copied }) {
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-4 flex flex-col gap-3">
      <div className="flex items-center justify-between">
        <Badge bg="#EEF2FF" color="#4338CA">{opt.label}</Badge>
        <button
          onClick={onCopy}
          className="flex items-center gap-1 text-xs font-semibold px-2.5 py-1.5 rounded-lg border border-slate-200 hover:bg-slate-50"
        >
          {copied ? <Check size={13} className="text-green-600" /> : <Copy size={13} />}
          {copied ? "Tersalin" : "Copy"}
        </button>
      </div>
      <p className="text-sm text-slate-700 leading-relaxed whitespace-pre-wrap">{opt.text}</p>
    </div>
  );
}

function FollowUpGenerator({ prospects }) {
  const [selectedId, setSelectedId] = useState("");
  const [form, setForm] = useState({
    tujuan: "harga", gaya: "Hangat & Personal", lastConversation: "", alasan: ALASAN_LIST[0],
  });
  const [messages, setMessages] = useState(null);
  const [copiedKey, setCopiedKey] = useState("");

  const selected = prospects.find((p) => p.id === selectedId);

  useEffect(() => {
    if (selected) {
      setForm((f) => ({ ...f, alasan: selected.alasanBelumClosing || f.alasan, lastConversation: selected.responTerakhir || "" }));
    }
  }, [selectedId]); // eslint-disable-line

  const handleGenerate = () => {
    const ctx = {
      nama: selected ? selected.nama : "",
      unit: selected ? selected.unit : "",
      tujuan: form.tujuan,
      gaya: form.gaya,
      alasan: form.alasan,
      lastConversation: form.lastConversation,
    };
    setMessages(generateMessages(ctx));
  };

  const copy = (key, text) => {
    navigator.clipboard?.writeText(text);
    setCopiedKey(key);
    setTimeout(() => setCopiedKey(""), 1500);
  };

  const analysis = SMART_ANALYSIS[form.alasan] || SMART_ANALYSIS.Lainnya;

  return (
    <div className="grid lg:grid-cols-2 gap-5">
      <div className="bg-white rounded-xl border border-slate-200 p-5 space-y-4 h-fit">
        <div className="font-bold flex items-center gap-2" style={{ fontFamily: "'Sora', sans-serif" }}><MessageSquareText size={16} /> Follow Up Generator</div>

        <Field label="Pilih Prospek (opsional, isi otomatis)">
          <select className={inputCls} value={selectedId} onChange={(e) => setSelectedId(e.target.value)}>
            <option value="">-- Isi manual --</option>
            {prospects.map((p) => <option key={p.id} value={p.id}>{p.nama} · {p.unit}</option>)}
          </select>
        </Field>

        {!selected && (
          <div className="grid grid-cols-2 gap-3">
            <Field label="Nama Customer">
              <input className={inputCls} value={form.customNama || ""} onChange={(e) => setForm({ ...form, customNama: e.target.value })} />
            </Field>
            <Field label="Unit">
              <select className={inputCls} value={form.customUnit || UNIT_LIST[0]} onChange={(e) => setForm({ ...form, customUnit: e.target.value })}>
                {UNIT_LIST.map((u) => <option key={u}>{u}</option>)}
              </select>
            </Field>
          </div>
        )}

        <Field label="Tujuan Follow Up">
          <select className={inputCls} value={form.tujuan} onChange={(e) => setForm({ ...form, tujuan: e.target.value })}>
            {TUJUAN_LIST.map((t) => <option key={t.id} value={t.id}>{t.label}</option>)}
          </select>
        </Field>

        <div className="grid grid-cols-2 gap-3">
          <Field label="Gaya Bahasa">
            <select className={inputCls} value={form.gaya} onChange={(e) => setForm({ ...form, gaya: e.target.value })}>
              {GAYA_BAHASA.map((g) => <option key={g}>{g}</option>)}
            </select>
          </Field>
          <Field label="Alasan Belum Closing">
            <select className={inputCls} value={form.alasan} onChange={(e) => setForm({ ...form, alasan: e.target.value })}>
              {ALASAN_LIST.map((a) => <option key={a}>{a}</option>)}
            </select>
          </Field>
        </div>

        <Field label="Last Conversation / Respon Terakhir">
          <textarea rows={2} className={inputCls} value={form.lastConversation} onChange={(e) => setForm({ ...form, lastConversation: e.target.value })} />
        </Field>

        <button
          onClick={() => {
            const ctx = {
              nama: selected ? selected.nama : form.customNama || "",
              unit: selected ? selected.unit : form.customUnit || UNIT_LIST[0],
              tujuan: form.tujuan, gaya: form.gaya, alasan: form.alasan,
            };
            setMessages(generateMessages(ctx));
          }}
          className="w-full py-2.5 rounded-lg text-sm font-semibold text-white flex items-center justify-center gap-2"
          style={{ background: "#B91C3C" }}
        >
          <Sparkles size={15} /> Generate Follow Up
        </button>

        <div className="rounded-xl p-4" style={{ background: "#F8FAFC" }}>
          <div className="flex items-center gap-1.5 text-xs font-bold uppercase tracking-wide text-slate-500 mb-2"><Sparkles size={13} /> Smart Analysis</div>
          <div className="text-sm space-y-1.5">
            <div><span className="font-semibold">Analisis: </span>{analysis.analisis}</div>
            <div><span className="font-semibold">Strategi: </span>{analysis.strategi}</div>
            <div><span className="font-semibold">Next Action: </span>{analysis.nextAction}</div>
          </div>
        </div>
      </div>

      <div className="space-y-3">
        {!messages && (
          <div className="bg-white rounded-xl border border-dashed border-slate-300 p-10 text-center text-slate-400 text-sm">
            Isi form di sebelah kiri lalu klik <span className="font-semibold text-slate-500">Generate Follow Up</span> untuk menghasilkan 3 pilihan pesan WhatsApp.
          </div>
        )}
        {messages && (
          <>
            <MessageCard opt={messages.A} onCopy={() => copy("A", messages.A.text)} copied={copiedKey === "A"} />
            <MessageCard opt={messages.B} onCopy={() => copy("B", messages.B.text)} copied={copiedKey === "B"} />
            <MessageCard opt={messages.C} onCopy={() => copy("C", messages.C.text)} copied={copiedKey === "C"} />
          </>
        )}
      </div>
    </div>
  );
}

/* ============================================================
   REACTIVATION PAGE
   ============================================================ */

function ReactivationPage({ prospects, openDetail }) {
  const [range, setRange] = useState("7");

  const ranges = [
    { id: "7", label: "7 hari" }, { id: "14", label: "14 hari" }, { id: "30", label: "30 hari" },
    { id: "60", label: "60 hari" }, { id: "90", label: "90+ hari" },
  ];

  const list = useMemo(() => {
    const min = Number(range);
    return prospects
      .filter((p) => p.statusProspek !== "LOST" && p.statusProspek !== "DELIVERY")
      .filter((p) => {
        const d = daysBetween(p.lastFollowUp);
        if (range === "90") return d >= 90;
        const idx = ranges.findIndex((r) => r.id === range);
        const nextMin = ranges[idx + 1] ? Number(ranges[idx + 1].id) : Infinity;
        return d >= min && d < nextMin;
      })
      .map((p) => ({ p, score: calculateScore(p), days: daysBetween(p.lastFollowUp) }))
      .sort((a, b) => b.score - a.score);
  }, [prospects, range]);

  return (
    <div>
      <div className="flex items-center gap-2 mb-4">
        {ranges.map((r) => (
          <button
            key={r.id}
            onClick={() => setRange(r.id)}
            className="px-3.5 py-1.5 rounded-lg text-sm font-medium border"
            style={range === r.id ? { background: "#12233F", color: "white", borderColor: "#12233F" } : { borderColor: "#E2E8F0", color: "#475569" }}
          >
            {r.label}
          </button>
        ))}
      </div>

      <div className="rounded-xl p-4 mb-4" style={{ background: "linear-gradient(90deg,#12233F,#1E3A5F)" }}>
        <div className="text-white font-bold flex items-center gap-2" style={{ fontFamily: "'Sora', sans-serif" }}><RotateCcw size={16} /> Prospek yang Layak Dihidupkan Kembali</div>
        <div className="text-white/60 text-sm mt-1">Diranking berdasarkan potensi closing (score tertinggi) dalam rentang waktu yang dipilih.</div>
      </div>

      <div className="bg-white rounded-xl border border-slate-200 divide-y divide-slate-50">
        {list.map(({ p, score, days }, i) => {
          const c = suhuColor(getSuhu(score));
          return (
            <div key={p.id} onClick={() => openDetail(p)} className="flex items-center gap-3 px-4 py-3 hover:bg-slate-50 cursor-pointer">
              <div className="w-6 text-xs text-slate-400 font-mono">#{i + 1}</div>
              <ScoreRing score={score} />
              <div className="flex-1 min-w-0">
                <div className="font-medium text-sm text-slate-800 truncate">{p.nama} <span className="text-slate-400 font-normal">· {p.unit}</span></div>
                <div className="text-xs text-slate-400">{p.sales} · Tidak dihubungi {days} hari</div>
              </div>
              <Badge bg={c.bg} color={c.text}>{getSuhu(score)}</Badge>
            </div>
          );
        })}
      </div>
      {list.length === 0 && <EmptyState title="Tidak ada prospek" desc="Tidak ada prospek pada rentang waktu ini." />}
    </div>
  );
}

/* ============================================================
   PERFORMANCE SALES PAGE
   ============================================================ */

function PerformancePage({ prospects }) {
  const rows = useMemo(() => {
    return SALES_LIST.map((s) => {
      const mine = prospects.filter((p) => p.sales === s);
      const total = mine.length;
      const fu = mine.filter((p) => p.jumlahFollowUp > 0).length;
      const respon = mine.filter((p) => p.sudahMerespon).length;
      const testDrive = mine.filter((p) => p.statusProspek === "TEST DRIVE" || STATUS_TO_FUNNEL[p.statusProspek] >= 4).length;
      const negosiasi = mine.filter((p) => STATUS_TO_FUNNEL[p.statusProspek] >= 5).length;
      const spk = mine.filter((p) => p.statusProspek === "SPK" || p.statusProspek === "DELIVERY").length;
      const hot = mine.filter((p) => getSuhu(calculateScore(p)) === "HOT").length;
      const closingRate = total ? Math.round((spk / total) * 100) : 0;

      let indicator = "🟢 Produktif";
      let note = "Performa follow up dan closing seimbang.";
      if (total > 0 && fu / total < 0.4) { indicator = "🔴 Follow Up Rendah"; note = "Banyak prospek tetapi sedikit follow up dilakukan."; }
      else if (fu > 0 && respon / fu < 0.4) { indicator = "🟡 Perlu Coaching"; note = "Banyak follow up tetapi sedikit respon customer."; }
      else if (respon > 0 && spk / respon < 0.15) { indicator = "🟡 Perlu Coaching"; note = "Banyak respon tetapi sedikit konversi ke SPK."; }
      else if (hot >= 2 && spk === 0) { indicator = "🟡 Perlu Coaching"; note = "Memiliki prospek HOT namun belum terkonversi."; }

      return { sales: s, total, fu, respon, testDrive, negosiasi, spk, closingRate, indicator, note };
    }).sort((a, b) => b.spk - a.spk);
  }, [prospects]);

  return (
    <div className="space-y-4">
      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left">
            <thead>
              <tr className="text-[11px] uppercase tracking-wide text-slate-400 border-b border-slate-100">
                <th className="py-2.5 px-3 font-semibold">Sales</th>
                <th className="py-2.5 px-3 font-semibold">Prospek</th>
                <th className="py-2.5 px-3 font-semibold">Follow Up</th>
                <th className="py-2.5 px-3 font-semibold">Respon</th>
                <th className="py-2.5 px-3 font-semibold">Test Drive</th>
                <th className="py-2.5 px-3 font-semibold">Negosiasi</th>
                <th className="py-2.5 px-3 font-semibold">SPK</th>
                <th className="py-2.5 px-3 font-semibold">Closing Rate</th>
                <th className="py-2.5 px-3 font-semibold">Indikator</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.sales} className="border-b border-slate-50">
                  <td className="py-2.5 px-3 font-semibold text-sm text-slate-800">{r.sales}</td>
                  <td className="py-2.5 px-3 text-sm">{r.total}</td>
                  <td className="py-2.5 px-3 text-sm">{r.fu}</td>
                  <td className="py-2.5 px-3 text-sm">{r.respon}</td>
                  <td className="py-2.5 px-3 text-sm">{r.testDrive}</td>
                  <td className="py-2.5 px-3 text-sm">{r.negosiasi}</td>
                  <td className="py-2.5 px-3 text-sm font-semibold text-green-700">{r.spk}</td>
                  <td className="py-2.5 px-3 text-sm">{r.closingRate}%</td>
                  <td className="py-2.5 px-3 text-sm">{r.indicator}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
      <div className="grid sm:grid-cols-2 gap-3">
        {rows.filter((r) => r.indicator !== "🟢 Produktif").map((r) => (
          <div key={r.sales} className="bg-white rounded-xl border border-amber-100 p-4">
            <div className="font-semibold text-sm">{r.sales} — {r.indicator}</div>
            <div className="text-xs text-slate-500 mt-1">{r.note}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ============================================================
   FUNNEL PAGE
   ============================================================ */

function FunnelPage({ prospects }) {
  const funnelData = useMemo(() => {
    return FUNNEL_STAGES.map((stage, idx) => ({
      stage,
      count: prospects.filter((p) => isActiveFunnel(p) && STATUS_TO_FUNNEL[p.statusProspek] >= idx).length,
    }));
  }, [prospects]);

  const conv = funnelData.map((f, i) => {
    if (i === 0) return 100;
    const prev = funnelData[i - 1].count || 1;
    return Math.round((f.count / prev) * 100);
  });

  return (
    <div className="bg-white rounded-xl border border-slate-200 p-6">
      <div className="font-bold mb-6" style={{ fontFamily: "'Sora', sans-serif" }}>Funnel Konversi Prospek</div>
      <div className="space-y-3">
        {funnelData.map((f, i) => {
          const max = funnelData[0].count || 1;
          const pct = Math.max(6, Math.round((f.count / max) * 100));
          return (
            <div key={f.stage} className="flex items-center gap-4">
              <div className="w-28 text-sm text-slate-600 font-medium shrink-0">{f.stage}</div>
              <div className="flex-1 h-10 bg-slate-50 rounded-lg overflow-hidden relative">
                <div
                  className="h-full rounded-lg flex items-center px-3 text-white text-sm font-bold"
                  style={{ width: `${pct}%`, background: `linear-gradient(90deg, #12233F, #B91C3C)` }}
                >
                  {f.count}
                </div>
              </div>
              <div className="w-16 text-right text-xs text-slate-400 shrink-0">{i > 0 ? `${conv[i]}%` : ""}</div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ============================================================
   HISTORY PAGE
   ============================================================ */

function HistoryPage({ prospects }) {
  const entries = useMemo(() => {
    const list = [];
    prospects.forEach((p) => {
      (p.history || []).forEach((h) => list.push({ ...h, nama: p.nama, sales: p.sales, unit: p.unit, id: p.id }));
    });
    return list.sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));
  }, [prospects]);

  return (
    <div className="bg-white rounded-xl border border-slate-200 divide-y divide-slate-50">
      {entries.map((h, i) => (
        <div key={i} className="px-4 py-3 flex items-start gap-3">
          <div className="w-9 h-9 rounded-full bg-slate-100 flex items-center justify-center shrink-0"><ClipboardList size={15} className="text-slate-500" /></div>
          <div className="flex-1 min-w-0">
            <div className="text-sm font-semibold text-slate-800">{h.nama} <span className="font-normal text-slate-400">· {h.unit} · {h.sales}</span></div>
            <div className="text-sm text-slate-600 mt-0.5">{h.isi}</div>
            {h.respon && <div className="text-xs text-slate-500 mt-0.5">Respon: {h.respon}</div>}
          </div>
          <div className="text-xs text-slate-400 shrink-0 text-right">{formatDate(h.tanggal)}<br />{h.jam} · {h.media}</div>
        </div>
      ))}
      {entries.length === 0 && <EmptyState title="Belum ada history" desc="Riwayat follow up akan muncul di sini." />}
    </div>
  );
}

/* ============================================================
   SETTINGS PAGE
   ============================================================ */

function SettingsPage() {
  return (
    <div className="grid md:grid-cols-2 gap-4">
      <div className="bg-white rounded-xl border border-slate-200 p-5">
        <div className="font-bold mb-3" style={{ fontFamily: "'Sora', sans-serif" }}>Tim Sales</div>
        <div className="space-y-1.5">
          {SALES_LIST.map((s) => (
            <div key={s} className="flex items-center justify-between text-sm px-3 py-2 rounded-lg bg-slate-50">
              <span>{s}</span>
              <Badge bg="#E2E8F0" color="#475569">Aktif</Badge>
            </div>
          ))}
        </div>
      </div>
      <div className="bg-white rounded-xl border border-slate-200 p-5">
        <div className="font-bold mb-3" style={{ fontFamily: "'Sora', sans-serif" }}>Sumber Prospek</div>
        <div className="flex flex-wrap gap-1.5">
          {SUMBER_LIST.map((s) => <Badge key={s} bg="#F1F5F9" color="#475569">{s}</Badge>)}
        </div>
      </div>
      <div className="bg-white rounded-xl border border-slate-200 p-5">
        <div className="font-bold mb-3" style={{ fontFamily: "'Sora', sans-serif" }}>Kategori Scoring</div>
        <div className="space-y-2 text-sm">
          <div className="flex items-center justify-between"><span>HOT</span><Badge bg="#FEE2E2" color="#B91C3C">Score 70 - 100</Badge></div>
          <div className="flex items-center justify-between"><span>WARM</span><Badge bg="#FEF3C7" color="#B45309">Score 40 - 69</Badge></div>
          <div className="flex items-center justify-between"><span>COLD</span><Badge bg="#E2E8F0" color="#475569">Score 0 - 39</Badge></div>
        </div>
      </div>
      <div className="bg-white rounded-xl border border-slate-200 p-5">
        <div className="font-bold mb-3" style={{ fontFamily: "'Sora', sans-serif" }}>Tentang Aplikasi</div>
        <div className="text-sm text-slate-500 leading-relaxed">
          Toyota Prospect Monitoring & Follow Up Generator membantu Sales, Supervisor, dan Kepala Cabang
          memonitor prospek, menentukan prioritas follow up, dan menghasilkan pesan WhatsApp follow up yang personal.
          Data disimpan secara otomatis di perangkat ini.
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   MAIN APP
   ============================================================ */

export default function App() {
  const [prospects, setProspects] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [active, setActive] = useState("dashboard");
  const [mobileOpen, setMobileOpen] = useState(false);
  const [detail, setDetail] = useState(null);

  useEffect(() => {
    let data = null;
    try {
      const raw = localStorage.getItem("toyota_prospects");
      if (raw) data = JSON.parse(raw);
    } catch (e) { /* not found or invalid */ }
    if (!data || !Array.isArray(data) || data.length === 0) {
      data = makeDemoData();
      try { localStorage.setItem("toyota_prospects", JSON.stringify(data)); } catch (e) {}
    }
    setProspects(data);
    setLoaded(true);
  }, []);

  useEffect(() => {
    if (!loaded) return;
    try { localStorage.setItem("toyota_prospects", JSON.stringify(prospects)); } catch (e) {}
  }, [prospects, loaded]);

  const addHistory = useCallback((id, note, nextDate) => {
    setProspects((prev) =>
      prev.map((p) => {
        if (p.id !== id) return p;
        const entry = {
          tanggal: todayISO(), jam: new Date().toTimeString().slice(0, 5),
          media: note.media, isi: note.isi, respon: note.respon, nextAction: note.nextAction, nextDate: note.nextDate,
        };
        return {
          ...p,
          history: [...(p.history || []), entry],
          lastFollowUp: todayISO(),
          nextFollowUp: nextDate || p.nextFollowUp,
          jumlahFollowUp: (p.jumlahFollowUp || 0) + 1,
          responTerakhir: note.respon || p.responTerakhir,
        };
      })
    );
    setDetail((d) => (d && d.id === id ? { ...d, jumlahFollowUp: (d.jumlahFollowUp || 0) + 1 } : d));
  }, []);

  const hotSekarangCount = useMemo(
    () => prospects.filter((p) => getSuhu(calculateScore(p)) === "HOT" && getPriority(p) === "SEKARANG").length,
    [prospects]
  );

  const openDetail = useCallback((p) => {
    const fresh = prospects.find((x) => x.id === p.id) || p;
    setDetail(fresh);
  }, [prospects]);

  if (!loaded) {
    return (
      <div className="min-h-[400px] flex items-center justify-center text-slate-400 text-sm">
        Memuat data prospek...
      </div>
    );
  }

  const titles = {
    dashboard: "Dashboard", prospek: "Database Prospek", hot: "Prospek HOT",
    followup: "Follow Up", generator: "Follow Up Generator", reactivation: "Reactivation Database",
    performance: "Performance Sales", funnel: "Funnel Konversi", history: "History Follow Up", settings: "Settings",
  };

  return (
    <div className="w-full min-h-screen flex" style={{ background: "#F5F6F8", fontFamily: "'Inter', sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Sora:wght@600;700;800&family=Inter:wght@400;500;600;700&display=swap');
      `}</style>

      <Sidebar active={active} setActive={setActive} mobileOpen={mobileOpen} setMobileOpen={setMobileOpen} alertCount={hotSekarangCount} />

      <div className="flex-1 min-w-0 flex flex-col">
        <div className="sticky top-0 z-20 bg-white/90 backdrop-blur border-b border-slate-200 px-4 sm:px-6 py-3.5 flex items-center gap-3">
          <button className="lg:hidden w-9 h-9 rounded-lg hover:bg-slate-100 flex items-center justify-center" onClick={() => setMobileOpen(true)}>
            <Menu size={19} />
          </button>
          <div>
            <div className="font-bold text-lg leading-tight" style={{ fontFamily: "'Sora', sans-serif", color: "#12233F" }}>{titles[active]}</div>
            <div className="text-xs text-slate-400">{new Date().toLocaleDateString("id-ID", { weekday: "long", day: "numeric", month: "long", year: "numeric" })}</div>
          </div>
        </div>

        <div className="p-4 sm:p-6 flex-1">
          {active === "dashboard" && <Dashboard prospects={prospects} setActiveTab={setActive} />}
          {active === "prospek" && <ProspectsPage prospects={prospects} setProspects={setProspects} openDetail={openDetail} />}
          {active === "hot" && <HotPage prospects={prospects} openDetail={openDetail} />}
          {active === "followup" && <FollowUpPage prospects={prospects} openDetail={openDetail} />}
          {active === "generator" && <FollowUpGenerator prospects={prospects} />}
          {active === "reactivation" && <ReactivationPage prospects={prospects} openDetail={openDetail} />}
          {active === "performance" && <PerformancePage prospects={prospects} />}
          {active === "funnel" && <FunnelPage prospects={prospects} />}
          {active === "history" && <HistoryPage prospects={prospects} />}
          {active === "settings" && <SettingsPage />}
        </div>
      </div>

      {detail && (
        <ProspectDetail
          p={detail}
          onClose={() => setDetail(null)}
          onAddHistory={addHistory}
        />
      )}
    </div>
  );
}
