/* =====================================================================
 * KONFIGURASI PORTAL AREK MANBILL (Aplikasi Rekap & Evaluasi Kinerja Manbill)
 * Cukup ubah file ini. Daftar aplikasi & pengumuman juga dapat dikelola
 * dari Spreadsheet APIK (sheet "Portal_Aplikasi" dan "Portal_Pengumuman");
 * bila sheet tersebut terisi, isinya menggantikan daftar "apps" di bawah.
 * ===================================================================== */
window.PORTAL_CONFIG = {
  name: "AREK Manbill",
  fullName: "Aplikasi Rekap & Evaluasi Kinerja Manbill",
  subtitle: "Portal terpadu Manbill Optimization UP 3 Jawa Timur. Rek, kabeh aplikasi ono kene: pilih aplikasi untuk mulai bekerja.",
  organization: "Manbill Optimization PLN UP 3 Jawa Timur",
  version: "1.0",

  // URL Web App APIK hasil Deploy (berakhiran /exec). Dipakai untuk membuka APIK
  // dan mengambil ringkasan, pengumuman, serta Buku Panduan.
  apikUrl: "https://script.google.com/macros/s/GANTI_DENGAN_ID_DEPLOYMENT/exec",

  // Tampilkan panel ringkasan APIK di halaman utama
  showSummary: true,

  // Panel "SLA Gabungan" di halaman utama: URL Web App Dashboard SLA (berakhiran /exec).
  // Kosongkan ("") untuk menyembunyikan panel.
  slaUrl: "https://script.google.com/macros/s/GANTI_DENGAN_ID_DEPLOYMENT_SLA/exec",
  // Alamat yang dibuka tombol "Buka Dashboard SLA" (kosong = sama dengan slaUrl)
  slaDashboardUrl: "",

  // Kontak bantuan (kosongkan yang tidak dipakai)
  contact: { name: "Admin APIK", whatsapp: "", email: "" },

  // Daftar aplikasi cadangan (dipakai bila sheet Portal_Aplikasi belum ada / APIK tidak dapat dihubungi)
  // status: "Aktif" | "Segera Hadir" | "Pemeliharaan"; url "#panduan" = unduh Buku Panduan APIK
  apps: [
    { name: "APIK Manbill", desc: "Penilaian & Insentif Kinerja", category: "Kinerja & Reward", status: "Aktif", url: "apik/", icon: "fa-ranking-star", color: "#0b5a7a", newTab: true },
    { name: "Buku Panduan APIK", desc: "Unduh panduan penggunaan (PDF)", category: "Dokumen", status: "Aktif", url: "#panduan", icon: "fa-book-open", color: "#b45309", isNew: true },
    { name: "Absensi Rapat", desc: "Kehadiran, Notulen & Eviden Rapat", category: "Operasional", status: "Segera Hadir", url: "", icon: "fa-clipboard-user", color: "#0f766e", newTab: true },
    { name: "Monitoring SLA", desc: "SLA Kontrak Manbill per Area", category: "Laporan", status: "Segera Hadir", url: "", icon: "fa-gauge-high", color: "#1d4ed8", newTab: true },
    { name: "Laporan Bulanan", desc: "Rekap Kinerja & Reward", category: "Laporan", status: "Segera Hadir", url: "", icon: "fa-chart-column", color: "#7c3aed", newTab: true },
    { name: "Arsip Dokumen", desc: "Dokumen & Berkas Manbill", category: "Dokumen", status: "Segera Hadir", url: "", icon: "fa-box-archive", color: "#7c4a1e", newTab: true }
  ]
};
