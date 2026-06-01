import React, { useState, useEffect } from 'react';
import {
  LayoutDashboard, Users, UserCog, HeartHandshake, Beef,
  ScanLine, Ticket, Award, LineChart,
  ChevronDown, RefreshCw, DownloadCloud, QrCode, X, Box, CheckCircle, AlertCircle, Plus, Eye
} from 'lucide-react';
import {
  PieChart, Pie, Cell, BarChart, Bar, XAxis, YAxis,
  CartesianGrid, Tooltip, ResponsiveContainer, Legend
} from 'recharts';
import { QRCodeSVG } from 'qrcode.react';
import QRCode from 'qrcode'; // Library baru untuk generate gambar QR ke PDF
import { Html5QrcodeScanner } from 'html5-qrcode';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';

// --- 1. DATA BAWAAN ---
const defaultWarga = [
  "Bpk. Ahmad", "Ibu Siti", "Bpk. Budi", "Ibu Ani", "Bpk. Joko", "Ibu Rina", "Bpk. Agus", "Ibu Sri",
"Bpk. Hendra", "Ibu Maya", "Bpk. Dedi", "Ibu Nina", "Bpk. Yudi", "Ibu Dewi", "Bpk. Eko", "Ibu Ratna",
"Bpk. Iwan", "Ibu Sari", "Bpk. Arif", "Ibu Lia", "Bpk. Rizal", "Ibu Ayu", "Bpk. Anton", "Ibu Fitri",
"Bpk. Wahyu", "Ibu Dian", "Bpk. Indra", "Ibu Tari", "Bpk. Gilang", "Ibu Wati", "Bpk. Rahmat"
].map((nama, i) => ({
  id: `W-${String(i + 1).padStart(3, '0')}`,
                    nama: nama,
                    rt: `RT 0${(i % 5) + 1}`,
}));

const defaultMudhohi = [
  { id: 'MD-01', nama: 'Kel. Bpk. H. Abdullah', hewan: 'Sapi', wa: '081234567890' },
{ id: 'MD-02', nama: 'Bpk. Budi Santoso', hewan: 'Kambing', wa: '081311223344' },
{ id: 'MD-03', nama: 'Ibu Siti Khadijah', hewan: 'Kambing', wa: '081322334455' },
{ id: 'MD-04', nama: 'Sdr. Andi Pratama', hewan: 'Kambing', wa: '081333445566' },
{ id: 'MD-05', nama: 'Bpk. Joko Susilo', hewan: 'Kambing', wa: '081344556677' }
];

const defaultPanitia = ["H. Fulan", "Ust. Hasan", "Bpk. Zainal", "Bpk. Anwar", "Bpk. Somad", "Bpk. Qasim", "Sdr. Ilham", "Sdr. Reza", "Bpk. Yasin", "Sdr. Iqbal"].map((nama, i) => ({
  id: `P-${String(i + 1).padStart(2, '0')}`,
                                                                                                                                                                                    nama: nama,
                                                                                                                                                                                    peran: i === 0 ? "Ketua Panitia" : "Anggota",
                                                                                                                                                                                    kontak: `0811002200${i}`
}));

const defaultHewan = [
  { id: "SAPI-01", jenis: 'Sapi', bobot: '350 kg', status: 'Selesai' },
{ id: "SAPI-02", jenis: 'Sapi', bobot: '320 kg', status: 'Dikuliti' },
{ id: "KMBG-01", jenis: 'Kambing', bobot: '25 kg', status: 'Selesai' },
{ id: "KMBG-02", jenis: 'Kambing', bobot: '28 kg', status: 'Selesai' },
{ id: "KMBG-03", jenis: 'Kambing', bobot: '24 kg', status: 'Disembelih' },
{ id: "KMBG-04", jenis: 'Kambing', bobot: '26 kg', status: 'Disembelih' },
{ id: "KMBG-05", jenis: 'Kambing', bobot: '27 kg', status: 'Menunggu' },
{ id: "KMBG-06", jenis: 'Kambing', bobot: '25 kg', status: 'Menunggu' },
{ id: "KMBG-07", jenis: 'Kambing', bobot: '30 kg', status: 'Menunggu' },
{ id: "KMBG-08", jenis: 'Kambing', bobot: '22 kg', status: 'Menunggu' },
{ id: "KMBG-09", jenis: 'Kambing', bobot: '29 kg', status: 'Menunggu' },
{ id: "KMBG-10", jenis: 'Kambing', bobot: '25 kg', status: 'Menunggu' },
{ id: "KMBG-11", jenis: 'Kambing', bobot: '23 kg', status: 'Menunggu' },
{ id: "KMBG-12", jenis: 'Kambing', bobot: '26 kg', status: 'Menunggu' }
];

const defaultKeuangan = [
  { id: 1, tanggal: '28/05/2026', keterangan: 'Saldo Kas Masjid & Donasi Jamaah', jenis: 'Masuk', nominal: 52500000 },
{ id: 2, tanggal: '01/06/2026', keterangan: 'DP Tim Jagal', jenis: 'Keluar', nominal: 32500000 }
];

const generateKupon = () => {
  const kuponWarga = defaultWarga.map(w => ({ id: `KP-W-${w.id}`, nama: w.nama, tipe: 'Warga', status: 'Belum Diambil' }));
  const kuponMudhohi = defaultMudhohi.map(m => ({ id: `KP-M-${m.id}`, nama: m.nama, tipe: 'Mudhohi', status: 'Belum Diambil' }));
  return [...kuponWarga, ...kuponMudhohi];
};

const menuItems = [
  { id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
{ id: 'warga', label: 'Data Warga', icon: Users },
{ id: 'panitia', label: 'Data Panitia', icon: UserCog },
{ id: 'mudhohi', label: 'Mudhohi', icon: HeartHandshake },
{ id: 'hewan', label: 'Hewan Qurban', icon: Beef },
{ id: 'scan', label: 'Scan Kupon', icon: ScanLine },
{ id: 'kupon', label: 'Kupon Digital', icon: Ticket },
{ id: 'sertifikat', label: 'Sertifikat', icon: Award },
{ id: 'uang', label: 'Laporan RAB', icon: LineChart },
];

export default function App() {
  const [activeTab, setActiveTab] = useState('dashboard');

  // --- 2. STATES DATABASE ---
  const [warga, setWarga] = useState(() => JSON.parse(localStorage.getItem('sys_warga')) || defaultWarga);
  const [panitia, setPanitia] = useState(() => JSON.parse(localStorage.getItem('sys_panitia')) || defaultPanitia);
  const [mudhohi, setMudhohi] = useState(() => JSON.parse(localStorage.getItem('sys_mudhohi')) || defaultMudhohi);
  const [hewan, setHewan] = useState(() => JSON.parse(localStorage.getItem('sys_hewan_v5')) || defaultHewan);
  const [kupon, setKupon] = useState(() => JSON.parse(localStorage.getItem('sys_kupon')) || generateKupon());
  const [keuangan, setKeuangan] = useState(() => JSON.parse(localStorage.getItem('sys_keuangan')) || defaultKeuangan);

  const [formHewan, setFormHewan] = useState({ jenis: 'Sapi', bobot: '' });
  const [formUang, setFormUang] = useState({ keterangan: '', jenis: 'Masuk', nominal: '' });

  const [qrModal, setQrModal] = useState({ isOpen: false, data: null });
  const [pdfPreview, setPdfPreview] = useState({ isOpen: false, url: '', title: '' });

  const [scanResult, setScanResult] = useState(null);
  const [manualCode, setManualCode] = useState('');

  useEffect(() => { localStorage.setItem('sys_warga', JSON.stringify(warga)); }, [warga]);
  useEffect(() => { localStorage.setItem('sys_panitia', JSON.stringify(panitia)); }, [panitia]);
  useEffect(() => { localStorage.setItem('sys_mudhohi', JSON.stringify(mudhohi)); }, [mudhohi]);
  useEffect(() => { localStorage.setItem('sys_hewan_v5', JSON.stringify(hewan)); }, [hewan]);
  useEffect(() => { localStorage.setItem('sys_kupon', JSON.stringify(kupon)); }, [kupon]);
  useEffect(() => { localStorage.setItem('sys_keuangan', JSON.stringify(keuangan)); }, [keuangan]);

  const totalSapi = hewan.filter(h => h.jenis === 'Sapi').length;
  const totalKambing = hewan.filter(h => h.jenis === 'Kambing').length;
  const totalMasuk = keuangan.filter(k => k.jenis === 'Masuk').reduce((acc, curr) => acc + curr.nominal, 0);
  const totalKeluar = keuangan.filter(k => k.jenis === 'Keluar').reduce((acc, curr) => acc + curr.nominal, 0);
  const saldo = totalMasuk - totalKeluar;

  const pieData = [
    { name: 'Selesai', value: hewan.filter(h=>h.status==='Selesai').length, color: '#10b981' },
    { name: 'Dikuliti', value: hewan.filter(h=>h.status==='Dikuliti').length, color: '#8b5cf6' },
    { name: 'Disembelih', value: hewan.filter(h=>h.status==='Disembelih').length, color: '#f59e0b' },
    { name: 'Menunggu', value: hewan.filter(h=>h.status==='Menunggu').length, color: '#94a3b8' },
  ].filter(d => d.value > 0);

  const barData = [
    { name: 'Masuk', Pemasukan: totalMasuk },
    { name: 'Keluar', Pengeluaran: totalKeluar },
  ];

  // --- 3. HANDLERS CRUD & EXPORT ---
  const simpanHewan = (e) => {
    e.preventDefault();
    const prefix = formHewan.jenis === 'Sapi' ? 'SAPI' : 'KMBG';
    const newId = `${prefix}-${String(hewan.filter(h=>h.jenis===formHewan.jenis).length + 1).padStart(2, '0')}`;
    setHewan([{ id: newId, jenis: formHewan.jenis, bobot: formHewan.bobot + ' kg', status: 'Menunggu' }, ...hewan]);
    setFormHewan({ jenis: 'Sapi', bobot: '' });
  };

  const updateStatusHewan = (id, newStatus) => {
    setHewan(hewan.map(h => h.id === id ? { ...h, status: newStatus } : h));
  };

  const simpanUang = (e) => {
    e.preventDefault();
    const newTx = { id: Date.now(), tanggal: new Date().toLocaleDateString('id-ID'), keterangan: formUang.keterangan, jenis: formUang.jenis, nominal: parseInt(formUang.nominal) };
    setKeuangan([...keuangan, newTx]);
    setFormUang({ keterangan: '', jenis: 'Masuk', nominal: '' });
  };

  const exportExcel = () => {
    const ws = XLSX.utils.json_to_sheet(keuangan);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "RAB");
    XLSX.writeFile(wb, "Laporan_RAB_Kurban.xlsx");
  };

  const exportPDF = () => {
    const doc = new jsPDF();
    doc.setFontSize(16); doc.text("Laporan Keuangan Panitia Kurban", 14, 20); doc.setFontSize(10);
    let y = 35;
    keuangan.forEach((k, i) => { doc.text(`${i+1}. ${k.tanggal} | ${k.keterangan} | ${k.jenis} | Rp ${k.nominal.toLocaleString('id-ID')}`, 14, y); y += 8; });
    doc.text(`Total Saldo Akhir: Rp ${saldo.toLocaleString('id-ID')}`, 14, y + 10);
    doc.save("Laporan_RAB.pdf");
  };

  // Unduh Kupon + Embed Gambar QR Code
  const prosesKuponDigital = async (k, isPreview = false) => {
    const doc = new jsPDF({ format: [80, 80] });

    // Border & Title
    doc.setDrawColor(16, 185, 129); doc.setLineWidth(1); doc.rect(2, 2, 76, 76);
    doc.setFontSize(14); doc.setTextColor(16, 185, 129); doc.text("KUPON QURBAN", 40, 12, { align: "center" });
    doc.setFontSize(9); doc.setTextColor(50, 50, 50); doc.text(`Tipe: ${k.tipe}`, 40, 18, { align: "center" });

    // Generate QR Code as Base64 Image
    try {
      const qrDataUrl = await QRCode.toDataURL(k.id, { margin: 1, width: 100 });
      // Paste the image into PDF (format, x, y, width, height)
      doc.addImage(qrDataUrl, 'PNG', 25, 22, 30, 30);
    } catch (err) {
      console.error("Gagal generate QR Code PDF", err);
    }

    // Identitas
    doc.setFontSize(11); doc.setTextColor(0, 0, 0); doc.text(k.nama, 40, 58, { align: "center" });
    doc.setFontSize(9); doc.setTextColor(50, 50, 50); doc.text(`ID: ${k.id}`, 40, 65, { align: "center" });
    doc.setFontSize(7); doc.setTextColor(100, 100, 100); doc.text("Tunjukkan kupon ini kepada Panitia", 40, 74, { align: "center" });

    if(isPreview) {
      setPdfPreview({ isOpen: true, url: doc.output('bloburl'), title: `Kupon ${k.id}` });
    } else {
      doc.save(`Kupon_${k.id}.pdf`);
    }
  };

  const prosesSertifikatMudhohi = (m, isPreview = false) => {
    const doc = new jsPDF('landscape');
    doc.setDrawColor(16, 185, 129); doc.setLineWidth(3); doc.rect(10, 10, 277, 190);
    doc.setFontSize(30); doc.setTextColor(16, 185, 129); doc.text("SERTIFIKAT QURBAN", 148, 50, { align: "center" });
    doc.setFontSize(14); doc.setTextColor(50, 50, 50); doc.text("Diberikan dengan penuh rasa syukur kepada:", 148, 80, { align: "center" });
    doc.setFontSize(28); doc.setTextColor(15, 23, 42); doc.text(m.nama, 148, 105, { align: "center" });
    doc.setFontSize(14); doc.setTextColor(50, 50, 50); doc.text(`Atas partisipasinya menyerahkan hewan qurban berupa ${m.hewan}.`, 148, 125, { align: "center" });
    doc.text("Semoga Allah SWT menerima amal ibadah qurban ini. Amin.", 148, 135, { align: "center" });

    if(isPreview) {
      setPdfPreview({ isOpen: true, url: doc.output('bloburl'), title: `Sertifikat ${m.nama}` });
    } else {
      doc.save(`Sertifikat_Mudhohi_${m.id}.pdf`);
    }
  };

  const prosesSertifikatPanitia = (p, isPreview = false) => {
    const doc = new jsPDF('landscape');
    doc.setDrawColor(59, 130, 246); doc.setLineWidth(3); doc.rect(10, 10, 277, 190);
    doc.setFontSize(30); doc.setTextColor(59, 130, 246); doc.text("PIAGAM PENGHARGAAN", 148, 50, { align: "center" });
    doc.setFontSize(14); doc.setTextColor(50, 50, 50); doc.text("Apresiasi setinggi-tingginya diberikan kepada:", 148, 80, { align: "center" });
    doc.setFontSize(28); doc.setTextColor(15, 23, 42); doc.text(p.nama, 148, 105, { align: "center" });
    doc.setFontSize(14); doc.setTextColor(50, 50, 50); doc.text(`Atas dedikasi dan kerja kerasnya bertugas sebagai ${p.peran}`, 148, 125, { align: "center" });
    doc.text("Pada penyelenggaraan ibadah Qurban 1447 H.", 148, 135, { align: "center" });

    if(isPreview) {
      setPdfPreview({ isOpen: true, url: doc.output('bloburl'), title: `Piagam ${p.nama}` });
    } else {
      doc.save(`Piagam_Panitia_${p.id}.pdf`);
    }
  };

  // --- LOGIKA SCANNER KAMERA ---
  useEffect(() => {
    if (activeTab === 'scan') {
      const scanner = new Html5QrcodeScanner("reader", { qrbox: { width: 250, height: 250 }, fps: 5 });
      scanner.render((decodedText) => prosesValidasiKupon(decodedText), (error) => {});
      return () => { scanner.clear().catch(e => console.error(e)); };
    }
  }, [activeTab]);

  const prosesValidasiKupon = (idKupon) => {
    const kuponDitemukan = kupon.find(k => k.id === idKupon);
    if (!kuponDitemukan) return setScanResult({ status: 'error', pesan: `Kupon ${idKupon} tidak terdaftar!` });
    if (kuponDitemukan.status === 'Sudah Diambil') return setScanResult({ status: 'error', pesan: `Kupon ${idKupon} atas nama ${kuponDitemukan.nama} SUDAH DIAMBIL!` });

    setKupon(kupon.map(k => k.id === idKupon ? { ...k, status: 'Sudah Diambil' } : k));
    setScanResult({ status: 'success', pesan: `Berhasil! Kupon ${idKupon} (${kuponDitemukan.tipe}) divalidasi.` });
    setManualCode('');
  };

  const handleManualScan = (e) => { e.preventDefault(); if(manualCode.trim() !== '') prosesValidasiKupon(manualCode.trim()); };

  return (
    <div className="flex h-screen bg-slate-50 font-sans overflow-hidden">

    {/* --- SIDEBAR --- */}
    <aside className="w-64 bg-[#059669] text-white flex flex-col shadow-xl z-20 overflow-hidden">
    <div className="p-6 flex items-center gap-3 border-b border-emerald-700/50">
    <div className="bg-amber-400 p-2 rounded-lg text-emerald-900 font-bold">Q</div>
    <div>
    <h1 className="text-lg font-bold leading-tight">QurbanPro</h1>
    <p className="text-xs text-emerald-100">Manajemen Qurban</p>
    </div>
    </div>

    <div className="flex-1 overflow-y-auto py-4 px-3 space-y-1 scrollbar-hide">
    {menuItems.map((item) => {
      const Icon = item.icon;
      const isActive = activeTab === item.id;

      return (
        <button
        key={item.id}
        onClick={() => { setActiveTab(item.id); setScanResult(null); }}
        className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all duration-200 ${
          isActive ? 'bg-white text-[#059669] font-bold shadow-sm' : 'text-emerald-50 hover:bg-emerald-800/40 hover:text-white'
        }`}
        >
        <Icon size={18} strokeWidth={isActive ? 2.5 : 2} />
        <span className="text-sm tracking-wide">{item.label}</span>
        </button>
      );
    })}
    </div>
    </aside>

    {/* --- MAIN CONTENT --- */}
    <main className="flex-1 flex flex-col h-full overflow-y-auto relative">
    <header className="bg-white border-b border-slate-200 px-8 py-4 flex justify-between items-center sticky top-0 z-10">
    <div>
    <h2 className="text-xl font-bold text-slate-800 capitalize">
    {menuItems.find(m => m.id === activeTab)?.label || 'Menu'}
    </h2>
    <p className="text-sm text-slate-500">Sistem Informasi Pengelola Real-Time</p>
    </div>
    <div className="flex items-center gap-2">
    <span className="bg-emerald-100 text-emerald-800 px-3 py-1 rounded-full text-xs font-bold flex items-center gap-2">
    <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span> Sistem Aktif
    </span>
    </div>
    </header>

    <div className="p-8">

    {/* TAB 1: DASHBOARD */}
    {activeTab === 'dashboard' && (
      <div className="space-y-6 animate-in fade-in duration-500">
      <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
      <Card title="Total Hewan Qurban" value={hewan.length} sub={`${totalSapi} Sapi, ${totalKambing} Kambing`} icon={<Beef size={20} className="text-emerald-500"/>} border="border-l-4 border-l-emerald-500" />
      <Card title="Kupon Diclaim" value={kupon.filter(k=>k.status==='Sudah Diambil').length} sub={`Dari total ${kupon.length} kupon`} icon={<ScanLine size={20} className="text-blue-500"/>} border="border-l-4 border-l-blue-500" />
      <Card title="Mudhohi" value={mudhohi.length} sub="Peserta kurban" icon={<HeartHandshake size={20} className="text-amber-500"/>} border="border-l-4 border-l-amber-500" />
      <Card title="Saldo Kas" value={`Rp ${(saldo/1000000).toFixed(1)}jt`} sub={`Masuk: Rp ${(totalMasuk/1000000).toFixed(1)}jt`} icon={<span className="text-purple-500 font-bold">$</span>} border="border-l-4 border-l-purple-500" />
      </div>
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
      <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm flex flex-col">
      <h4 className="font-semibold text-slate-800 mb-2">Status Penyembelihan Hewan</h4>
      <div className="flex-1 min-h-[250px]">
      <ResponsiveContainer width="100%" height="100%">
      <PieChart>
      <Pie data={pieData} cx="50%" cy="50%" innerRadius={60} outerRadius={90} paddingAngle={2} dataKey="value" label={({name, value}) => `${name}: ${value}`}>
      {pieData.map((entry, index) => <Cell key={`cell-${index}`} fill={entry.color} />)}
      </Pie>
      <Tooltip />
      <Legend />
      </PieChart>
      </ResponsiveContainer>
      </div>
      </div>
      <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
      <h4 className="font-semibold text-slate-800 mb-6">Arus Kas (RAB)</h4>
      <div className="h-64">
      <ResponsiveContainer width="100%" height="100%">
      <BarChart data={barData} barSize={60}>
      <CartesianGrid strokeDasharray="3 3" vertical={false} />
      <XAxis dataKey="name" axisLine={false} tickLine={false} />
      <YAxis axisLine={false} tickLine={false} tickFormatter={(value) => `${value / 1000000}jt`} />
      <Tooltip cursor={{fill: 'transparent'}} formatter={(value) => `Rp ${value.toLocaleString('id-ID')}`}/>
      <Bar dataKey="Pemasukan" fill="#10b981" radius={[4, 4, 0, 0]} />
      <Bar dataKey="Pengeluaran" fill="#ef4444" radius={[4, 4, 0, 0]} />
      </BarChart>
      </ResponsiveContainer>
      </div>
      </div>
      </div>
      </div>
    )}

    {/* TAB 2: DATA WARGA */}
    {activeTab === 'warga' && (
      <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm animate-in fade-in duration-300">
      <h3 className="font-bold text-xl mb-6">Database Warga (Penerima Daging)</h3>
      <div className="overflow-x-auto max-h-[60vh]">
      <table className="w-full text-left text-sm border-collapse">
      <thead className="bg-slate-900 text-white sticky top-0">
      <tr><th className="p-3">ID Warga</th><th className="p-3">Nama Lengkap</th><th className="p-3">Wilayah</th></tr>
      </thead>
      <tbody className="divide-y divide-slate-100">
      {warga.map(w => (
        <tr key={w.id} className="hover:bg-slate-50">
        <td className="p-3 font-mono font-bold text-slate-600">{w.id}</td>
        <td className="p-3 font-semibold">{w.nama}</td>
        <td className="p-3">{w.rt}</td>
        </tr>
      ))}
      </tbody>
      </table>
      </div>
      </div>
    )}

    {/* TAB 3: DATA PANITIA */}
    {activeTab === 'panitia' && (
      <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm animate-in fade-in duration-300">
      <h3 className="font-bold text-xl mb-6">Susunan Panitia</h3>
      <table className="w-full text-left text-sm border-collapse">
      <thead className="bg-slate-900 text-white">
      <tr><th className="p-3">ID Panitia</th><th className="p-3">Nama Lengkap</th><th className="p-3">Peran / Jabatan</th><th className="p-3">Kontak WA</th></tr>
      </thead>
      <tbody className="divide-y divide-slate-100">
      {panitia.map(p => (
        <tr key={p.id} className="hover:bg-slate-50">
        <td className="p-3 font-mono font-bold text-slate-600">{p.id}</td>
        <td className="p-3 font-semibold">{p.nama}</td>
        <td className="p-3"><span className="bg-amber-50 text-amber-600 px-2 py-1 rounded text-xs font-bold border border-amber-100">{p.peran}</span></td>
        <td className="p-3 text-slate-500">{p.kontak}</td>
        </tr>
      ))}
      </tbody>
      </table>
      </div>
    )}

    {/* TAB 4: MUDHOHI */}
    {activeTab === 'mudhohi' && (
      <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm animate-in fade-in duration-300">
      <h3 className="font-bold text-xl mb-6">Data Mudhohi (Peserta Kurban)</h3>
      <div className="overflow-x-auto max-h-[60vh]">
      <table className="w-full text-left text-sm border-collapse">
      <thead className="bg-slate-900 text-white sticky top-0">
      <tr><th className="p-3">ID Mudhohi</th><th className="p-3">Nama / Kelompok</th><th className="p-3">Jenis Hewan</th><th className="p-3">Kontak WA</th></tr>
      </thead>
      <tbody className="divide-y divide-slate-100">
      {mudhohi.map(m => (
        <tr key={m.id} className="hover:bg-slate-50">
        <td className="p-3 font-mono font-bold text-slate-600">{m.id}</td>
        <td className="p-3 font-semibold">{m.nama}</td>
        <td className="p-3"><span className="bg-slate-200 px-2 py-1 rounded text-xs font-bold">{m.hewan}</span></td>
        <td className="p-3">{m.wa}</td>
        </tr>
      ))}
      </tbody>
      </table>
      </div>
      </div>
    )}

    {/* TAB 5: HEWAN QURBAN */}
    {activeTab === 'hewan' && (
      <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm animate-in fade-in duration-300">
      <h3 className="font-bold text-xl mb-4">Pencatatan Hewan Qurban</h3>
      <form onSubmit={simpanHewan} className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-6 bg-slate-50 p-4 rounded-lg">
      <select value={formHewan.jenis} onChange={e=>setFormHewan({...formHewan, jenis: e.target.value})} className="col-span-1 p-2.5 border rounded outline-none focus:ring-2 focus:ring-emerald-500">
      <option value="Sapi">Sapi</option>
      <option value="Kambing">Kambing</option>
      </select>
      <input type="number" placeholder="Perkiraan Bobot (Angka)" value={formHewan.bobot} onChange={e=>setFormHewan({...formHewan, bobot: e.target.value})} className="col-span-2 p-2.5 border rounded outline-none focus:ring-2 focus:ring-emerald-500" required />
      <button type="submit" className="bg-[#059669] text-white p-2.5 rounded font-bold hover:bg-[#047857] flex justify-center items-center gap-2"><Plus size={16}/> Tambah Hewan</button>
      </form>

      <div className="overflow-x-auto max-h-[50vh]">
      <table className="w-full text-left text-sm border-collapse">
      <thead className="bg-slate-900 text-white sticky top-0">
      <tr>
      <th className="p-3">ID Hewan</th>
      <th className="p-3">Jenis</th>
      <th className="p-3">Bobot</th>
      <th className="p-3">Status Pemotongan</th>
      <th className="p-3 text-center">Aksi (QR Code)</th>
      </tr>
      </thead>
      <tbody className="divide-y divide-slate-100">
      {hewan.map(h => (
        <tr key={h.id} className="hover:bg-slate-50">
        <td className="p-3 font-mono font-bold text-slate-800">{h.id}</td>
        <td className="p-3"><span className="bg-slate-200 px-2 py-1 rounded text-xs font-bold">{h.jenis}</span></td>
        <td className="p-3 font-medium text-slate-600">{h.bobot}</td>
        <td className="p-3">
        <select value={h.status} onChange={(e) => updateStatusHewan(h.id, e.target.value)} className="p-1 border rounded text-xs font-bold outline-none cursor-pointer">
        <option value="Menunggu">Menunggu</option>
        <option value="Disembelih">Disembelih</option>
        <option value="Dikuliti">Dikuliti</option>
        <option value="Selesai">Selesai</option>
        </select>
        </td>
        <td className="p-3 flex justify-center">
        <button onClick={() => setQrModal({ isOpen: true, data: { id: h.id, nama: `Identitas ${h.jenis}` } })} className="bg-slate-800 text-white px-3 py-1.5 rounded flex items-center gap-2 hover:bg-slate-900 text-xs">
        <QrCode size={14} /> Label QR
        </button>
        </td>
        </tr>
      ))}
      </tbody>
      </table>
      </div>
      </div>
    )}

    {/* TAB 6: SCAN KUPON */}
    {activeTab === 'scan' && (
      <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm animate-in fade-in duration-300">
      <h3 className="font-bold text-xl mb-6">Validasi & Scan Kupon</h3>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
      <div className="flex flex-col items-center border border-slate-200 p-4 rounded-xl bg-slate-50">
      <p className="font-bold text-slate-700 mb-4 flex items-center gap-2"><ScanLine size={18}/> Arahkan QR Code ke Kamera</p>
      <div id="reader" className="w-full max-w-sm rounded-lg overflow-hidden border-2 border-dashed border-emerald-500 bg-black"></div>
      </div>

      <div className="flex flex-col justify-center space-y-6">
      <div className="bg-slate-50 p-6 rounded-xl border border-slate-200">
      <p className="font-bold text-slate-700 mb-3">Atau Input ID Kupon Manual:</p>
      <form onSubmit={handleManualScan} className="flex gap-2">
      <input type="text" value={manualCode} onChange={(e)=>setManualCode(e.target.value)} placeholder="Contoh: KP-W-W-001" className="flex-1 p-3 border border-slate-300 rounded-lg outline-none focus:ring-2 focus:ring-emerald-500 font-mono" />
      <button type="submit" className="bg-[#059669] text-white px-6 py-3 rounded-lg font-bold hover:bg-[#047857]">Validasi</button>
      </form>
      </div>

      {scanResult && (
        <div className={`p-4 rounded-xl border flex gap-4 items-start animate-in zoom-in duration-300 ${scanResult.status === 'success' ? 'bg-emerald-50 border-emerald-200 text-emerald-800' : 'bg-red-50 border-red-200 text-red-800'}`}>
        {scanResult.status === 'success' ? <CheckCircle className="mt-1 flex-shrink-0" /> : <AlertCircle className="mt-1 flex-shrink-0" />}
        <div>
        <h4 className="font-bold">{scanResult.status === 'success' ? 'Validasi Berhasil' : 'Validasi Gagal'}</h4>
        <p className="text-sm mt-1">{scanResult.pesan}</p>
        </div>
        </div>
      )}
      </div>
      </div>
      </div>
    )}

    {/* TAB 7: KUPON DIGITAL */}
    {activeTab === 'kupon' && (
      <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm animate-in fade-in duration-300">
      <h3 className="font-bold text-xl mb-2">Manajemen Kupon Digital</h3>
      <p className="text-slate-500 text-sm mb-6">Daftar seluruh kupon, bisa ditampilkan QR-nya atau diunduh PDF.</p>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
      <div>
      <h4 className="font-bold text-emerald-700 bg-emerald-50 p-3 rounded-t-lg border-b-2 border-emerald-500">Daftar Kupon Warga</h4>
      <div className="overflow-y-auto max-h-[50vh] border border-slate-200 rounded-b-lg">
      <table className="w-full text-left text-sm">
      <thead className="bg-slate-100 text-slate-700 sticky top-0">
      <tr><th className="p-3">ID Kupon</th><th className="p-3">Pemilik</th><th className="p-3">Status</th><th className="p-3">Aksi</th></tr>
      </thead>
      <tbody className="divide-y divide-slate-100">
      {kupon.filter(k=>k.tipe==='Warga').map(k => (
        <tr key={k.id} className="hover:bg-slate-50">
        <td className="p-3 font-mono text-xs">{k.id}</td>
        <td className="p-3 font-medium">{k.nama}</td>
        <td className="p-3"><span className={`px-2 py-1 rounded text-[10px] font-bold ${k.status==='Sudah Diambil'?'bg-emerald-100 text-emerald-700':'bg-orange-100 text-orange-700'}`}>{k.status}</span></td>
        <td className="p-3 flex gap-2">
        <button onClick={()=>prosesKuponDigital(k, true)} className="bg-slate-100 text-slate-600 p-1.5 rounded hover:bg-slate-200" title="Preview"><Eye size={14}/></button>
        <button onClick={()=>prosesKuponDigital(k, false)} className="bg-emerald-100 text-emerald-600 p-1.5 rounded hover:bg-emerald-200" title="Unduh Kupon"><DownloadCloud size={14}/></button>
        </td>
        </tr>
      ))}
      </tbody>
      </table>
      </div>
      </div>

      <div>
      <h4 className="font-bold text-blue-700 bg-blue-50 p-3 rounded-t-lg border-b-2 border-blue-500">Daftar Kupon Mudhohi</h4>
      <div className="overflow-y-auto max-h-[50vh] border border-slate-200 rounded-b-lg">
      <table className="w-full text-left text-sm">
      <thead className="bg-slate-100 text-slate-700 sticky top-0">
      <tr><th className="p-3">ID Kupon</th><th className="p-3">Pemilik</th><th className="p-3">Status</th><th className="p-3">Aksi</th></tr>
      </thead>
      <tbody className="divide-y divide-slate-100">
      {kupon.filter(k=>k.tipe==='Mudhohi').map(k => (
        <tr key={k.id} className="hover:bg-slate-50">
        <td className="p-3 font-mono text-xs">{k.id}</td>
        <td className="p-3 font-medium">{k.nama}</td>
        <td className="p-3"><span className={`px-2 py-1 rounded text-[10px] font-bold ${k.status==='Sudah Diambil'?'bg-emerald-100 text-emerald-700':'bg-orange-100 text-orange-700'}`}>{k.status}</span></td>
        <td className="p-3 flex gap-2">
        <button onClick={()=>prosesKuponDigital(k, true)} className="bg-slate-100 text-slate-600 p-1.5 rounded hover:bg-slate-200" title="Preview"><Eye size={14}/></button>
        <button onClick={()=>prosesKuponDigital(k, false)} className="bg-emerald-100 text-emerald-600 p-1.5 rounded hover:bg-emerald-200" title="Unduh Kupon"><DownloadCloud size={14}/></button>
        </td>
        </tr>
      ))}
      </tbody>
      </table>
      </div>
      </div>
      </div>
      </div>
    )}

    {/* TAB 8: SERTIFIKAT DENGAN PREVIEW */}
    {activeTab === 'sertifikat' && (
      <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm animate-in fade-in duration-300">
      <h3 className="font-bold text-xl mb-2">Cetak Sertifikat & Piagam</h3>
      <p className="text-slate-500 text-sm mb-6">Pratinjau atau unduh sertifikat untuk Mudhohi dan Piagam untuk Panitia.</p>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
      <div>
      <h4 className="font-bold text-emerald-700 bg-emerald-50 p-3 rounded-t-lg border-b-2 border-emerald-500">Sertifikat Mudhohi</h4>
      <div className="overflow-y-auto max-h-[50vh] border border-slate-200 rounded-b-lg">
      <table className="w-full text-left text-sm">
      <tbody className="divide-y divide-slate-100">
      {mudhohi.map(m => (
        <tr key={m.id} className="hover:bg-slate-50 flex justify-between items-center p-3">
        <div>
        <p className="font-bold text-slate-800">{m.nama}</p>
        <p className="text-xs text-slate-500">Qurban: {m.hewan}</p>
        </div>
        <div className="flex gap-2">
        <button onClick={()=>prosesSertifikatMudhohi(m, true)} className="bg-slate-100 text-slate-600 px-3 py-1.5 rounded-lg text-xs font-bold hover:bg-slate-200 flex items-center gap-1">
        <Eye size={14}/> Preview
        </button>
        <button onClick={()=>prosesSertifikatMudhohi(m, false)} className="bg-emerald-600 text-white px-3 py-1.5 rounded-lg text-xs font-bold hover:bg-emerald-700 flex items-center gap-1">
        <DownloadCloud size={14}/> Unduh
        </button>
        </div>
        </tr>
      ))}
      </tbody>
      </table>
      </div>
      </div>

      <div>
      <h4 className="font-bold text-blue-700 bg-blue-50 p-3 rounded-t-lg border-b-2 border-blue-500">Piagam Panitia</h4>
      <div className="overflow-y-auto max-h-[50vh] border border-slate-200 rounded-b-lg">
      <table className="w-full text-left text-sm">
      <tbody className="divide-y divide-slate-100">
      {panitia.map(p => (
        <tr key={p.id} className="hover:bg-slate-50 flex justify-between items-center p-3">
        <div>
        <p className="font-bold text-slate-800">{p.nama}</p>
        <p className="text-xs text-slate-500">Jabatan: {p.peran}</p>
        </div>
        <div className="flex gap-2">
        <button onClick={()=>prosesSertifikatPanitia(p, true)} className="bg-slate-100 text-slate-600 px-3 py-1.5 rounded-lg text-xs font-bold hover:bg-slate-200 flex items-center gap-1">
        <Eye size={14}/> Preview
        </button>
        <button onClick={()=>prosesSertifikatPanitia(p, false)} className="bg-blue-600 text-white px-3 py-1.5 rounded-lg text-xs font-bold hover:bg-blue-700 flex items-center gap-1">
        <DownloadCloud size={14}/> Unduh
        </button>
        </div>
        </tr>
      ))}
      </tbody>
      </table>
      </div>
      </div>
      </div>
      </div>
    )}

    {/* TAB 9: KEUANGAN (RAB) */}
    {activeTab === 'uang' && (
      <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm animate-in fade-in duration-300">
      <div className="flex justify-between items-center mb-6">
      <h3 className="font-bold text-xl">Laporan RAB Transparan</h3>
      <div className="flex gap-2">
      <button onClick={exportExcel} className="bg-emerald-600 text-white px-4 py-2 rounded text-sm font-bold flex items-center gap-2 hover:bg-emerald-700"><DownloadCloud size={16}/> Excel</button>
      <button onClick={exportPDF} className="bg-red-500 text-white px-4 py-2 rounded text-sm font-bold flex items-center gap-2 hover:bg-red-600"><DownloadCloud size={16}/> PDF</button>
      </div>
      </div>

      <form onSubmit={simpanUang} className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-6 bg-slate-50 p-4 rounded-lg">
      <input type="text" placeholder="Uraian Transaksi" value={formUang.keterangan} onChange={e=>setFormUang({...formUang, keterangan: e.target.value})} className="col-span-2 p-2.5 border rounded outline-none focus:ring-2 focus:ring-emerald-500" required />
      <select value={formUang.jenis} onChange={e=>setFormUang({...formUang, jenis: e.target.value})} className="p-2.5 border rounded outline-none focus:ring-2 focus:ring-emerald-500">
      <option value="Masuk">Masuk (Debit)</option>
      <option value="Keluar">Keluar (Kredit)</option>
      </select>
      <input type="number" placeholder="Nominal Rp" value={formUang.nominal} onChange={e=>setFormUang({...formUang, nominal: e.target.value})} className="p-2.5 border rounded outline-none focus:ring-2 focus:ring-emerald-500" required />
      <div className="col-span-4 flex justify-end"><button type="submit" className="bg-[#059669] text-white px-6 py-2.5 rounded font-bold hover:bg-[#047857] flex items-center gap-2"><Plus size={16}/> Simpan Transaksi</button></div>
      </form>

      <table className="w-full text-left text-sm border-collapse">
      <thead className="bg-slate-100 text-slate-700">
      <tr><th className="p-3">Tanggal</th><th className="p-3">Uraian</th><th className="p-3">Jenis</th><th className="p-3 text-right">Nominal</th></tr>
      </thead>
      <tbody className="divide-y divide-slate-100">
      {keuangan.map((k) => (
        <tr key={k.id} className="hover:bg-slate-50">
        <td className="p-3">{k.tanggal}</td>
        <td className="p-3">{k.keterangan}</td>
        <td className="p-3"><span className={`px-2 py-1 rounded text-xs font-bold ${k.jenis==='Masuk'?'bg-emerald-100 text-emerald-700':'bg-red-100 text-red-700'}`}>{k.jenis}</span></td>
        <td className={`p-3 text-right font-bold ${k.jenis==='Masuk'?'text-emerald-600':'text-red-600'}`}>Rp {k.nominal.toLocaleString('id-ID')}</td>
        </tr>
      ))}
      </tbody>
      <tfoot className="bg-slate-900 text-white font-bold text-lg">
      <tr><td colSpan="3" className="p-4 text-right">TOTAL SALDO:</td><td className="p-4 text-right text-emerald-400">Rp {saldo.toLocaleString('id-ID')}</td></tr>
      </tfoot>
      </table>
      </div>
    )}
    </div>
    </main>

    {/* --- MODAL QR CODE --- */}
    {qrModal.isOpen && (
      <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center z-50 p-4">
      <div className="bg-white p-8 rounded-2xl shadow-2xl max-w-sm w-full text-center relative animate-in zoom-in duration-200">
      <button onClick={() => setQrModal({ isOpen: false, data: null })} className="absolute top-4 right-4 text-slate-400 hover:text-red-500"><X size={24} /></button>
      <h3 className="font-bold text-xl mb-2">QR Code {qrModal.data.tipe ? `Kupon ${qrModal.data.tipe}` : 'Label Hewan'}</h3>
      <p className="text-sm text-slate-500 mb-6">{qrModal.data.nama} <br/><span className="font-mono text-xs font-bold">{qrModal.data.id}</span></p>
      <div className="flex justify-center p-4 bg-white border-2 border-slate-100 rounded-xl mb-6 shadow-inner">
      <QRCodeSVG value={qrModal.data.id} size={180} />
      </div>
      <button onClick={() => setQrModal({ isOpen: false, data: null })} className="w-full bg-slate-800 text-white py-3 rounded-lg font-bold hover:bg-slate-900">Tutup</button>
      </div>
      </div>
    )}

    {/* --- MODAL PDF PREVIEW --- */}
    {pdfPreview.isOpen && (
      <div className="fixed inset-0 bg-slate-900/80 backdrop-blur-sm flex items-center justify-center z-50 p-4 md:p-8">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-4xl h-full max-h-[85vh] flex flex-col overflow-hidden relative animate-in zoom-in duration-200">
      <div className="p-4 border-b border-slate-200 flex justify-between items-center bg-slate-50">
      <h3 className="font-bold text-lg text-slate-800">Preview: {pdfPreview.title}</h3>
      <button onClick={() => setPdfPreview({ isOpen: false, url: '', title: '' })} className="text-slate-500 hover:text-red-500 bg-white rounded-lg p-1.5 shadow-sm border border-slate-200"><X size={20} /></button>
      </div>
      <div className="flex-1 w-full bg-slate-200 p-2 md:p-4">
      <iframe src={pdfPreview.url} className="w-full h-full border-0 rounded bg-white shadow-inner" title="PDF Preview" />
      </div>
      </div>
      </div>
    )}

    </div>
  );
}

function Card({ title, value, sub, icon, border }) {
  return (
    <div className={`bg-white p-5 rounded-xl border border-slate-200 shadow-sm ${border} flex flex-col justify-between`}>
    <div className="flex justify-between items-start mb-4">
    <div><p className="text-slate-500 text-sm font-medium">{title}</p><h3 className="text-3xl font-bold text-slate-800 mt-2">{value}</h3></div>
    <div className="p-2 bg-slate-50 rounded-lg border border-slate-100">{icon}</div>
    </div>
    <p className="text-xs text-slate-400">{sub}</p>
    </div>
  );
}
