import React, { useState, useEffect } from 'react';
import { 
  LayoutDashboard, Users, UserCog, HeartHandshake, Beef, 
  ScanLine, Ticket, Award, LineChart, 
  X, Plus, Eye, DownloadCloud, Menu, CheckCircle, AlertCircle
} from 'lucide-react';
import { 
  PieChart, Pie, Cell, BarChart, Bar, XAxis, YAxis, 
  CartesianGrid, Tooltip, ResponsiveContainer, Legend 
} from 'recharts';
import { QRCodeSVG } from 'qrcode.react';
import QRCode from 'qrcode';
import { Html5QrcodeScanner } from 'html5-qrcode';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';

// --- DATA BAWAAN ---
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
  { id: "KMBG-02", jenis: 'Kambing', bobot: '28 kg', status: 'Menunggu' },
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
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  
  const [warga, setWarga] = useState(() => JSON.parse(localStorage.getItem('sys_warga')) || defaultWarga);
  const [panitia, setPanitia] = useState(() => JSON.parse(localStorage.getItem('sys_panitia')) || defaultPanitia);
  const [mudhohi, setMudhohi] = useState(() => JSON.parse(localStorage.getItem('sys_mudhohi')) || defaultMudhohi);
  const [hewan, setHewan] = useState(() => JSON.parse(localStorage.getItem('sys_hewan_v6')) || defaultHewan);
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
  useEffect(() => { localStorage.setItem('sys_hewan_v6', JSON.stringify(hewan)); }, [hewan]);
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

  const prosesKuponDigital = async (k, isPreview = false) => {
    const doc = new jsPDF({ format: [80, 80] });
    doc.setDrawColor(16, 185, 129); doc.setLineWidth(1); doc.rect(2, 2, 76, 76);
    doc.setFontSize(14); doc.setTextColor(16, 185, 129); doc.text("KUPON QURBAN", 40, 12, { align: "center" });
    doc.setFontSize(9); doc.setTextColor(50, 50, 50); doc.text(`Tipe: ${k.tipe}`, 40, 18, { align: "center" });
    try {
      const qrDataUrl = await QRCode.toDataURL(k.id, { margin: 1, width: 100 });
      doc.addImage(qrDataUrl, 'PNG', 25, 22, 30, 30);
    } catch (err) { console.error(err); }
    doc.setFontSize(11); doc.setTextColor(0, 0, 0); doc.text(k.nama, 40, 58, { align: "center" });
    doc.setFontSize(9); doc.setTextColor(50, 50, 50); doc.text(`ID: ${k.id}`, 40, 65, { align: "center" });
    doc.setFontSize(7); doc.setTextColor(100, 100, 100); doc.text("Tunjukkan kupon ini kepada Panitia", 40, 74, { align: "center" });
    if(isPreview) { setPdfPreview({ isOpen: true, url: doc.output('bloburl'), title: `Kupon ${k.id}` }); } 
    else { doc.save(`Kupon_${k.id}.pdf`); }
  };

  const prosesSertifikatMudhohi = (m, isPreview = false) => {
    const doc = new jsPDF('landscape');
    doc.setDrawColor(16, 185, 129); doc.setLineWidth(3); doc.rect(10, 10, 277, 190);
    doc.setFontSize(30); doc.setTextColor(16, 185, 129); doc.text("SERTIFIKAT QURBAN", 148, 50, { align: "center" });
    doc.setFontSize(14); doc.setTextColor(50, 50, 50); doc.text("Diberikan dengan penuh rasa syukur kepada:", 148, 80, { align: "center" });
    doc.setFontSize(28); doc.setTextColor(15, 23, 42); doc.text(m.nama, 148, 105, { align: "center" });
    doc.setFontSize(14); doc.setTextColor(50, 50, 50); doc.text(`Atas partisipasinya menyerahkan hewan qurban berupa ${m.hewan}.`, 148, 125, { align: "center" });
    if(isPreview) { setPdfPreview({ isOpen: true, url: doc.output('bloburl'), title: `Sertifikat ${m.nama}` }); } 
    else { doc.save(`Sertifikat_Mudhohi_${m.id}.pdf`); }
  };

  const prosesSertifikatPanitia = (p, isPreview = false) => {
    const doc = new jsPDF('landscape');
    doc.setDrawColor(59, 130, 246); doc.setLineWidth(3); doc.rect(10, 10, 277, 190);
    doc.setFontSize(30); doc.setTextColor(59, 130, 246); doc.text("PIAGAM PENGHARGAAN", 148, 50, { align: "center" });
    doc.setFontSize(28); doc.setTextColor(15, 23, 42); doc.text(p.nama, 148, 105, { align: "center" });
    doc.setFontSize(14); doc.setTextColor(50, 50, 50); doc.text(`Atas dedikasi bertugas sebagai ${p.peran}`, 148, 125, { align: "center" });
    if(isPreview) { setPdfPreview({ isOpen: true, url: doc.output('bloburl'), title: `Piagam ${p.nama}` }); } 
    else { doc.save(`Piagam_Panitia_${p.id}.pdf`); }
  };

  useEffect(() => {
    if (activeTab === 'scan') {
      const scanner = new Html5QrcodeScanner("reader", { qrbox: { width: 200, height: 200 }, fps: 5 });
      scanner.render((decodedText) => prosesValidasiKupon(decodedText), (error) => {});
      return () => { scanner.clear().catch(e => console.error(e)); };
    }
  }, [activeTab]);

  const prosesValidasiKupon = (idKupon) => {
    const kuponDitemukan = kupon.find(k => k.id === idKupon);
    if (!kuponDitemukan) return setScanResult({ status: 'error', pesan: `Kupon ${idKupon} tidak terdaftar!` });
    if (kuponDitemukan.status === 'Sudah Diambil') return setScanResult({ status: 'error', pesan: `Kupon ${idKupon} SUDAH DIAMBIL!` });
    setKupon(kupon.map(k => k.id === idKupon ? { ...k, status: 'Sudah Diambil' } : k));
    setScanResult({ status: 'success', pesan: `Berhasil! Kupon ${idKupon} atas nama ${kuponDitemukan.nama} divalidasi.` });
    setManualCode('');
  };

  const handleManualScan = (e) => { e.preventDefault(); if(manualCode.trim() !== '') prosesValidasiKupon(manualCode.trim()); };

  return (
    <div className="flex h-screen bg-slate-50 font-sans overflow-hidden relative">
      
      {/* --- MENU UTAMA RESPONSIF FIX --- */}
      <aside className={`fixed md:sticky top-0 bottom-0 left-0 w-64 bg-[#059669] text-white flex flex-col shadow-xl z-30 transition-transform duration-300 ease-in-out h-full ${isSidebarOpen ? 'translate-x-0' : '-translate-x-full md:translate-x-0'} overflow-hidden`}>
        <div className="p-6 flex items-center justify-between border-b border-emerald-700/50 flex-shrink-0">
          <div className="flex items-center gap-3">
            <div className="bg-amber-400 p-2 rounded-lg text-emerald-900 font-bold">Q</div>
            <div>
              <h1 className="text-lg font-bold leading-tight">QurbanPro</h1>
              <p className="text-xs text-emerald-100">Manajemen Qurban</p>
            </div>
          </div>
          <button onClick={() => setIsSidebarOpen(false)} className="md:hidden text-white hover:text-amber-300 transition"><X size={22} /></button>
        </div>
        
        <div className="flex-1 overflow-y-auto py-4 px-3 space-y-1">
          {menuItems.map((item) => {
            const Icon = item.icon;
            const isActive = activeTab === item.id;
            return (
              <button key={item.id} onClick={() => { setActiveTab(item.id); setScanResult(null); setIsSidebarOpen(false); }} className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all duration-200 ${isActive ? 'bg-white text-[#059669] font-bold shadow-sm' : 'text-emerald-50 hover:bg-emerald-800/40 hover:text-white'}`}>
                <Icon size={18} strokeWidth={isActive ? 2.5 : 2} /> <span className="text-sm tracking-wide">{item.label}</span>
              </button>
            );
          })}
        </div>
      </aside>

      {/* Latar gelap saat menu kebuka di HP */}
      {isSidebarOpen && <div onClick={() => setIsSidebarOpen(false)} className="fixed inset-0 bg-black/50 z-20 md:hidden animate-in fade-in duration-200"></div>}

      {/* --- AREA KONTEN --- */}
      <main className="flex-1 flex flex-col h-full overflow-y-auto w-full min-w-0">
        <header className="bg-white border-b border-slate-200 px-4 md:px-8 py-4 flex justify-between items-center sticky top-0 z-10 flex-shrink-0">
          <div className="flex items-center gap-3">
            <button onClick={() => setIsSidebarOpen(true)} className="md:hidden p-1.5 text-slate-600 hover:bg-slate-100 rounded-lg transition"><Menu size={24} /></button>
            <div>
              <h2 className="text-base md:text-xl font-bold text-slate-800 capitalize">{menuItems.find(m => m.id === activeTab)?.label || 'Menu'}</h2>
            </div>
          </div>
          <span className="bg-emerald-100 text-emerald-800 px-3 py-1 rounded-full text-xs font-bold flex items-center gap-1.5"><span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span> Aktif</span>
        </header>

        <div className="p-4 md:p-8 flex-1">
          
          {/* DASHBOARD */}
          {activeTab === 'dashboard' && (
            <div className="space-y-6 animate-in fade-in">
              <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 md:gap-6">
                <Card title="Hewan Qurban" value={hewan.length} sub={`${totalSapi} Sapi, ${totalKambing} Kambing`} icon={<Beef size={18} className="text-emerald-500"/>} border="border-l-4 border-l-emerald-500" />
                <Card title="Kupon Diambil" value={kupon.filter(k=>k.status==='Sudah Diambil').length} sub={`Dari ${kupon.length} kupon`} icon={<ScanLine size={18} className="text-blue-500"/>} border="border-l-4 border-l-blue-500" />
                <Card title="Total Mudhohi" value={mudhohi.length} sub="Peserta aktif" icon={<HeartHandshake size={18} className="text-amber-500"/>} border="border-l-4 border-l-amber-500" />
                <Card title="Saldo Kas" value={`Rp ${(saldo/1000000).toFixed(1)}jt`} sub={`Masuk: ${(totalMasuk/1000000).toFixed(1)}jt`} icon={<span className="text-purple-500 font-bold">$</span>} border="border-l-4 border-l-purple-500" />
              </div>
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm h-[320px]"><h4 className="font-semibold text-sm text-slate-800 mb-2">Penyembelihan</h4><ResponsiveContainer width="100%" height="90%"><PieChart><Pie data={pieData} cx="50%" cy="50%" innerRadius={50} outerRadius={75} paddingAngle={2} dataKey="value" label={({name, value}) => `${name}:${value}`}><{pieData.map((entry, i) => <Cell key={i} fill={entry.color} />)}</Pie><Tooltip /></PieChart></ResponsiveContainer></div>
                <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm h-[320px]"><h4 className="font-semibold text-sm text-slate-800 mb-4">Arus Kas (RAB)</h4><ResponsiveContainer width="100%" height="85%"><BarChart data={barData} barSize={45}><CartesianGrid strokeDasharray="3 3" vertical={false} /><XAxis dataKey="name" /><YAxis tickFormatter={(v) => `${v/1000000}jt`} /><Tooltip formatter={(v) => `Rp ${v.toLocaleString()}`}/><Bar dataKey="Pemasukan" fill="#10b981" /><Bar dataKey="Pengeluaran" fill="#ef4444" /></BarChart></ResponsiveContainer></div>
              </div>
            </div>
          )}

          {/* DATA WARGA */}
          {activeTab === 'warga' && (
            <div className="bg-white p-4 md:p-6 rounded-xl border border-slate-200 shadow-sm max-w-full overflow-hidden">
              <h3 className="font-bold text-lg mb-4">Database Warga ({warga.length} Jiwa)</h3>
              <div className="overflow-x-auto max-h-[60vh] border border-slate-100 rounded-lg">
                <table className="w-full text-left text-sm border-collapse min-w-[500px]">
                  <thead className="bg-slate-900 text-white sticky top-0"><tr><th className="p-3">ID</th><th className="p-3">Nama</th><th className="p-3">Wilayah</th></tr></thead>
                  <tbody className="divide-y divide-slate-100">{warga.map(w => <tr key={w.id} className="hover:bg-slate-50"><td className="p-3 font-mono font-bold text-slate-500">{w.id}</td><td className="p-3 font-semibold">{w.nama}</td><td className="p-3">{w.rt}</td></tr>)}</tbody>
                </table>
              </div>
            </div>
          )}

          {/* DATA PANITIA */}
          {activeTab === 'panitia' && (
            <div className="bg-white p-4 md:p-6 rounded-xl border border-slate-200 shadow-sm overflow-hidden">
              <h3 className="font-bold text-lg mb-4">Susunan Panitia</h3>
              <div className="overflow-x-auto border border-slate-100 rounded-lg">
                <table className="w-full text-left text-sm min-w-[500px]">
                  <thead className="bg-slate-900 text-white"><tr><th className="p-3">ID</th><th className="p-3">Nama</th><th className="p-3">Jabatan</th><th className="p-3">WA</th></tr></thead>
                  <tbody className="divide-y divide-slate-100">{panitia.map(p => <tr key={p.id} className="hover:bg-slate-50"><td className="p-3 font-mono text-slate-500">{p.id}</td><td className="p-3 font-semibold">{p.nama}</td><td className="p-3"><span className="bg-amber-50 text-amber-600 px-2 py-0.5 rounded text-xs font-bold border border-amber-100">{p.peran}</span></td><td className="p-3 text-slate-500">{p.kontak}</td></tr>)}</tbody>
                </table>
              </div>
            </div>
          )}

          {/* MUDHOHI */}
          {activeTab === 'mudhohi' && (
            <div className="bg-white p-4 md:p-6 rounded-xl border border-slate-200 shadow-sm overflow-hidden">
              <h3 className="font-bold text-lg mb-4">Data Mudhohi</h3>
              <div className="overflow-x-auto border border-slate-100 rounded-lg">
                <table className="w-full text-left text-sm min-w-[500px]">
                  <thead className="bg-slate-900 text-white"><tr><th className="p-3">ID</th><th className="p-3">Nama</th><th className="p-3">Hewan</th><th className="p-3">WA</th></tr></thead>
                  <tbody className="divide-y divide-slate-100">{mudhohi.map(m => <tr key={m.id} className="hover:bg-slate-50"><td className="p-3 font-mono text-slate-500">{m.id}</td><td className="p-3 font-semibold">{m.nama}</td><td className="p-3"><span className="bg-slate-100 text-slate-700 px-2 py-0.5 rounded text-xs font-bold">{m.hewan}</span></td><td className="p-3 text-slate-500">{m.wa}</td></tr>)}</tbody>
                </table>
              </div>
            </div>
          )}

          {/* HEWAN QURBAN */}
          {activeTab === 'hewan' && (
            <div className="bg-white p-4 md:p-6 rounded-xl border border-slate-200 shadow-sm overflow-hidden">
              <h3 className="font-bold text-lg mb-4">Pencatatan Hewan Qurban</h3>
              <form onSubmit={simpanHewan} className="grid grid-cols-1 md:grid-cols-4 gap-3 mb-6 bg-slate-50 p-4 rounded-lg text-sm">
                <select value={formHewan.jenis} onChange={e=>setFormHewan({...formHewan, jenis: e.target.value})} className="p-2 border rounded bg-white"><option value="Sapi">Sapi</option><option value="Kambing">Kambing</option></select>
                <input type="number" placeholder="Bobot (kg)" value={formHewan.bobot} onChange={e=>setFormHewan({...formHewan, bobot: e.target.value})} className="p-2 border rounded md:col-span-2 bg-white" required />
                <button type="submit" className="bg-[#059669] text-white p-2 rounded font-bold hover:bg-[#047857] flex justify-center items-center gap-1"><Plus size={16}/> Tambah</button>
              </form>
              <div className="overflow-x-auto max-h-[45vh] border border-slate-100 rounded-lg">
                <table className="w-full text-left text-sm min-w-[500px]">
                  <thead className="bg-slate-900 text-white sticky top-0"><tr><th class="p-3">ID</th><th class="p-3">Jenis</th><th class="p-3">Bobot</th><th class="p-3">Status</th><th class="p-3 text-center">QR</th></tr></thead>
                  <tbody className="divide-y divide-slate-100">{hewan.map(h => <tr key={h.id} className="hover:bg-slate-50"><td className="p-3 font-mono font-bold">{h.id}</td><td className="p-3"><span className="bg-slate-100 px-2 py-0.5 rounded text-xs font-bold">{h.jenis}</span></td><td className="p-3 text-slate-600 font-medium">{h.bobot}</td><td className="p-3"><select value={h.status} onChange={(e) => updateStatusHewan(h.id, e.target.value)} className="p-1 border rounded text-xs font-bold outline-none"><option value="Menunggu">Menunggu</option><option value="Disembelih">Disembelih</option><option value="Dikuliti">Dikuliti</option><option value="Selesai">Selesai</option></select></td><td className="p-3 flex justify-center"><button onClick={() => setQrModal({ isOpen: true, data: { id: h.id, nama: `Identitas ${h.jenis}` } })} className="bg-slate-800 text-white p-1.5 rounded hover:bg-slate-900"><QrCode size={14} /></button></td></tr>)}</tbody>
                </table>
              </div>
            </div>
          )}

          {/* KUPON DIGITAL */}
          {activeTab === 'kupon' && (
            <div className="bg-white p-4 md:p-6 rounded-xl border border-slate-200 shadow-sm animate-in fade-in">
              <h3 className="font-bold text-lg mb-4">Manajemen Kupon Digital</h3>
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                {['Warga', 'Mudhohi'].map(tipe => (
                  <div key={tipe} className="border border-slate-200 rounded-xl overflow-hidden shadow-xs">
                    <h4 className={`font-bold p-3 border-b-2 ${tipe==='Warga'?'text-emerald-700 bg-emerald-50 border-emerald-500':'text-blue-700 bg-blue-50 border-blue-500'}`}>Kupon {tipe}</h4>
                    <div className="overflow-y-auto max-h-[45vh] text-xs"><table className="w-full text-left border-collapse"><thead className="bg-slate-50 text-slate-600 border-b"><tr><th class="p-3">ID Kupon</th><th class="p-3">Nama</th><th class="p-3">Status</th><th class="p-3 text-right">Aksi</th></tr></thead><tbody className="divide-y bg-white">{kupon.filter(k=>k.tipe===tipe).map(k => <tr key={k.id} className="hover:bg-slate-50"><td className="p-3 font-mono font-bold text-slate-500">{k.id}</td><td className="p-3 font-medium">{k.nama}</td><td className="p-3"><span className={`px-2 py-0.5 rounded text-[10px] font-bold ${k.status==='Sudah Diambil'?'bg-emerald-100 text-emerald-700':'bg-orange-100 text-orange-700'}`}>{k.status}</span></td><td className="p-3 flex justify-end gap-1.5"><button onClick={()=>prosesKuponDigital(k, true)} className="p-1 border rounded text-slate-600 bg-slate-50 hover:bg-slate-100"><Eye size={12}/></button><button onClick={()=>prosesKuponDigital(k, false)} className="p-1 border rounded text-emerald-600 bg-emerald-50 hover:bg-emerald-100"><DownloadCloud size={12}/></button></td></tr>)}</tbody></table></div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* SCAN KUPON */}
          {activeTab === 'scan' && (
            <div className="bg-white p-4 md:p-6 rounded-xl border border-slate-200 shadow-sm">
              <h3 className="font-bold text-lg mb-6">Validasi & Scan Kupon</h3>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="flex flex-col items-center border border-slate-200 p-4 rounded-xl bg-slate-50"><p className="font-bold text-xs text-slate-600 mb-3 flex items-center gap-1"><ScanLine size={14}/> Arahkan QR ke Kamera</p><div id="reader" className="w-full max-w-[280px] rounded-lg overflow-hidden border border-dashed border-emerald-500 bg-black"></div></div>
                <div className="flex flex-col justify-center space-y-4">
                  <div className="bg-slate-50 p-4 rounded-xl border border-slate-200"><p className="font-bold text-sm text-slate-700 mb-2">Input ID Manual:</p><form onSubmit={handleManualScan} className="flex gap-2"><input type="text" value={manualCode} onChange={(e)=>setManualCode(e.target.value)} placeholder="KP-W-W-001" className="flex-1 p-2 border border-slate-300 rounded-lg outline-none text-sm font-mono" /><button type="submit" className="bg-[#059669] text-white px-4 py-2 rounded-lg text-sm font-bold hover:bg-[#047857]">Cek</button></form></div>
                  {scanResult && <div className={`p-4 rounded-xl border flex gap-3 text-xs font-medium ${scanResult.status === 'success' ? 'bg-emerald-50 border-emerald-200 text-emerald-800' : 'bg-red-50 border-red-200 text-red-800'}`}>{scanResult.status === 'success' ? <CheckCircle size={16} className="flex-shrink-0" /> : <AlertCircle size={16} className="flex-shrink-0" />}<p>{scanResult.pesan}</p></div>}
                </div>
              </div>
            </div>
          )}

          {/* SERTIFIKAT */}
          {activeTab === 'sertifikat' && (
            <div className="bg-white p-4 md:p-6 rounded-xl border border-slate-200 shadow-sm animate-in fade-in">
              <h3 className="font-bold text-lg mb-4">Cetak Sertifikat & Piagam</h3>
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                {['Mudhohi', 'Panitia'].map(tipe => (
                  <div key={tipe} className="border border-slate-200 rounded-xl overflow-hidden shadow-xs">
                    <h4 className={`font-bold p-3 border-b-2 ${tipe==='Mudhohi'?'text-emerald-700 bg-emerald-50 border-emerald-500':'text-blue-700 bg-blue-50 border-blue-500'}`}>Berkas {tipe}</h4>
                    <div className="overflow-y-auto max-h-[45vh] text-xs divide-y bg-white">
                      {(tipe==='Mudhohi'?mudhohi:panitia).map(item => (
                        <div key={item.id} className="hover:bg-slate-50 flex justify-between items-center p-3">
                          <div className="font-semibold text-slate-800">{item.nama}</div>
                          <div className="flex gap-1">
                            <button onClick={()=>(tipe==='Mudhohi'?prosesSertifikatMudhohi(item, true):prosesSertifikatPanitia(item, true))} className="p-1 border rounded text-slate-600 bg-slate-50 hover:bg-slate-100 flex items-center gap-1 font-medium"><Eye size={12}/> Preview</button>
                            <button onClick={()=>(tipe==='Mudhohi'?prosesSertifikatMudhohi(item, false):prosesSertifikatPanitia(item, false))} className="p-1 border rounded text-emerald-600 bg-emerald-50 hover:bg-emerald-100 flex items-center gap-1 font-medium"><DownloadCloud size={12}/> Unduh</button>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* LAPORAN RAB */}
          {activeTab === 'uang' && (
            <div className="bg-white p-4 md:p-6 rounded-xl border border-slate-200 shadow-sm overflow-hidden">
              <div className="flex justify-between items-center mb-4"><h3 className="font-bold text-lg">Laporan RAB</h3><div className="flex gap-1.5"><button onClick={exportExcel} className="bg-emerald-600 text-white px-2.5 py-1.5 rounded text-xs font-bold flex items-center gap-1 hover:bg-emerald-700"><DownloadCloud size={12}/> Excel</button><button onClick={exportPDF} className="bg-red-500 text-white px-2.5 py-1.5 rounded text-xs font-bold flex items-center gap-1 hover:bg-red-600"><DownloadCloud size={12}/> PDF</button></div></div>
              <form onSubmit={simpanUang} className="grid grid-cols-1 md:grid-cols-4 gap-2.5 mb-6 bg-slate-50 p-4 rounded-lg text-sm">
                <input type="text" placeholder="Uraian" value={formUang.keterangan} onChange={e=>setFormUang({...formUang, keterangan: e.target.value})} className="p-2 border rounded bg-white md:col-span-2" required />
                <select value={formUang.jenis} onChange={e=>setFormUang({...formUang, jenis: e.target.value})} className="p-2 border rounded bg-white"><option value="Masuk">Masuk (Debit)</option><option value="Keluar">Keluar (Kredit)</option></select>
                <input type="number" placeholder="Nominal" value={formUang.nominal} onChange={e=>setFormUang({...formUang, nominal: e.target.value})} className="p-2 border rounded bg-white" required />
                <div className="md:col-span-4 flex justify-end"><button type="submit" className="bg-[#059669] text-white px-4 py-2 rounded text-xs font-bold hover:bg-[#047857]">Simpan</button></div>
              </form>
              <div className="overflow-x-auto border border-slate-100 rounded-lg text-xs">
                <table className="w-full text-left min-w-[500px]"><thead className="bg-slate-100 text-slate-700"><tr><th class="p-3">Tanggal</th><th class="p-3">Uraian</th><th class="p-3">Jenis</th><th class="p-3 text-right">Nominal</th></tr></thead><tbody className="divide-y">{keuangan.map((k) => <tr key={k.id} className="hover:bg-slate-50"><td className="p-3 text-slate-500">{k.tanggal}</td><td className="p-3 font-medium text-slate-800">{k.keterangan}</td><td className="p-3"><span className={`px-2 py-0.5 rounded font-bold text-[10px] ${k.jenis==='Masuk'?'bg-emerald-100 text-emerald-700':'bg-red-100 text-red-700'}`}>{k.jenis}</span></td><td className={`p-3 text-right font-bold ${k.jenis==='Masuk'?'text-emerald-600':'text-red-600'}`}>Rp {k.nominal.toLocaleString('id-ID')}</td></tr>)}</tbody><tfoot className="bg-slate-900 text-white font-bold text-sm"><tr><td colSpan="3" className="p-3 text-right">TOTAL SALDO:</td><td className="p-3 text-right text-emerald-400">Rp {saldo.toLocaleString('id-ID')}</td></tr></tfoot></table>
              </div>
            </div>
          )}

        </div>
      </main>

      {/* MODAL QR CODE */}
      {qrModal.isOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center z-50 p-4">
          <div className="bg-white p-6 rounded-2xl shadow-2xl max-w-xs w-full text-center relative animate-in zoom-in duration-200">
            <button onClick={() => setQrModal({ isOpen: false, data: null })} className="absolute top-4 right-4 text-slate-400 hover:text-red-500"><X size={20} /></button>
            <h4 className="font-bold text-base mb-1">QR Code {qrModal.data.id}</h4>
            <p className="text-xs text-slate-500 mb-4">{qrModal.data.nama}</p>
            <div className="flex justify-center p-3 bg-white border border-slate-100 rounded-xl mb-4"><QRCodeSVG value={qrModal.data.id} size={150} /></div>
            <button onClick={() => setQrModal({ isOpen: false, data: null })} className="w-full bg-slate-800 text-white py-2 rounded-lg text-sm font-bold hover:bg-slate-900">Tutup</button>
          </div>
        </div>
      )}

      {/* MODAL PDF PREVIEW */}
      {pdfPreview.isOpen && (
        <div className="fixed inset-0 bg-slate-900/80 backdrop-blur-xs flex items-center justify-center z-50 p-3 md:p-6">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-4xl h-full max-h-[80vh] flex flex-col overflow-hidden relative animate-in zoom-in duration-200">
            <div className="p-3 border-b border-slate-200 flex justify-between items-center bg-slate-50"><h3 className="font-bold text-sm text-slate-800">Preview: {pdfPreview.title}</h3><button onClick={() => setPdfPreview({ isOpen: false, url: '', title: '' })} className="text-slate-500 hover:text-red-500 bg-white rounded-md p-1 border border-slate-200"><X size={16} /></button></div>
            <div className="flex-1 w-full bg-slate-100 p-2"><iframe src={pdfPreview.url} className="w-full h-full border-0 rounded bg-white shadow-inner" title="PDF Preview" /></div>
          </div>
        </div>
      )}

    </div>
  );
}

function Card({ title, value, sub, icon, border }) {
  return (
    <div className={`bg-white p-4 rounded-xl border border-slate-200 shadow-sm ${border} flex flex-col justify-between`}>
      <div className="flex justify-between items-start mb-2">
        <div><p className="text-slate-500 text-[10px] md:text-sm font-medium">{title}</p><h3 className="text-lg md:text-3xl font-bold text-slate-800 mt-1">{value}</h3></div>
        <div className="p-1 md:p-2 bg-slate-50 rounded-lg border border-slate-100">{icon}</div>
      </div>
      <p className="text-[9px] md:text-xs text-slate-400 truncate">{sub}</p>
    </div>
  );
}
