import React, { useState, useEffect } from 'react';
import { 
  LayoutDashboard, Users, UserCog, HeartHandshake, Beef, 
  ScanLine, Ticket, Award, LineChart, 
  X, Plus, Eye, DownloadCloud, Menu, CheckCircle, AlertCircle
} from 'lucide-react';
import { 
  PieChart, Pie, Cell, BarChart, Bar, XAxis, YAxis, 
  CartesianGrid, Tooltip, ResponsiveContainer 
} from 'recharts';
import { QRCodeSVG } from 'qrcode.react';
import QRCode from 'qrcode';
import { Html5QrcodeScanner } from 'html5-qrcode';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';

// --- DATA BAWAAN ---
const defaultWarga = Array.from({ length: 31 }, (_, i) => ({
  id: `W-${String(i + 1).padStart(3, '0')}`,
  nama: ["Bpk. Ahmad", "Ibu Siti", "Bpk. Budi", "Ibu Ani", "Bpk. Joko", "Ibu Rina", "Bpk. Agus", "Ibu Sri", "Bpk. Hendra", "Ibu Maya", "Bpk. Dedi", "Ibu Nina", "Bpk. Yudi", "Ibu Dewi", "Bpk. Eko", "Ibu Ratna", "Bpk. Iwan", "Ibu Sari", "Bpk. Arif", "Ibu Lia", "Bpk. Rizal", "Ibu Ayu", "Bpk. Anton", "Ibu Fitri", "Bpk. Wahyu", "Ibu Dian", "Bpk. Indra", "Ibu Tari", "Bpk. Gilang", "Ibu Wati", "Bpk. Rahmat"][i] || `Warga ${i+1}`,
  rt: `RT 0${(i % 5) + 1}`,
  wa: `0812${Math.floor(10000000 + Math.random() * 90000000)}` // Nomor dummy random
}));

const defaultMudhohi = [
  { id: 'MD-01', nama: 'Kel. Bpk. H. Abdullah', hewan: 'Sapi', wa: '081234567890' },
  { id: 'MD-02', nama: 'Bpk. Budi Santoso', hewan: 'Kambing', wa: '081311223344' },
  { id: 'MD-03', nama: 'Ibu Siti Khadijah', hewan: 'Kambing', wa: '081322334455' },
  { id: 'MD-04', nama: 'Sdr. Andi Pratama', hewan: 'Kambing', wa: '081333445566' },
  { id: 'MD-05', nama: 'Bpk. Joko Susilo', hewan: 'Kambing', wa: '081344556677' }
];

const defaultPanitia = Array.from({ length: 10 }, (_, i) => ({
  id: `P-${String(i + 1).padStart(2, '0')}`,
  nama: ["H. Fulan", "Ust. Hasan", "Bpk. Zainal", "Bpk. Anwar", "Bpk. Somad", "Bpk. Qasim", "Sdr. Ilham", "Sdr. Reza", "Bpk. Yasin", "Sdr. Iqbal"][i],
  peran: i === 0 ? "Ketua Panitia" : i === 3 ? "Bendahara" : "Anggota",
  wa: `0813${Math.floor(10000000 + Math.random() * 90000000)}`
}));

const defaultHewan = [
  { id: "SAPI-01", jenis: 'Sapi', bobot: '350 kg', status: 'Selesai' },
  { id: "SAPI-02", jenis: 'Sapi', bobot: '320 kg', status: 'Dikuliti' },
  ...Array.from({ length: 12 }, (_, i) => ({
    id: `KMBG-${String(i + 1).padStart(2, '0')}`,
    jenis: 'Kambing',
    bobot: `${20 + (i % 10)} kg`,
    status: i < 2 ? 'Selesai' : i < 4 ? 'Disembelih' : 'Menunggu'
  }))
];

const defaultKeuangan = [
  { id: 1, tanggal: '28/05/2026', keterangan: 'Saldo Kas Masjid', jenis: 'Masuk', nominal: 52500000 },
  { id: 2, tanggal: '01/06/2026', keterangan: 'DP Tim Jagal', jenis: 'Keluar', nominal: 32500000 }
];

const generateKuponInit = () => {
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
  
  const [warga] = useState(() => JSON.parse(localStorage.getItem('sys_warga_v9')) || defaultWarga);
  const [panitia] = useState(() => JSON.parse(localStorage.getItem('sys_panitia_v9')) || defaultPanitia);
  const [mudhohi] = useState(() => JSON.parse(localStorage.getItem('sys_mudhohi_v9')) || defaultMudhohi);
  const [hewan, setHewan] = useState(() => JSON.parse(localStorage.getItem('sys_hewan_v9')) || defaultHewan);
  const [kupon, setKupon] = useState(() => JSON.parse(localStorage.getItem('sys_kupon_v9')) || generateKuponInit());
  const [keuangan, setKeuangan] = useState(() => JSON.parse(localStorage.getItem('sys_keuangan_v9')) || defaultKeuangan);
  
  const [formHewan, setFormHewan] = useState({ jenis: 'Sapi', bobot: '' });
  const [formUang, setFormUang] = useState({ keterangan: '', jenis: 'Masuk', nominal: '' });
  
  const [qrModal, setQrModal] = useState({ isOpen: false, data: null });
  const [pdfPreview, setPdfPreview] = useState({ isOpen: false, url: '', title: '' });
  const [scanResult, setScanResult] = useState(null);
  const [manualCode, setManualCode] = useState('');

  useEffect(() => {
    localStorage.setItem('sys_warga_v9', JSON.stringify(warga));
    localStorage.setItem('sys_panitia_v9', JSON.stringify(panitia));
    localStorage.setItem('sys_mudhohi_v9', JSON.stringify(mudhohi));
    localStorage.setItem('sys_hewan_v9', JSON.stringify(hewan));
    localStorage.setItem('sys_kupon_v9', JSON.stringify(kupon));
    localStorage.setItem('sys_keuangan_v9', JSON.stringify(keuangan));
  }, [warga, panitia, mudhohi, hewan, kupon, keuangan]);

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
    { name: 'Pemasukan', v: totalMasuk },
    { name: 'Pengeluaran', v: totalKeluar }
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
    doc.setDrawColor(16, 185, 129); doc.rect(2, 2, 76, 76);
    doc.setFontSize(12); doc.setTextColor(16, 185, 129); doc.text("KUPON QURBAN", 40, 12, { align: "center" });
    try {
      const qrDataUrl = await QRCode.toDataURL(k.id, { margin: 1 });
      doc.addImage(qrDataUrl, 'PNG', 25, 20, 30, 30);
    } catch (e) {}
    doc.setFontSize(10); doc.setTextColor(0,0,0); doc.text(k.nama, 40, 58, { align: "center" });
    doc.text(k.id, 40, 65, { align: "center" });
    if(isPreview) { setPdfPreview({ isOpen: true, url: doc.output('bloburl'), title: k.id }); }
    else { doc.save(`${k.id}.pdf`); }
  };

  const prosesSertifikat = (name, role, isMudhohi, isPreview = false) => {
    const doc = new jsPDF('landscape');
    const color = isMudhohi ? [16, 185, 129] : [59, 130, 246];
    doc.setDrawColor(color[0], color[1], color[2]); doc.setLineWidth(2); doc.rect(10, 10, 277, 190);
    doc.setFontSize(30); doc.setTextColor(color[0], color[1], color[2]);
    doc.text(isMudhohi ? "SERTIFIKAT QURBAN" : "PIAGAM PENGHARGAAN", 148, 50, { align: "center" });
    doc.setFontSize(24); doc.setTextColor(0,0,0); doc.text(name, 148, 100, { align: "center" });
    doc.setFontSize(14); doc.text(isMudhohi ? `Atas partisipasi Qurban 1447 H` : `Dedikasi sebagai ${role}`, 148, 120, { align: "center" });
    if(isPreview) { setPdfPreview({ isOpen: true, url: doc.output('bloburl'), title: name }); }
    else { doc.save(`${name}.pdf`); }
  };

  const prosesValidasiKupon = (idKupon) => {
    const found = kupon.find(k => k.id === idKupon);
    if (!found) return setScanResult({ status: 'error', pesan: 'Tidak Terdaftar' });
    if (found.status === 'Sudah Diambil') return setScanResult({ status: 'error', pesan: 'SUDAH DIAMBIL' });
    
    setKupon(prev => prev.map(k => k.id === idKupon ? { ...k, status: 'Sudah Diambil' } : k));
    setScanResult({ status: 'success', pesan: `Berhasil: ${found.nama}` });
    setManualCode('');
  };

  // PEMBENAHAN CRASH ELEMEN #READER
  useEffect(() => {
    let scanner = null;
    if (activeTab === 'scan' && document.getElementById("reader")) {
      scanner = new Html5QrcodeScanner("reader", { qrbox: 200, fps: 10 });
      scanner.render((txt) => {
        const found = kupon.find(k => k.id === txt);
        if(!found) setScanResult({ status:'error', pesan:'Tidak Terdaftar' });
        else if(found.status==='Sudah Diambil') setScanResult({ status:'error', pesan:'SUDAH DIAMBIL' });
        else {
          setKupon(prev => prev.map(k => k.id === txt ? {...k, status:'Sudah Diambil'} : k));
          setScanResult({ status:'success', pesan:`Berhasil: ${found.nama}` });
        }
      }, () => {});
    }
    return () => { if (scanner) { scanner.clear().catch(()=>{}); } };
  }, [activeTab, kupon]);

  return (
    <div className="flex h-screen bg-slate-50 font-sans overflow-hidden relative">
      
      {/* SIDEBAR */}
      <aside className={`fixed md:sticky top-0 bottom-0 left-0 w-64 bg-[#059669] text-white flex flex-col shadow-xl z-30 transition-transform duration-300 ease-in-out h-full ${isSidebarOpen ? 'translate-x-0' : '-translate-x-full md:translate-x-0'}`}>
        <div className="p-6 flex items-center justify-between border-b border-emerald-700/50 flex-shrink-0">
          <div className="flex items-center gap-2">
            <div className="bg-amber-400 p-1.5 rounded-lg text-emerald-900 font-bold">Q</div>
            <h1 className="font-bold text-lg">QurbanPro</h1>
          </div>
          <button onClick={() => setIsSidebarOpen(false)} className="md:hidden text-white"><X size={20}/></button>
        </div>
        <div className="flex-1 overflow-y-auto p-3 space-y-1">
          {menuItems.map(m => (
            <button key={m.id} onClick={()=>{setActiveTab(m.id); setIsSidebarOpen(false)}} className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl transition ${activeTab===m.id?'bg-white text-[#059669] font-bold shadow-sm':'text-emerald-50 hover:bg-emerald-800/40'}`}>
              <m.icon size={18}/> <span className="text-sm">{m.label}</span>
            </button>
          ))}
        </div>
      </aside>

      {isSidebarOpen && <div onClick={()=>setIsSidebarOpen(false)} className="fixed inset-0 bg-black/50 z-20 md:hidden"></div>}

      {/* MAIN CONTENT */}
      <main className="flex-1 flex flex-col min-w-0 overflow-y-auto w-full">
        <header className="bg-white border-b p-4 px-6 flex justify-between items-center sticky top-0 z-10 flex-shrink-0">
          <div className="flex items-center gap-3">
            <button onClick={()=>setIsSidebarOpen(true)} className="md:hidden p-1.5 text-slate-600 hover:bg-slate-100 rounded-lg transition"><Menu size={24}/></button>
            <h2 className="text-base md:text-xl font-bold text-slate-800 capitalize">{activeTab}</h2>
          </div>
          <div className="flex items-center gap-2 bg-emerald-50 text-emerald-700 px-3 py-1 rounded-full text-xs font-bold">
            <div className="w-2 h-2 bg-emerald-500 rounded-full animate-pulse"></div> Aktif
          </div>
        </header>

        <div className="p-4 md:p-8 flex-1">
          {activeTab === 'dashboard' && (
            <div className="space-y-6 animate-in fade-in duration-300">
              <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 md:gap-6">
                <Card title="Total Hewan Qurban" value={hewan.length} sub={`${totalSapi} Sapi, ${totalKambing} Kambing`} color="emerald"/>
                <Card title="Kupon Diclaim" value={kupon.filter(k=>k.status==='Sudah Diambil').length} sub={`Dari total ${kupon.length}`} color="blue"/>
                <Card title="Mudhohi" value={mudhohi.length} sub="Peserta aktif" color="amber"/>
                <Card title="Saldo Kas" value={`Rp ${(saldo/1000000).toFixed(1)}jt`} sub={`Masuk: Rp ${(totalMasuk/1000000).toFixed(1)}jt`} color="purple"/>
              </div>
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                <div className="bg-white p-4 rounded-xl border h-[320px] shadow-sm flex flex-col">
                  <h4 className="font-semibold text-sm text-slate-800 mb-2">Status Penyembelihan</h4>
                  <div className="flex-1 min-h-0">
                    <ResponsiveContainer width="100%" height="100%">
                      <PieChart>
                        <Pie data={pieData} cx="50%" cy="50%" innerRadius={50} outerRadius={75} paddingAngle={2} dataKey="value" label={({name, value}) => `${name}:${value}`}>
                          {pieData.map((e,i)=><Cell key={i} fill={e.color}/>)}
                        </Pie>
                      </PieChart>
                    </ResponsiveContainer>
                  </div>
                </div>
                <div className="bg-white p-4 rounded-xl border h-[320px] shadow-sm flex flex-col">
                  <h4 className="font-semibold text-sm text-slate-800 mb-2">Arus Kas Keuangan</h4>
                  <div className="flex-1 min-h-0">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={barData} barSize={50}>
                        <XAxis dataKey="name" />
                        <YAxis tickFormatter={(v) => `${(v/1000000).toFixed(0)}jt`} />
                        <Bar dataKey="v" fill="#059669" />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              </div>
            </div>
          )}

          {activeTab === 'warga' && (
  <div className="bg-white rounded-xl border shadow-sm overflow-hidden">
    <div className="p-4 border-b font-bold text-slate-800">Database Warga</div>
    <div className="overflow-x-auto max-h-[60vh]"><table className="w-full text-left text-sm border-collapse min-w-[500px]">
      <thead className="bg-slate-50 border-b text-slate-600"><tr><th className="p-3">ID</th><th className="p-3">Nama</th><th className="p-3">RT</th><th className="p-3">WhatsApp</th></tr></thead>
      <tbody className="divide-y">{warga.map(w=><tr key={w.id} className="hover:bg-slate-50"><td className="p-3 font-mono text-slate-400 font-bold">{w.id}</td><td className="p-3 font-semibold text-slate-800">{w.nama}</td><td className="p-3 text-slate-600">{w.rt}</td>
      <td className="p-3"><a href={`https://wa.me/${w.wa.replace('0','62')}`} target="_blank" className="text-emerald-600 font-bold hover:underline">{w.wa}</a></td></tr>)}</tbody>
    </table></div>
  </div>
)}

          {activeTab === 'panitia' && (
  <div className="bg-white rounded-xl border shadow-sm overflow-hidden">
    <div className="p-4 border-b font-bold text-slate-800">Struktur Panitia</div>
    <table className="w-full text-left text-sm border-collapse">
      <thead className="bg-slate-50 border-b text-slate-600"><tr><th className="p-3">Nama</th><th className="p-3">Jabatan</th><th className="p-3">Kontak</th></tr></thead>
      <tbody className="divide-y">{panitia.map(p=><tr key={p.id} className="hover:bg-slate-50"><td className="p-3 font-semibold text-slate-800">{p.nama}</td><td className="p-3 text-blue-600">{p.peran}</td>
      <td className="p-3"><a href={`https://wa.me/${p.wa.replace('0','62')}`} target="_blank" className="bg-emerald-100 text-emerald-700 px-2 py-1 rounded text-xs font-bold hover:bg-emerald-200">Chat WA</a></td></tr>)}</tbody>
    </table>
  </div>
)}

          {activeTab === 'mudhohi' && (
            <div className="bg-white rounded-xl border shadow-sm overflow-hidden">
              <div className="p-4 border-b font-bold text-slate-800">Daftar Peserta Mudhohi</div>
              <table className="w-full text-left text-sm border-collapse">
                <thead className="bg-slate-50 border-b text-slate-600"><tr><th className="p-3 font-semibold">Nama Mudhohi</th><th className="p-3 font-semibold">Patungan Hewan</th></tr></thead>
                <tbody className="divide-y">{mudhohi.map(m=><tr key={m.id} className="hover:bg-slate-50"><td className="p-3 font-semibold text-slate-800">{m.nama}</td><td className="p-3"><span className="bg-emerald-50 text-emerald-700 px-2.5 py-0.5 rounded-md text-xs font-bold border border-emerald-100">{m.hewan}</span></td></tr>)}</tbody>
              </table>
            </div>
          )}

          {activeTab === 'hewan' && (
            <div className="bg-white rounded-xl border shadow-sm overflow-hidden">
              <div className="p-4 border-b font-bold text-slate-800">Monitoring Hewan Qurban</div>
              <div className="p-4 bg-slate-50 border-b flex gap-2 text-xs">
                <form onSubmit={simpanHewan} className="flex gap-2">
                  <select value={formHewan.jenis} onChange={e=>setFormHewan({...formHewan, jenis: e.target.value})} className="border p-1 rounded bg-white"><option value="Sapi">Sapi</option><option value="Kambing">Kambing</option></select>
                  <input type="number" placeholder="kg" value={formHewan.bobot} onChange={e=>setFormHewan({...formHewan, bobot: e.target.value})} className="border p-1 rounded w-20 bg-white" required />
                  <button type="submit" className="bg-emerald-600 text-white px-3 rounded font-bold">Tambah</button>
                </form>
              </div>
              <div className="overflow-x-auto max-h-[50vh]"><table className="w-full text-left text-sm min-w-[400px]">
                <thead className="bg-slate-50 border-b text-slate-600"><tr><th className="p-3 font-semibold">ID</th><th className="p-3 font-semibold">Jenis</th><th className="p-3 font-semibold">Status Lapangan</th></tr></thead>
                <tbody className="divide-y">{hewan.map(h=><tr key={h.id} className="hover:bg-slate-50"><td className="p-3 font-mono font-bold text-slate-900">{h.id}</td><td className="p-3 font-medium text-slate-700">{h.jenis} ({h.bobot})</td><td className="p-3"><select value={h.status} onChange={(e) => updateStatusHewan(h.id, e.target.value)} className="text-xs border rounded p-1 bg-white">{['Menunggu','Disembelih','Dikuliti','Selesai'].map(s=><option key={s} value={s}>{s}</option>)}</select></td></tr>)}</tbody>
              </table></div>
            </div>
          )}

          {activeTab === 'kupon' && (
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              {['Warga', 'Mudhohi'].map(t => (
                <div key={t} className="bg-white rounded-xl border shadow-sm overflow-hidden">
                  <div className={`p-3 border-b font-bold ${t === 'Warga' ? 'bg-emerald-50 text-emerald-800' : 'bg-blue-50 text-blue-800'}`}>Kupon Pembagian {t}</div>
                  <div className="max-h-96 overflow-y-auto text-xs"><table className="w-full text-left border-collapse">
                    <tbody className="divide-y">{kupon.filter(k=>k.tipe===t).map(k=><tr key={k.id} className="hover:bg-slate-50">
                      <td className="p-3 font-mono font-bold text-slate-400">{k.id}</td><td className="p-3 font-medium text-slate-800">{k.nama}</td>
                      <td className="p-3 text-right flex justify-end gap-1.5"><button onClick={()=>prosesKuponDigital(k,true)} className="p-1.5 border rounded-lg text-slate-500 bg-slate-50 hover:bg-slate-100"><Eye size={14}/></button><button onClick={()=>prosesKuponDigital(k,false)} className="p-1.5 border rounded-lg bg-emerald-50 text-emerald-600 hover:bg-emerald-100"><DownloadCloud size={14}/></button></td>
                    </tr>)}</tbody>
                  </table></div>
                </div>
              ))}
            </div>
          )}

          {activeTab === 'scan' && (
            <div className="bg-white p-6 rounded-xl border shadow-sm text-center max-w-xl mx-auto">
              <div id="reader" className="mx-auto max-w-xs overflow-hidden rounded-xl bg-black mb-6 border-2 border-dashed border-emerald-400"></div>
              <div className="mt-4"><form onSubmit={(e)=>{e.preventDefault(); prosesValidasiKupon(manualCode);}} className="flex gap-2 max-w-xs mx-auto"><input type="text" value={manualCode} onChange={e=>setManualCode(e.target.value)} placeholder="Masukkan Kode Kupon" className="flex-1 border p-2.5 rounded-lg outline-none text-sm font-mono border-slate-300 bg-white"/><button type="submit" className="bg-[#059669] text-white px-5 rounded-lg font-bold text-sm">Validasi</button></form></div>
              {scanResult && <div className={`mt-4 p-3 rounded-xl border text-sm font-bold ${scanResult.status==='success'?'bg-emerald-50 text-emerald-800 border-emerald-200':'bg-red-50 text-red-800 border-red-200'}`}>{scanResult.pesan}</div>}
            </div>
          )}

          {activeTab === 'sertifikat' && (
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              <div className="bg-white rounded-xl border shadow-sm p-5 h-96 flex flex-col">
                <h4 className="font-bold border-b pb-2 mb-3 text-emerald-800 flex items-center gap-1.5">Berkas Sertifikat Mudhohi</h4>
                <div className="overflow-y-auto divide-y flex-1 text-xs">{mudhohi.map(m=><div key={m.id} className="py-2.5 flex justify-between items-center"><div><p className="font-bold text-slate-800 text-sm">{m.nama}</p></div><div className="flex gap-1.5"><button onClick={()=>prosesSertifikat(m.nama, '', true, true)} className="p-1.5 border rounded-lg bg-slate-50 flex items-center gap-1">Preview</button><button onClick={()=>prosesSertifikat(m.nama, '', true, false)} className="p-1.5 border rounded-lg bg-emerald-50 text-emerald-600 font-bold flex items-center gap-1"><DownloadCloud size={12}/> Unduh</button></div></div>)}</div>
              </div>
              <div className="bg-white rounded-xl border shadow-sm p-5 h-96 flex flex-col">
                <h4 className="font-bold border-b pb-2 mb-3 text-blue-800 flex items-center gap-1.5">Piagam Penghargaan Panitia</h4>
                <div className="overflow-y-auto divide-y flex-1 text-xs">{panitia.map(p=><div key={p.id} className="py-2.5 flex justify-between items-center"><div><p className="font-bold text-slate-800 text-sm">{p.nama}</p><p className="text-blue-600 font-medium text-[10px]">{p.peran}</p></div><div className="flex gap-1.5"><button onClick={()=>prosesSertifikat(p.nama, p.peran, false, true)} className="p-1.5 border rounded-lg bg-slate-50 flex items-center gap-1">Preview</button><button onClick={()=>prosesSertifikat(p.nama, p.peran, false, false)} className="p-1.5 border rounded-lg bg-blue-50 text-blue-600 font-bold flex items-center gap-1"><DownloadCloud size={12}/> Unduh</button></div></div>)}</div>
              </div>
            </div>
          )}

          {activeTab === 'uang' && (
            <div className="bg-white rounded-xl border shadow-sm overflow-hidden">
              <div className="p-4 bg-slate-900 text-white flex justify-between items-center">
                <div><h3 className="font-bold">Laporan Anggaran Keuangan</h3></div>
                <div className="text-right"><span className="text-emerald-400 font-black text-lg">Rp {saldo.toLocaleString('id-ID')}</span></div>
              </div>
              <div className="p-4 bg-slate-50 border-b flex justify-end gap-2"><button onClick={exportExcel} className="bg-emerald-600 text-white px-3.5 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5"><DownloadCloud size={14}/> Unduh Excel</button></div>
              <div className="overflow-x-auto"><table className="w-full text-left text-sm border-collapse min-w-[500px]">
                <thead className="bg-slate-100 border-b text-slate-600"><tr><th className="p-3 font-semibold">Tanggal</th><th className="p-3 font-semibold">Uraian Keterangan</th><th className="p-3 font-semibold text-right">Nominal Arus</th></tr></thead>
                <tbody className="divide-y bg-white">{keuangan.map(k=><tr key={k.id} className="hover:bg-slate-50"><td className="p-3 font-mono text-slate-400 text-xs">{k.tanggal}</td><td className="p-3 font-medium text-slate-800">{k.keterangan}</td><td className={`p-3 text-right font-bold ${k.jenis==='Masuk'?'text-emerald-600':'text-red-500'}`}>{k.jenis === 'Masuk' ? '+' : '-'} Rp {k.nominal.toLocaleString('id-ID')}</td></tr>)}</tbody>
              </table></div>
            </div>
          )}
        </div>
      </main>

      {/* PDF PREVIEW MODAL */}
      {pdfPreview.isOpen && (
        <div className="fixed inset-0 bg-black/80 flex flex-col z-50 p-3 md:p-10">
          <div className="flex justify-between bg-white p-3.5 rounded-t-xl border-b max-w-4xl w-full mx-auto shadow-xl">
            <h3 className="font-bold text-slate-800 truncate text-sm">Pratinjau Dokumen</h3>
            <button onClick={()=>setPdfPreview({isOpen:false,url:'',title:''})} className="text-slate-400 hover:text-red-500"><X size={20}/></button>
          </div>
          <div className="flex-1 max-w-4xl w-full mx-auto bg-slate-100 rounded-b-xl overflow-hidden p-1">
            <iframe src={pdfPreview.url} className="w-full h-full border-0 bg-white" title="PDF Live Preview" />
          </div>
        </div>
      )}
    </div>
  );
}

function Card({ title, value, sub, color }) {
  const colors = { emerald: 'border-l-emerald-500', blue: 'border-l-blue-500', amber: 'border-l-amber-500', purple: 'border-l-purple-500' };
  return (
    <div className={`bg-white p-4 rounded-xl border border-slate-200 border-l-4 ${colors[color]} shadow-sm flex flex-col justify-between min-w-0`}>
      <div>
        <p className="text-[10px] uppercase font-bold text-slate-400 tracking-wider truncate">{title}</p>
        <h3 className="text-lg md:text-2xl font-black text-slate-800 my-0.5 md:my-1 truncate">{value}</h3>
      </div>
      <p className="text-[9px] md:text-xs text-slate-500 mt-1 truncate">{sub}</p>
    </div>
  );
}
