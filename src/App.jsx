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
const defaultWarga = Array.from({ length: 31 }, (_, i) => ({
  id: `W-${String(i + 1).padStart(3, '0')}`,
  nama: ["Bpk. Ahmad", "Ibu Siti", "Bpk. Budi", "Ibu Ani", "Bpk. Joko", "Ibu Rina", "Bpk. Agus", "Ibu Sri", "Bpk. Hendra", "Ibu Maya", "Bpk. Dedi", "Ibu Nina", "Bpk. Yudi", "Ibu Dewi", "Bpk. Eko", "Ibu Ratna", "Bpk. Iwan", "Ibu Sari", "Bpk. Arif", "Ibu Lia", "Bpk. Rizal", "Ibu Ayu", "Bpk. Anton", "Ibu Fitri", "Bpk. Wahyu", "Ibu Dian", "Bpk. Indra", "Ibu Tari", "Bpk. Gilang", "Ibu Wati", "Bpk. Rahmat"][i] || `Warga ${i+1}`,
  rt: `RT 0${(i % 5) + 1}`,
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
  kontak: `0811002200${i}`
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
  
  const [warga] = useState(() => JSON.parse(localStorage.getItem('sys_warga')) || defaultWarga);
  const [panitia] = useState(() => JSON.parse(localStorage.getItem('sys_panitia')) || defaultPanitia);
  const [mudhohi] = useState(() => JSON.parse(localStorage.getItem('sys_mudhohi')) || defaultMudhohi);
  const [hewan, setHewan] = useState(() => JSON.parse(localStorage.getItem('sys_hewan_v7')) || defaultHewan);
  const [kupon, setKupon] = useState(() => JSON.parse(localStorage.getItem('sys_kupon_v7')) || generateKuponInit());
  const [keuangan, setKeuangan] = useState(() => JSON.parse(localStorage.getItem('sys_keuangan')) || defaultKeuangan);
  
  const [qrModal, setQrModal] = useState({ isOpen: false, data: null });
  const [pdfPreview, setPdfPreview] = useState({ isOpen: false, url: '', title: '' });
  const [scanResult, setScanResult] = useState(null);
  const [manualCode, setManualCode] = useState('');

  useEffect(() => {
    localStorage.setItem('sys_hewan_v7', JSON.stringify(hewan));
    localStorage.setItem('sys_kupon_v7', JSON.stringify(kupon));
    localStorage.setItem('sys_keuangan', JSON.stringify(keuangan));
  }, [hewan, kupon, keuangan]);

  const totalSapi = hewan.filter(h => h.jenis === 'Sapi').length;
  const totalKambing = hewan.filter(h => h.jenis === 'Kambing').length;
  const saldo = keuangan.reduce((acc, curr) => curr.jenis === 'Masuk' ? acc + curr.nominal : acc - curr.nominal, 0);

  const pieData = [
    { name: 'Selesai', value: hewan.filter(h=>h.status==='Selesai').length, color: '#10b981' },
    { name: 'Proses', value: hewan.filter(h=>['Disembelih','Dikuliti'].includes(h.status)).length, color: '#8b5cf6' },
    { name: 'Menunggu', value: hewan.filter(h=>h.status==='Menunggu').length, color: '#94a3b8' },
  ].filter(d => d.value > 0);

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

  useEffect(() => {
    if (activeTab === 'scan') {
      const scanner = new Html5QrcodeScanner("reader", { qrbox: 200, fps: 10 });
      scanner.render((txt) => {
        const found = kupon.find(k => k.id === txt);
        if(!found) setScanResult({ status:'error', pesan:'Tidak Terdaftar' });
        else if(found.status==='Sudah Diambil') setScanResult({ status:'error', pesan:'SUDAH DIAMBIL' });
        else {
          setKupon(prev => prev.map(k => k.id === txt ? {...k, status:'Sudah Diambil'} : k));
          setScanResult({ status:'success', pesan:`Berhasil: ${found.nama}` });
        }
      }, () => {});
      return () => { scanner.clear().catch(()=>{}); };
    }
  }, [activeTab]);

  return (
    <div className="flex h-screen bg-slate-50 font-sans overflow-hidden">
      {/* SIDEBAR */}
      <aside className={`fixed md:sticky top-0 bottom-0 left-0 w-64 bg-[#059669] text-white flex flex-col shadow-xl z-30 transition-transform duration-300 ${isSidebarOpen ? 'translate-x-0' : '-translate-x-full md:translate-x-0'}`}>
        <div className="p-6 flex items-center justify-between border-b border-emerald-700/50">
          <div className="flex items-center gap-2"><div className="bg-amber-400 p-1.5 rounded-lg text-emerald-900 font-bold">Q</div><h1 className="font-bold text-lg">QurbanPro</h1></div>
          <button onClick={() => setIsSidebarOpen(false)} className="md:hidden"><X size={20}/></button>
        </div>
        <div className="flex-1 overflow-y-auto p-3 space-y-1">
          {menuItems.map(m => (
            <button key={m.id} onClick={()=>{setActiveTab(m.id); setIsSidebarOpen(false)}} className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl transition ${activeTab===m.id?'bg-white text-[#059669] font-bold shadow-sm':'hover:bg-emerald-800/40'}`}>
              <m.icon size={18}/> <span className="text-sm">{m.label}</span>
            </button>
          ))}
        </div>
      </aside>

      {isSidebarOpen && <div onClick={()=>setIsSidebarOpen(false)} className="fixed inset-0 bg-black/50 z-20 md:hidden"></div>}

      {/* MAIN */}
      <main className="flex-1 flex flex-col min-w-0 overflow-y-auto">
        <header className="bg-white border-b p-4 px-6 flex justify-between items-center sticky top-0 z-10">
          <div className="flex items-center gap-3">
            <button onClick={()=>setIsSidebarOpen(true)} className="md:hidden p-1 text-slate-600"><Menu size={24}/></button>
            <h2 className="font-bold text-slate-800 capitalize">{activeTab}</h2>
          </div>
          <div className="flex items-center gap-2 bg-emerald-50 text-emerald-700 px-3 py-1 rounded-full text-xs font-bold">
            <div className="w-2 h-2 bg-emerald-500 rounded-full animate-pulse"></div> Aktif
          </div>
        </header>

        <div className="p-4 md:p-8">
          {activeTab === 'dashboard' && (
            <div className="space-y-6">
              <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
                <Card title="Hewan Qurban" value={hewan.length} sub={`${totalSapi} Sapi, ${totalKambing} Kambing`} color="emerald"/>
                <Card title="Kupon Claim" value={kupon.filter(k=>k.status==='Sudah Diambil').length} sub={`Total ${kupon.length}`} color="blue"/>
                <Card title="Mudhohi" value={mudhohi.length} sub="Peserta" color="amber"/>
                <Card title="Saldo" value={`Rp ${(saldo/1000000).toFixed(1)}jt`} sub="Kas Panitia" color="purple"/>
              </div>
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                <div className="bg-white p-4 rounded-xl border h-[300px] shadow-sm"><ResponsiveContainer><PieChart><Pie data={pieData} cx="50%" cy="50%" innerRadius={40} outerRadius={70} dataKey="value" label>{pieData.map((e,i)=><Cell key={i} fill={e.color}/>)}</Pie><Tooltip/></PieChart></ResponsiveContainer></div>
                <div className="bg-white p-4 rounded-xl border h-[300px] shadow-sm"><ResponsiveContainer><BarChart data={[{n:'Saldo', v:saldo}]}><XAxis dataKey="n"/><YAxis/><Tooltip/><Bar dataKey="v" fill="#10b981"/></BarChart></ResponsiveContainer></div>
              </div>
            </div>
          )}

          {activeTab === 'warga' && (
            <div className="bg-white rounded-xl border shadow-sm overflow-hidden">
              <div className="p-4 border-b font-bold">Database Warga</div>
              <table className="w-full text-left text-sm">
                <thead className="bg-slate-50"><tr><th className="p-3">ID</th><th className="p-3">Nama</th><th className="p-3">RT</th></tr></thead>
                <tbody className="divide-y">{warga.map(w=><tr key={w.id}> <td className="p-3 font-mono text-slate-400">{w.id}</td><td className="p-3 font-semibold">{w.nama}</td><td className="p-3">{w.rt}</td></tr>)}</tbody>
              </table>
            </div>
          )}

          {activeTab === 'panitia' && (
            <div className="bg-white rounded-xl border shadow-sm overflow-hidden">
              <div className="p-4 border-b font-bold">Data Panitia</div>
              <table className="w-full text-left text-sm">
                <thead className="bg-slate-50"><tr><th className="p-3">Nama</th><th className="p-3">Jabatan</th></tr></thead>
                <tbody className="divide-y">{panitia.map(p=><tr key={p.id}><td className="p-3 font-semibold">{p.nama}</td><td className="p-3 text-blue-600">{p.peran}</td></tr>)}</tbody>
              </table>
            </div>
          )}

          {activeTab === 'hewan' && (
            <div className="bg-white rounded-xl border shadow-sm overflow-hidden">
              <div className="p-4 border-b font-bold">Hewan Qurban</div>
              <div className="overflow-x-auto"><table className="w-full text-left text-sm">
                <thead className="bg-slate-50"><tr><th className="p-3">ID</th><th className="p-3">Jenis</th><th className="p-3">Status</th></tr></thead>
                <tbody className="divide-y">{hewan.map(h=><tr key={h.id}><td className="p-3 font-bold">{h.id}</td><td className="p-3">{h.jenis}</td><td className="p-3"><span className="bg-slate-100 px-2 py-1 rounded text-xs">{h.status}</span></td></tr>)}</tbody>
              </table></div>
            </div>
          )}

          {activeTab === 'kupon' && (
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              {['Warga', 'Mudhohi'].map(t => (
                <div key={t} className="bg-white rounded-xl border shadow-sm overflow-hidden">
                  <div className="p-3 bg-slate-50 border-b font-bold">Kupon {t}</div>
                  <div className="max-h-96 overflow-y-auto"><table className="w-full text-xs">
                    <tbody className="divide-y">{kupon.filter(k=>k.tipe===t).map(k=><tr key={k.id}>
                      <td className="p-3 font-mono">{k.id}</td><td className="p-3">{k.nama}</td>
                      <td className="p-3"><button onClick={()=>prosesKuponDigital(k,true)} className="p-1 border rounded mr-1"><Eye size={12}/></button><button onClick={()=>prosesKuponDigital(k,false)} className="p-1 border rounded bg-emerald-50 text-emerald-600"><DownloadCloud size={12}/></button></td>
                    </tr>)}</tbody>
                  </table></div>
                </div>
              ))}
            </div>
          )}

          {activeTab === 'scan' && (
            <div className="bg-white p-6 rounded-xl border shadow-sm text-center">
              <div id="reader" className="mx-auto max-w-xs overflow-hidden rounded-lg bg-black mb-6"></div>
              <div className="flex gap-2 max-w-xs mx-auto mb-4">
                <input type="text" value={manualCode} onChange={e=>setManualCode(e.target.value)} placeholder="ID Kupon" className="flex-1 border p-2 rounded text-sm font-mono"/>
                <button onClick={()=>prosesValidasiKupon(manualCode)} className="bg-emerald-600 text-white px-4 py-2 rounded text-sm">Cek</button>
              </div>
              {scanResult && <div className={`p-3 rounded-lg font-bold text-sm ${scanResult.status==='success'?'bg-emerald-50 text-emerald-700':'bg-red-50 text-red-700'}`}>{scanResult.pesan}</div>}
            </div>
          )}

          {activeTab === 'sertifikat' && (
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div className="bg-white rounded-xl border shadow-sm p-4 h-96 flex flex-col">
                <h4 className="font-bold mb-4">Mudhohi</h4>
                <div className="overflow-y-auto divide-y">{mudhohi.map(m=><div key={m.id} className="py-2 flex justify-between items-center text-sm"><span>{m.nama}</span><div className="flex gap-1"><button onClick={()=>prosesSertifikat(m.nama, '', true, true)} className="p-1 border rounded"><Eye size={14}/></button><button onClick={()=>prosesSertifikat(m.nama, '', true, false)} className="p-1 border rounded bg-emerald-50 text-emerald-600"><DownloadCloud size={14}/></button></div></div>)}</div>
              </div>
              <div className="bg-white rounded-xl border shadow-sm p-4 h-96 flex flex-col">
                <h4 className="font-bold mb-4">Panitia</h4>
                <div className="overflow-y-auto divide-y">{panitia.map(p=><div key={p.id} className="py-2 flex justify-between items-center text-sm"><span>{p.nama}</span><div className="flex gap-1"><button onClick={()=>prosesSertifikat(p.nama, p.peran, false, true)} className="p-1 border rounded"><Eye size={14}/></button><button onClick={()=>prosesSertifikat(p.nama, p.peran, false, false)} className="p-1 border rounded bg-blue-50 text-blue-600"><DownloadCloud size={14}/></button></div></div>)}</div>
              </div>
            </div>
          )}

          {activeTab === 'uang' && (
            <div className="bg-white rounded-xl border shadow-sm overflow-hidden">
              <div className="p-4 bg-slate-900 text-white flex justify-between"><span>Laporan Keuangan</span><span className="text-emerald-400 font-bold">Saldo: Rp {saldo.toLocaleString()}</span></div>
              <table className="w-full text-sm">
                <thead className="bg-slate-50 border-b"><tr><th className="p-3">Uraian</th><th className="p-3 text-right">Nominal</th></tr></thead>
                <tbody className="divide-y">{keuangan.map(k=><tr key={k.id}><td className="p-3">{k.keterangan}</td><td className={`p-3 text-right font-bold ${k.jenis==='Masuk'?'text-emerald-600':'text-red-500'}`}>Rp {k.nominal.toLocaleString()}</td></tr>)}</tbody>
              </table>
            </div>
          )}
        </div>
      </main>

      {/* PDF MODAL */}
      {pdfPreview.isOpen && (
        <div className="fixed inset-0 bg-black/80 flex flex-col z-50 p-2 md:p-10">
          <div className="flex justify-between bg-white p-3 rounded-t-xl border-b">
            <h3 className="font-bold truncate">{pdfPreview.title}</h3>
            <button onClick={()=>setPdfPreview({isOpen:false,url:'',title:''})}><X/></button>
          </div>
          <iframe src={pdfPreview.url} className="flex-1 w-full bg-white rounded-b-xl border-0" />
        </div>
      )}
    </div>
  );
}

function Card({ title, value, sub, color }) {
  const colors = { emerald: 'border-l-emerald-500', blue: 'border-l-blue-500', amber: 'border-l-amber-500', purple: 'border-l-purple-500' };
  return (
    <div className={`bg-white p-4 rounded-xl border border-slate-200 border-l-4 ${colors[color]} shadow-sm`}>
      <p className="text-[10px] uppercase font-bold text-slate-500">{title}</p>
      <h3 className="text-xl md:text-2xl font-black text-slate-800 my-1">{value}</h3>
      <p className="text-[9px] text-slate-400">{sub}</p>
    </div>
  );
}
