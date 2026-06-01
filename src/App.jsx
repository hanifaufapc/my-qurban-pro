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
  const [hewan, setHewan] = useState(() => JSON.parse(localStorage.getItem('sys_hewan_vfinal')) || defaultHewan);
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
  useEffect(() => { localStorage.setItem('sys_hewan_vfinal', JSON.stringify(hewan)); }, [hewan]);
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
    try {
      const qrDataUrl = await QRCode.toDataURL(k.id, { margin: 1, width: 100 });
      doc.addImage(qrDataUrl, 'PNG', 25, 20, 30, 30);
    } catch (err) { console.error(err); }
    doc.setFontSize(10); doc.setTextColor(0, 0, 0); doc.text(k.nama, 40, 58, { align: "center" });
    doc.text(`ID: ${k.id}`, 40, 65, { align: "center" });
    if(isPreview) { setPdfPreview({ isOpen: true, url: doc.output('bloburl'), title: `Kupon ${k.id}` }); } 
    else { doc.save(`Kupon_${k.id}.pdf`); }
  };

  const prosesSertifikatMudhohi = (m, isPreview = false) => {
    const doc = new jsPDF('landscape');
    doc.setDrawColor(16, 185, 129); doc.setLineWidth(3); doc.rect(10, 10, 277, 190);
    doc.setFontSize(30); doc.setTextColor(16, 185, 129); doc.text("SERTIFIKAT QURBAN", 148, 50, { align: "center" });
    doc.setFontSize(28); doc.setTextColor(15, 23, 42); doc.text(m.nama, 148, 105, { align: "center" });
    if(isPreview) { setPdfPreview({ isOpen: true, url: doc.output('bloburl'), title: `Sertifikat ${m.nama}` }); } 
    else { doc.save(`Sertifikat_${m.id}.pdf`); }
  };

  const prosesSertifikatPanitia = (p, isPreview = false) => {
    const doc = new jsPDF('landscape');
    doc.setDrawColor(59, 130, 246); doc.setLineWidth(3); doc.rect(10, 10, 277, 190);
    doc.setFontSize(30); doc.setTextColor(59, 130, 246); doc.text("PIAGAM PENGHARGAAN", 148, 50, { align: "center" });
    doc.setFontSize(28); doc.setTextColor(15, 23, 42); doc.text(p.nama, 148, 105, { align: "center" });
    if(isPreview) { setPdfPreview({ isOpen: true, url: doc.output('bloburl'), title: `Piagam ${p.nama}` }); } 
    else { doc.save(`Piagam_${p.id}.pdf`); }
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
    if (kupon
