/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef, useMemo } from 'react';
import { 
  collection, 
  doc, 
  setDoc, 
  deleteDoc, 
  onSnapshot, 
  query, 
  getDocFromServer 
} from 'firebase/firestore';
import { 
  signInWithPopup,
  GoogleAuthProvider,
  onAuthStateChanged,
  User as FirebaseUser
} from 'firebase/auth';
import { 
  Search, Plus, Book, Check, X, Edit2, Trash2, User, Loader2, Sparkles, 
  Image as ImageIcon, Camera, ScanLine, Library, LayoutGrid, List as ListIcon, 
  Settings, ChevronDown, ChevronUp, Layers, Palette, Download, SortAsc, LogOut
} from 'lucide-react';
import { db, auth } from './firebase';
import firebaseConfig from '../firebase-applet-config.json';

import { Html5Qrcode, Html5QrcodeSupportedFormats } from 'html5-qrcode';
import * as XLSX from 'xlsx';

// --- Constants ---
const CATEGORIES = ["Novel", "Komik", "Biografi", "Non-Fiksi", "Pelajaran"];
const appId = firebaseConfig.projectId;

// --- Error Handling ---
enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId: string | undefined;
    email: string | null | undefined;
    emailVerified: boolean | undefined;
    isAnonymous: boolean | undefined;
    tenantId: string | null | undefined;
    providerInfo: {
      providerId: string;
      displayName: string | null;
      email: string | null;
      photoUrl: string | null;
    }[];
  }
}

function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null) {
  const errInfo: FirestoreErrorInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: auth.currentUser?.uid,
      email: auth.currentUser?.email,
      emailVerified: auth.currentUser?.emailVerified,
      isAnonymous: auth.currentUser?.isAnonymous,
      tenantId: auth.currentUser?.tenantId,
      providerInfo: auth.currentUser?.providerData.map(provider => ({
        providerId: provider.providerId,
        displayName: provider.displayName,
        email: provider.email,
        photoUrl: provider.photoURL
      })) || []
    },
    operationType,
    path
  }
  console.error('Firestore Error: ', JSON.stringify(errInfo));
  throw new Error(JSON.stringify(errInfo));
}

// --- Theme Definitions ---
const THEMES: Record<string, any> = {
  pink: { 
    name: 'Pastel Pink',
    bg: 'bg-pink-50/50', header: 'from-pink-300 to-purple-300', primary: 'bg-pink-400', 
    primaryHover: 'hover:bg-pink-500', text: 'text-pink-600', border: 'border-pink-100', accent: 'bg-pink-100' 
  },
  lavender: { 
    name: 'Lavender Dream',
    bg: 'bg-purple-50/50', header: 'from-purple-300 to-indigo-300', primary: 'bg-purple-400', 
    primaryHover: 'hover:bg-purple-500', text: 'text-purple-600', border: 'border-purple-100', accent: 'bg-purple-100' 
  },
  rosegold: { 
    name: 'Rose Gold Elegance',
    bg: 'bg-rose-50/50', header: 'from-rose-300 to-orange-300', primary: 'bg-rose-400', 
    primaryHover: 'hover:bg-rose-500', text: 'text-rose-600', border: 'border-rose-100', accent: 'bg-rose-100' 
  },
  sky: { 
    name: 'Sky Blossom',
    bg: 'bg-sky-50/50', header: 'from-sky-300 to-blue-300', primary: 'bg-sky-400', 
    primaryHover: 'hover:bg-sky-500', text: 'text-sky-600', border: 'border-sky-100', accent: 'bg-sky-100' 
  }
};

// --- Book App ---
interface BookData {
  id: string;
  title: string;
  author: string;
  category: string;
  seriesName: string;
  isRead: boolean;
  page: string;
  coverUrl: string;
  price: number;
  synopsis: string;
  addedAt: number;
  updatedAt: number;
}

function BookApp() {
  // State Data & UI
  const [user, setUser] = useState<FirebaseUser | null>(null);
  const [books, setBooks] = useState<BookData[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [searchTerm, setSearchTerm] = useState('');
  const [viewMode, setViewMode] = useState<'list' | 'grid'>('grid');
  const [activeTheme, setActiveTheme] = useState('pink');
  const [sortBy, setSortBy] = useState<'title' | 'addedAt' | 'isRead' | 'author'>('addedAt');
  
  // State Modals & Forms
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [isThemeModalOpen, setIsThemeModalOpen] = useState(false);
  const [isSynopsisModalOpen, setIsSynopsisModalOpen] = useState(false);
  const [isNetWorthModalOpen, setIsNetWorthModalOpen] = useState(false);
  const [isSearchingApi, setIsSearchingApi] = useState(false);
  const [isScanning, setIsScanning] = useState(false);
  const [isManualScanning, setIsManualScanning] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [currentBook, setCurrentBook] = useState<BookData>(getEmptyBook());
  const [isbnInput, setIsbnInput] = useState('');
  const [expandedSeries, setExpandedSeries] = useState<Record<string, boolean>>({});
  
  // Synopsis Modal State
  const [randomBooks, setRandomBooks] = useState<BookData[]>([]);
  const [synopsisMap, setSynopsisMap] = useState<Record<string, { text: string; loading: boolean }>>({});
  
  // Context Menu State
  const [actionMenuBook, setActionMenuBook] = useState<BookData | null>(null);
  const pressTimer = useRef<NodeJS.Timeout | null>(null);
  const scannerRef = useRef<Html5Qrcode | null>(null);

  const t = THEMES[activeTheme];

  const handleManualCapture = async () => {
    if (!scannerRef.current || !isScanning) return;
    
    setIsManualScanning(true);
    try {
      const video = document.querySelector('#reader video') as HTMLVideoElement;
      if (video) {
        const canvas = document.createElement('canvas');
        canvas.width = video.videoWidth;
        canvas.height = video.videoHeight;
        const ctx = canvas.getContext('2d');
        if (ctx) {
          ctx.drawImage(video, 0, 0);
          canvas.toBlob(async (blob) => {
            if (blob) {
              const file = new File([blob], "barcode.jpg", { type: "image/jpeg" });
              
              // Create temp container for scan
              const tempId = "reader-manual-temp";
              let tempElem = document.getElementById(tempId);
              if (!tempElem) {
                tempElem = document.createElement('div');
                tempElem.id = tempId;
                tempElem.style.display = 'none';
                document.body.appendChild(tempElem);
              }

              const html5QrCode = new Html5Qrcode(tempId, {
                formatsToSupport: [ 
                  Html5QrcodeSupportedFormats.EAN_13, 
                  Html5QrcodeSupportedFormats.EAN_8, 
                  Html5QrcodeSupportedFormats.UPC_A, 
                  Html5QrcodeSupportedFormats.UPC_E,
                  Html5QrcodeSupportedFormats.CODE_128,
                  Html5QrcodeSupportedFormats.CODE_39
                ],
                useBarCodeDetectorIfSupported: true,
                verbose: false,
                experimentalFeatures: {
                  useBarCodeDetectorIfSupported: true
                }
              });
              try {
                const decodedText = await html5QrCode.scanFile(file, true);
                setIsbnInput(decodedText);
                stopScanner();
                handleSmartSearch('isbn', decodedText);
              } catch (err) {
                alert("Barcode tidak terdeteksi. Pastikan barcode terlihat jelas dan berada tepat di kotak tengah, lalu coba lagi.");
              } finally {
                if (tempElem.parentNode) tempElem.parentNode.removeChild(tempElem);
              }
            }
            setIsManualScanning(false);
          }, 'image/jpeg', 0.95);
        }
      } else {
        setIsManualScanning(false);
      }
    } catch (err) {
      console.error("Manual capture failed", err);
      setIsManualScanning(false);
    }
  };

  const startScanner = async () => {
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
      alert("Kamera tidak didukung atau tidak dapat diakses di browser ini. Pastikan Anda membuka aplikasi di tab baru jika di dalam editor.");
      return;
    }

    if (scannerRef.current) {
      await stopScanner();
    }

    setIsScanning(true);
    
    // Tunggu elemen "reader" muncul di DOM
    let attempts = 0;
    const maxAttempts = 20;
    
    const checkAndStart = async () => {
      const element = document.getElementById("reader");
      if (element) {
        try {
          const html5QrCode = new Html5Qrcode("reader", {
            formatsToSupport: [ 
              Html5QrcodeSupportedFormats.EAN_13, 
              Html5QrcodeSupportedFormats.EAN_8, 
              Html5QrcodeSupportedFormats.UPC_A, 
              Html5QrcodeSupportedFormats.UPC_E,
              Html5QrcodeSupportedFormats.CODE_128,
              Html5QrcodeSupportedFormats.CODE_39
            ],
            useBarCodeDetectorIfSupported: true,
            verbose: false,
            experimentalFeatures: {
              useBarCodeDetectorIfSupported: true
            }
          });
          scannerRef.current = html5QrCode;
          await html5QrCode.start(
            { facingMode: "environment" },
            { 
              fps: 15, 
              qrbox: { width: 300, height: 150 }
            },
            (decodedText) => {
              setIsbnInput(decodedText);
              stopScanner();
              handleSmartSearch('isbn', decodedText);
            },
            () => {} // error callback (ignore common scanning failures)
          );
        } catch (err: any) {
          console.error("Scanner failed to start", err);
          alert("Gagal mengakses kamera: " + (err.message || err));
          setIsScanning(false);
          scannerRef.current = null;
        }
      } else if (attempts < maxAttempts) {
        attempts++;
        setTimeout(checkAndStart, 100);
      } else {
        setIsScanning(false);
        alert("Elemen kamera gagal dimuat.");
      }
    };
    
    // Initial delay untuk memberikan waktu mounting modal
    setTimeout(checkAndStart, 300);
  };

  const stopScanner = async () => {
    if (scannerRef.current) {
      try {
        if (scannerRef.current.isScanning) {
          await scannerRef.current.stop();
        }
        await scannerRef.current.clear();
      } catch (err) {
        console.error("Stop scanner failed", err);
      }
      scannerRef.current = null;
    }
    setIsScanning(false);
  };

  const handleBarcodeUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    
    setIsSearchingApi(true);
    // Create a temporary element for scanning
    const tempId = "reader-file-scan";
    let tempElem = document.getElementById(tempId);
    if (!tempElem) {
      tempElem = document.createElement('div');
      tempElem.id = tempId;
      tempElem.style.display = 'none';
      document.body.appendChild(tempElem);
    }

    const html5QrCode = new Html5Qrcode(tempId);
    try {
      const decodedText = await html5QrCode.scanFile(file, true);
      setIsbnInput(decodedText);
      await handleSmartSearch('isbn', decodedText);
    } catch (err) {
      alert("Gagal membaca barcode dari gambar. Pastikan gambar jelas dan berisi barcode ISBN.");
      console.error(err);
    } finally {
      setIsSearchingApi(false);
      if (tempElem.parentNode) tempElem.parentNode.removeChild(tempElem);
      e.target.value = '';
    }
  };

  const handleExportExcel = () => {
    if (books.length === 0) {
      alert("Tidak ada data untuk diekspor.");
      return;
    }

    const exportData = books.map(book => ({
      'Judul Buku': book.title,
      'Penulis': book.author,
      'Kategori': book.category,
      'Nama Seri': book.seriesName || '-',
      'Status': book.isRead ? 'Selesai' : 'Belum Selesai',
      'Halaman Terakhir': book.isRead ? '-' : (book.page || '0'),
      'Tanggal Ditambahkan': new Date(book.addedAt).toLocaleDateString('id-ID', {
        day: '2-digit',
        month: 'long',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
      }),
      'Link Cover': book.coverUrl || '-'
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    
    // Auto-width adjustment
    const colWidths = [
      { wch: 30 }, // Judul
      { wch: 20 }, // Penulis
      { wch: 15 }, // Kategori
      { wch: 20 }, // Seri
      { wch: 15 }, // Status
      { wch: 15 }, // Halaman
      { wch: 25 }, // Tanggal
      { wch: 50 }, // Cover
    ];
    worksheet['!cols'] = colWidths;

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Koleksi Buku");
    
    XLSX.writeFile(workbook, `Koleksi_Buku_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const handleUpdatePrice = async (bookId: string, newPrice: number) => {
    if (!user) return;
    const path = `artifacts/${appId}/users/${user.uid}/books`;
    try {
      await setDoc(doc(db, path, bookId), { price: newPrice, updatedAt: Date.now() }, { merge: true });
    } catch (error) {
      console.error("Gagal update harga", error);
    }
  };

  const totalNetWorth = useMemo(() => {
    return books.reduce((sum, book) => sum + (book.price || 0), 0);
  }, [books]);

  const openSynopsisModal = () => {
    const unread = books.filter(b => !b.isRead);
    if (unread.length === 0) {
      alert("Tidak ada buku yang sedang dibaca (belum selesai).");
      return;
    }
    const shuffled = [...unread].sort(() => 0.5 - Math.random());
    const selected = shuffled.slice(0, 5);
    setRandomBooks(selected);
    setIsSynopsisModalOpen(true);

    // Fetch synopsis for each selected book
    selected.forEach(async (book) => {
      // Use cached synopsis if available
      if (book.synopsis && book.synopsis !== "Sinopsis tidak ditemukan untuk buku ini.") {
        setSynopsisMap(prev => ({
          ...prev,
          [book.id]: { text: book.synopsis, loading: false }
        }));
        return;
      }

      // Initialize loading state
      setSynopsisMap(prev => ({
        ...prev,
        [book.id]: { text: '', loading: true }
      }));

      try {
        const queryStr = book.author && book.author !== 'Anonim' 
          ? `intitle:${book.title}+inauthor:${book.author}` 
          : `intitle:${book.title}`;
          
        const apiKey = import.meta.env.VITE_GOOGLE_BOOKS_API_KEY;
        const url = `https://www.googleapis.com/books/v1/volumes?q=${encodeURIComponent(queryStr)}&maxResults=1${apiKey ? `&key=${apiKey}` : ''}`;
        
        const response = await fetch(url);

        if (response.status === 429) {
          throw new Error("Limit tercapai. Butuh API Key.");
        }

        const data = await response.json();
        
        let synopsisText = "Sinopsis tidak ditemukan untuk buku ini.";
        if (data.items && data.items.length > 0) {
          const description = data.items[0].volumeInfo.description;
          if (description) {
            synopsisText = description;
          }
        }
        
        // Cache the found synopsis to Firestore
        if (user && synopsisText !== "Sinopsis tidak ditemukan untuk buku ini.") {
          const path = `artifacts/${appId}/users/${user.uid}/books`;
          await setDoc(doc(db, path, book.id), { synopsis: synopsisText, updatedAt: Date.now() }, { merge: true });
        }

        setSynopsisMap(prev => ({
          ...prev,
          [book.id]: { text: synopsisText, loading: false }
        }));
      } catch (err: any) {
        console.error("Failed to fetch synopsis", err);
        setSynopsisMap(prev => ({
          ...prev,
          [book.id]: { text: "Gagal memuat sinopsis: " + (err.message || "Masalah koneksi"), loading: false }
        }));
      }
    });
  };

  const renderSharedModals = () => (
    <>
      {/* Scanner Modal */}
      {isScanning && (
        <div className="fixed inset-0 z-[110] bg-black flex flex-col">
          <div className="flex justify-between items-center p-4 text-white">
            <h3 className="font-bold">Scan Barcode ISBN</h3>
            <button onClick={stopScanner} className="p-2 bg-white/10 rounded-full active:scale-95 transition-all"><X /></button>
          </div>
          
          <div className="flex-1 flex flex-col relative overflow-hidden bg-gray-900">
            <div id="reader" className="w-full h-full"></div>
            
            {/* Guide Overlay */}
            <div className="absolute inset-0 pointer-events-none flex flex-col items-center justify-center">
              <div className="w-64 h-40 border-2 border-white/50 rounded-2xl relative shadow-[0_0_0_400px_rgba(0,0,0,0.5)]">
                <div className="absolute top-0 left-0 w-4 h-4 border-t-4 border-l-4 border-white -translate-x-1 -translate-y-1"></div>
                <div className="absolute top-0 right-0 w-4 h-4 border-t-4 border-r-4 border-white translate-x-1 -translate-y-1"></div>
                <div className="absolute bottom-0 left-0 w-4 h-4 border-b-4 border-l-4 border-white -translate-x-1 translate-y-1"></div>
                <div className="absolute bottom-0 right-0 w-4 h-4 border-b-4 border-r-4 border-white translate-x-1 translate-y-1"></div>
              </div>
              <p className="text-white font-bold text-xs mt-6 bg-black/40 px-4 py-2 rounded-full backdrop-blur-md">Posisikan barcode dalam kotak</p>
            </div>
          </div>

          <div className="p-8 bg-black flex flex-col items-center gap-4">
            <button 
              onClick={handleManualCapture}
              disabled={isManualScanning}
              className={`w-full max-w-sm py-4 rounded-3xl font-black transition-all flex items-center justify-center gap-2 active:scale-95 ${isManualScanning ? 'bg-gray-700 text-gray-400' : 'bg-white text-black shadow-lg shadow-white/20'}`}
            >
              {isManualScanning ? (
                <>
                  <Loader2 className="w-6 h-6 animate-spin" /> Sedang Memindai...
                </>
              ) : (
                <>
                  <ScanLine className="w-6 h-6" /> AMBIL BARCODE SEKARANG
                </>
              )}
            </button>
            <p className="text-white/40 text-[10px] uppercase font-black tracking-widest">Gunakan tombol jika deteksi otomatis lambat</p>
          </div>
        </div>
      )}

      {/* Shared Modals */}
      {actionMenuBook && (
        <div className="fixed inset-0 z-[100] flex items-end justify-center bg-gray-900/40 backdrop-blur-sm p-4 animate-in fade-in duration-200">
          <div className="bg-white w-full max-w-md rounded-[2.5rem] p-6 shadow-2xl transform transition-transform animate-in slide-in-from-bottom-10 mb-2">
            <h3 className="font-bold text-center text-gray-800 mb-2 truncate px-4">{actionMenuBook.title}</h3>
            <div className="grid grid-cols-2 gap-4 mt-6">
              <button onClick={() => { 
                  setCurrentBook({...actionMenuBook}); 
                  setActionMenuBook(null); 
                  setIsModalOpen(true); 
                }}
                className="bg-blue-50 text-blue-600 p-4 rounded-2xl font-bold flex flex-col items-center gap-2 active:scale-95 transition-all">
                <Edit2 className="w-6 h-6" /> Edit
              </button>
              <button onClick={() => handleDeleteBook(actionMenuBook.id)}
                className="bg-red-50 text-red-500 p-4 rounded-2xl font-bold flex flex-col items-center gap-2 active:scale-95 transition-all">
                <Trash2 className="w-6 h-6" /> Hapus
              </button>
            </div>
            <button onClick={() => setActionMenuBook(null)} className="w-full mt-4 p-4 text-gray-400 font-bold active:bg-gray-50 rounded-2xl transition-colors">Batal</button>
          </div>
        </div>
      )}

      {isSettingsOpen && (
        <div className="fixed inset-0 z-[100] flex items-end justify-center bg-gray-900/40 backdrop-blur-sm p-4 animate-in fade-in duration-200">
          <div className="bg-white w-full max-w-md rounded-[2.5rem] p-6 shadow-2xl animate-in slide-in-from-bottom-10 mb-2">
            <div className="flex items-center justify-between mb-6">
              <h3 className="font-bold text-lg flex items-center gap-2 text-gray-800"><Settings className={`w-5 h-5 ${t.text}`} /> Pengaturan</h3>
              <button onClick={() => setIsSettingsOpen(false)} className="bg-gray-100 p-2 rounded-full text-gray-500 active:bg-gray-200"><X className="w-4 h-4" /></button>
            </div>
            
            <div className="space-y-6">
              {/* User Profile Section */}
              <div className="bg-gray-50 p-4 rounded-3xl flex items-center gap-4">
                {user?.photoURL ? (
                  <img src={user.photoURL} alt="Profile" className="w-14 h-14 rounded-full border-2 border-white shadow-sm" />
                ) : (
                  <div className={`w-14 h-14 rounded-full ${t.accent} flex items-center justify-center text-gray-400`}>
                    <User className="w-8 h-8" />
                  </div>
                )}
                <div className="min-w-0">
                  <p className="font-bold text-gray-800 truncate">{user?.displayName || 'Pengguna'}</p>
                  <p className="text-xs text-gray-500 truncate">{user?.email}</p>
                </div>
              </div>

              <div className="space-y-2">
                <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-2">Menu Utama</p>
                <div className="grid grid-cols-1 gap-2">
                  <button onClick={() => { setIsSettingsOpen(false); setIsThemeModalOpen(true); }}
                    className="w-full p-4 bg-white border border-gray-100 rounded-2xl flex items-center gap-3 active:bg-gray-50 transition-all">
                    <div className={`p-2 rounded-xl ${t.accent} ${t.text}`}><Palette className="w-5 h-5" /></div>
                    <div className="text-left">
                      <p className="font-bold text-sm text-gray-800">Tema Aplikasi</p>
                      <p className="text-[10px] text-gray-400">Ganti warna suasana aplikasi</p>
                    </div>
                  </button>

                  <button onClick={handleExportExcel}
                    className="w-full p-4 bg-white border border-gray-100 rounded-2xl flex items-center gap-3 active:bg-gray-50 transition-all">
                    <div className="p-2 rounded-xl bg-green-50 text-green-600"><Download className="w-5 h-5" /></div>
                    <div className="text-left">
                      <p className="font-bold text-sm text-gray-800">Ekspor ke Excel</p>
                      <p className="text-[10px] text-gray-400">Download semua data buku</p>
                    </div>
                  </button>
                </div>
              </div>

              <div className="pt-4 border-t border-gray-100">
                <button onClick={handleLogout} className="w-full p-4 bg-red-50 text-red-500 rounded-2xl font-bold flex items-center justify-center gap-2 active:scale-95 transition-all">
                  <LogOut className="w-5 h-5" /> Keluar Akun
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {isThemeModalOpen && (
        <div className="fixed inset-0 z-[110] flex items-end justify-center bg-gray-900/40 backdrop-blur-sm p-4 animate-in fade-in duration-200">
          <div className="bg-white w-full max-w-md rounded-[2.5rem] p-6 shadow-2xl animate-in slide-in-from-bottom-10 mb-2">
            <div className="flex items-center justify-between mb-6">
              <h3 className="font-bold text-lg flex items-center gap-2 text-gray-800"><Palette className={`w-5 h-5 ${t.text}`} /> Pilih Tema</h3>
              <button onClick={() => setIsThemeModalOpen(false)} className="bg-gray-100 p-2 rounded-full text-gray-500 active:bg-gray-200"><X className="w-4 h-4" /></button>
            </div>
            <div className="space-y-3">
              {Object.keys(THEMES).map(key => (
                <button key={key} onClick={() => { setActiveTheme(key); setIsThemeModalOpen(false); }} 
                  className={`w-full p-4 rounded-2xl flex items-center justify-between border-2 transition-all ${activeTheme === key ? (THEMES[key].border) + ' ' + (THEMES[key].bg) : 'border-gray-50 bg-white'}`}>
                  <span className="font-bold text-gray-700">{THEMES[key].name}</span>
                  <div className={`w-6 h-6 rounded-full bg-gradient-to-br ${THEMES[key].header}`}></div>
                </button>
              ))}
            </div>
          </div>
        </div>
      )}

      {isSynopsisModalOpen && (
        <div className="fixed inset-0 z-[110] flex flex-col bg-white animate-in slide-in-from-bottom-full duration-300 sm:max-w-md sm:mx-auto">
          <header className={`p-6 bg-gradient-to-br ${t.header} text-white shadow-lg flex justify-between items-center rounded-b-[2.5rem] sticky top-0 z-10`}>
            <div>
              <h3 className="text-xl font-bold flex items-center gap-2 italic"><Sparkles className="w-6 h-6" /> Intip Bacaanmu</h3>
              <p className="text-white/80 text-xs">Mengingat kembali apa yang sedang kau baca...</p>
            </div>
            <button onClick={() => setIsSynopsisModalOpen(false)} className="bg-white/20 p-2 rounded-full backdrop-blur-sm active:scale-95 transition-all"><X className="w-5 h-5" /></button>
          </header>

          <div className={`flex-1 overflow-y-auto p-6 space-y-8 pb-32 ${t.bg}`}>
            {randomBooks.map((book, idx) => (
              <div key={book.id} className="animate-in fade-in slide-in-from-bottom-10" style={{ animationDelay: `${idx * 150}ms` }}>
                 <div className="flex gap-4 items-start mb-3">
                   <div className="w-20 h-28 shrink-0 rounded-2xl bg-white shadow-lg overflow-hidden border-2 border-white">
                     {book.coverUrl ? <img src={book.coverUrl} className="w-full h-full object-cover" referrerPolicy="no-referrer" /> : <div className={`w-full h-full flex items-center justify-center ${t.accent}`}><Book className={`w-8 h-8 ${t.text} opacity-20`} /></div>}
                   </div>
                   <div className="flex-1 py-1">
                     <h4 className="font-bold text-gray-800 leading-tight line-clamp-2">{book.title}</h4>
                     <p className="text-xs text-gray-400 mt-1 font-medium">{book.author}</p>
                     <div className={`mt-3 inline-block px-3 py-1 rounded-full ${t.accent} ${t.text} text-[10px] font-bold uppercase`}>Hal: {book.page || '0'}</div>
                   </div>
                 </div>
                 
                 <div className="bg-white p-5 rounded-[2rem] shadow-sm border border-gray-50 relative">
                   <div className={`absolute -top-3 left-6 w-6 h-6 ${t.primary} rounded-full flex items-center justify-center text-white text-[10px] font-bold`}>{idx + 1}</div>
                   {synopsisMap[book.id]?.loading ? (
                     <div className="flex flex-col items-center justify-center py-6 gap-2">
                       <Loader2 className={`w-6 h-6 animate-spin ${t.text}`} />
                       <p className="text-[10px] font-bold text-gray-400">Memuat isi buku...</p>
                     </div>
                   ) : (
                     <p className="text-sm leading-relaxed text-gray-600 line-clamp-[8] text-justify">
                       {synopsisMap[book.id]?.text}
                     </p>
                   )}
                 </div>
              </div>
            ))}
          </div>

          <div className="absolute bottom-10 left-0 right-0 px-10 flex justify-center">
             <button onClick={() => setIsSynopsisModalOpen(false)} className={`w-full ${t.primary} text-white font-bold py-4 rounded-2xl shadow-xl active:scale-95 transition-all`}>
               Lanjutkan Membaca
             </button>
          </div>
        </div>
      )}

      {isNetWorthModalOpen && (
        <div className="fixed inset-0 z-[110] flex flex-col bg-white animate-in slide-in-from-bottom-full duration-300 sm:max-w-md sm:mx-auto">
          <header className={`p-6 bg-gradient-to-br ${t.header} text-white shadow-lg flex justify-between items-center rounded-b-[2.5rem] sticky top-0 z-10`}>
            <div>
              <h3 className="text-xl font-bold flex items-center gap-2 italic"><Sparkles className="w-6 h-6" /> Investasi Buku</h3>
              <p className="text-white/80 text-xs">Nilai koleksi perpustakaanmu...</p>
            </div>
            <button onClick={() => setIsNetWorthModalOpen(false)} className="bg-white/20 p-2 rounded-full backdrop-blur-sm active:scale-95 transition-all"><X className="w-5 h-5" /></button>
          </header>

          <div className={`flex-1 overflow-y-auto p-6 space-y-6 ${t.bg}`}>
            {/* Total Value Card at Top */}
            <div className={`bg-white rounded-[2rem] p-6 shadow-sm border ${t.border} animate-in fade-in slide-in-from-top-4`}>
               <p className="text-center text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">Total Nilai Koleksi</p>
               <h2 className={`text-center text-2xl font-black ${t.text}`}>
                 Rp {totalNetWorth.toLocaleString('id-ID')}
               </h2>
            </div>

            <div className="space-y-4">
              {books.length === 0 ? (
                <div className="text-center py-20 text-gray-400">Belum ada buku dalam koleksi.</div>
              ) : (
                books.map((book) => (
                  <div key={book.id} className="bg-white p-4 rounded-3xl shadow-sm border border-gray-50 flex items-center gap-4 animate-in fade-in slide-in-from-bottom-2">
                    <div className="w-12 h-16 shrink-0 rounded-lg bg-gray-50 overflow-hidden border border-gray-100">
                      {book.coverUrl ? <img src={book.coverUrl} className="w-full h-full object-cover" referrerPolicy="no-referrer" /> : <div className={`w-full h-full flex items-center justify-center ${t.accent}`}><Book className={`w-8 h-8 ${t.text} opacity-20`} /></div>}
                    </div>
                    <div className="flex-1 min-w-0">
                      <h4 className="font-bold text-gray-800 text-sm truncate">{book.title}</h4>
                      <div className="flex items-center gap-2 mt-1">
                        <span className="text-[10px] font-bold text-gray-400 uppercase">Rp</span>
                        <input 
                          type="number" 
                          defaultValue={book.price || 0}
                          onBlur={(e) => {
                            const val = parseInt(e.target.value);
                            if (!isNaN(val) && val !== book.price) {
                              handleUpdatePrice(book.id, val);
                            }
                          }}
                          className={`bg-gray-50 rounded-lg px-2 py-1 text-xs font-bold ${t.text} outline-none focus:ring-2 focus:ring-gray-200 transition-all w-24`}
                          placeholder="Harga..."
                        />
                      </div>
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>
        </div>
      )}

      {isModalOpen && (
        <div className={`fixed inset-0 z-[100] flex flex-col bg-white animate-in slide-in-from-bottom-full duration-300 sm:max-w-md sm:mx-auto sm:shadow-2xl`}>
          <header className={`flex justify-between items-center p-5 bg-gradient-to-br ${t.header} shadow-sm sticky top-0 z-10 rounded-b-3xl`}>
            <h2 className="text-xl font-bold text-white">{currentBook.id ? 'Edit Buku' : 'Tambah Buku'}</h2>
            <button onClick={closeModal} className="bg-white/20 p-2 rounded-full text-white active:bg-white/30"><X className="w-5 h-5" /></button>
          </header>

          <form onSubmit={handleSaveBook} className={`flex-1 overflow-y-auto p-6 space-y-6 pb-32 ${t.bg}`}>
            <div className="flex flex-col items-center gap-2">
              <div className={`w-32 h-44 rounded-2xl bg-white shadow-md border-2 border-dashed ${t.border} overflow-hidden flex items-center justify-center relative`}>
                {currentBook.coverUrl ? <img src={currentBook.coverUrl} className="w-full h-full object-cover" referrerPolicy="no-referrer" /> : <ImageIcon className={`w-10 h-10 ${t.text} opacity-30`} />}
              </div>
              <div className="flex gap-2">
                <label className="bg-white shadow-sm border border-gray-100 rounded-full px-3 py-1.5 text-[10px] font-bold text-gray-600 cursor-pointer flex items-center gap-1 active:scale-95 transition-all hover:bg-gray-50">
                  <ImageIcon className="w-3 h-3" /> Galeri
                  <input type="file" accept="image/*" onChange={handleImageUpload} className="hidden" />
                </label>
                <label className="bg-white shadow-sm border border-gray-100 rounded-full px-3 py-1.5 text-[10px] font-bold text-gray-600 cursor-pointer flex items-center gap-1 active:scale-95 transition-all hover:bg-gray-50">
                  <Camera className="w-3 h-3" /> Kamera
                  <input type="file" accept="image/*" capture="environment" onChange={handleImageUpload} className="hidden" />
                </label>
              </div>
            </div>

            {!currentBook.id && (
              <div className="bg-white p-4 rounded-3xl shadow-sm border border-gray-50 space-y-3">
                <p className="text-xs font-bold text-gray-500 flex items-center gap-1"><ScanLine className="w-4 h-4" /> Cari Cepat Barcode/ISBN</p>
                <div className="flex gap-2">
                  <input type="text" placeholder="Ketik ISBN..." value={isbnInput} onChange={(e) => setIsbnInput(e.target.value)} className="flex-1 bg-gray-50 rounded-xl px-3 py-2 text-sm outline-none border border-gray-100" />
                  <div className="flex gap-1">
                    <button type="button" onClick={startScanner} title="Scan Kamera" className={`bg-gray-100 ${t.text} p-2 rounded-xl active:bg-gray-200 transition-colors`}>
                      <Camera className="w-5 h-5" />
                    </button>
                    <label className={`bg-gray-100 ${t.text} p-2 rounded-xl active:bg-gray-200 transition-colors cursor-pointer flex items-center justify-center`}>
                      <ImageIcon className="w-5 h-5" />
                      <input type="file" accept="image/*" className="hidden" onChange={handleBarcodeUpload} />
                    </label>
                  </div>
                  <button type="button" onClick={() => handleSmartSearch('isbn')} className={`bg-gray-100 ${t.text} px-3 rounded-xl font-bold text-sm`}>Cari</button>
                </div>
              </div>
            )}

            <div className="bg-white p-5 rounded-3xl shadow-sm border border-gray-50 space-y-4">
              <div className="space-y-1.5">
                <label className="text-xs font-bold text-gray-500 uppercase">Judul Buku</label>
                <div className="flex gap-2">
                  <input type="text" required value={currentBook.title} onChange={(e) => setCurrentBook({...currentBook, title: e.target.value})} className="flex-1 bg-gray-50 rounded-xl p-3 outline-none focus:ring-2 focus:ring-gray-200 text-sm font-bold text-gray-800" />
                  {!currentBook.id && (
                    <button type="button" onClick={() => handleSmartSearch('title')} className={`${t.accent} ${t.text} p-3 rounded-xl active:scale-95 flex items-center justify-center`}>
                      {isSearchingApi ? <Loader2 className="w-4 h-4 animate-spin" /> : <Sparkles className="w-4 h-4" />}
                    </button>
                  )}
                </div>
              </div>
              
              <div className="space-y-1.5">
                <label className="text-xs font-bold text-gray-500 uppercase">Nama Seri (Opsional)</label>
                <input type="text" placeholder="Contoh: Harry Potter" value={currentBook.seriesName} onChange={(e) => setCurrentBook({...currentBook, seriesName: e.target.value})} className="w-full bg-gray-50 rounded-xl p-3 outline-none focus:ring-2 focus:ring-gray-200 text-sm font-medium" />
                <p className="text-[10px] text-gray-400">Isi ini agar buku otomatis dikelompokkan dalam satu folder seri.</p>
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-gray-500 uppercase">Penulis</label>
                  <input type="text" value={currentBook.author} onChange={(e) => setCurrentBook({...currentBook, author: e.target.value})} className="w-full bg-gray-50 rounded-xl p-3 outline-none text-sm font-medium" />
                </div>
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-gray-500 uppercase">Kategori</label>
                  <select value={currentBook.category} onChange={(e) => setCurrentBook({...currentBook, category: e.target.value})} className="w-full bg-gray-50 rounded-xl p-3 outline-none text-sm font-medium">
                    {CATEGORIES.map(cat => <option key={cat} value={cat}>{cat}</option>)}
                  </select>
                </div>
              </div>
            </div>

            <div className="bg-white border border-gray-50 rounded-[1.5rem] p-5 shadow-sm">
              <label className="text-xs font-bold text-gray-500 uppercase mb-3 block">Status Membaca</label>
              <div className="flex gap-3">
                <button type="button" onClick={() => setCurrentBook({...currentBook, isRead: false})} className={`flex-1 py-3 rounded-xl text-sm font-bold border-2 transition-all ${!currentBook.isRead ? (t.border + ' ' + t.accent + ' ' + t.text) : 'border-transparent bg-gray-50 text-gray-400'}`}>Belum Selesai</button>
                <button type="button" onClick={() => setCurrentBook({...currentBook, isRead: true})} className={`flex-1 py-3 rounded-xl text-sm font-bold border-2 transition-all ${currentBook.isRead ? 'border-green-200 bg-green-50 text-green-600' : 'border-transparent bg-gray-50 text-gray-400'}`}>Selesai</button>
              </div>
              {!currentBook.isRead && (
                <div className="mt-3 flex items-center gap-3 bg-gray-50 p-2 rounded-xl border border-gray-100">
                  <span className="text-xs font-bold text-gray-500 pl-2">Halaman:</span>
                  <input type="number" value={currentBook.page} onChange={(e) => setCurrentBook({...currentBook, page: e.target.value})} className={`flex-1 bg-white border border-gray-100 rounded-lg p-2 outline-none font-bold ${t.text} text-sm text-center`} />
                </div>
              )}
            </div>

            <button type="submit" disabled={isSaving} className={`w-full ${t.primary} ${t.primaryHover} active:scale-[0.98] disabled:opacity-50 disabled:scale-100 text-white font-bold text-lg p-4 rounded-2xl shadow-xl transition-all flex justify-center items-center gap-2 mt-4`}>
              {isSaving ? <Loader2 className="w-6 h-6 animate-spin" /> : <Check className="w-6 h-6" />}
              {isSaving ? 'Menyimpan...' : 'Simpan Data'}
            </button>
          </form>
        </div>
      )}
    </>
  );

  // --- Auth & Data Fetching ---
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (currentUser) => {
      setUser(currentUser);
      setIsLoading(false);
    });
    return () => {
      unsubscribe();
      if (scannerRef.current) {
        stopScanner();
      }
    };
  }, []);

  const handleLogin = async () => {
    const provider = new GoogleAuthProvider();
    try {
      await signInWithPopup(auth, provider);
    } catch (error) {
      console.error("Login error:", error);
      alert("Gagal masuk dengan Google. Pastikan pop-up diizinkan.");
    }
  };

  const handleLogout = async () => {
    try {
      setIsSettingsOpen(false);
      await auth.signOut();
    } catch (error: any) {
      console.error("Logout error:", error);
      alert("Terjadi kesalahan saat keluar: " + (error.message || error));
    }
  };

  useEffect(() => {
    if (!user) return;
    
    // Test connection
    const testConnection = async () => {
      try {
        await getDocFromServer(doc(db, 'test', 'connection'));
      } catch (error) {
        if(error instanceof Error && error.message.includes('the client is offline')) {
          console.error("Please check your Firebase configuration.");
        }
      }
    };
    testConnection();

    setIsLoading(true);
    const path = `artifacts/${appId}/users/${user.uid}/books`;
    const q = query(collection(db, path));
    
    const unsubscribe = onSnapshot(q, (snapshot) => {
      const bookData: BookData[] = [];
      snapshot.forEach((doc) => {
        bookData.push({ id: doc.id, ...doc.data() } as BookData);
      });
      bookData.sort((a, b) => b.updatedAt - a.updatedAt);
      setBooks(bookData);
      setIsLoading(false);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, path);
    });

    return () => unsubscribe();
  }, [user]);

  function getEmptyBook(): BookData {
    return { 
      id: '', 
      title: '', 
      author: '', 
      category: 'Novel', 
      seriesName: '', 
      isRead: false, 
      page: '', 
      coverUrl: '', 
      price: 0,
      synopsis: '',
      addedAt: Date.now(),
      updatedAt: Date.now()
    };
  }

  // --- Handlers DB ---
  const handleSaveBook = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!user || !currentBook.title) return;

    setIsSaving(true);
    const isEditing = !!currentBook.id;
    const bookId = isEditing ? currentBook.id : Date.now().toString();
    const path = `artifacts/${appId}/users/${user.uid}/books`;
    
    const bookData: any = {
      title: currentBook.title,
      author: currentBook.author || 'Anonim',
      category: currentBook.category,
      seriesName: currentBook.seriesName || '',
      isRead: currentBook.isRead,
      page: currentBook.isRead ? '' : (currentBook.page || '0'),
      coverUrl: currentBook.coverUrl || '',
      price: currentBook.price || 0,
      synopsis: currentBook.synopsis || '',
      updatedAt: Date.now(),
      addedAt: currentBook.addedAt || Date.now()
    };
    
    try {
      await setDoc(doc(db, path, bookId), bookData, { merge: true });
      closeModal();
    } catch (error) {
      alert("Gagal menyimpan data. Silakan coba lagi.");
      handleFirestoreError(error, OperationType.WRITE, `${path}/${bookId}`);
    } finally {
      setIsSaving(false);
    }
  };

  const handleDeleteBook = async (bookId: string) => {
    if (!user || !bookId) return;
    const path = `artifacts/${appId}/users/${user.uid}/books`;
    try {
      await deleteDoc(doc(db, path, bookId));
      setActionMenuBook(null);
    } catch (error) {
      handleFirestoreError(error, OperationType.DELETE, `${path}/${bookId}`);
    }
  };

  const closeModal = () => {
    setIsModalOpen(false);
    setCurrentBook(getEmptyBook());
    setIsbnInput('');
  };

  // --- API Pencarian Cerdas ---
  const handleSmartSearch = async (type: 'title' | 'isbn' = 'title', customIsbn?: string) => {
    const queryStr = type === 'isbn' ? `isbn:${customIsbn || isbnInput}` : `intitle:${currentBook.title}`;
    if (!queryStr) {
      alert('Masukkan kata kunci pencarian terlebih dahulu.');
      return;
    }
    
    setIsSearchingApi(true);
    try {
      const apiKey = import.meta.env.VITE_GOOGLE_BOOKS_API_KEY;
      const url = `https://www.googleapis.com/books/v1/volumes?q=${encodeURIComponent(queryStr)}&maxResults=1${apiKey ? `&key=${apiKey}` : ''}`;
      
      const response = await fetch(url);
      
      if (response.status === 429) {
        throw new Error("Terlalu banyak permintaan (Limit tercapai). Harap masukkan API Key di pengaturan atau tunggu beberapa saat.");
      }
      
      if (!response.ok) {
        throw new Error(`Gagal menghubungi layanan Google (Status: ${response.status})`);
      }
      
      const data = await response.json();
      
      if (data.items && data.items.length > 0) {
        const item = data.items[0];
        const info = item.volumeInfo;
        const sale = item.saleInfo;

        let fetchedPrice = 0;
        if (sale?.listPrice) {
          const amount = sale.listPrice.amount;
          const currency = sale.listPrice.currencyCode;
          if (currency === 'IDR') {
            fetchedPrice = amount;
          } else if (currency === 'USD') {
            fetchedPrice = amount * 17000;
          } else {
            fetchedPrice = amount * 17000;
          }
        }

        setCurrentBook(prev => ({
          ...prev,
          title: info.title || prev.title,
          author: info.authors ? info.authors.join(', ') : (prev.author === 'Anonim' ? '' : prev.author),
          coverUrl: info.imageLinks?.thumbnail?.replace('http:', 'https:') || prev.coverUrl,
          price: fetchedPrice || prev.price,
          synopsis: info.description || prev.synopsis,
        }));
        if (type === 'isbn') setIsbnInput('');
      } else { 
        alert('Data buku tidak ditemukan di Google Books.'); 
      }
    } catch (error: any) { 
      console.error(error); 
      alert('Terjadi kesalahan pencarian: ' + (error.message || 'Masalah koneksi internet'));
    } finally { 
      setIsSearchingApi(false); 
    }
  };

  const handleImageUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement('canvas');
        const MAX_WIDTH = 300;
        canvas.width = MAX_WIDTH;
        canvas.height = img.height * (MAX_WIDTH / img.width);
        const ctx = canvas.getContext('2d');
        if (ctx) {
          ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
          setCurrentBook({...currentBook, coverUrl: canvas.toDataURL('image/jpeg', 0.8)});
        }
      };
      if (event.target?.result) {
        img.src = event.target.result as string;
      }
    };
    reader.readAsDataURL(file);
  };

  // --- Long Press ---
  const handleTouchStart = (book: BookData) => {
    pressTimer.current = setTimeout(() => {
      setActionMenuBook(book);
      if (navigator.vibrate) navigator.vibrate(50);
    }, 500);
  };
  const handleTouchEnd = () => { 
    if (pressTimer.current) clearTimeout(pressTimer.current); 
  };

  // --- Series Grouping & Filtering ---
  const toggleSeriesExpand = (seriesName: string) => {
    setExpandedSeries(prev => ({ ...prev, [seriesName]: !prev[seriesName] }));
  };

  const { standaloneBooks, groupedSeries } = useMemo(() => {
    let filtered = searchTerm 
      ? books.filter(b => b.title.toLowerCase().includes(searchTerm.toLowerCase()) || b.author.toLowerCase().includes(searchTerm.toLowerCase()))
      : books;
      
    // Sorting logic
    filtered = [...filtered].sort((a, b) => {
      if (sortBy === 'title') return a.title.localeCompare(b.title);
      if (sortBy === 'author') return a.author.localeCompare(b.author);
      if (sortBy === 'addedAt') return b.addedAt - a.addedAt;
      if (sortBy === 'isRead') {
        if (a.isRead === b.isRead) return b.addedAt - a.addedAt;
        return a.isRead ? -1 : 1;
      }
      return 0;
    });

    const standalone: BookData[] = [];
    const groups: Record<string, BookData[]> = {};

    filtered.forEach(book => {
      if (book.seriesName && book.seriesName.trim() !== '') {
        const sName = book.seriesName.trim();
        if (!groups[sName]) groups[sName] = [];
        groups[sName].push(book);
      } else {
        standalone.push(book);
      }
    });

    return { standaloneBooks: standalone, groupedSeries: groups };
  }, [books, searchTerm, sortBy]);

  const stats = useMemo(() => {
    return {
      total: books.length,
      read: books.filter(b => b.isRead).length,
      categories: new Set(books.map(b => b.category)).size
    };
  }, [books]);

  // --- Render UI Components ---
  const renderBookCard = (book: BookData) => {
    if (viewMode === 'grid') {
      return (
        <div key={book.id} 
          onMouseDown={() => handleTouchStart(book)} 
          onMouseUp={handleTouchEnd} 
          onMouseLeave={handleTouchEnd} 
          onTouchStart={() => handleTouchStart(book)} 
          onTouchEnd={handleTouchEnd}
          className="relative bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden active:scale-95 transition-all animate-in fade-in zoom-in group aspect-[2/3] flex flex-col">
          <div className={`absolute top-0 left-0 right-0 h-1 z-10 ${book.isRead ? 'bg-green-400' : t.primary}`}></div>
          {book.coverUrl ? (
             <img src={book.coverUrl} alt={book.title} className="w-full h-full object-cover" referrerPolicy="no-referrer" />
          ) : (
             <div className={`w-full h-full flex flex-col items-center justify-center ${t.bg} p-2 text-center`}>
               <Book className={`w-8 h-8 mb-2 ${t.text} opacity-50`} />
               <p className="text-[10px] font-bold text-gray-600 line-clamp-3 leading-tight">{book.title}</p>
             </div>
          )}
          <div className="absolute bottom-2 right-2 flex gap-1">
            {book.isRead && <div className="bg-green-500 text-white p-1 rounded-full shadow-md"><Check className="w-3 h-3" /></div>}
          </div>
        </div>
      );
    }

    return (
      <div key={book.id} 
        onMouseDown={() => handleTouchStart(book)} 
        onMouseUp={handleTouchEnd} 
        onMouseLeave={handleTouchEnd} 
        onTouchStart={() => handleTouchStart(book)} 
        onTouchEnd={handleTouchEnd}
        className="bg-white p-3.5 rounded-[1.5rem] shadow-sm border border-gray-50 active:scale-[0.98] transition-all flex gap-4 relative overflow-hidden animate-in fade-in slide-in-from-bottom-2">
        <div className={`absolute left-0 top-0 bottom-0 w-1.5 ${book.isRead ? 'bg-green-300' : t.primary}`}></div>
        <div className="w-16 h-24 shrink-0 rounded-xl bg-gray-100 overflow-hidden shadow-sm border border-gray-100 relative">
          {book.coverUrl ? <img src={book.coverUrl} alt="Cover" className="w-full h-full object-cover" referrerPolicy="no-referrer" /> : <div className={`w-full h-full flex items-center justify-center ${t.bg}`}><Book className={`w-6 h-6 ${t.text} opacity-50`} /></div>}
        </div>
        <div className="flex-1 flex flex-col justify-center min-w-0">
          <h3 className="font-bold text-gray-800 leading-tight mb-1 truncate pr-2">{book.title}</h3>
          <div className="flex items-center text-xs text-gray-500 gap-1.5 mb-2"><User className="w-3.5 h-3.5 shrink-0" /><span className="truncate">{book.author}</span></div>
          <div className="mt-auto flex items-center justify-between">
             <span className={`bg-gray-100 text-gray-600 px-2 py-0.5 rounded-md text-[10px] font-bold`}>{book.category}</span>
             {book.isRead ? <div className="text-green-600 font-bold text-xs flex items-center gap-1"><Check className="w-3 h-3"/> Selesai</div> : <div className={`${t.text} font-bold text-xs`}>Hal: {book.page || '?'}</div>}
          </div>
        </div>
      </div>
    );
  };

  if (isLoading) return <div className={`min-h-screen ${t.bg} flex items-center justify-center`}><Loader2 className={`w-10 h-10 ${t.text} animate-spin`} /></div>;

  if (!user) {
    return (
      <div className={`min-h-screen ${t.bg} flex flex-col items-center justify-center p-6 text-center sm:max-w-md sm:mx-auto sm:shadow-2xl bg-white`}>
        <div className={`w-24 h-24 ${t.accent} rounded-full flex items-center justify-center mb-6 animate-bounce`}>
          <Book className={`w-12 h-12 ${t.text}`} />
        </div>
        <h1 className="text-3xl font-bold text-gray-800 mb-2">BukuKu</h1>
        <p className="text-gray-500 mb-8">Simpan dan kelola koleksi buku pribadimu dengan mudah.</p>
        <button 
          onClick={handleLogin}
          className={`${t.primary} ${t.primaryHover} text-white px-8 py-4 rounded-2xl font-bold shadow-xl active:scale-95 transition-all flex items-center gap-3`}
        >
          <User className="w-5 h-5" /> Masuk dengan Google
        </button>
      </div>
    );
  }

  return (
    <div className={`min-h-screen ${t.bg} text-gray-800 font-sans sm:max-w-md sm:mx-auto sm:shadow-2xl relative flex flex-col transition-colors duration-500`}>
      
      {/* Header & Search */}
      <header className={`bg-gradient-to-br ${t.header} p-6 rounded-b-[2.5rem] shadow-md z-10 sticky top-0 transition-all duration-500`}>
        <div className="flex justify-between items-center mb-5">
          <div>
            <h1 className="text-2xl font-bold text-white tracking-wide">BukuKu</h1>
            <p className="text-white/90 text-sm mt-0.5">Perpustakaan pribadimu 🌸</p>
          </div>
          <button onClick={() => setIsNetWorthModalOpen(true)} className="bg-white/20 p-3 rounded-full backdrop-blur-sm shadow-inner active:scale-95 transition-all">
            <Library className="text-white w-6 h-6" />
          </button>
        </div>
        
        <div className="flex gap-2">
          <div className="relative group flex-1">
            <div className="absolute inset-y-0 left-4 flex items-center pointer-events-none"><Search className={`h-5 w-5 ${t.text} opacity-50`} /></div>
            <input type="text" placeholder="Cari buku..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full bg-white/95 backdrop-blur-sm text-gray-800 rounded-full py-3.5 pl-12 pr-4 outline-none shadow-sm focus:ring-4 focus:ring-white/30 transition-all font-medium text-sm" />
          </div>
          
          <div className="relative">
            <select 
              value={sortBy} 
              onChange={(e) => setSortBy(e.target.value as any)}
              className="appearance-none bg-white/95 backdrop-blur-sm p-3.5 pl-10 pr-4 rounded-full shadow-sm active:scale-95 transition-all text-gray-600 outline-none font-bold text-xs cursor-pointer"
            >
              <option value="addedAt">Terbaru</option>
              <option value="title">Judul</option>
              <option value="author">Penulis</option>
              <option value="isRead">Status</option>
            </select>
            <SortAsc className="absolute left-3.5 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400 pointer-events-none" />
          </div>

          <button onClick={() => setViewMode(v => v === 'list' ? 'grid' : 'list')} className="bg-white/95 backdrop-blur-sm p-3.5 rounded-full shadow-sm active:scale-95 transition-all text-gray-600 flex items-center justify-center shrink-0">
            {viewMode === 'list' ? <LayoutGrid className="w-5 h-5" /> : <ListIcon className="w-5 h-5" />}
          </button>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="flex-1 px-5 pt-6 pb-28 overflow-y-auto">
        {books.length > 0 && (
          <div className="flex gap-4 mb-5 px-1 animate-in fade-in duration-700">
            <div className="flex flex-col">
              <span className="text-[9px] font-bold text-gray-400 uppercase tracking-[0.2em] mb-0.5">Total Buku</span>
              <span className={`text-xs font-black ${t.text}`}>{stats.total}</span>
            </div>
            <div className="w-px h-6 bg-gray-100 self-end mb-1"></div>
            <div className="flex flex-col">
              <span className="text-[9px] font-bold text-gray-400 uppercase tracking-[0.2em] mb-0.5">Sudah Baca</span>
              <span className="text-xs font-black text-green-500">{stats.read}</span>
            </div>
            <div className="w-px h-6 bg-gray-100 self-end mb-1"></div>
            <div className="flex flex-col">
              <span className="text-[9px] font-bold text-gray-400 uppercase tracking-[0.2em] mb-0.5">Belum Baca</span>
              <span className={`text-xs font-black ${t.text}`}>{stats.total - stats.read}</span>
            </div>
          </div>
        )}
        {books.length === 0 ? (
          <div className="text-center mt-16 text-gray-400 flex flex-col items-center animate-in fade-in duration-500">
            <div className={`w-24 h-24 ${t.accent} rounded-full flex items-center justify-center mb-4`}><Book className={`w-10 h-10 ${t.text}`} /></div>
            <p className="font-medium text-gray-500">Koleksi masih kosong.</p>
          </div>
        ) : (
          <div className="space-y-6">
            {Object.keys(groupedSeries).map(seriesName => (
              <div key={seriesName} className="bg-white rounded-3xl shadow-sm border border-gray-100 overflow-hidden">
                <button onClick={() => toggleSeriesExpand(seriesName)} className={`w-full p-4 flex items-center justify-between bg-gradient-to-r from-gray-50 to-white active:bg-gray-100 transition-colors`}>
                  <div className="flex items-center gap-3">
                    <div className={`p-2 rounded-xl ${t.accent} ${t.text}`}><Layers className="w-5 h-5" /></div>
                    <div className="text-left">
                      <h3 className="font-bold text-gray-800 text-sm">{seriesName}</h3>
                      <p className="text-[10px] font-bold text-gray-400 uppercase tracking-wide">{groupedSeries[seriesName].length} Buku</p>
                    </div>
                  </div>
                  {expandedSeries[seriesName] ? <ChevronUp className="w-5 h-5 text-gray-400" /> : <ChevronDown className="w-5 h-5 text-gray-400" />}
                </button>
                
                {expandedSeries[seriesName] && (
                  <div className={`p-4 pt-2 border-t border-gray-50 bg-gray-50/30 ${viewMode === 'grid' ? 'grid grid-cols-3 gap-3' : 'space-y-3'}`}>
                    {groupedSeries[seriesName].map(book => renderBookCard(book))}
                  </div>
                )}
              </div>
            ))}

            {standaloneBooks.length > 0 && (
              <div className={viewMode === 'grid' ? 'grid grid-cols-3 gap-3' : 'space-y-3'}>
                {standaloneBooks.map(book => renderBookCard(book))}
              </div>
            )}
          </div>
        )}
      </main>

      {/* Bottom Navigation Bar */}
      <div className="fixed sm:absolute bottom-0 left-0 right-0 bg-white border-t border-gray-100 shadow-[0_-10px_40px_-15px_rgba(0,0,0,0.1)] z-20 px-6 py-4 flex justify-between items-center rounded-t-3xl">
        <button onClick={openSynopsisModal} className="flex flex-col items-center gap-1 active:scale-95 transition-all outline-none">
          <Library className={`w-6 h-6 ${t.text}`} />
        </button>
        
        <button onClick={() => { setCurrentBook(getEmptyBook()); setIsModalOpen(true); }}
          className={`${t.primary} ${t.primaryHover} text-white p-4 rounded-full shadow-lg -translate-y-8 active:scale-95 transition-all border-4 border-white`}>
          <Plus className="w-8 h-8" />
        </button>

        <button onClick={() => setIsSettingsOpen(true)} className="flex flex-col items-center gap-1 active:scale-95 transition-all">
          <Settings className={`w-6 h-6 ${t.text}`} />
        </button>
      </div>

      {renderSharedModals()}
    </div>
  );
}

export default function App() {
  return <BookApp />;
}
