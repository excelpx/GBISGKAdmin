const DB_NAME = 'GerejaDataHubDB_v1';
const STORE_NAME = 'files';

let activeTab = 'dashboard';
let dbRef = null;
let currentUser = null;
let driveFiles = [];
let activeFileObj = null;
let currentSheetName = '';
let searchQuery = '';
let selectedMonth = '';
let activeDocPreviewUrl = null;
let photoDragMode = false;
let currentDraggedPhotoId = null;

let stats = {
    totalRows: 0,
    filesCount: 0,
    photosCount: 0
};

function initAppDatabase() {
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(DB_NAME, 1);
        request.onupgradeneeded = () => {
            const db = request.result;
            if (!db.objectStoreNames.contains(STORE_NAME)) {
                db.createObjectStore(STORE_NAME, { keyPath: 'id' });
            }
        };
        request.onsuccess = () => {
            dbRef = request.result;
            resolve(dbRef);
        };
        request.onerror = () => reject(request.error);
    });
}

function generateGradientPhotoBlob(title, color1, color2) {
    return new Promise((resolve) => {
        const canvas = document.createElement('canvas');
        canvas.width = 800;
        canvas.height = 600;
        const ctx = canvas.getContext('2d');
        if (ctx) {
            const gradient = ctx.createLinearGradient(0, 0, 800, 600);
            gradient.addColorStop(0, color1);
            gradient.addColorStop(1, color2);
            ctx.fillStyle = gradient;
            ctx.fillRect(0, 0, 800, 600);

            ctx.fillStyle = 'rgba(255, 255, 255, 0.05)';
            ctx.beginPath();
            ctx.arc(400, 300, 200, 0, Math.PI * 2);
            ctx.fill();

            ctx.strokeStyle = 'rgba(255, 255, 255, 0.2)';
            ctx.lineWidth = 4;
            ctx.beginPath();
            ctx.moveTo(400, 150);
            ctx.lineTo(400, 450);
            ctx.moveTo(300, 250);
            ctx.lineTo(500, 250);
            ctx.stroke();

            ctx.fillStyle = '#ffffff';
            ctx.font = 'bold 36px sans-serif';
            ctx.textAlign = 'center';
            ctx.textBaseline = 'middle';
            ctx.fillText(title, 400, 300);

            ctx.fillStyle = 'rgba(255, 255, 255, 0.7)';
            ctx.font = '16px monospace';
            ctx.fillText('DATABASE DIGITAL GEREJA INTERNAL', 400, 520);
        }
        canvas.toBlob((blob) => {
            resolve(blob || new Blob());
        }, 'image/jpeg', 0.85);
    });
}

// Parse SheetJS arrays helpers
function generateMockExcelBlob() {
    const dataSheet1 = [
        { 'No Anggota': 'J-001', 'Nama Lengkap': 'Yohanes Prasetyo', Gender: 'Laki-laki', 'Status Pernikahan': 'Menikah', 'Sektor Jemaat': 'Sektor Kedamaian', 'Nomor Telepon': '0812-3456-7890', 'Tanggal Lahir': '1985-04-12' },
        { 'No Anggota': 'J-002', 'Nama Lengkap': 'Maria Elizabeth', Gender: 'Perempuan', 'Status Pernikahan': 'Menikah', 'Sektor Jemaat': 'Sektor Kedamaian', 'Nomor Telepon': '0813-9876-5432', 'Tanggal Lahir': '1988-09-21' },
        { 'No Anggota': 'J-003', 'Nama Lengkap': 'Andreas Wijaya', Gender: 'Laki-laki', 'Status Pernikahan': 'Lajang', 'Sektor Jemaat': 'Sektor Kasih', 'Nomor Telepon': '0856-4321-8765', 'Tanggal Lahir': '1995-11-03' },
        { 'No Anggota': 'J-004', 'Nama Lengkap': 'Sarah Angelica', Gender: 'Perempuan', 'Status Pernikahan': 'Lajang', 'Sektor Jemaat': 'Sektor Kasih', 'Nomor Telepon': '0878-1122-3344', 'Tanggal Lahir': '2000-07-18' },
        { 'No Anggota': 'J-005', 'Nama Lengkap': 'Budi Hartono', Gender: 'Laki-laki', 'Status Pernikahan': 'Menikah', 'Sektor Jemaat': 'Sektor Pengharapan', 'Nomor Telepon': '0811-2233-4455', 'Tanggal Lahir': '1972-12-30' },
        { 'No Anggota': 'J-006', 'Nama Lengkap': 'Christina Natalia', Gender: 'Perempuan', 'Status Pernikahan': 'Lajang', 'Sektor Jemaat': 'Sektor Pengharapan', 'Nomor Telepon': '0852-5555-4433', 'Tanggal Lahir': '1998-12-25' }
    ];

    const dataSheet2 = [
        { Tanggal: '2026-05-01', 'Uraian Kas': 'Saldo Awal Mei', Kategori: 'Saldo', Penerimaan: 15450000, Pengeluaran: 0 },
        { Tanggal: '2026-05-03', 'Uraian Kas': 'Kolekte Kebaktian Minggu', Kategori: 'Kolekte', Penerimaan: 3240000, Pengeluaran: 0 },
        { Tanggal: '2026-05-07', 'Uraian Kas': 'Pembayaran Listrik & Wifi', Kategori: 'Operasional', Penerimaan: 0, Pengeluaran: 850000 },
        { Tanggal: '2026-05-10', 'Uraian Kas': 'Dana Diakonia Sakit', Kategori: 'Diakonia', Penerimaan: 0, Pengeluaran: 1000000 },
        { Tanggal: '2026-05-14', 'Uraian Kas': 'Sumbangan Persepuluhan', Kategori: 'Sumbangan', Penerimaan: 5000000, Pengeluaran: 0 }
    ];

    const workbook = XLSX.utils.book_new();
    const ws1 = XLSX.utils.json_to_sheet(dataSheet1);
    XLSX.utils.book_append_sheet(workbook, ws1, 'Daftar Jemaat');
    const ws2 = XLSX.utils.json_to_sheet(dataSheet2);
    XLSX.utils.book_append_sheet(workbook, ws2, 'Laporan Kas Keuangan');

    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    return new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

async function seedMockDatabase() {
    const db = dbRef;
    return new Promise((resolve, reject) => {
        const transaction = db.transaction(STORE_NAME, 'readwrite');
        const store = transaction.objectStore(STORE_NAME);
        const countRequest = store.count();

        countRequest.onsuccess = async () => {
            if (countRequest.result === 0) {
                console.log("Database kosong, menanamkan bank data awal...");

                const excelBlob = generateMockExcelBlob();
                store.put({
                    id: 'mock-excel-1',
                    name: 'Database_Jemaat_Utama.xlsx',
                    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    size: String(excelBlob.size),
                    createdTime: new Date().toISOString(),
                    parentFolderId: 'excels',
                    blob: excelBlob
                });

                const photoSeeds = [
                    { name: 'Struktur_Majelis_2026.jpg', title: 'Struktur Organisasi Majelis', c1: '#1e3a8a', c2: '#312e81' },
                    { name: 'Kegiatan_Sosial_Jemaat.jpg', title: 'Aksi Sosial & Kasih Jemaat', c1: '#064e3b', c2: '#115e59' },
                    { name: 'Perayaan_Paskah_Gereja.jpg', title: 'Dokumentasi Perayaan Paskah', c1: '#581c87', c2: '#6b21a8' }
                ];

                for (let i = 0; i < photoSeeds.length; i++) {
                    const seed = photoSeeds[i];
                    const photoBlob = await generateGradientPhotoBlob(seed.title, seed.c1, seed.c2);
                    store.put({
                        id: `mock-photo-${i + 1}`,
                        name: seed.name,
                        mimeType: 'image/jpeg',
                        size: String(photoBlob.size),
                        createdTime: new Date(Date.now() - i * 86400000).toISOString(),
                        parentFolderId: 'photos',
                        displayOrder: 1000 - i,
                        blob: photoBlob
                    });
                }
            }
            resolve();
        };
        countRequest.onerror = () => reject(countRequest.error);
    });
}

function toast(text, type = 'success') {
    const container = document.getElementById('feedback-alert');
    const iconEl = document.getElementById('feedback-icon');
    const textEl = document.getElementById('feedback-text');

    textEl.innerText = text;
    if (type === 'success') {
        container.className = "fixed bottom-6 right-6 z-50 max-w-sm w-full bg-slate-900 text-white border border-slate-800 rounded-2xl p-4 shadow-2xl flex items-start gap-3 transition-all duration-300 translate-y-0 opacity-100";
        iconEl.setAttribute('data-lucide', 'check-circle');
        iconEl.className = "w-5 h-5 text-emerald-400 flex-shrink-0";
    } else {
        container.className = "fixed bottom-6 right-6 z-50 max-w-sm w-full bg-red-950 text-white border border-red-900 rounded-2xl p-4 shadow-2xl flex items-start gap-3 transition-all duration-300 translate-y-0 opacity-100";
        iconEl.setAttribute('data-lucide', 'alert-circle');
        iconEl.className = "w-5 h-5 text-red-400 flex-shrink-0";
    }
    lucide.createIcons();
    container.classList.remove('hidden');

    setTimeout(() => {
        container.classList.add('opacity-0');
        container.classList.add('translate-y-2');
        setTimeout(() => {
            container.classList.add('hidden');
        }, 300);
    }, 3500);
}

window.addEventListener('DOMContentLoaded', async () => {
    // Setup Date Banner
    const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    const today = new Date();
    document.getElementById('banner-date').innerText = today.toLocaleDateString('id-ID', options);
    document.getElementById('footer-year').innerText = today.getFullYear();

    // Setup DB
    await initAppDatabase();
    await seedMockDatabase();

    // Run Icons
    lucide.createIcons();

    // Session check
    checkSession();
    
    // Handle form login
    document.getElementById('login-form').addEventListener('submit', handleLoginAttempt);

    const photoDropzone = document.getElementById('photo-dropzone');
    if (photoDropzone) {
        photoDropzone.addEventListener('dragover', handlePhotoDropzoneDragOver);
        photoDropzone.addEventListener('dragenter', handlePhotoDropzoneDragEnter);
        photoDropzone.addEventListener('dragleave', handlePhotoDropzoneDragLeave);
        photoDropzone.addEventListener('drop', handlePhotoDropzoneDrop);
    }
    updatePhotoDragModeButton();
});

function checkSession() {
    const userStr = localStorage.getItem('gereja_user');
    const tokenStr = localStorage.getItem('gereja_token');
    if (userStr && tokenStr) {
        currentUser = JSON.parse(userStr);
        document.getElementById('auth-gate').classList.add('hidden');
        document.getElementById('app-workspace').classList.remove('hidden');
        document.getElementById('user-profile-name').innerText = currentUser.displayName;
        document.getElementById('user-profile-email').innerText = currentUser.email;
        document.getElementById('banner-email').innerText = currentUser.email;
        document.getElementById('user-avatar-initial').innerText = currentUser.displayName.slice(0,1);
        loadDashboardStats();
        fetchExcelFilesList();
        fetchPhotosList();
        fetchDocsList();
    } else {
        document.getElementById('auth-gate').classList.remove('hidden');
        document.getElementById('app-workspace').classList.add('hidden');
    }
}

async function handleLoginAttempt(event) {
    event.preventDefault();
    const usernameInput = document.getElementById('login-username').value;
    const passwordInput = document.getElementById('login-password').value;
    const errorContainer = document.getElementById('login-error-container');
    const errorText = document.getElementById('login-error-text');
    const submitBtn = document.getElementById('submit-login-btn');

    errorContainer.classList.add('hidden');
    
    if (usernameInput.toLowerCase() === 'admin' && (passwordInput === 'gereja123' || passwordInput === 'admin')) {
        submitBtn.disabled = true;
        submitBtn.innerHTML = `<i data-lucide="refresh-cw" class="w-4 h-4 animate-spin text-white"></i><span>Memverifikasi sandi...</span>`;
        lucide.createIcons();

        setTimeout(() => {
            currentUser = {
                displayName: "Admin Jemaat",
                email: "admin@gereja-datahub.id"
            };
            localStorage.setItem('gereja_user', JSON.stringify(currentUser));
            localStorage.setItem('gereja_token', 'dummy-token-vscode');

            submitBtn.disabled = false;
            submitBtn.innerText = "Masuk Sekarang";

            checkSession();
            toast('Syalom! Anda berhasil masuk ke sistem jemaat.', 'success');
        }, 900);
    } else {
        errorText.innerText = 'Username atau password salah. Coba lagi.';
        errorContainer.classList.remove('hidden');
    }
}

function triggerLogout() {
    const confirmed = window.confirm("Apakah Anda yakin ingin keluar dari sistem internal gereja?");
    if (confirmed) {
        localStorage.removeItem('gereja_user');
        localStorage.removeItem('gereja_token');
        currentUser = null;
        activeFileObj = null;
        document.getElementById('login-username').value = '';
        document.getElementById('login-password').value = '';
        checkSession();
        toast('Berhasil keluar dari sesi.', 'success');
    }
}

function switchTab(tabId) {
    activeTab = tabId;
    const screens = ['dashboard', 'excel', 'photo'];
    screens.forEach((sc) => {
        const layer = document.getElementById(`view-${sc}`);
        if (sc === tabId) {
            layer.classList.remove('hidden');
        } else {
            layer.classList.add('hidden');
        }
    });
    const btns = document.querySelectorAll('.tab-btn');
    btns.forEach((btn) => {
        const isThis = btn.id === `tab-btn-${tabId}`;
        const ico = btn.querySelector('i');
        if (isThis) {
            btn.className = "tab-btn px-4 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-2 transition-all cursor-pointer bg-white text-blue-600 shadow-xs";
            if (ico) ico.className = "w-3.5 h-3.5 text-blue-600";
        } else {
            btn.className = "tab-btn px-4 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-2 transition-all cursor-pointer text-slate-600 hover:bg-slate-50";
            if (ico) ico.className = "w-3.5 h-3.5 text-slate-400";
        }
    });
    const mobBtns = document.querySelectorAll('.mobile-tab-btn');
    mobBtns.forEach((btn) => {
        const isThis = btn.id === `mobile-tab-btn-${tabId}`;
        const ico = btn.querySelector('i');
        if (isThis) {
            btn.className = "mobile-tab-btn text-left w-full px-4 py-3 rounded-xl text-xs font-semibold flex items-center gap-3 transition-all bg-blue-600 text-white shadow-xs";
            if (ico) ico.className = "w-4 h-4 text-white";
        } else {
            btn.className = "mobile-tab-btn text-left w-full px-4 py-3 rounded-xl text-xs font-semibold flex items-center gap-3 transition-all text-slate-600 hover:bg-slate-50";
            if (ico) ico.className = "w-4 h-4 text-slate-400";
        }
    });
    loadDashboardStats();
    lucide.createIcons();
}

let isMobileMenuOpen = false;
function toggleMobileMenu(forceState) {
    isMobileMenuOpen = (forceState !== undefined) ? forceState : !isMobileMenuOpen;
    const drawer = document.getElementById('mobile-menu-drawer');
    const icon = document.getElementById('mobile-menu-icon');
    const menuButton = document.querySelector('[aria-controls="mobile-menu-drawer"]');
    if (isMobileMenuOpen) {
        drawer.classList.remove('hidden');
        icon.setAttribute('data-lucide', 'x');
        try { document.body.style.overflow = 'hidden'; } catch(e){}
        if (menuButton) menuButton.setAttribute('aria-expanded', 'true');
    } else {
        drawer.classList.add('hidden');
        icon.setAttribute('data-lucide', 'menu');
        try { document.body.style.overflow = ''; } catch(e){}
        if (menuButton) menuButton.setAttribute('aria-expanded', 'false');
    }
    lucide.createIcons();
}

async function loadDashboardStats() {
    if (!dbRef) return;
    const db = dbRef;
    const transaction = db.transaction(STORE_NAME, 'readonly');
    const store = transaction.objectStore(STORE_NAME);
    const cursorRequest = store.openCursor();
    let filesCount = 0;
    let photosCount = 0;
    let rowsNum = 0;
    cursorRequest.onsuccess = async () => {
        const cursor = cursorRequest.result;
        if (cursor) {
            const val = cursor.value;
            if (val.parentFolderId === 'excels') {
                filesCount++;
                if (activeFileObj && val.id === activeFileObj.id && activeFileObj.parsedData) {
                    const activeSheet = activeFileObj.parsedData.sheets[currentSheetName];
                    rowsNum = activeSheet ? activeSheet.rows.length : 0;
                }
            } else if (val.parentFolderId === 'photos') {
                photosCount++;
            }
            cursor.continue();
        } else {
            stats.filesCount = filesCount;
            stats.photosCount = photosCount;
            document.getElementById('metric-files-count').innerText = filesCount;
            document.getElementById('metric-photos-count').innerText = photosCount;
            if (activeFileObj && activeFileObj.parsedData) {
                const activeSheet = activeFileObj.parsedData.sheets[currentSheetName];
                stats.totalRows = activeSheet ? activeSheet.rows.length : 0;
            } else {
                stats.totalRows = 0;
            }
            document.getElementById('metric-total-rows').innerText = stats.totalRows;
        }
    };
}

async function fetchExcelFilesList() {
    if (!dbRef) return;
    document.getElementById('excel-list-loading').classList.remove('hidden');
    const container = document.getElementById('excel-files-list');
    container.innerHTML = '';
    const transaction = dbRef.transaction(STORE_NAME, 'readonly');
    const store = transaction.objectStore(STORE_NAME);
    const cursorReq = store.openCursor();
    const items = [];
    cursorReq.onsuccess = () => {
        const cursor = cursorReq.result;
        if (cursor) {
            if (cursor.value.parentFolderId === 'excels') {
                items.push(cursor.value);
            }
            cursor.continue();
        } else {
            document.getElementById('excel-list-loading').classList.add('hidden');
            if (items.length === 0) {
                container.innerHTML = `<p class="text-[10px] text-slate-400 font-medium text-center py-4">Belum ada file Excel.</p>`;
                return;
            }
            items.sort((a,b) => new Date(b.createdTime).getTime() - new Date(a.createdTime).getTime());
            items.forEach((item) => {
                const isSelected = activeFileObj && activeFileObj.id === item.id;
                const card = document.createElement('button');
                card.onclick = () => selectExcelFileFromList(item.id);
                card.className = `w-full text-left p-2.5 rounded-xl text-xs font-medium flex items-center justify-between gap-2 border transition cursor-pointer ${
                    isSelected 
                        ? 'bg-blue-50 border-blue-200 text-blue-950' 
                        : 'bg-white hover:bg-slate-50 border-slate-200/70 text-slate-700'
                }`;
                card.innerHTML = `
                    <span class="flex items-center gap-2 truncate">
                        <i data-lucide="file-spreadsheet" class="w-3.5 h-3.5 flex-shrink-0 ${isSelected ? 'text-blue-600' : 'text-slate-400'}"></i>
                        <span class="truncate font-semibold">${item.name}</span>
                    </span>
                    <i data-lucide="chevron-right" class="w-3 h-3 text-slate-400 flex-shrink-0"></i>
                `;
                container.appendChild(card);
            });
            lucide.createIcons();
        }
    };
}

function triggerExcelInput() {
    document.getElementById('excel-file-hidden-input').click();
}

async function handleLocalExcelUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (e) => {
        try {
            const data = e.target.result;
            const workbook = XLSX.read(data, { type: 'array' });
            const parsed = {
                sheetNames: workbook.SheetNames,
                sheets: {}
            };
            workbook.SheetNames.forEach((sheetName) => {
                const sheet = workbook.Sheets[sheetName];
                const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
                let headers = [];
                if (rows.length > 0) {
                    headers = Object.keys(rows[0]);
                }
                parsed.sheets[sheetName] = { headers, rows };
            });
            const id = 'file-' + Math.random().toString(36).substring(2, 9);
            const fileObj = {
                id,
                name: file.name,
                mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                size: String(file.size),
                createdTime: new Date().toISOString(),
                parentFolderId: 'excels',
                blob: file
            };
            const transaction = dbRef.transaction(STORE_NAME, 'readwrite');
            transaction.objectStore(STORE_NAME).put(fileObj);
            transaction.oncomplete = async () => {
                activeFileObj = {
                    id: fileObj.id,
                    name: fileObj.name,
                    parsedData: parsed
                };
                currentSheetName = parsed.sheetNames[0];
                await fetchExcelFilesList();
                renderExcelWorkspace();
                toast(`File "${file.name}" berhasil diunggah!`, 'success');
            };
        } catch (err) {
            console.error(err);
            toast('File Excel rusak atau format tidak didukung.', 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

async function selectExcelFileFromList(fileId) {
    const transaction = dbRef.transaction(STORE_NAME, 'readonly');
    const request = transaction.objectStore(STORE_NAME).get(fileId);
    request.onsuccess = () => {
        const item = request.result;
        if (item) {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const arrayBuffer = e.target.result;
                    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
                    const parsed = {
                        sheetNames: workbook.SheetNames,
                        sheets: {}
                    };
                    workbook.SheetNames.forEach((sheetName) => {
                        const sheet = workbook.Sheets[sheetName];
                        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
                        let headers = [];
                        if (rows.length > 0) {
                            headers = Object.keys(rows[0]);
                        }
                        parsed.sheets[sheetName] = { headers, rows };
                    });
                    activeFileObj = {
                        id: item.id,
                        name: item.name,
                        parsedData: parsed
                    };
                    currentSheetName = parsed.sheetNames[0] || '';
                    renderExcelWorkspace();
                    fetchExcelFilesList();
                    toast(`File "${item.name}" berhasil dimuat!`, 'success');
                } catch (err) {
                    console.error(err);
                    toast('Gagal memproses file Excel.', 'error');
                }
            };
            reader.readAsArrayBuffer(item.blob);
        }
    };
}

function renderExcelWorkspace() {
    const emptyBlock = document.getElementById('excel-no-file-selected');
    const activeBlock = document.getElementById('excel-active-workspace');
    if (!activeFileObj || !activeFileObj.parsedData) {
        emptyBlock.classList.remove('hidden');
        activeBlock.classList.add('hidden');
        return;
    }
    emptyBlock.classList.add('hidden');
    activeBlock.classList.remove('hidden');
    document.getElementById('active-excel-filename').innerText = activeFileObj.name;
    const tabContainer = document.getElementById('excel-sheets-tabs');
    tabContainer.innerHTML = '';
    activeFileObj.parsedData.sheetNames.forEach((name) => {
        const isActive = name === currentSheetName;
        const btn = document.createElement('button');
        btn.onclick = () => {
            currentSheetName = name;
            renderExcelWorkspace();
        };
        btn.className = `px-3 py-1.5 rounded-lg text-xs font-semibold whitespace-nowrap transition cursor-pointer ${
            isActive
                ? 'bg-blue-600 text-white shadow-xs'
                : 'bg-transparent text-slate-600 hover:bg-slate-100'
        }`;
        btn.innerText = name;
        tabContainer.appendChild(btn);
    });
    renderExcelDataGrid();
    loadDashboardStats();
}

function renderExcelDataGrid() {
    const tableHead = document.getElementById('excel-table-head');
    const tableBody = document.getElementById('excel-table-body');
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';
    if (!activeFileObj || !currentSheetName) return;
    const sheet = activeFileObj.parsedData.sheets[currentSheetName];
    if (!sheet) return;
    const headers = sheet.headers;
    const headTr = document.createElement('tr');
    const firstTh = document.createElement('th');
    firstTh.className = "py-3 px-4 w-12";
    firstTh.innerText = "No";
    headTr.appendChild(firstTh);
    headers.forEach((h) => {
        const th = document.createElement('th');
        th.className = "py-3 px-4 font-semibold";
        th.innerText = h;
        headTr.appendChild(th);
    });
    const actionTh = document.createElement('th');
    actionTh.className = "py-3 px-4 text-right w-24";
    actionTh.innerText = "Aksi";
    headTr.appendChild(actionTh);
    tableHead.appendChild(headTr);
    const q = searchQuery.toLowerCase().trim();
    const filteredRows = sheet.rows.filter((row) => {
        const matchesSearch = !q || Object.values(row).some((val) => String(val).toLowerCase().includes(q));
        const matchesMonth = selectedMonth === '' || isRowInSelectedMonth(row, selectedMonth);
        return matchesSearch && matchesMonth;
    });
    const monthLabel = selectedMonth === '' ? 'Semua bulan' : getMonthName(parseInt(selectedMonth, 10));
    document.getElementById('excel-grid-rows-indicator').innerText = `Total baris: ${filteredRows.length} dari ${sheet.rows.length} (${monthLabel})`;
    if (filteredRows.length === 0) {
        const rowTr = document.createElement('tr');
        rowTr.innerHTML = `<td colspan="${headers.length + 2}" class="py-8 text-center text-slate-400 text-xs">Tidak ditemukan baris data jemaat yang cocok.</td>`;
        tableBody.appendChild(rowTr);
        return;
    }
    filteredRows.forEach((row, index) => {
        const actualIndex = sheet.rows.indexOf(row);
        const tr = document.createElement('tr');
        tr.className = "hover:bg-slate-50/70 border-b border-slate-100 hover:shadow-xs transition-all";
        const numTd = document.createElement('td');
        numTd.className = "py-3.5 px-4 font-mono text-[10px] text-slate-400";
        numTd.innerText = String(index + 1);
        tr.appendChild(numTd);
        headers.forEach((h) => {
            const td = document.createElement('td');
            td.className = "py-3.5 px-4 font-medium text-slate-700 font-sans break-all max-w-[200px]";
            td.innerText = row[h] !== undefined ? row[h] : "";
            tr.appendChild(td);
        });
        const actTd = document.createElement('td');
        actTd.className = "py-3.5 px-4 text-right space-x-1 whitespace-nowrap";
        actTd.innerHTML = `
            <button onclick="openEditRecordModal(${actualIndex})" class="p-1 px-2 hover:bg-blue-50 text-blue-600 rounded-md font-bold text-[10px] cursor-pointer inline-block">Edit</button>
            <button onclick="deleteRecordFromActiveGrid(${actualIndex})" class="p-1 px-2 hover:bg-red-50 text-red-500 rounded-md font-bold text-[10px] cursor-pointer inline-block">Hapus</button>
        `;
        tr.appendChild(actTd);
        tableBody.appendChild(tr);
    });
}

function filterExcelTableGrid(query) {
    searchQuery = query;
    renderExcelDataGrid();
}

function filterExcelByMonth(month) {
    selectedMonth = month;
    renderExcelDataGrid();
}

function getMonthName(monthIndex) {
    const monthNames = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    return monthNames[monthIndex] || 'Semua bulan';
}

function parseDateValue(value) {
    if (value === undefined || value === null || value === '') return null;
    const raw = String(value).trim();
    const direct = new Date(raw);
    if (!isNaN(direct.getTime())) return direct;
    const parts = raw.split(/[\/\-. ]+/);
    if (parts.length === 3) {
        const [p1, p2, p3] = parts.map((part) => part.replace(/[^0-9]/g, ''));
        if (p1.length === 4) {
            const year = parseInt(p1, 10);
            const month = parseInt(p2, 10) - 1;
            const day = parseInt(p3, 10);
            return new Date(year, month, day);
        }
        const day = parseInt(p1, 10);
        const month = parseInt(p2, 10) - 1;
        const year = parseInt(p3, 10);
        if (!Number.isNaN(day) && !Number.isNaN(month) && !Number.isNaN(year)) {
            return new Date(year, month, day);
        }
    }
    return null;
}

function isRowInSelectedMonth(row, month) {
    const monthIndex = parseInt(month, 10);
    return Object.values(row).some((val) => {
        const date = parseDateValue(val);
        return date instanceof Date && !isNaN(date.getTime()) && date.getMonth() === monthIndex;
    });
}

async function triggerDeleteSelectedExcel() {
    if (!activeFileObj) return;
    const confirmed = window.confirm(`Apakah Anda yakin ingin menghapus file Excel "${activeFileObj.name}" dari sistem lokal secara permanen?`);
    if (confirmed) {
        const transaction = dbRef.transaction(STORE_NAME, 'readwrite');
        transaction.objectStore(STORE_NAME).delete(activeFileObj.id);
        transaction.oncomplete = async () => {
            activeFileObj = null;
            await fetchExcelFilesList();
            renderExcelWorkspace();
            toast('File Excel berhasil dihapus.', 'success');
        };
    }
}

async function saveActiveChangesToIndexedDB() {
    if (!activeFileObj) return;
    const saveBtn = document.getElementById('save-excel-changes-btn');
    saveBtn.disabled = true;
    saveBtn.innerHTML = `<i data-lucide="refresh-cw" class="w-3.5 h-3.5 animate-spin"></i><span>Menyimpan...</span>`;
    lucide.createIcons();
    try {
        const workbook = XLSX.utils.book_new();
        activeFileObj.parsedData.sheetNames.forEach((sheetName) => {
            const sheet = activeFileObj.parsedData.sheets[sheetName];
            const worksheet = XLSX.utils.json_to_sheet(sheet.rows);
            XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        });
        const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const updatedBlob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const transaction = dbRef.transaction(STORE_NAME, 'readwrite');
        const store = transaction.objectStore(STORE_NAME);
        const getRequest = store.get(activeFileObj.id);
        getRequest.onsuccess = () => {
            const record = getRequest.result;
            if (record) {
                record.blob = updatedBlob;
                record.size = String(updatedBlob.size);
                record.createdTime = new Date().toISOString();
                store.put(record);
            }
        };
        transaction.oncomplete = () => {
            saveBtn.disabled = false;
            saveBtn.innerHTML = `<i data-lucide="save" class="w-3.5 h-3.5 text-white"></i><span>Simpan Terkini</span>`;
            lucide.createIcons();
            toast('Perubahan database jemaat berhasil disimpan!', 'success');
        };
    } catch (err) {
        console.error(err);
        saveBtn.disabled = false;
        saveBtn.innerHTML = `<i data-lucide="save" class="w-3.5 h-3.5 text-white"></i><span>Simpan Terkini</span>`;
        lucide.createIcons();
        toast('Gagal melakukan penyimpanan database.', 'error');
    }
}

function downloadActiveExcelBackup() {
    if (!activeFileObj) return;
    try {
        const workbook = XLSX.utils.book_new();
        activeFileObj.parsedData.sheetNames.forEach((sheetName) => {
            const sheet = activeFileObj.parsedData.sheets[sheetName];
            const worksheet = XLSX.utils.json_to_sheet(sheet.rows);
            XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        });
        XLSX.writeFile(workbook, activeFileObj.name);
        toast('File cadangan lokal berhasil diunduh.', 'success');
    } catch (err) {
        console.error(err);
        toast('Gagal memulai unduhan file.', 'error');
    }
}

function openEditRecordModal(rowIndex) {
    if (!activeFileObj || !currentSheetName) return;
    const sheet = activeFileObj.parsedData.sheets[currentSheetName];
    if (!sheet) return;
    const row = sheet.rows[rowIndex];
    document.getElementById('row-modal-title').innerText = "Edit Baris Data Jemaat";
    document.getElementById('raw-modal-index-hidden').value = String(rowIndex);
    const container = document.getElementById('row-modal-fields-mapped');
    container.innerHTML = '';
    sheet.headers.forEach((h) => {
        const formGroup = document.createElement('div');
        formGroup.className = "space-y-1";
        formGroup.innerHTML = `
            <label class="text-[10px] font-bold text-slate-500 uppercase tracking-wide block">${h}</label>
            <input 
                type="text" 
                name="${h}"
                value="${row[h] !== undefined ? row[h] : ''}" 
                class="w-full px-4 py-2 text-xs bg-slate-50 border border-slate-200 focus:border-blue-500 rounded-lg outline-none transition text-slate-800 font-medium"
            />
        `;
        container.appendChild(formGroup);
    });
    document.getElementById('row-modal').classList.remove('hidden');
}

function openAddRecordModal() {
    if (!activeFileObj || !currentSheetName) return;
    const sheet = activeFileObj.parsedData.sheets[currentSheetName];
    if (!sheet) return;
    document.getElementById('row-modal-title').innerText = "Tambah Baris Database Baru";
    document.getElementById('raw-modal-index-hidden').value = "__NEW__";
    const container = document.getElementById('row-modal-fields-mapped');
    container.innerHTML = '';
    sheet.headers.forEach((h) => {
        const formGroup = document.createElement('div');
        formGroup.className = "space-y-1";
        formGroup.innerHTML = `
            <label class="text-[10px] font-bold text-slate-500 uppercase tracking-wide block">${h}</label>
            <input 
                type="text" 
                name="${h}"
                placeHolder="Masukkan nilai kolom ${h}"
                class="w-full px-4 py-2 text-xs bg-slate-50 border border-slate-200 focus:border-blue-500 rounded-lg outline-none transition text-slate-800 font-medium"
            />
        `;
        container.appendChild(formGroup);
    });
    document.getElementById('row-modal').classList.remove('hidden');
}

function closeRowDialogueModal() {
    document.getElementById('row-modal').classList.add('hidden');
}

function handleRowModalFormSubmit(event) {
    event.preventDefault();
    const targetIndexStr = document.getElementById('raw-modal-index-hidden').value;
    const sheet = activeFileObj.parsedData.sheets[currentSheetName];
    if (!sheet) return;
    const inputs = document.getElementById('row-modal-form').querySelectorAll('input[name]');
    const compiledRow = {};
    inputs.forEach((input) => {
        compiledRow[input.name] = input.value;
    });
    if (targetIndexStr === '__NEW__') {
        sheet.rows.push(compiledRow);
        toast('Satu baris data jemaat berhasil ditambahkan!', 'success');
    } else {
        const idx = parseInt(targetIndexStr, 10);
        sheet.rows[idx] = compiledRow;
        toast('Perubahan baris berhasil diubah!', 'success');
    }
    closeRowDialogueModal();
    renderExcelDataGrid();
    loadDashboardStats();
}

function deleteRecordFromActiveGrid(rowIndex) {
    if (!activeFileObj || !currentSheetName) return;
    const sheet = activeFileObj.parsedData.sheets[currentSheetName];
    if (!sheet) return;
    const confirmed = window.confirm("Apakah Anda yakin ingin menghapus baris data jemaat ini?");
    if (confirmed) {
        sheet.rows.splice(rowIndex, 1);
        renderExcelDataGrid();
        loadDashboardStats();
        toast('Satu baris jemaat dihapus dari kisi.', 'success');
    }
}

async function fetchPhotosList() {
    if (!dbRef) return;
    const grid = document.getElementById('gallery-items-grid');
    const alertBox = document.getElementById('gallery-empty-alert');
    grid.innerHTML = '';
    const transaction = dbRef.transaction(STORE_NAME, 'readonly');
    const store = transaction.objectStore(STORE_NAME);
    const cursorReq = store.openCursor();
    const photoList = [];
    cursorReq.onsuccess = () => {
        const cursor = cursorReq.result;
        if (cursor) {
            if (cursor.value.parentFolderId === 'photos') {
                photoList.push(cursor.value);
            }
            cursor.continue();
        } else {
            if (photoList.length === 0) {
                alertBox.classList.remove('hidden');
                grid.classList.add('hidden');
                return;
            }
            alertBox.classList.add('hidden');
            grid.classList.remove('hidden');
            photoList.sort((a,b) => getPhotoOrderValue(b) - getPhotoOrderValue(a));
            photoList.forEach((pic) => {
                const cell = document.createElement('div');
                cell.className = "bg-white border border-slate-200/80 rounded-2xl overflow-hidden hover:shadow-md transition-all flex flex-col group relative";
                let imageSrcUrl = '';
                try {
                    imageSrcUrl = URL.createObjectURL(pic.blob);
                } catch (e) {
                    imageSrcUrl = '#';
                }
                if (photoDragMode) {
                    cell.draggable = true;
                    cell.dataset.photoId = pic.id;
                    cell.classList.add('drag-enabled');
                    cell.addEventListener('dragstart', (event) => {
                        currentDraggedPhotoId = pic.id;
                        event.dataTransfer.setData('text/plain', pic.id);
                        event.dataTransfer.effectAllowed = 'move';
                        cell.classList.add('opacity-50');
                    });
                    cell.addEventListener('dragend', () => {
                        currentDraggedPhotoId = null;
                        cell.classList.remove('opacity-50');
                    });
                    cell.addEventListener('dragover', (event) => {
                        event.preventDefault();
                        event.dataTransfer.dropEffect = 'move';
                        cell.classList.add('drag-over-target');
                    });
                    cell.addEventListener('dragleave', () => {
                        cell.classList.remove('drag-over-target');
                    });
                    cell.addEventListener('drop', async (event) => {
                        event.preventDefault();
                        cell.classList.remove('drag-over-target');
                        const targetPhotoId = pic.id;
                        if (!currentDraggedPhotoId || currentDraggedPhotoId === targetPhotoId) return;
                        const reordered = reorderPhotoList(photoList, currentDraggedPhotoId, targetPhotoId);
                        await savePhotoOrderSequence(reordered);
                        fetchPhotosList();
                        toast('Urutan galeri berhasil diperbarui.', 'success');
                    });
                }
                cell.innerHTML = `
                    <div class="h-40 overflow-hidden bg-slate-50 relative">
                        <img src="${imageSrcUrl}" alt="Arsip" class="w-full h-full object-cover group-hover:scale-105 transition duration-300" />
                        <div class="absolute inset-0 bg-black/40 opacity-0 group-hover:opacity-100 transition flex items-center justify-center gap-2">
                            <button onclick="previewPhotoLightbox('${pic.id}', '${pic.name}', '${imageSrcUrl}')" class="p-2 bg-white rounded-lg text-slate-700 hover:text-black hover:scale-105 transition cursor-pointer shadow">
                                <i data-lucide="eye" class="w-4 h-4"></i>
                            </button>
                            <button onclick="triggerPhotoLocalDownload('${pic.id}', '${pic.name}')" class="p-2 bg-white rounded-lg text-slate-705 hover:text-black hover:scale-105 transition cursor-pointer shadow">
                                <i data-lucide="download" class="w-4 h-4"></i>
                            </button>
                        </div>
                    </div>
                    <div class="p-3.5 flex items-center justify-between gap-2 border-t border-slate-100 bg-white">
                        <div class="truncate">
                            <h4 class="text-xs font-bold text-slate-800 truncate leading-snug">${pic.name}</h4>
                            <p class="text-[9px] text-slate-400 font-mono mt-0.5">${(parseInt(pic.size, 10)/1024).toFixed(1)} KB</p>
                        </div>
                        <button onclick="triggerRemovePhotoImage('${pic.id}')" class="p-1.5 rounded-lg border border-slate-100 hover:border-red-100 hover:bg-red-50 text-slate-450 hover:text-red-500 transition cursor-pointer" title="Hapus Gambar">
                            <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
                        </button>
                    </div>
                `;
                grid.appendChild(cell);
            });
            lucide.createIcons();
        }
    };
}

function getPhotoOrderValue(photo) {
    return typeof photo.displayOrder === 'number' ? photo.displayOrder : new Date(photo.createdTime).getTime();
}

function reorderPhotoList(photoList, draggedId, targetId) {
    const draggedIndex = photoList.findIndex((item) => item.id === draggedId);
    const targetIndex = photoList.findIndex((item) => item.id === targetId);
    if (draggedIndex === -1 || targetIndex === -1 || draggedIndex === targetIndex) return photoList;
    const [movedItem] = photoList.splice(draggedIndex, 1);
    photoList.splice(targetIndex, 0, movedItem);
    return photoList;
}

async function savePhotoOrderSequence(photoList) {
    if (!dbRef) return;
    const transaction = dbRef.transaction(STORE_NAME, 'readwrite');
    const store = transaction.objectStore(STORE_NAME);
    const count = photoList.length;
    photoList.forEach((item, index) => {
        const updated = { ...item, displayOrder: count - index };
        store.put(updated);
    });
    return new Promise((resolve, reject) => {
        transaction.oncomplete = resolve;
        transaction.onerror = reject;
        transaction.onabort = reject;
    });
}

function updatePhotoDragModeButton() {
    const btn = document.getElementById('toggle-drag-mode-btn');
    if (!btn) return;
    if (photoDragMode) {
        btn.className = 'px-3 py-2 text-[10px] font-semibold rounded-xl border border-blue-200 bg-blue-600 text-white hover:bg-blue-700 transition';
        btn.innerText = 'Nonaktifkan Drag Mode';
    } else {
        btn.className = 'px-3 py-2 text-[10px] font-semibold rounded-xl border border-slate-200 text-slate-600 bg-slate-50 hover:bg-slate-100 transition';
        btn.innerText = 'Aktifkan Drag Mode';
    }
}

function togglePhotoDragMode() {
    photoDragMode = !photoDragMode;
    updatePhotoDragModeButton();
    fetchPhotosList();
    toast(photoDragMode ? 'Drag Mode aktif — seret kartu foto untuk mengubah urutan.' : 'Drag Mode nonaktif.', 'success');
}

function handlePhotoDropzoneDragOver(event) {
    event.preventDefault();
    event.dataTransfer.dropEffect = 'copy';
    event.currentTarget.classList.add('drag-over');
}

function handlePhotoDropzoneDragEnter(event) {
    event.preventDefault();
    event.currentTarget.classList.add('drag-over');
}

function handlePhotoDropzoneDragLeave(event) {
    event.currentTarget.classList.remove('drag-over');
}

function handlePhotoDropzoneDrop(event) {
    event.preventDefault();
    event.currentTarget.classList.remove('drag-over');
    const files = event.dataTransfer.files;
    if (files.length > 0) {
        handlePhotoLocalUpload({ target: { files } });
    }
}

function triggerPhotoSecretInput() {
    document.getElementById('photo-hidden-file-input').click();
}

function triggerDocUploadInput() {
    document.getElementById('doc-hidden-file-input').click();
}

async function handleDocLocalUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const transaction = dbRef.transaction(STORE_NAME, 'readwrite');
    const record = {
        id: 'doc-' + Math.random().toString(36).substring(2, 9),
        name: file.name,
        mimeType: file.type || 'application/octet-stream',
        size: String(file.size),
        createdTime: new Date().toISOString(),
        parentFolderId: 'docs',
        blob: file
    };
    transaction.objectStore(STORE_NAME).put(record);
    transaction.oncomplete = () => {
        fetchDocsList();
        loadDashboardStats();
        toast(`Dokumen "${file.name}" berhasil diunggah!`, 'success');
    };
}

async function fetchDocsList() {
    if (!dbRef) return;
    const list = document.getElementById('docs-files-list');
    list.innerHTML = '';

    const transaction = dbRef.transaction(STORE_NAME, 'readonly');
    const store = transaction.objectStore(STORE_NAME);
    const cursorReq = store.openCursor();
    const docs = [];

    cursorReq.onsuccess = () => {
        const cursor = cursorReq.result;
        if (cursor) {
            if (cursor.value.parentFolderId === 'docs') {
                docs.push(cursor.value);
            }
            cursor.continue();
        } else {
            if (docs.length === 0) {
                list.innerHTML = `<div class="text-[10px] text-slate-400 text-center py-6 border border-dashed border-slate-200 rounded-2xl">Belum ada dokumen pendukung.</div>`;
                return;
            }

            docs.sort((a,b) => new Date(b.createdTime).getTime() - new Date(a.createdTime).getTime());
            docs.forEach((doc) => {
                const item = document.createElement('div');
                item.className = 'flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 p-4 rounded-2xl border border-slate-200 bg-slate-50';
                item.innerHTML = `
                    <div class="min-w-0">
                        <p class="text-xs font-bold text-slate-900 truncate">${doc.name}</p>
                        <p class="text-[9px] text-slate-500 mt-1">${doc.mimeType || 'Dokumen'} • ${(parseInt(doc.size, 10) / 1024).toFixed(1)} KB</p>
                    </div>
                    <div class="flex items-center gap-2">
                        <button onclick="previewDocFile('${doc.id}')" class="px-3 py-2 bg-slate-200 hover:bg-slate-300 text-slate-900 rounded-xl text-[10px] font-semibold transition">Pratinjau</button>
                        <button onclick="triggerDocLocalDownload('${doc.id}', '${doc.name}')" class="px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-[10px] font-semibold transition">Unduh</button>
                        <button onclick="triggerRemoveDocFile('${doc.id}')" class="px-3 py-2 bg-red-50 hover:bg-red-100 text-red-600 rounded-xl text-[10px] font-semibold transition">Hapus</button>
                    </div>
                `;
                list.appendChild(item);
            });
        }
    };
}

async function triggerRemoveDocFile(docId) {
    const confirmed = window.confirm("Apakah Anda yakin ingin menghapus dokumen ini dari arsip pendukung?");
    if (confirmed) {
        const transaction = dbRef.transaction(STORE_NAME, 'readwrite');
        transaction.objectStore(STORE_NAME).delete(docId);
        transaction.oncomplete = () => {
            fetchDocsList();
            loadDashboardStats();
            toast('Dokumen berhasil dihapus.', 'success');
        };
    }
}

async function triggerDocLocalDownload(docId, filename) {
    const transaction = dbRef.transaction(STORE_NAME, 'readonly');
    transaction.objectStore(STORE_NAME).get(docId).onsuccess = (e) => {
        const record = e.target.result;
        if (record) {
            const downloadUrl = URL.createObjectURL(record.blob);
            const link = document.createElement('a');
            link.href = downloadUrl;
            link.download = filename;
            link.click();
            toast('Mengunduh dokumen...', 'success');
        }
    };
}

function previewDocFile(docId) {
    const transaction = dbRef.transaction(STORE_NAME, 'readonly');
    transaction.objectStore(STORE_NAME).get(docId).onsuccess = (e) => {
        const record = e.target.result;
        const previewContent = document.getElementById('doc-preview-content');
        const titleEl = document.getElementById('doc-preview-title');
        const infoEl = document.getElementById('doc-preview-info');
        const openBtn = document.getElementById('doc-preview-open-btn');

        if (!record) {
            previewContent.innerHTML = `<div class="p-6 text-center text-slate-500 text-sm">Dokumen tidak ditemukan.</div>`;
            titleEl.innerText = 'Preview Dokumen';
            infoEl.innerText = 'Tidak ada konten yang dapat ditampilkan.';
            openBtn.classList.add('hidden');
            document.getElementById('doc-preview-modal').classList.remove('hidden');
            return;
        }

        if (activeDocPreviewUrl) {
            URL.revokeObjectURL(activeDocPreviewUrl);
            activeDocPreviewUrl = null;
        }

        const blobUrl = URL.createObjectURL(record.blob);
        activeDocPreviewUrl = blobUrl;
        titleEl.innerText = record.name;
        infoEl.innerText = `${record.mimeType || 'Dokumen'} • ${(parseInt(record.size, 10) / 1024).toFixed(1)} KB`;
        openBtn.classList.remove('hidden');
        openBtn.onclick = openDocInNewTab;

        const previewableImage = record.mimeType?.startsWith('image/');
        const previewablePdf = record.mimeType === 'application/pdf';
        const previewableHtml = record.mimeType === 'text/html';

        if (previewableImage) {
            previewContent.innerHTML = `<div class="flex items-center justify-center h-full p-4"><img src="${blobUrl}" alt="${record.name}" class="max-w-full max-h-[76vh] object-contain rounded-2xl" /></div>`;
        } else if (previewablePdf) {
            previewContent.innerHTML = `<iframe src="${blobUrl}" class="w-full h-full border-0" title="Preview PDF"></iframe>`;
        } else if (previewableHtml) {
            previewContent.innerHTML = `<iframe src="${blobUrl}" class="w-full h-full border-0" title="Preview Dokumen"></iframe>`;
        } else {
            previewContent.innerHTML = `<div class="p-6 text-center text-slate-500 text-sm space-y-3"><p class="font-semibold text-slate-900">Pratinjau tidak tersedia untuk jenis file ini.</p><p>Silakan buka di tab baru atau unduh file untuk melihatnya secara lengkap.</p></div>`;
        }

        document.getElementById('doc-preview-modal').classList.remove('hidden');
    };
}

function openDocInNewTab() {
    if (!activeDocPreviewUrl) return;
    window.open(activeDocPreviewUrl, '_blank');
}

function closeDocPreviewModal() {
    const previewContent = document.getElementById('doc-preview-content');
    previewContent.innerHTML = '';
    document.getElementById('doc-preview-modal').classList.add('hidden');
    if (activeDocPreviewUrl) {
        URL.revokeObjectURL(activeDocPreviewUrl);
        activeDocPreviewUrl = null;
    }
}

async function handlePhotoLocalUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    const iconUploader = document.getElementById('photo-uploader-status-icon');
    iconUploader.innerHTML = `<i data-lucide="refresh-cw" class="w-5 h-5 text-blue-500 animate-spin"></i>`;
    lucide.createIcons();
    setTimeout(() => {
        const id = 'photo-' + Math.random().toString(36).substring(2, 9);
        const record = {
            id,
            name: file.name,
            mimeType: file.type || 'image/jpeg',
            size: String(file.size),
            createdTime: new Date().toISOString(),
            parentFolderId: 'photos',
            displayOrder: Date.now(),
            blob: file
        };
        const transaction = dbRef.transaction(STORE_NAME, 'readwrite');
        transaction.objectStore(STORE_NAME).put(record);
        transaction.oncomplete = () => {
            iconUploader.innerHTML = `<i data-lucide="plus" class="w-5 h-5 text-blue-500"></i>`;
            lucide.createIcons();
            fetchPhotosList();
            loadDashboardStats();
            toast(`Foto "${file.name}" berhasil diunggah!`, 'success');
        };
    }, 600);
}

async function triggerRemovePhotoImage(photoId) {
    const confirmed = window.confirm("Apakah Anda yakin ingin menghapus foto dokumentasi ini dari arsip internal?");
    if (confirmed) {
        const transaction = dbRef.transaction(STORE_NAME, 'readwrite');
        transaction.objectStore(STORE_NAME).delete(photoId);
        transaction.oncomplete = () => {
            fetchPhotosList();
            loadDashboardStats();
            toast('Foto berhasil dihapus dari arsip.', 'success');
        };
    }
}

async function triggerPhotoLocalDownload(photoId, filename) {
    const transaction = dbRef.transaction(STORE_NAME, 'readonly');
    transaction.objectStore(STORE_NAME).get(photoId).onsuccess = (e) => {
        const record = e.target.result;
        if (record) {
            const downloadUrl = URL.createObjectURL(record.blob);
            const link = document.createElement('a');
            link.href = downloadUrl;
            link.download = filename;
            link.click();
            toast('Mengunduh foto...', 'success');
        }
    };
}

function previewPhotoLightbox(id, name, blobUrl) {
    document.getElementById('lightbox-item-title').innerText = name;
    document.getElementById('lightbox-main-img').src = blobUrl;
    document.getElementById('lightbox-modal').classList.remove('hidden');
    document.getElementById('lightbox-download-action-btn').onclick = () => {
        const link = document.createElement('a');
        link.href = blobUrl;
        link.download = name;
        link.click();
    };
}

function closeLightboxModal() {
    document.getElementById('lightbox-modal').classList.add('hidden');
}
