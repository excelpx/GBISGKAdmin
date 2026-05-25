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
let currentExcelPage = 1;
let excelRowsPerPage = 30;
let activeDocPreviewUrl = null;
let currentDraggedPhotoId = null;
let pendingDocFiles = [];
let docCustomCategories = [];
let selectedDocCategory = null;
let selectedDocFilterCategory = 'semua';
let currentDocsFetchId = 0;
let uploadInProgress = false;
let selectedDocSearchQuery = '';
let externalDocLinks = [];
let externalYouthDocLinks = [];
let driveFileCache = null;
let driveFileCacheTimestamp = 0;
let driveFileCachePromise = null;
const DRIVE_FILE_CACHE_TTL = 5000;

let stats = {
    totalRows: 0,
    filesCount: 0,
    photosCount: 0
};

// ================= GOOGLE DRIVE DATABASE CONFIG =================
// 1) Deploy Code.gs sebagai Web App
// 2) Paste URL Web App Google Apps Script di bawah ini
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycby6gFd4FiYx3fuHQy5J3KZqtrG68KWvdXG0RMCLSpUt8tCZqyiKx9lbwc1PPHjBnnZUZQ/exec";
const DRIVE_FOLDER_ID = "1B1Sk6QJraYeWYDycrualpTlSr30IaNTB";

async function driveApi(action, payload = {}) {
    if (!APPS_SCRIPT_URL || APPS_SCRIPT_URL.includes('PASTE_URL')) {
        throw new Error('URL Google Apps Script belum diisi di script.js');
    }
    const response = await fetch(APPS_SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action, ...payload })
    });
    const text = await response.text();
    let result;
    try { result = JSON.parse(text); } catch (e) { throw new Error('Response Apps Script bukan JSON: ' + text); }
    if (!result.success) throw new Error(result.error || 'Terjadi kesalahan Google Drive');
    return result;
}

function fileToBase64(file) {
    // Optional second argument is onProgress callback: (loaded, total) => {}
    return function _fileToBase64(fileParam, onProgress) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(String(reader.result).split(',')[1]);
            reader.onerror = reject;
            if (onProgress && typeof onProgress === 'function') {
                reader.onprogress = (ev) => {
                    try { onProgress(ev.loaded, ev.total); } catch (e) {}
                };
            }
            reader.readAsDataURL(fileParam);
        });
    };
}

function base64ToBlob(base64, mimeType) {
    const bytes = atob(base64);
    const array = new Uint8Array(bytes.length);
    for (let i = 0; i < bytes.length; i++) array[i] = bytes.charCodeAt(i);
    return new Blob([array], { type: mimeType || 'application/octet-stream' });
}

async function uploadFileToDrive(file, category, extra = {}, onProgress) {
    // fileToBase64 was converted to a factory; call it with progress callback
    const converter = fileToBase64();
    const base64 = await converter(file, onProgress);
    const result = await driveApi('upload', {
        fileName: file.name,
        mimeType: file.type || 'application/octet-stream',
        category,
        displayOrder: extra.displayOrder || Date.now(),
        file: base64
    });
    invalidateDriveFileCache();
    return result.file;
}

function invalidateDriveFileCache() {
    driveFileCache = null;
    driveFileCacheTimestamp = 0;
    driveFileCachePromise = null;
}

async function getCachedDriveFiles({ forceRefresh = false } = {}) {
    if (!forceRefresh && driveFileCache && (Date.now() - driveFileCacheTimestamp) < DRIVE_FILE_CACHE_TTL) {
        return driveFileCache;
    }
    if (!forceRefresh && driveFileCachePromise) {
        return driveFileCachePromise;
    }
    driveFileCachePromise = driveApi('list')
        .then((result) => {
            driveFileCache = result.files || [];
            driveFileCacheTimestamp = Date.now();
            driveFileCachePromise = null;
            return driveFileCache;
        })
        .catch((err) => {
            driveFileCachePromise = null;
            throw err;
        });
    return driveFileCachePromise;
}

async function getDriveFileBlob(fileId) {
    const result = await driveApi('get', { fileId });
    return {
        ...result.file,
        blob: base64ToBlob(result.file.base64, result.file.mimeType)
    };
}
// ================================================================


function initAppDatabase() {
    // Database sekarang menggunakan Google Drive melalui Google Apps Script.
    dbRef = true;
    return Promise.resolve(true);
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
    // Tidak menanam data dummy lokal lagi, supaya data murni dari Google Drive.
    return Promise.resolve();
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

function showGlobalLoader(message = 'Memuat...') {
    const overlay = document.getElementById('global-loading-overlay');
    const messageEl = document.getElementById('global-loading-message');
    if (messageEl) messageEl.innerText = message;
    if (overlay) overlay.classList.remove('hidden');
}

function hideGlobalLoader() {
    const overlay = document.getElementById('global-loading-overlay');
    if (overlay) overlay.classList.add('hidden');
}

function setButtonState(button, { loading = false, text = null } = {}) {
    if (!button) return;
    button.disabled = loading;
    if (loading) {
        button.innerHTML = `<i data-lucide="refresh-cw" class="w-4 h-4 animate-spin"></i><span>${text || button.dataset.loadingText || button.innerText}</span>`;
    } else if (text !== null) {
        button.innerHTML = text;
    }
    lucide.createIcons();
}

function waitForNextFrame() {
    return new Promise((resolve) => requestAnimationFrame(resolve));
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
});

async function checkSession() {
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
        
        // Load custom categories and sanitize invalid values
        const stored = localStorage.getItem('doc_custom_categories');
        const loaded = stored ? JSON.parse(stored) : [];
        docCustomCategories = Array.isArray(loaded)
            ? loaded.filter((cat) => typeof cat === 'string' && cat.trim() && !['null', 'undefined'].includes(cat.trim().toLowerCase()))
            : [];
        localStorage.setItem('doc_custom_categories', JSON.stringify(docCustomCategories));
        updateDocCategoryFilters();

        const externalStored = localStorage.getItem('external_doc_links');
        externalDocLinks = externalStored ? JSON.parse(externalStored) : [];
        if (!Array.isArray(externalDocLinks)) externalDocLinks = [];

        const externalYouthStored = localStorage.getItem('external_youth_doc_links');
        externalYouthDocLinks = externalYouthStored ? JSON.parse(externalYouthStored) : [];
        if (!Array.isArray(externalYouthDocLinks)) externalYouthDocLinks = [];

        await Promise.all([
            loadDashboardStats(),
            fetchExcelFilesList(),
            fetchPhotosList(),
            fetchDocsList()
        ]);
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
    const screens = ['dashboard', 'excel', 'photo', 'youth'];

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

    if (tabId === 'photo' || tabId === 'youth') {
        selectedDocFilterCategory = 'semua';
        selectedDocSearchQuery = '';

        const docsSearchInput = document.getElementById('docs-search-input');
        if (docsSearchInput) docsSearchInput.value = '';
        const youthSearchInput = document.getElementById('youth-search-input');
        if (youthSearchInput) youthSearchInput.value = '';

        initializeDocCategoryFilter();
        fetchDocsList();
    }

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
    try {
        const files = await getCachedDriveFiles();
        const excels = files.filter(f => f.category === 'excels');
        const photos = files.filter(f => f.category === 'photos');
        stats.filesCount = excels.length;
        stats.photosCount = photos.length;
        if (activeFileObj && activeFileObj.parsedData) {
            const activeSheet = activeFileObj.parsedData.sheets[currentSheetName];
            stats.totalRows = activeSheet ? activeSheet.rows.length : 0;
        } else {
            stats.totalRows = 0;
        }
        document.getElementById('metric-files-count').innerText = stats.filesCount;
        document.getElementById('metric-photos-count').innerText = stats.photosCount;
        document.getElementById('metric-total-rows').innerText = stats.totalRows;
    } catch (err) {
        console.error(err);
        toast(err.message || 'Gagal memuat statistik Google Drive.', 'error');
    }
}


async function fetchExcelFilesList() {
    document.getElementById('excel-list-loading').classList.remove('hidden');
    const container = document.getElementById('excel-files-list');
    container.innerHTML = '';
    try {
        const files = await getCachedDriveFiles();
        const items = files.filter(f => f.category === 'excels');
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
                isSelected ? 'bg-blue-50 border-blue-200 text-blue-950' : 'bg-white hover:bg-slate-50 border-slate-200/70 text-slate-700'
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
    } catch (err) {
        document.getElementById('excel-list-loading').classList.add('hidden');
        console.error(err);
        toast(err.message || 'Gagal memuat daftar Excel dari Google Drive.', 'error');
    }
}


function triggerExcelInput() {
    document.getElementById('excel-file-hidden-input').click();
}

async function handleLocalExcelUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    showGlobalLoader('Memproses file Excel...');
    await waitForNextFrame();
    try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array' });
        const parsed = { sheetNames: workbook.SheetNames, sheets: {} };
        workbook.SheetNames.forEach((sheetName) => {
            const sheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
            const headers = rows.length > 0 ? Object.keys(rows[0]) : [];
            parsed.sheets[sheetName] = { headers, rows };
        });
        const uploaded = await uploadFileToDrive(file, 'excels');
        activeFileObj = { id: uploaded.id, name: uploaded.name, parsedData: parsed };
        currentSheetName = parsed.sheetNames[0];
        await fetchExcelFilesList();
        renderExcelWorkspace();
        loadDashboardStats();
        toast(`File "${file.name}" berhasil diunggah ke Google Drive!`, 'success');
    } catch (err) {
        console.error(err);
        toast(err.message || 'File Excel gagal diunggah ke Google Drive.', 'error');
    } finally {
        hideGlobalLoader();
        event.target.value = '';
    }
}


async function selectExcelFileFromList(fileId) {
    showGlobalLoader('Memuat data Excel...');
    await waitForNextFrame();
    try {
        const item = await getDriveFileBlob(fileId);
        const arrayBuffer = await item.blob.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const parsed = { sheetNames: workbook.SheetNames, sheets: {} };
        workbook.SheetNames.forEach((sheetName) => {
            const sheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
            const headers = rows.length > 0 ? Object.keys(rows[0]) : [];
            parsed.sheets[sheetName] = { headers, rows };
        });
        activeFileObj = { id: item.id, name: item.name, parsedData: parsed };
        currentSheetName = parsed.sheetNames[0] || '';
        renderExcelWorkspace();
        fetchExcelFilesList();
        toast(`File "${item.name}" berhasil dimuat dari Google Drive!`, 'success');
    } catch (err) {
        console.error(err);
        toast(err.message || 'Gagal memproses file Excel dari Google Drive.', 'error');
    } finally {
        hideGlobalLoader();
    }
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
    currentExcelPage = 1;
    renderExcelDataGrid();
    loadDashboardStats();
}

function goToExcelPage(pageNumber) {
    currentExcelPage = Math.max(1, Math.min(pageNumber, 999999));
    renderExcelDataGrid();
}

function setExcelRowsPerPage(count) {
    excelRowsPerPage = count;
    currentExcelPage = 1;
    renderExcelDataGrid();
}

function renderExcelPagination(totalPages, currentPage) {
    const container = document.getElementById('excel-pagination-controls');
    if (!container) return;
    container.innerHTML = '';

    const pageInfo = document.createElement('div');
    pageInfo.className = 'text-[10px] text-slate-500';
    pageInfo.innerText = `Halaman ${currentPage} dari ${totalPages}`;
    container.appendChild(pageInfo);

    const pager = document.createElement('div');
    pager.className = 'flex items-center gap-2 flex-wrap';

    const prevBtn = document.createElement('button');
    prevBtn.type = 'button';
    prevBtn.innerText = 'Sebelumnya';
    prevBtn.disabled = currentPage <= 1;
    prevBtn.className = `px-3 py-1.5 rounded-xl text-[10px] font-semibold transition ${currentPage <= 1 ? 'bg-slate-100 text-slate-400 cursor-not-allowed' : 'bg-white border border-slate-200 text-slate-700 hover:bg-slate-50'}`;
    prevBtn.onclick = () => goToExcelPage(currentPage - 1);
    pager.appendChild(prevBtn);

    const nextBtn = document.createElement('button');
    nextBtn.type = 'button';
    nextBtn.innerText = 'Berikutnya';
    nextBtn.disabled = currentPage >= totalPages;
    nextBtn.className = `px-3 py-1.5 rounded-xl text-[10px] font-semibold transition ${currentPage >= totalPages ? 'bg-slate-100 text-slate-400 cursor-not-allowed' : 'bg-white border border-slate-200 text-slate-700 hover:bg-slate-50'}`;
    nextBtn.onclick = () => goToExcelPage(currentPage + 1);
    pager.appendChild(nextBtn);

    const perPage = document.createElement('select');
    perPage.className = 'text-[10px] border border-slate-200 rounded-xl px-3 py-1.5 bg-slate-50';
    perPage.onchange = (event) => setExcelRowsPerPage(parseInt(event.target.value, 10));
    [10, 20, 30, 50, 100].forEach((size) => {
        const option = document.createElement('option');
        option.value = size;
        option.innerText = `${size} baris`;
        if (size === excelRowsPerPage) option.selected = true;
        perPage.appendChild(option);
    });
    pager.appendChild(perPage);

    container.appendChild(pager);
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

    if (selectedMonth !== '') {
        filteredRows.sort((a, b) => {
            const dateA = getEarliestDateInMonth(a, selectedMonth);
            const dateB = getEarliestDateInMonth(b, selectedMonth);
            if (dateA && dateB) return dateA.getDate() - dateB.getDate();
            if (dateA) return -1;
            if (dateB) return 1;
            return 0;
        });
    }

    const totalRows = filteredRows.length;
    const totalPages = Math.max(1, Math.ceil(totalRows / excelRowsPerPage));
    if (currentExcelPage > totalPages) currentExcelPage = totalPages;
    const startIndex = (currentExcelPage - 1) * excelRowsPerPage;
    const pageRows = filteredRows.slice(startIndex, startIndex + excelRowsPerPage);
    const monthLabel = selectedMonth === '' ? 'Semua bulan' : getMonthName(parseInt(selectedMonth, 10));
    const showingLabel = totalRows === 0 ? '0' : `${startIndex + 1}-${Math.min(startIndex + excelRowsPerPage, totalRows)}`;
    document.getElementById('excel-grid-rows-indicator').innerText = `Menampilkan ${showingLabel} dari ${totalRows} baris (${monthLabel})`;
    renderExcelPagination(totalPages, currentExcelPage);
    if (totalRows === 0) {
        const rowTr = document.createElement('tr');
        rowTr.innerHTML = `<td colspan="${headers.length + 2}" class="py-8 text-center text-slate-400 text-xs">Tidak ditemukan baris data jemaat yang cocok.</td>`;
        tableBody.appendChild(rowTr);
        return;
    }
    pageRows.forEach((row, pageIndex) => {
        const actualIndex = sheet.rows.indexOf(row);
        const tr = document.createElement('tr');
        const rowNumber = startIndex + pageIndex + 1;
        tr.className = "hover:bg-slate-50/70 border-b border-slate-100 hover:shadow-xs transition-all";
        const numTd = document.createElement('td');
        numTd.className = "py-3.5 px-4 font-mono text-[10px] text-slate-400";
        numTd.innerText = String(rowNumber);
        tr.appendChild(numTd);
        headers.forEach((h) => {
            const td = document.createElement('td');
            td.className = "py-3.5 px-4 font-medium text-slate-700 font-sans";
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
    currentExcelPage = 1;
    renderExcelDataGrid();
}

function filterExcelByMonth(month) {
    selectedMonth = month;
    currentExcelPage = 1;
    renderExcelDataGrid();
}

function getMonthName(monthIndex) {
    const monthNames = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    return monthNames[monthIndex] || 'Semua bulan';
}

function parseDateValue(value) {
    if (value === undefined || value === null || value === '') return null;
    const raw = String(value).trim();
    const cleaned = raw.replace(/\s+/g, ' ');

    const numericDateMatch = cleaned.match(/^([0-9]{1,2})[\/\-.]([0-9]{1,2})[\/\-.]([0-9]{2,4})$/);
    if (numericDateMatch) {
        const day = parseInt(numericDateMatch[1], 10);
        const month = parseInt(numericDateMatch[2], 10) - 1;
        let year = parseInt(numericDateMatch[3], 10);
        if (year < 100) year += year >= 70 ? 1900 : 2000;
        if (!Number.isNaN(day) && !Number.isNaN(month) && !Number.isNaN(year)) {
            return new Date(year, month, day);
        }
    }

    const isoDateMatch = cleaned.match(/^([0-9]{4})[\/\-.]([0-9]{1,2})[\/\-.]([0-9]{1,2})$/);
    if (isoDateMatch) {
        const year = parseInt(isoDateMatch[1], 10);
        const month = parseInt(isoDateMatch[2], 10) - 1;
        const day = parseInt(isoDateMatch[3], 10);
        if (!Number.isNaN(day) && !Number.isNaN(month) && !Number.isNaN(year)) {
            return new Date(year, month, day);
        }
    }

    const direct = new Date(cleaned);
    if (!isNaN(direct.getTime())) return direct;

    return null;
}

function getEarliestDateInMonth(row, month) {
    const monthIndex = parseInt(month, 10);
    const dates = Object.values(row)
        .map((val) => parseDateValue(val))
        .filter((date) => date instanceof Date && !isNaN(date.getTime()) && date.getMonth() === monthIndex);
    if (dates.length === 0) return null;
    dates.sort((a, b) => a.getTime() - b.getTime());
    return dates[0];
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
    const confirmed = window.confirm(`Apakah Anda yakin ingin menghapus file Excel "${activeFileObj.name}" dari Google Drive?`);
    if (confirmed) {
        try {
            await driveApi('delete', { fileId: activeFileObj.id });
            activeFileObj = null;
            await fetchExcelFilesList();
            renderExcelWorkspace();
            loadDashboardStats();
            toast('File Excel berhasil dihapus dari Google Drive.', 'success');
        } catch (err) {
            console.error(err);
            toast(err.message || 'Gagal menghapus file Excel.', 'error');
        }
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
        const converter = fileToBase64();
        const base64 = await converter(new File([updatedBlob], activeFileObj.name, { type: updatedBlob.type }));
        const result = await driveApi('update', {
            fileId: activeFileObj.id,
            fileName: activeFileObj.name,
            mimeType: updatedBlob.type,
            category: 'excels',
            file: base64
        });
        activeFileObj.id = result.file.id;
        await fetchExcelFilesList();
        toast('Perubahan database jemaat berhasil disimpan ke Google Drive!', 'success');
    } catch (err) {
        console.error(err);
        toast(err.message || 'Gagal menyimpan database ke Google Drive.', 'error');
    } finally {
        saveBtn.disabled = false;
        saveBtn.innerHTML = `<i data-lucide="save" class="w-3.5 h-3.5 text-white"></i><span>Simpan Terkini</span>`;
        lucide.createIcons();
    }
}


function downloadActiveExcelBackup() {
    if (!activeFileObj) return;
    const downloadBtn = document.getElementById('download-excel-backup-btn');
    setButtonState(downloadBtn, { loading: true, text: 'Menyiapkan...' });
    showGlobalLoader('Menyiapkan unduhan...');
    setTimeout(async () => {
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
        } finally {
            hideGlobalLoader();
            setButtonState(downloadBtn, { loading: false, text: '<i data-lucide="download" class="w-3.5 h-3.5 text-slate-500"></i><span>Backup Lokal</span>' });
        }
    }, 50);
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
    const grid = document.getElementById('gallery-items-grid');
    const alertBox = document.getElementById('gallery-empty-alert');
    grid.innerHTML = '';
    try {
        const result = await driveApi('list', { category: 'photos', includeContent: true });
        const photoList = (result.files || []).map(pic => ({
            ...pic,
            blob: pic.base64 ? base64ToBlob(pic.base64, pic.mimeType) : new Blob()
        }));
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
            let imageSrcUrl = URL.createObjectURL(pic.blob);
            cell.innerHTML = `
                <div class="h-40 overflow-hidden bg-slate-50 relative">
                    <img src="${imageSrcUrl}" alt="Arsip" class="w-full h-full object-cover group-hover:scale-105 transition duration-300" />
                    <div class="absolute inset-0 bg-black/40 opacity-0 group-hover:opacity-100 transition flex items-center justify-center gap-2">
                        <button onclick="previewPhotoLightbox('${pic.id}', '${pic.name}', '${imageSrcUrl}')" class="p-2 bg-white rounded-lg text-slate-700 hover:text-black hover:scale-105 transition cursor-pointer shadow"><i data-lucide="eye" class="w-4 h-4"></i></button>
                        <button onclick="triggerPhotoLocalDownload('${pic.id}', '${pic.name}')" class="p-2 bg-white rounded-lg text-slate-705 hover:text-black hover:scale-105 transition cursor-pointer shadow"><i data-lucide="download" class="w-4 h-4"></i></button>
                    </div>
                </div>
                <div class="p-3.5 flex items-center justify-between gap-2 border-t border-slate-100 bg-white">
                    <div class="truncate"><h4 class="text-xs font-bold text-slate-800 truncate leading-snug">${pic.name}</h4><p class="text-[9px] text-slate-400 font-mono mt-0.5">${(parseInt(pic.size, 10)/1024).toFixed(1)} KB</p></div>
                    <button onclick="triggerRemovePhotoImage('${pic.id}')" class="p-1.5 rounded-lg border border-slate-100 hover:border-red-100 hover:bg-red-50 text-slate-450 hover:text-red-500 transition cursor-pointer" title="Hapus Gambar"><i data-lucide="trash-2" class="w-3.5 h-3.5"></i></button>
                </div>`;
            grid.appendChild(cell);
        });
        lucide.createIcons();
    } catch (err) {
        console.error(err);
        toast(err.message || 'Gagal memuat foto dari Google Drive.', 'error');
    }
}


function getPhotoOrderValue(photo) {
    return typeof photo.displayOrder === 'number' ? photo.displayOrder : new Date(photo.createdTime).getTime();
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
    // If opened from Youth tab, remember selected youth category to preselect in modal
    try {
        if (activeTab === 'youth') {
            const sel = document.getElementById('youth-upload-category');
            window.pendingUploadCategoryFromUI = sel ? sel.value : null;
            window.pendingUploadFromYouth = true;
        } else {
            window.pendingUploadCategoryFromUI = null;
            window.pendingUploadFromYouth = false;
        }
    } catch (e) { window.pendingUploadCategoryFromUI = null; window.pendingUploadFromYouth = false; }

    const uploadInputId = activeTab === 'youth' ? 'youth-doc-hidden-file-input' : 'doc-hidden-file-input';
    const uploadInput = document.getElementById(uploadInputId);
    if (uploadInput) uploadInput.click();
}

async function handleDocLocalUpload(event) {
    const files = Array.from(event.target.files || []);
    if (!files.length) return;

    // Simpan semua file yang tertunda dan tampilkan modal kategori sekali
    pendingDocFiles = files;
    const pre = window.pendingUploadCategoryFromUI || null;
    loadCategoryModal(pre);
    // clear pending UI hint
    window.pendingUploadCategoryFromUI = null;

    event.target.value = '';
}


function loadCategoryModal(preselectCategory) {
    // Muat kategori custom dari localStorage dan hapus nilai tidak valid
    const stored = localStorage.getItem('doc_custom_categories');
    const loaded = stored ? JSON.parse(stored) : [];
    docCustomCategories = Array.isArray(loaded)
        ? loaded.filter((cat) => typeof cat === 'string' && cat.trim() && !['null', 'undefined'].includes(cat.trim().toLowerCase()))
        : [];
    localStorage.setItem('doc_custom_categories', JSON.stringify(docCustomCategories));
    
    // Update filter buttons dengan kategori custom
    updateDocCategoryFilters();
    
    // Reset pilihan kategori di modal
    selectedDocCategory = preselectCategory || null;
    document.querySelectorAll('.category-option').forEach(btn => {
        btn.classList.remove('bg-blue-100', 'border-blue-400');
    });

    // Reset tombol aksi bila bukan mode ubah kategori
    const actionBtn = document.getElementById('doc-category-action-btn');
    if (actionBtn && !window.isCategoryChangeMode) {
        actionBtn.disabled = false;
        actionBtn.textContent = 'Lanjutkan Upload';
        actionBtn.onclick = proceedUploadDocWithCategory;
    }
    
    // Update tampilan custom categories
    const customContainer = document.getElementById('doc-custom-categories');
    customContainer.innerHTML = '';
    if (docCustomCategories.length === 0) {
        const emptyMessage = document.createElement('div');
        emptyMessage.className = 'text-[10px] text-slate-400 px-3 py-2 rounded-xl border border-dashed border-slate-200 bg-slate-50';
        emptyMessage.innerText = 'Belum ada kategori custom. Tambahkan kategori baru untuk melihat opsi hapus.';
        customContainer.appendChild(emptyMessage);
    }
    docCustomCategories.forEach((cat) => {
        const item = document.createElement('div');
        item.className = 'flex items-center justify-between gap-2';

        const btn = document.createElement('button');
        btn.onclick = () => selectCategory(cat);
        btn.className = 'category-option w-full text-left p-3 rounded-xl border border-slate-200 hover:border-blue-300 hover:bg-blue-50 transition flex items-center gap-3';
        btn.setAttribute('data-value', cat);
        btn.innerHTML = `
            <i data-lucide="tag" class="w-4 h-4 text-slate-500"></i>
            <div>
                <p class="text-xs font-semibold text-slate-900">${cat}</p>
                <p class="text-[9px] text-slate-500">Kategori custom</p>
            </div>
        `;

        const removeBtn = document.createElement('button');
        removeBtn.type = 'button';
        removeBtn.className = 'inline-flex items-center justify-center gap-1 px-3 h-8 rounded-xl border border-slate-200 bg-white text-slate-500 hover:bg-red-50 hover:text-red-600 transition text-[10px] font-semibold';
        removeBtn.innerHTML = `<i data-lucide="trash-2" class="w-4 h-4"></i><span>Hapus</span>`;
        removeBtn.onclick = (event) => {
            event.stopPropagation();
            deleteCustomCategory(cat);
        };

        item.appendChild(btn);
        item.appendChild(removeBtn);
        customContainer.appendChild(item);
    });

    // If a preselectCategory is provided and it's not in docCustomCategories or builtins, show a temporary option
    if (preselectCategory) {
        const existsCustom = docCustomCategories.some(c => String(c).toLowerCase() === String(preselectCategory).toLowerCase());
        const builtinMatch = Array.from(document.querySelectorAll('.category-option')).some(b => b.getAttribute('data-value') === preselectCategory);
        if (!existsCustom && !builtinMatch) {
            const btn = document.createElement('button');
            btn.onclick = () => selectCategory(preselectCategory);
            btn.className = 'category-option w-full text-left p-3 rounded-xl border border-slate-200 hover:border-blue-300 hover:bg-blue-50 transition flex items-center gap-3 mb-2';
            btn.setAttribute('data-value', preselectCategory);
            btn.innerHTML = `
                <i data-lucide="tag" class="w-4 h-4 text-slate-500"></i>
                <div>
                    <p class="text-xs font-semibold text-slate-900">${escapeHtml(preselectCategory)}</p>
                    <p class="text-[9px] text-slate-500">Kategori dari Youth</p>
                </div>`;
            customContainer.insertBefore(btn, customContainer.firstChild);
        }
    }
    
    lucide.createIcons();
    // Highlight preselected category if any
    if (selectedDocCategory) selectCategory(selectedDocCategory);
    document.getElementById('doc-category-modal').classList.remove('hidden');
}


function closeCategoryModal() {
    document.getElementById('doc-category-modal').classList.add('hidden');
    pendingDocFiles = [];
    selectedDocCategory = null;
    uploadInProgress = false;
    window.pendingCategoryChangeDocId = null;
    window.pendingCategoryChangeDocName = null;
    window.isCategoryChangeMode = false;
    // reset any pending youth upload flag
    window.pendingUploadFromYouth = false;
    
    // Reset tombol aksi ke mode upload
    const actionBtn = document.getElementById('doc-category-action-btn');
    if (actionBtn) {
        actionBtn.disabled = false;
        actionBtn.textContent = 'Lanjutkan Upload';
        actionBtn.onclick = proceedUploadDocWithCategory;
    }

    // Hide and clear progress area if present
    const progressArea = document.getElementById('doc-upload-progress-area');
    const progressList = document.getElementById('doc-upload-progress-list');
    if (progressArea) progressArea.classList.add('hidden');
    if (progressList) progressList.innerHTML = '';
}


function selectCategory(categoryValue) {
    selectedDocCategory = categoryValue;
    // Highlight pilihan
    document.querySelectorAll('.category-option').forEach(btn => {
        btn.classList.remove('bg-blue-100', 'border-blue-400');
        const btnValue = btn.getAttribute('data-value');
        if (btnValue === categoryValue) {
            btn.classList.add('bg-blue-100', 'border-blue-400');
        }
    });
    const actionBtn = document.getElementById('doc-category-action-btn');
    if (actionBtn) {
        actionBtn.disabled = false;
    }
}


function toggleAddCustomCategory() {
    const inputArea = document.getElementById('doc-custom-input-area');
    const currentShow = !inputArea.classList.contains('hidden');
    if (currentShow) {
        inputArea.classList.add('hidden');
        document.getElementById('doc-custom-category-input').value = '';
    } else {
        inputArea.classList.remove('hidden');
        document.getElementById('doc-custom-category-input').focus();
    }
}


async function saveCustomCategory() {
    const input = document.getElementById('doc-custom-category-input');
    const newCategory = input.value.trim();
    
    if (!newCategory) {
        toast('Nama kategori tidak boleh kosong.', 'error');
        return;
    }
    
    const normalizedNewCategory = newCategory.trim().toLowerCase();
    if (docCustomCategories.some((cat) => cat.trim().toLowerCase() === normalizedNewCategory)) {
        toast('Kategori sudah ada.', 'error');
        return;
    }
    
    // Simpan dengan nama asli (tanpa normalisasi), normalisasi terjadi saat upload
    docCustomCategories.push(newCategory);
    localStorage.setItem('doc_custom_categories', JSON.stringify(docCustomCategories));
    
    // Update filter buttons dengan kategori baru
    updateDocCategoryFilters();
    
    // Reload modal dan pilih kategori baru
    input.value = '';
    updateDocCategoryFilters();
    loadCategoryModal();
    selectCategory(newCategory);
    toast(`Kategori "${newCategory}" berhasil dibuat.`, 'success');
}


function deleteCustomCategory(categoryName) {
    docCustomCategories = docCustomCategories.filter((cat) => cat !== categoryName);
    localStorage.setItem('doc_custom_categories', JSON.stringify(docCustomCategories));
    updateDocCategoryFilters();
    loadCategoryModal();

    const normalized = categoryName.toLowerCase().replace(/\s+/g, '_');
    if (selectedDocFilterCategory === normalized) {
        selectedDocFilterCategory = 'semua';
        filterDocsByCategory('semua');
    } else {
        fetchDocsList();
    }

    if (selectedDocCategory === categoryName) {
        selectedDocCategory = null;
    }
}


function updateDocCategoryFilters() {
    const filterContainerId = activeTab === 'youth' ? 'youth-docs-custom-category-filters' : 'docs-custom-category-filters';
    const filterContainer = document.getElementById(filterContainerId);
    if (!filterContainer) return;
    filterContainer.innerHTML = '';
    
    docCustomCategories.forEach((cat) => {
        const normalized = cat.toLowerCase().replace(/\s+/g, '_');
        const btn = document.createElement('button');
        btn.onclick = () => filterDocsByCategory(normalized);
        btn.className = 'doc-filter-btn text-[10px] font-semibold px-3 py-1.5 rounded-lg border border-slate-200 bg-slate-50 text-slate-600 hover:bg-blue-50 hover:border-blue-200 transition';
        btn.setAttribute('data-category', normalized);
        btn.textContent = cat;
        filterContainer.appendChild(btn);
    });
}


async function proceedUploadDocWithCategory() {
    if (uploadInProgress) return;

    if (!selectedDocCategory || ['null', 'undefined'].includes(String(selectedDocCategory).toLowerCase())) {
        toast('Silakan pilih kategori dokumen yang valid terlebih dahulu.', 'error');
        return;
    }

    if (!pendingDocFiles || pendingDocFiles.length === 0) {
        closeCategoryModal();
        return;
    }

    const filesToUpload = pendingDocFiles.slice();
    const actionBtn = document.getElementById('doc-category-action-btn');
    if (actionBtn) actionBtn.disabled = true;

    // Prepare progress UI
    const progressArea = document.getElementById('doc-upload-progress-area');
    const progressList = document.getElementById('doc-upload-progress-list');
    if (progressList) progressList.innerHTML = '';
    if (progressArea) progressArea.classList.remove('hidden');

    uploadInProgress = true;

    try {
        let normalizedCategory = String(selectedDocCategory)
            .trim()
            .toLowerCase()
            .replace(/\s+/g, '_');

        if (activeTab === 'youth' || window.pendingUploadFromYouth) {
            if (!normalizedCategory.startsWith('youth_')) {
                normalizedCategory = `youth_${normalizedCategory}`;
            }
        }

        for (const file of filesToUpload) {
            // create progress entry
            const itemEl = document.createElement('div');
            itemEl.className = 'p-2 bg-slate-50 border border-slate-100 rounded-xl flex items-center gap-3';
            itemEl.innerHTML = `<div class="flex-1 min-w-0"><div class="text-xs font-medium text-slate-900 truncate">${escapeHtml(file.name)}</div><div class="w-full bg-slate-200 rounded-full h-2 mt-2 overflow-hidden"><div class="doc-progress-bar bg-blue-600 h-2 w-0"></div></div></div><div class="ml-3 text-[10px] text-slate-500 upload-status">0%</div>`;
            if (progressList) progressList.appendChild(itemEl);

            const progressBar = itemEl.querySelector('.doc-progress-bar');
            const statusText = itemEl.querySelector('.upload-status');

            try {
                await uploadFileToDrive(file, normalizedCategory, {}, (loaded, total) => {
                    const percent = total ? Math.round((loaded / total) * 100) : 0;
                    if (progressBar) progressBar.style.width = percent + '%';
                    if (statusText) statusText.innerText = percent + '%';
                });
                if (progressBar) progressBar.style.width = '100%';
                if (statusText) statusText.innerText = 'Selesai';
                itemEl.classList.add('opacity-100');
            } catch (errFile) {
                console.error('Upload gagal untuk file', file.name, errFile);
                if (statusText) statusText.innerText = 'Gagal';
                if (progressBar) progressBar.classList.remove('bg-blue-600');
                if (progressBar) progressBar.classList.add('bg-red-600');
            }
        }

        await refreshDriveLists();
        closeCategoryModal();

        const label = (activeTab === 'youth' || window.pendingUploadFromYouth) ? 'Youth Dokumen' : 'Dokumen';
        const message = filesToUpload.length === 1
            ? `Dokumen "${filesToUpload[0].name}" berhasil diunggah ke ${label}!`
            : `${filesToUpload.length} dokumen berhasil diunggah ke ${label}!`;

        toast(message, 'success');
    } catch (err) {
        console.error(err);
        toast(err.message || 'Gagal upload dokumen ke Google Drive.', 'error');
    } finally {
        uploadInProgress = false;
        if (actionBtn) actionBtn.disabled = false;
    }

    window.pendingUploadFromYouth = false;
}


async function fetchDocsList() {
    const list = document.getElementById(activeTab === 'youth' ? 'youth-docs-files-list' : 'docs-files-list');
    const fetchId = ++currentDocsFetchId;

    if (!list) return;
    list.innerHTML = `<div class="text-[10px] text-slate-400 text-center py-6 border border-dashed border-slate-200 rounded-2xl">Memuat dokumen...</div>`;

    try {
        const allFiles = await getCachedDriveFiles();

        if (fetchId !== currentDocsFetchId) return;

        let docs = allFiles.filter(f => {
            if (!f.category) return false;
            if (f.category === 'photos' || f.category === 'excels') return false;

            const isYouthDoc = String(f.category).startsWith('youth_');

            if (activeTab === 'youth') {
                return isYouthDoc;
            }

            return !isYouthDoc;
        });

        const seen = new Set();
        docs = docs.filter((doc) => {
            if (seen.has(doc.id)) return false;
            seen.add(doc.id);
            return true;
        });

        if (selectedDocFilterCategory !== 'semua') {
            if (activeTab === 'youth') {
                docs = docs.filter(f => f.category === `youth_${selectedDocFilterCategory}`);
            } else {
                docs = docs.filter(f => f.category === selectedDocFilterCategory);
            }
        }

        const query = selectedDocSearchQuery.trim().toLowerCase();
        if (query) {
            docs = docs.filter(f => (f.name || '').toLowerCase().includes(query));
        }

        if (docs.length === 0) {
            const label = activeTab === 'youth' ? 'Youth Dokumen' : 'Dokumen Gereja';
            list.innerHTML = `<div class="text-[10px] text-slate-400 text-center py-6 border border-dashed border-slate-200 rounded-2xl">Belum ada ${label}.</div>`;
            if (activeTab === 'youth') renderExternalYouthDocs();
            return;
        }

        docs.sort((a, b) => new Date(b.createdTime).getTime() - new Date(a.createdTime).getTime());

        docs.forEach((doc) => {
            let rawCategory = doc.category || '';
            let isYouthDoc = false;

            if (rawCategory.startsWith('youth_')) {
                isYouthDoc = true;
                rawCategory = rawCategory.replace(/^youth_/, '');
            }

            const categoryDisplay = getCategoryDisplay(rawCategory);

            const item = document.createElement('div');
            item.className = 'flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 p-4 rounded-2xl border border-slate-200 bg-slate-50';

            item.innerHTML = `
                <div class="min-w-0 flex-1">
                    <div class="flex items-center gap-2 mb-1">
                        <p class="text-xs font-bold text-slate-900 truncate">${escapeHtml(doc.name)}</p>
                        <span class="inline-block px-2 py-1 bg-blue-100 text-blue-700 text-[8px] font-semibold rounded-lg whitespace-nowrap">${categoryDisplay}</span>
                        ${isYouthDoc ? '<span class="inline-block px-2 py-1 bg-purple-100 text-purple-700 text-[8px] font-semibold rounded-lg whitespace-nowrap">Youth</span>' : ''}
                    </div>
                    <p class="text-[9px] text-slate-500">${escapeHtml(doc.mimeType || 'Dokumen')} • ${(parseInt(doc.size, 10) / 1024).toFixed(1)} KB</p>
                </div>

                <div class="flex items-center gap-2 flex-wrap sm:flex-nowrap">
                    <button type="button" class="doc-preview-btn px-3 py-2 bg-slate-200 hover:bg-slate-300 text-slate-900 rounded-xl text-[10px] font-semibold transition">Pratinjau</button>
                    <button type="button" class="doc-change-category-btn px-3 py-2 bg-yellow-50 hover:bg-yellow-100 text-yellow-700 rounded-xl text-[10px] font-semibold transition">Ubah Kategori</button>
                    <button type="button" class="doc-download-btn px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-[10px] font-semibold transition">Unduh</button>
                    <button type="button" class="doc-remove-btn px-3 py-2 bg-red-50 hover:bg-red-100 text-red-600 rounded-xl text-[10px] font-semibold transition">Hapus</button>
                </div>
            `;

            item.querySelector('.doc-preview-btn').addEventListener('click', () => previewDocFile(doc.id));
            item.querySelector('.doc-change-category-btn').addEventListener('click', () => triggerChangeDocCategory(doc.id, doc.name));
            item.querySelector('.doc-download-btn').addEventListener('click', () => triggerDocLocalDownload(doc.id, doc.name));
            item.querySelector('.doc-remove-btn').addEventListener('click', () => triggerRemoveDocFile(doc.id));

            list.appendChild(item);
        });

        lucide.createIcons();
        if (activeTab === 'youth') renderExternalYouthDocs();
    } catch (err) {
        console.error(err);
        toast(err.message || 'Gagal memuat dokumen dari Google Drive.', 'error');
    }
}

async function refreshDriveLists() {
    invalidateDriveFileCache();
    await Promise.all([
        loadDashboardStats(),
        fetchExcelFilesList(),
        fetchDocsList()
    ]);
}


function filterDocsByCategory(category) {
    selectedDocFilterCategory = category;
    
    // Update button styling
    document.querySelectorAll('.doc-filter-btn').forEach(btn => {
        btn.classList.remove('bg-blue-600', 'text-white', 'border-blue-600');
        btn.classList.add('bg-slate-50', 'text-slate-600', 'border-slate-200');
        if (btn.getAttribute('data-category') === category) {
            btn.classList.remove('bg-slate-50', 'text-slate-600', 'border-slate-200');
            btn.classList.add('bg-blue-600', 'text-white', 'border-blue-600');
        }
    });
    
    // Reload dokumen dengan filter baru
    fetchDocsList();
}

function setDocSearchQuery(value) {
    selectedDocSearchQuery = String(value || '');
    fetchDocsList();
}

function saveExternalDocLinks() {
    localStorage.setItem('external_doc_links', JSON.stringify(externalDocLinks));
}

function getLinkSource(url) {
    if (!url) return 'Link Eksternal';
    const lower = url.toLowerCase();
    if (lower.includes('drive.google.com') || lower.includes('docs.google.com') || lower.includes('spreadsheets.google.com')) return 'Google Drive';
    if (lower.includes('speedseet') || lower.includes('speedsheet')) return 'SpeedSheet';
    return 'Link Eksternal';
}

function normalizeExternalLink(url) {
    return String(url || '').trim();
}

function getGoogleDriveFileId(url) {
    const patterns = [
        /drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/,
        /docs\.google\.com\/document\/d\/([a-zA-Z0-9_-]+)/,
        /docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/,
        /drive\.google\.com\/open\?id=([a-zA-Z0-9_-]+)/
    ];
    for (const pattern of patterns) {
        const match = url.match(pattern);
        if (match && match[1]) return match[1];
    }
    return null;
}

function getGoogleDrivePreviewUrl(url) {
    const id = getGoogleDriveFileId(url);
    if (!id) return url;
    return `https://drive.google.com/file/d/${id}/preview`;
}

function getGoogleDriveDownloadUrl(url) {
    const id = getGoogleDriveFileId(url);
    if (!id) return url;
    return `https://drive.google.com/uc?export=download&id=${id}`;
}

function addExternalDocLink(rawUrl) {
    const url = normalizeExternalLink(rawUrl);
    if (!url) {
        toast('Silakan masukkan link yang valid.', 'error');
        return;
    }

    const existing = externalDocLinks.some(link => link.url === url);
    if (existing) {
        toast('Link sudah terdaftar.', 'error');
        return;
    }

    externalDocLinks.push({
        url,
        title: url,
        source: getLinkSource(url),
        createdAt: new Date().toISOString()
    });
    saveExternalDocLinks();
    fetchDocsList();
    toast('Link dokumen eksternal berhasil ditambahkan.', 'success');
}

function handleAddExternalDocLink() {
    const input = document.getElementById('external-doc-link-input');
    if (!input) return;
    addExternalDocLink(input.value);
    input.value = '';
}

function previewExternalLink(link) {
    const previewContent = document.getElementById('doc-preview-content');
    const titleEl = document.getElementById('doc-preview-title');
    const infoEl = document.getElementById('doc-preview-info');
    const openBtn = document.getElementById('doc-preview-open-btn');
    try {
        if (activeDocPreviewUrl) URL.revokeObjectURL(activeDocPreviewUrl);
        const previewUrl = link.source === 'Google Drive' ? getGoogleDrivePreviewUrl(link.url) : link.url;
        activeDocPreviewUrl = previewUrl;
        titleEl.innerText = link.title || link.url;
        infoEl.innerText = `${link.source} • ${link.url}`;
        openBtn.classList.remove('hidden');
        openBtn.onclick = openDocInNewTab;
        previewContent.innerHTML = `<iframe src="${previewUrl}" class="w-full h-full border-0" title="Preview Link Dokumen"></iframe>`;
        document.getElementById('doc-preview-modal').classList.remove('hidden');
    } catch (err) {
        console.error(err);
        previewContent.innerHTML = `<div class="p-6 text-center text-slate-500 text-sm">Tidak dapat memuat preview untuk link ini. Silakan buka di tab baru.</div>`;
        titleEl.innerText = link.title || 'Preview Link Eksternal';
        infoEl.innerText = `${link.source} • ${link.url}`;
        openBtn.classList.remove('hidden');
        openBtn.onclick = openDocInNewTab;
        document.getElementById('doc-preview-modal').classList.remove('hidden');
    }
}

function downloadExternalLink(link) {
    const downloadUrl = link.source === 'Google Drive' ? getGoogleDriveDownloadUrl(link.url) : link.url;
    const anchor = document.createElement('a');
    anchor.href = downloadUrl;
    anchor.target = '_blank';
    anchor.rel = 'noreferrer noopener';
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
}

function removeExternalDocLink(url) {
    const idx = externalDocLinks.findIndex(l => l.url === url);
    if (idx === -1) {
        toast('Link eksternal tidak ditemukan.', 'error');
        return;
    }
    externalDocLinks.splice(idx, 1);
    saveExternalDocLinks();
    fetchDocsList();
    toast('Link eksternal berhasil dihapus.', 'success');
}


// ===== EXTERNAL YOUTH DOCUMENT LINKS =====
function saveExternalYouthDocLinks() {
    localStorage.setItem('external_youth_doc_links', JSON.stringify(externalYouthDocLinks));
}

function renderExternalYouthDocs() {
    const list = document.getElementById('youth-external-docs-list');
    if (!list) return;
    
    list.innerHTML = '';
    if (externalYouthDocLinks.length === 0) return;
    
    externalYouthDocLinks.forEach((link) => {
        const item = document.createElement('div');
        item.className = 'flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 p-4 rounded-2xl border border-slate-200 bg-slate-50';
        
        item.innerHTML = `
            <div class="min-w-0 flex-1">
                <div class="flex items-center gap-2 mb-1">
                    <p class="text-xs font-bold text-slate-900 truncate">${escapeHtml(link.title || link.url)}</p>
                    <span class="inline-block px-2 py-1 bg-blue-100 text-blue-700 text-[8px] font-semibold rounded-lg whitespace-nowrap">${link.source || 'External'}</span>
                </div>
                <p class="text-[9px] text-slate-500 truncate">${escapeHtml(link.url)}</p>
            </div>
            
            <div class="flex items-center gap-2 flex-wrap sm:flex-nowrap">
                <button type="button" class="youth-external-preview-btn px-3 py-2 bg-slate-200 hover:bg-slate-300 text-slate-900 rounded-xl text-[10px] font-semibold transition">Pratinjau</button>
                <button type="button" class="youth-external-download-btn px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-[10px] font-semibold transition">Unduh</button>
                <button type="button" class="youth-external-remove-btn px-3 py-2 bg-red-50 hover:bg-red-100 text-red-600 rounded-xl text-[10px] font-semibold transition">Hapus</button>
            </div>
        `;
        
        item.querySelector('.youth-external-preview-btn').addEventListener('click', () => previewExternalLink(link));
        item.querySelector('.youth-external-download-btn').addEventListener('click', () => downloadExternalLink(link));
        item.querySelector('.youth-external-remove-btn').addEventListener('click', () => removeExternalYouthDocLink(link.url));
        
        list.appendChild(item);
    });
    
    lucide.createIcons();
}

function handleAddExternalYouthDocLink() {
    const input = document.getElementById('youth-external-doc-link-input');
    if (!input) return;
    
    const url = normalizeExternalLink(input.value);
    if (!url) {
        toast('Silakan masukkan link yang valid.', 'error');
        return;
    }
    
    const existing = externalYouthDocLinks.some(link => link.url === url);
    if (existing) {
        toast('Link sudah terdaftar.', 'error');
        return;
    }
    
    externalYouthDocLinks.push({
        url,
        title: url,
        source: getLinkSource(url),
        createdAt: new Date().toISOString()
    });
    
    saveExternalYouthDocLinks();
    renderExternalYouthDocs();
    input.value = '';
    toast('Link dokumen eksternal berhasil ditambahkan.', 'success');
}

function removeExternalYouthDocLink(url) {
    const idx = externalYouthDocLinks.findIndex(l => l.url === url);
    if (idx === -1) {
        toast('Link eksternal tidak ditemukan.', 'error');
        return;
    }
    externalYouthDocLinks.splice(idx, 1);
    saveExternalYouthDocLinks();
    renderExternalYouthDocs();
    toast('Link eksternal berhasil dihapus.', 'success');
}
// ==========================================


function initializeDocCategoryFilter() {
    // Set default filter to "semua" dan highlight button
    filterDocsByCategory('semua');
}


function escapeHtml(value) {
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function getCategoryDisplay(category) {
    if (!category) return 'Dokumen';
    const displayMap = {
        'dokumen': 'Dokumen',
        'warta_gereja': 'Warta Gereja',
        'surat_menyurat': 'Surat Menyurat'
    };
    
    // Jika kategori adalah preset, tampilkan display name
    if (displayMap[category]) return displayMap[category];
    
    // Jika custom category, coba match dengan yang disimpan di docCustomCategories
    const normalized = (str) => str.toLowerCase().replace(/\s+/g, '_');
    for (let customCat of docCustomCategories) {
        if (normalized(customCat) === category) {
            return customCat;
        }
    }
    
    // Fallback: format dengan capitalize
    return category.replace(/_/g, ' ').split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
}


async function triggerRemoveDocFile(docId) {
    const confirmed = window.confirm("Apakah Anda yakin ingin menghapus dokumen ini dari Google Drive?");
    if (confirmed) {
        try {
            await driveApi('delete', { fileId: docId });
            invalidateDriveFileCache();
            await Promise.all([fetchDocsList(), loadDashboardStats()]);
            toast('Dokumen berhasil dihapus dari Google Drive.', 'success');
        } catch (err) {
            console.error(err);
            toast(err.message || 'Gagal menghapus dokumen.', 'error');
        }
    }
}


async function triggerChangeDocCategory(docId, docName) {
    // Simpan dokumen yang akan diubah kategorinya
    window.pendingCategoryChangeDocId = docId;
    window.pendingCategoryChangeDocName = docName;
    window.isCategoryChangeMode = true;
    
    // Bersihkan pilihan kategori sebelumnya
    selectedDocCategory = null;
    
    // Load dan tampilkan modal kategori
    loadCategoryModal();
    
    // Ubah tombol action
    const actionBtn = document.getElementById('doc-category-action-btn');
    if (actionBtn) {
        actionBtn.textContent = 'Ubah Kategori';
        actionBtn.onclick = proceedChangeDocCategory;
    }
}


async function proceedChangeDocCategory() {
    if (!selectedDocCategory) {
        toast('Silakan pilih kategori terlebih dahulu.', 'error');
        return;
    }
    
    const docId = window.pendingCategoryChangeDocId;
    const docName = window.pendingCategoryChangeDocName;
    
    if (!docId) {
        closeCategoryModal();
        return;
    }
    
    try {
        const normalizedCategory = selectedDocCategory.toLowerCase().replace(/\s+/g, '_');
        await driveApi('updateCategory', {
            fileId: docId,
            category: normalizedCategory
        });
        invalidateDriveFileCache();
        await fetchDocsList();
        closeCategoryModal();
        toast(`Kategori dokumen "${docName}" berhasil diubah menjadi "${selectedDocCategory}".`, 'success');
    } catch (err) {
        console.error(err);
        toast(err.message || 'Gagal mengubah kategori dokumen.', 'error');
    }
}


async function triggerDocLocalDownload(docId, filename) {
    try {
        const record = await getDriveFileBlob(docId);
        const downloadUrl = URL.createObjectURL(record.blob);
        const link = document.createElement('a');
        link.href = downloadUrl;
        link.download = filename;
        link.click();
        toast('Mengunduh dokumen dari Google Drive...', 'success');
    } catch (err) {
        console.error(err);
        toast(err.message || 'Gagal mengunduh dokumen.', 'error');
    }
}


async function previewDocFile(docId) {
    const previewContent = document.getElementById('doc-preview-content');
    const titleEl = document.getElementById('doc-preview-title');
    const infoEl = document.getElementById('doc-preview-info');
    const openBtn = document.getElementById('doc-preview-open-btn');
    showGlobalLoader('Memuat pratinjau dokumen...');
    await waitForNextFrame();
    previewContent.innerHTML = `<div class="p-6 text-center text-slate-500 text-sm">Memuat pratinjau dokumen...</div>`;
    try {
        const record = await getDriveFileBlob(docId);
        if (activeDocPreviewUrl) URL.revokeObjectURL(activeDocPreviewUrl);
        const blobUrl = URL.createObjectURL(record.blob);
        activeDocPreviewUrl = blobUrl;
        titleEl.innerText = record.name;
        infoEl.innerText = `${record.mimeType || 'Dokumen'} • ${(parseInt(record.size, 10) / 1024).toFixed(1)} KB`;
        openBtn.classList.remove('hidden');
        openBtn.onclick = openDocInNewTab;
        const previewableImage = record.mimeType?.startsWith('image/');
        const previewablePdf = record.mimeType === 'application/pdf';
        const previewableHtml = record.mimeType === 'text/html';
        if (previewableImage) previewContent.innerHTML = `<div class="flex items-center justify-center h-full p-4"><img src="${blobUrl}" alt="${record.name}" class="max-w-full max-h-[76vh] object-contain rounded-2xl" /></div>`;
        else if (previewablePdf || previewableHtml) previewContent.innerHTML = `<iframe src="${blobUrl}" class="w-full h-full border-0" title="Preview Dokumen"></iframe>`;
        else previewContent.innerHTML = `<div class="p-6 text-center text-slate-500 text-sm space-y-3"><p class="font-semibold text-slate-900">Pratinjau tidak tersedia untuk jenis file ini.</p><p>Silakan buka di tab baru atau unduh file untuk melihatnya secara lengkap.</p></div>`;
        document.getElementById('doc-preview-modal').classList.remove('hidden');
    } catch (err) {
        console.error(err);
        previewContent.innerHTML = `<div class="p-6 text-center text-slate-500 text-sm">Dokumen tidak ditemukan.</div>`;
        titleEl.innerText = 'Preview Dokumen';
        infoEl.innerText = 'Tidak ada konten yang dapat ditampilkan.';
        openBtn.classList.add('hidden');
        document.getElementById('doc-preview-modal').classList.remove('hidden');
    } finally {
        hideGlobalLoader();
    }
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
    try {
        await uploadFileToDrive(file, 'photos', { displayOrder: Date.now() });
        fetchPhotosList();
        loadDashboardStats();
        toast(`Foto "${file.name}" berhasil diunggah ke Google Drive!`, 'success');
    } catch (err) {
        console.error(err);
        toast(err.message || 'Gagal upload foto ke Google Drive.', 'error');
    } finally {
        iconUploader.innerHTML = `<i data-lucide="plus" class="w-5 h-5 text-blue-500"></i>`;
        lucide.createIcons();
        event.target.value = '';
    }
}


async function triggerRemovePhotoImage(photoId) {
    const confirmed = window.confirm("Apakah Anda yakin ingin menghapus foto dokumentasi ini dari Google Drive?");
    if (confirmed) {
        try {
            await driveApi('delete', { fileId: photoId });
            invalidateDriveFileCache();
            await Promise.all([fetchPhotosList(), loadDashboardStats()]);
            toast('Foto berhasil dihapus dari Google Drive.', 'success');
        } catch (err) {
            console.error(err);
            toast(err.message || 'Gagal menghapus foto.', 'error');
        }
    }
}


async function triggerPhotoLocalDownload(photoId, filename) {
    showGlobalLoader('Menyiapkan unduhan foto...');
    await waitForNextFrame();
    try {
        const record = await getDriveFileBlob(photoId);
        const downloadUrl = URL.createObjectURL(record.blob);
        const link = document.createElement('a');
        link.href = downloadUrl;
        link.download = filename;
        link.click();
        toast('Mengunduh foto dari Google Drive...', 'success');
    } catch (err) {
        console.error(err);
        toast(err.message || 'Gagal mengunduh foto.', 'error');
    } finally {
        hideGlobalLoader();
    }
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
