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
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result).split(',')[1]);
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}

function base64ToBlob(base64, mimeType) {
    const bytes = atob(base64);
    const array = new Uint8Array(bytes.length);
    for (let i = 0; i < bytes.length; i++) array[i] = bytes.charCodeAt(i);
    return new Blob([array], { type: mimeType || 'application/octet-stream' });
}

async function uploadFileToDrive(file, category, extra = {}) {
    const base64 = await fileToBase64(file);
    const result = await driveApi('upload', {
        fileName: file.name,
        mimeType: file.type || 'application/octet-stream',
        category,
        displayOrder: extra.displayOrder || Date.now(),
        file: base64
    });
    return result.file;
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
    try {
        const result = await driveApi('list');
        const files = result.files || [];
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
        const result = await driveApi('list', { category: 'excels' });
        const items = result.files || [];
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
        event.target.value = '';
    }
}


async function selectExcelFileFromList(fileId) {
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
        const base64 = await fileToBase64(new File([updatedBlob], activeFileObj.name, { type: updatedBlob.type }));
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
                cell.addEventListener('dragend', () => { currentDraggedPhotoId = null; cell.classList.remove('opacity-50'); });
                cell.addEventListener('dragover', (event) => { event.preventDefault(); event.dataTransfer.dropEffect = 'move'; cell.classList.add('drag-over-target'); });
                cell.addEventListener('dragleave', () => cell.classList.remove('drag-over-target'));
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

function reorderPhotoList(photoList, draggedId, targetId) {
    const draggedIndex = photoList.findIndex((item) => item.id === draggedId);
    const targetIndex = photoList.findIndex((item) => item.id === targetId);
    if (draggedIndex === -1 || targetIndex === -1 || draggedIndex === targetIndex) return photoList;
    const [movedItem] = photoList.splice(draggedIndex, 1);
    photoList.splice(targetIndex, 0, movedItem);
    return photoList;
}

async function savePhotoOrderSequence(photoList) {
    const count = photoList.length;
    const orders = photoList.map((item, index) => ({ id: item.id, displayOrder: count - index }));
    await driveApi('reorder', { orders });
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
    try {
        await uploadFileToDrive(file, 'docs');
        fetchDocsList();
        loadDashboardStats();
        toast(`Dokumen "${file.name}" berhasil diunggah ke Google Drive!`, 'success');
    } catch (err) {
        console.error(err);
        toast(err.message || 'Gagal upload dokumen ke Google Drive.', 'error');
    } finally {
        event.target.value = '';
    }
}


async function fetchDocsList() {
    const list = document.getElementById('docs-files-list');
    list.innerHTML = '';
    try {
        const result = await driveApi('list', { category: 'docs' });
        const docs = result.files || [];
        if (docs.length === 0) {
            list.innerHTML = `<div class="text-[10px] text-slate-400 text-center py-6 border border-dashed border-slate-200 rounded-2xl">Belum ada dokumen pendukung.</div>`;
            return;
        }
        docs.sort((a,b) => new Date(b.createdTime).getTime() - new Date(a.createdTime).getTime());
        docs.forEach((doc) => {
            const item = document.createElement('div');
            item.className = 'flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 p-4 rounded-2xl border border-slate-200 bg-slate-50';
            item.innerHTML = `
                <div class="min-w-0"><p class="text-xs font-bold text-slate-900 truncate">${doc.name}</p><p class="text-[9px] text-slate-500 mt-1">${doc.mimeType || 'Dokumen'} • ${(parseInt(doc.size, 10) / 1024).toFixed(1)} KB</p></div>
                <div class="flex items-center gap-2">
                    <button onclick="previewDocFile('${doc.id}')" class="px-3 py-2 bg-slate-200 hover:bg-slate-300 text-slate-900 rounded-xl text-[10px] font-semibold transition">Pratinjau</button>
                    <button onclick="triggerDocLocalDownload('${doc.id}', '${doc.name}')" class="px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-[10px] font-semibold transition">Unduh</button>
                    <button onclick="triggerRemoveDocFile('${doc.id}')" class="px-3 py-2 bg-red-50 hover:bg-red-100 text-red-600 rounded-xl text-[10px] font-semibold transition">Hapus</button>
                </div>`;
            list.appendChild(item);
        });
    } catch (err) {
        console.error(err);
        toast(err.message || 'Gagal memuat dokumen dari Google Drive.', 'error');
    }
}


async function triggerRemoveDocFile(docId) {
    const confirmed = window.confirm("Apakah Anda yakin ingin menghapus dokumen ini dari Google Drive?");
    if (confirmed) {
        try {
            await driveApi('delete', { fileId: docId });
            fetchDocsList();
            loadDashboardStats();
            toast('Dokumen berhasil dihapus dari Google Drive.', 'success');
        } catch (err) {
            console.error(err);
            toast(err.message || 'Gagal menghapus dokumen.', 'error');
        }
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
            fetchPhotosList();
            loadDashboardStats();
            toast('Foto berhasil dihapus dari Google Drive.', 'success');
        } catch (err) {
            console.error(err);
            toast(err.message || 'Gagal menghapus foto.', 'error');
        }
    }
}


async function triggerPhotoLocalDownload(photoId, filename) {
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
