const API_URL = "https://script.google.com/macros/s/AKfycbx20gIWQIbNBa2f6o1dr3elRkf1Us8G0Ypb1K8x53fTIwndiSsa6mGuZSWtpfSy8EAk/exec";
const TEACHER_API_URL = 'https://script.google.com/macros/s/AKfycbz6yVDqJbjY_mud1Yg7j5bwL7vQIU-8zHso2o6ycNoyEsD_5STlvxgkJHC-Bl3oJY3oUw/exec';

const TEACHER_LIST = [
    "Soferet Sefatja Domakubun.S.Pd", "Marthina Kormasela.S.PdK", "Oktovina Ratsina.S.Pd",
    "Oktovina L Watmanlussy.S.Pd", "Melianus Ratuanik.S.Pd", "Miryam Yuliana Lololuan.S.PdK",
    "Adelheid Renjaan.S.Pd", "Costantina Ratila.S.Pd", "Enike Benselina Ibur.S.Sos","Wanti Slarmanat.S.Pd",
    "Korlina Wulandari Domakubun.S.Pd", "Sulce Wessy.S.Pd",
    "Edward Domakubun", "Yoseph Yohanes Saiselar"
];
const KEPSEK_INFO = { nama: "SOFERET S DOMAKUBUN,S.Pd", nip: "19680606 199111 1 001" };
let teacherGlobalData = [];
let currentUser = null;
const { jsPDF } = window.jspdf;

// --- CUSTOM ALERT FUNCTION ---
function showAlert(msg, title = "Informasi") {
    const overlay = document.getElementById('custom-alert-overlay');
    const boxMsg = document.getElementById('alert-msg');
    const boxTitle = document.getElementById('alert-title');

    boxMsg.innerText = msg;
    boxTitle.innerText = title;
    overlay.style.display = 'flex';

    void overlay.offsetWidth;
    overlay.classList.add('show');
}

function closeAlert() {
    const overlay = document.getElementById('custom-alert-overlay');
    overlay.classList.remove('show');
    setTimeout(() => {
        overlay.style.display = 'none';
    }, 300);
}

document.addEventListener("DOMContentLoaded", function () {
    checkAuth();
    initTeacherFeature();
});

function handleLogin(e) {
    e.preventDefault();
    const btn = document.getElementById('btn-login'); 
    const oldText = btn.innerHTML;
    btn.disabled = true; 
    btn.innerHTML = "<i class='fas fa-circle-notch fa-spin'></i> MEMPROSES...";
    
    fetch(API_URL, { 
        method: 'POST', 
        body: JSON.stringify({ 
            action: 'login', 
            role: 'admin', 
            username: document.getElementById('u_name').value.trim(), 
            password: document.getElementById('u_pass').value.trim() 
        }) 
    })
    .then(r => r.json())
    .then(res => {
        if (res.status === 'success') { 
            localStorage.setItem('siakad_sess', JSON.stringify(res)); 
            checkAuth(); 
        } else { 
            showAlert(res.message, "Gagal Login"); 
        }
    })
    .catch(err => showAlert("Terjadi kesalahan koneksi.", "Error"))
    .finally(() => { 
        btn.disabled = false; 
        btn.innerHTML = oldText; 
    });
}

function checkAuth() {
    const sess = JSON.parse(localStorage.getItem('siakad_sess'));
    if (sess && sess.role === 'admin') {
        currentUser = sess;
        document.getElementById('login-overlay').style.display = 'none';
        document.getElementById('app-wrapper').style.display = 'flex';
        document.getElementById('lbl-user').innerText = sess.user.nama;
        document.getElementById('user-initial').innerText = sess.user.nama.charAt(0);
        document.getElementById('lbl-role').innerText = 'Administrator';
        document.getElementById('lbl-ta').innerText = sess.config.tahun_ajar;
        document.getElementById('lbl-smt').innerText = sess.config.semester;

        navTo('cetak-guru', document.querySelector('#menu-admin .nav-item')); 
        document.getElementById('cfg-ta').value = sess.config.tahun_ajar; 
        document.getElementById('cfg-smt').value = sess.config.semester;
    } else { 
        document.getElementById('login-overlay').style.display = 'flex'; 
        document.getElementById('app-wrapper').style.display = 'none'; 
    }
}

function navTo(page, el) {
    document.querySelectorAll('.page-section').forEach(p => p.classList.add('hidden'));
    document.getElementById('page-' + page).classList.remove('hidden');
    document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
    if (el) el.classList.add('active');

    let title = "Dashboard";
    if (page === 'cetak-guru') title = "Laporan Absensi Guru";
    if (page === 'admin-config') title = "Pengaturan";

    document.getElementById('page-heading').innerText = title;
    if (window.innerWidth < 900) document.getElementById('sidebar').classList.remove('show');
}

function doLogout() { 
    if (confirm("Keluar dari aplikasi?")) { 
        localStorage.removeItem('siakad_sess'); 
        location.reload(); 
    } 
}

// --- PENGATURAN ADMIN ---
function saveConfig() {
    const btn = event.target; btn.innerText = "Menyimpan...";
    fetch(API_URL, { method: 'POST', body: JSON.stringify({ action: 'update_config', tahun_ajar: document.getElementById('cfg-ta').value, semester: document.getElementById('cfg-smt').value }) })
        .then(r => r.json()).then(res => { showAlert(res.message, "Sukses"); if (res.status == 'success') location.reload(); }).finally(() => btn.innerText = "SIMPAN PERUBAHAN");
}

// =================================================================================
//  LOGIKA BARU: CETAK ABSENSI GURU (INTEGRATED)
// =================================================================================

function initTeacherFeature() {
    const selector = document.getElementById('t_guruSelector');
    if (selector && selector.options.length <= 1) {
        TEACHER_LIST.forEach(name => {
            let opt = document.createElement('option');
            opt.value = name; opt.text = name; selector.appendChild(opt);
        });
    }
}

function switchTeacherTab(tabName) {
    document.querySelectorAll('.teacher-tab-content').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.teacher-tabs .tab-btn').forEach(el => el.classList.remove('active'));
    document.getElementById('tab-teacher-' + tabName).classList.add('active');

    const btns = document.querySelectorAll('.teacher-tabs .tab-btn');
    if (tabName === 'individual') btns[0].classList.add('active'); 
    else if (tabName === 'master') btns[1].classList.add('active');
    else if (tabName === 'edit') btns[2].classList.add('active');
}

async function fetchTeacherData() {
    const btn = document.getElementById('t_btnFetch');
    const loading = document.getElementById('t_loading');
    btn.disabled = true; loading.style.display = 'block';

    try {
        const response = await fetch(TEACHER_API_URL);
        const data = await response.json();
        teacherGlobalData = data;
        loading.style.display = 'none'; btn.disabled = false;

        if (data.length > 0) {
            showAlert(`Berhasil! ${data.length} data guru diambil.`, "Sukses");
            renderEditTable(); // Merender data HTML setelah fetch berhasil
        } else {
            showAlert('Data 0. Cek Spreadsheet Guru.', "Info");
            renderEditTable();
        }

    } catch (error) {
        console.error(error);
        loading.style.display = 'none'; btn.disabled = false;
        showAlert('Gagal mengambil data guru.', "Error");
    }
}

function dateToTextID(dateObj) {
    let dd = String(dateObj.getDate()).padStart(2, '0');
    let mm = String(dateObj.getMonth() + 1).padStart(2, '0');
    let yyyy = dateObj.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
}

function timeToMinutes(timeStr) {
    if (!timeStr || timeStr === "-") return NaN;
    let parts = timeStr.split(':');
    if(parts.length < 2) return NaN;
    return (parseInt(parts[0]) * 60) + parseInt(parts[1]);
}

function processSmartMatrix(rawData, targetDatesObj) {
    let matrix = {};
    TEACHER_LIST.forEach(name => matrix[name] = {});
    let targetStrings = targetDatesObj.map(d => dateToTextID(d));

    rawData.forEach(item => {
        if (!TEACHER_LIST.includes(item.nama)) return;
        if (targetStrings.includes(item.tgl)) {
            let dateKey = item.tgl;
            if (!matrix[item.nama][dateKey]) matrix[item.nama][dateKey] = { fixedMasuk: null, fixedPulang: null };
            let entry = matrix[item.nama][dateKey];
            let ket = (item.keterangan || "").toLowerCase().trim();
            let timeStr = item.jam || "-";
            
            // Prioritas membaca Izin, Sakit, Alpa, Cuti, Dinas (Memunculkan teks pada laporan)
            if (ket.includes("sakit") || ket.includes("izin") || ket.includes("alpa") || ket.includes("cuti") || ket.includes("dinas")) {
                let displayKet = ket.charAt(0).toUpperCase() + ket.slice(1);
                entry.fixedMasuk = displayKet;
                entry.fixedPulang = "-";
            } else if (ket.includes("masuk") || ket.includes("terlambat") || ket.includes("datang")) {
                if (!entry.fixedMasuk || isNaN(timeToMinutes(entry.fixedMasuk))) { 
                    entry.fixedMasuk = timeStr; 
                } else if (!isNaN(timeToMinutes(timeStr))) { 
                    if (timeToMinutes(timeStr) < timeToMinutes(entry.fixedMasuk)) entry.fixedMasuk = timeStr; 
                }
            } else if (ket.includes("pulang") || ket.includes("keluar") || ket.includes("selesai")) {
                if (!entry.fixedPulang || isNaN(timeToMinutes(entry.fixedPulang))) { 
                    entry.fixedPulang = timeStr; 
                } else if (!isNaN(timeToMinutes(timeStr))) { 
                    if (timeToMinutes(timeStr) > timeToMinutes(entry.fixedPulang)) entry.fixedPulang = timeStr; 
                }
            } else {
                if (!entry.fixedMasuk) entry.fixedMasuk = timeStr;
                else if (!entry.fixedPulang) entry.fixedPulang = timeStr;
            }
        }
    });
    return matrix;
}

function renderEditTable() {
    const tbody = document.getElementById('tbody-edit-raw');
    if (!tbody) return;
    tbody.innerHTML = '';
    if (teacherGlobalData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" align="center">Data kosong.</td></tr>';
        return;
    }

    let html = '';
    teacherGlobalData.forEach((row, idx) => {
        html += createEditRowHTML(row.tgl, row.jam, row.nama, row.keterangan, idx);
    });
    tbody.innerHTML = html;
}

function createEditRowHTML(tgl, jam, nama, ket, id) {
    let teacherOptions = `<option value="">- Pilih Guru -</option>`;
    TEACHER_LIST.forEach(t => {
        teacherOptions += `<option value="${t}" ${t === nama ? 'selected' : ''}>${t}</option>`;
    });

    return `
    <tr data-id="${id}">
        <td><input type="text" class="form-control tgl-edit" value="${tgl || ''}" placeholder="DD/MM/YYYY" style="min-width: 110px;"></td>
        <td><input type="text" class="form-control jam-edit" value="${jam || ''}" placeholder="HH:MM" style="min-width: 80px;"></td>
        <td><select class="form-control nama-edit" style="min-width: 150px;">${teacherOptions}</select></td>
        <td><input type="text" class="form-control ket-edit" value="${ket || ''}" placeholder="Keterangan" style="min-width: 120px;"></td>
        <td align="center"><button class="btn btn-primary" onclick="this.closest('tr').remove()" style="background-color: #ef476f; border-radius: 6px; padding: 6px 10px;"><i class="fas fa-trash"></i></button></td>
    </tr>
    `;
}

function addEditRow() {
    const tbody = document.getElementById('tbody-edit-raw');
    if (tbody.querySelector('td[colspan]')) tbody.innerHTML = ''; 
    
    const newRow = createEditRowHTML('', '', '', '', 'new');
    tbody.insertAdjacentHTML('afterbegin', newRow);
}

function saveEditData() {
    const rows = document.querySelectorAll('#tbody-edit-raw tr');
    let newData = [];
    
    rows.forEach(tr => {
        if (tr.querySelector('.tgl-edit')) {
            const tgl = tr.querySelector('.tgl-edit').value.trim();
            const jam = tr.querySelector('.jam-edit').value.trim();
            const nama = tr.querySelector('.nama-edit').value;
            const ket = tr.querySelector('.ket-edit').value.trim();
            
            if (tgl && nama) { 
                newData.push({
                    tgl: tgl,
                    jam: jam,
                    nama: nama,
                    keterangan: ket
                });
            }
        }
    });
    
    teacherGlobalData = newData; // Simpan pembaruan ke dalam Memori Global Array
    showAlert("Data mentah berhasil diperbarui di memori. Laporan PDF akan menggunakan data ini.", "Tersimpan");
}

function finalizeTeacherPDF(doc, filename, mode) {
    if (mode === 'preview') {
        const previewDiv = document.getElementById('teacher-pdf-preview');
        previewDiv.innerHTML = `<iframe width='100%' height='100%' src='${doc.output('datauristring')}' style='border:none'></iframe>`;
        previewDiv.scrollIntoView({ behavior: "smooth" });
    } else {
        doc.save(filename);
    }
}

function generateIndividualPDF(mode) {
    if (teacherGlobalData.length === 0) { showAlert("Ambil data dulu.", "Peringatan"); return; }
    const selectedGuru = document.getElementById('t_guruSelector').value;
    if (!selectedGuru) { showAlert("Pilih nama guru terlebih dahulu.", "Peringatan"); return; }
    const selectedMonth = parseInt(document.getElementById('t_month').value);
    const selectedYear = parseInt(document.getElementById('t_year').value);
    const doc = new jsPDF({ orientation: "landscape", format: 'a4' });
    renderIndividualMonthPage(doc, selectedGuru, selectedMonth, selectedYear, true);
    const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

    finalizeTeacherPDF(doc, `Individu_${selectedGuru.replace(/ /g, '_')}_${monthNames[selectedMonth]}.pdf`, mode);
}

function generateAnnualIndividualPDF(mode) {
    if (teacherGlobalData.length === 0) { showAlert("Ambil data dulu.", "Peringatan"); return; }
    const selectedGuru = document.getElementById('t_guruSelector').value;
    if (!selectedGuru) { showAlert("Pilih nama guru terlebih dahulu.", "Peringatan"); return; }
    const selectedYear = parseInt(document.getElementById('t_year').value);
    const doc = new jsPDF({ orientation: "landscape", format: 'a4' });
    let isFirstPage = true; let dataFound = false;
    for (let m = 0; m < 12; m++) {
        let weeksBatches = generateWeeksBatches(selectedYear, m);
        let allTargetDates = weeksBatches.flat();
        let targetStrings = allTargetDates.map(d => dateToTextID(d));
        let hasData = teacherGlobalData.some(item => item.nama === selectedGuru && targetStrings.includes(item.tgl));
        if (!hasData) continue;
        dataFound = true;
        renderIndividualMonthPage(doc, selectedGuru, m, selectedYear, isFirstPage);
        isFirstPage = false;
    }
    if (!dataFound) {
        showAlert(`Tidak ada data absensi untuk ${selectedGuru} di tahun ${selectedYear}`, "Info");
    } else {
        finalizeTeacherPDF(doc, `Individu_Tahunan_${selectedGuru.replace(/ /g, '_')}_${selectedYear}.pdf`, mode);
    }
}

function renderIndividualMonthPage(doc, teacherName, monthIndex, year, isFirstPage) {
    if (!isFirstPage) doc.addPage();
    const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    const monthName = monthNames[monthIndex];
    let weeksBatches = generateWeeksBatches(year, monthIndex);
    let allTargetDates = weeksBatches.flat();
    let matrixData = processSmartMatrix(teacherGlobalData, allTargetDates);

    doc.setFont("helvetica", "bold"); doc.setFontSize(10);
    doc.text("DAFTAR KEHADIRAN GURU DAN STAFF SD INPRES LELINGLUAN", 148.5, 12, null, null, "center");
    doc.text(`TAHUN AJARAN ${year}/${year + 1}`, 148.5, 17, null, null, "center");
    doc.setFontSize(8); doc.text(`Bulan: ${monthName}`, 14, 25); doc.text(`Tahun: ${year}`, 14, 29);

    let headRow1 = [{ content: 'No', rowSpan: 2, styles: { valign: 'middle', halign: 'center' } }, { content: 'Nama', rowSpan: 2, styles: { valign: 'middle', halign: 'center' } }, { content: 'Hari Efektif', colSpan: 12, styles: { halign: 'center', fontStyle: 'bold' } }];
    const dayNamesIndo = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
    let headRow2 = []; dayNamesIndo.forEach(day => { headRow2.push({ content: day, colSpan: 2, styles: { halign: 'center', fontStyle: 'bold' } }); });

    let tableBody = [];
    let totalWeeks = weeksBatches.length;
    let rowsPerWeek = 4; let totalRowSpan = totalWeeks * rowsPerWeek;

    weeksBatches.forEach((batch, wIndex) => {
        let rowLabel = [];
        if (wIndex === 0) {
            rowLabel.push({ content: '1', rowSpan: totalRowSpan, styles: { valign: 'middle', halign: 'center' } });
            rowLabel.push({ content: teacherName, rowSpan: totalRowSpan, styles: { valign: 'middle', halign: 'center' } });
        }
        rowLabel.push({ content: `Minggu ${wIndex + 1}`, colSpan: 12, styles: { fontStyle: 'bold', fillColor: [240, 240, 240], halign: 'left' } });
        tableBody.push(rowLabel);

        let weekDatesMap = new Array(6).fill(null);
        batch.forEach(d => { let dayIdx = d.getDay(); if (dayIdx >= 1 && dayIdx <= 6) weekDatesMap[dayIdx - 1] = d; });

        let rowDates = []; weekDatesMap.forEach(d => { let txt = d ? dateToTextID(d) : ""; rowDates.push({ content: txt, colSpan: 2, styles: { halign: 'center', fontStyle: 'italic', fontSize: 7 } }); });
        tableBody.push(rowDates);

        let rowHeaders = []; for (let i = 0; i < 6; i++) { rowHeaders.push({ content: 'Masuk', styles: { halign: 'center', fontSize: 6 } }); rowHeaders.push({ content: 'Pulang', styles: { halign: 'center', fontSize: 6 } }); }
        tableBody.push(rowHeaders);

        let rowData = []; weekDatesMap.forEach(d => { let masuk = ""; let pulang = ""; if (d) { let dateKey = dateToTextID(d); let data = matrixData[teacherName] && matrixData[teacherName][dateKey]; masuk = (data && data.fixedMasuk) ? data.fixedMasuk : ""; pulang = (data && data.fixedPulang) ? data.fixedPulang : ""; } rowData.push({ content: masuk, styles: { halign: 'center' } }); rowData.push({ content: pulang, styles: { halign: 'center' } }); });
        tableBody.push(rowData);
    });

    doc.autoTable({ startY: 32, head: [headRow1, headRow2], body: tableBody, theme: 'grid', styles: { font: 'helvetica', fontSize: 7, cellPadding: 1, lineColor: [0, 0, 0], lineWidth: 0.1, textColor: [0, 0, 0] }, headStyles: { fillColor: [255, 255, 255], textColor: [0, 0, 0], fontStyle: 'bold', lineWidth: 0.1, lineColor: [0, 0, 0] }, bodyStyles: { fillColor: [255, 255, 255], textColor: [0, 0, 0] }, columnStyles: { 0: { cellWidth: 8 }, 1: { cellWidth: 45 } }, margin: { left: 10, right: 10 } });
    addSignature(doc, monthIndex, year);
}

function generateMasterPDF(mode) {
    if (teacherGlobalData.length === 0) { showAlert("Ambil data dulu.", "Peringatan"); return; }
    const selectedMonth = parseInt(document.getElementById('t_month').value);
    const selectedYear = parseInt(document.getElementById('t_year').value);
    const selectedWeek = parseInt(document.getElementById('t_week').value);
    const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    let datesInWeek = []; let firstDate = new Date(selectedYear, selectedMonth, 1); let dayOfFirst = firstDate.getDay(); let distToMon = (dayOfFirst + 6) % 7; let startMon = new Date(firstDate); startMon.setDate(firstDate.getDate() - distToMon); let targetMon = new Date(startMon); targetMon.setDate(startMon.getDate() + ((selectedWeek - 1) * 7));
    for (let i = 0; i < 6; i++) { let d = new Date(targetMon); d.setDate(targetMon.getDate() + i); datesInWeek.push(d); }
    let matrixData = processSmartMatrix(teacherGlobalData, datesInWeek);
    const doc = new jsPDF({ orientation: "landscape", format: 'a4' });
    createStrictMatrixPage(doc, datesInWeek, matrixData, selectedMonth, selectedYear, true);

    finalizeTeacherPDF(doc, `Rekap_Minggu_${selectedWeek}_${monthNames[selectedMonth]}.pdf`, mode);
}

function generateMonthlyFullPDF(mode) {
    if (teacherGlobalData.length === 0) { showAlert("Ambil data dulu.", "Peringatan"); return; }
    const selectedMonth = parseInt(document.getElementById('t_month').value);
    const selectedYear = parseInt(document.getElementById('t_year').value);
    const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    let weeksBatches = generateWeeksBatches(selectedYear, selectedMonth);
    let allTargetDates = weeksBatches.flat();
    let matrixData = processSmartMatrix(teacherGlobalData, allTargetDates);
    const doc = new jsPDF({ orientation: "landscape", format: 'a4' });
    weeksBatches.forEach((batch, index) => { let isFirst = (index === 0); createStrictMatrixPage(doc, batch, matrixData, selectedMonth, selectedYear, isFirst); });

    finalizeTeacherPDF(doc, `Absensi_Full_${monthNames[selectedMonth]}.pdf`, mode);
}

function generateAnnualMasterPDF(mode) {
    if (teacherGlobalData.length === 0) { showAlert("Ambil data dulu.", "Peringatan"); return; }
    const selectedYear = parseInt(document.getElementById('t_year').value);
    const doc = new jsPDF({ orientation: "landscape", format: 'a4' });
    let isGlobalFirstPage = true; let dataFound = false;
    for (let m = 0; m < 12; m++) {
        let weeksBatches = generateWeeksBatches(selectedYear, m);
        let allTargetDates = weeksBatches.flat();
        let targetStrings = allTargetDates.map(d => dateToTextID(d));
        let hasDataInMonth = teacherGlobalData.some(item => targetStrings.includes(item.tgl) && TEACHER_LIST.includes(item.nama));
        if (!hasDataInMonth) continue;
        dataFound = true;
        let matrixData = processSmartMatrix(teacherGlobalData, allTargetDates);
        weeksBatches.forEach((batch) => { createStrictMatrixPage(doc, batch, matrixData, m, selectedYear, isGlobalFirstPage); isGlobalFirstPage = false; });
    }
    if (!dataFound) {
        showAlert("Tidak ada data absensi ditemukan untuk tahun " + selectedYear, "Info");
    } else {
        finalizeTeacherPDF(doc, `Absensi_Tahunan_Master_${selectedYear}.pdf`, mode);
    }
}

function createStrictMatrixPage(doc, datesSubset, matrixData, monthIndex, year, isFirstPage, customTeacherList = null) {
    if (!isFirstPage) doc.addPage();
    const listToRender = customTeacherList || TEACHER_LIST;
    const dayNamesIndo = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
    const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    const monthName = monthNames[monthIndex];
    doc.setFont("helvetica", "bold"); doc.setFontSize(10);
    doc.text("DAFTAR KEHADIRAN GURU DAN STAFF SD INPRES LELINGLUAN", 148.5, 12, null, null, "center");
    doc.text(`TAHUN AJARAN ${year}/${year + 1}`, 148.5, 17, null, null, "center");
    doc.setFontSize(8); doc.text(`Bulan: ${monthName}`, 14, 25); doc.text(`Tahun: ${year}`, 14, 29);
    let row1 = [{ content: 'No', rowSpan: 4, styles: { valign: 'middle', halign: 'center' } }, { content: 'Nama', rowSpan: 4, styles: { valign: 'middle', halign: 'center' } }, { content: 'Hari Efektif', colSpan: datesSubset.length * 2, styles: { halign: 'center', fontStyle: 'bold' } }];
    let row2 = [], row3 = [], row4 = [];
    datesSubset.forEach(d => { let hari = dayNamesIndo[d.getDay()]; let tglString = dateToTextID(d); row2.push({ content: hari, colSpan: 2, styles: { halign: 'center', fontStyle: 'bold' } }); row3.push({ content: tglString, colSpan: 2, styles: { halign: 'center', fontStyle: 'italic' } }); row4.push({ content: 'Masuk', styles: { halign: 'center' } }); row4.push({ content: 'Pulang', styles: { halign: 'center' } }); });
    let tableBody = [];
    listToRender.forEach((name, idx) => {
        let noUrut = idx + 1; let row = [noUrut, name];
        datesSubset.forEach(d => { let dateKey = dateToTextID(d); let data = matrixData[name] && matrixData[name][dateKey]; row.push((data && data.fixedMasuk) ? data.fixedMasuk : ""); row.push((data && data.fixedPulang) ? data.fixedPulang : ""); });
        tableBody.push(row);
    });
    doc.autoTable({ startY: 32, head: [row1, row2, row3, row4], body: tableBody, theme: 'grid', styles: { font: 'helvetica', fontSize: 7, cellPadding: 1, lineColor: [0, 0, 0], lineWidth: 0.1, textColor: [0, 0, 0] }, headStyles: { fillColor: [255, 255, 255], textColor: [0, 0, 0], fontStyle: 'bold', lineWidth: 0.1, lineColor: [0, 0, 0] }, bodyStyles: { fillColor: [255, 255, 255], textColor: [0, 0, 0] }, columnStyles: { 0: { cellWidth: 8 }, 1: { cellWidth: 45 } }, margin: { left: 10, right: 10 } });
    addSignature(doc, monthIndex, year);
}

function addSignature(doc, monthIndex, year) {
    let finalY = doc.lastAutoTable.finalY + 10;
    if (finalY > 170) { doc.addPage(); finalY = 20; }
    const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    let lastDate = new Date(year, monthIndex + 1, 0);
    let tglTerakhir = lastDate.getDate();
    let titimangsa = `Lelingluan, ${tglTerakhir} ${monthNames[monthIndex]} ${year}`;
    let signX = 200;
    doc.setFont("helvetica", "normal"); doc.setFontSize(9); doc.text(titimangsa, signX, finalY); doc.text("Kepala Sekolah", signX, finalY + 5);
    let nameY = finalY + 30; doc.setFont("helvetica", "bold"); doc.text(KEPSEK_INFO.nama, signX, nameY);
    let textWidth = doc.getTextWidth(KEPSEK_INFO.nama); doc.setLineWidth(0.2); doc.line(signX, nameY + 1, signX + textWidth, nameY + 1);
    doc.setFont("helvetica", "normal"); doc.text(`NIP. ${KEPSEK_INFO.nip}`, signX, nameY + 5);
}

function generateWeeksBatches(year, month) {
    let weeksBatches = []; let currentBatch = []; let iterDate = new Date(year, month, 1);
    while (iterDate.getMonth() === month) {
        let day = iterDate.getDay(); if (day !== 0) { currentBatch.push(new Date(iterDate)); }
        let isSaturday = (day === 6); let nextDay = new Date(iterDate); nextDay.setDate(iterDate.getDate() + 1); let isLastDayOfMonth = (nextDay.getMonth() !== month);
        if ((isSaturday || isLastDayOfMonth) && currentBatch.length > 0) { weeksBatches.push(currentBatch); currentBatch = []; }
        iterDate.setDate(iterDate.getDate() + 1);
    }
    return weeksBatches;
}
