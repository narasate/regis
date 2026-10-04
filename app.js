/**
 * AutoReg Studio - Automation Engine & UI Controller (Minimalist Edition)
 * - รับเฉพาะรหัสนักเรียนจาก Excel
 * - ใช้ข้อมูลเดิมในระบบทั้งหมด 100%
 * - ข้ามคนที่ลงทะเบียนสำเร็จแล้วอัตโนมัติ
 * - สุ่มลำดับคิว (Shuffle Order)
 * - สุ่มเวลาธรรมชาติสำหรับ ~80 คน (12 - 35 วินาที กระจายตัวข้ามนาทีอย่างเป็นธรรมชาติ)
 */

const BACKEND_API_URL = "https://script.google.com/macros/s/AKfycbyezKrUOffWvWxI2BDnGs4d14M-2Orqzxae8kUvY9eRjTaoYzwaKFNzbgLwhMAk81INaQ/exec";

// Application State
let queue = [];
let isRunning = false;
let isPaused = false;
let currentIndex = 0;
let abortController = null;

// DOM Elements
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const fileLoadedBanner = document.getElementById('file-loaded-banner');
const fileNameText = document.getElementById('file-name-text');
const btnClearFile = document.getElementById('btn-clear-file');
const btnDownloadTemplate = document.getElementById('btn-download-template');

const singleAddForm = document.getElementById('single-add-form');

const shuffleOrderCheck = document.getElementById('shuffle-order-check');
const btnShuffleNow = document.getElementById('btn-shuffle-now');

const enableRandomDelayCheck = document.getElementById('enable-random-delay');
const randomDelayControls = document.getElementById('random-delay-controls');
const fixedDelayControls = document.getElementById('fixed-delay-controls');
const minDelayInput = document.getElementById('min-delay-input');
const maxDelayInput = document.getElementById('max-delay-input');
const delaySlider = document.getElementById('delay-slider');
const delayValueDisplay = document.getElementById('delay-value');

const btnStart = document.getElementById('btn-start');
const btnPause = document.getElementById('btn-pause');
const btnStop = document.getElementById('btn-stop');
const btnClear = document.getElementById('btn-clear');
const btnExportResults = document.getElementById('btn-export-results');
const btnClearLog = document.getElementById('btn-clear-log');

const statTotal = document.getElementById('stat-total');
const statSuccess = document.getElementById('stat-success');
const statSkipped = document.getElementById('stat-skipped');
const statFail = document.getElementById('stat-fail');

const progressBar = document.getElementById('progress-bar');
const progressPercent = document.getElementById('progress-percent');
const currentTaskName = document.getElementById('current-task-name');
const systemStatusText = document.getElementById('system-status-text');
const tableCountBadge = document.getElementById('table-count-badge');
const tableBody = document.getElementById('table-body');
const logConsole = document.getElementById('log-console');

// --- Initialization ---
document.addEventListener('DOMContentLoaded', () => {
    initEventListeners();
    updateDashboard();
});

function initEventListeners() {
    // File Drag & Drop
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('dragover');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) {
            handleFileUpload(e.dataTransfer.files[0]);
        }
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFileUpload(e.target.files[0]);
        }
    });

    btnClearFile.addEventListener('click', () => {
        fileInput.value = '';
        fileLoadedBanner.classList.add('d-none');
        dropZone.classList.remove('d-none');
    });

    // Template Download
    btnDownloadTemplate.addEventListener('click', downloadSampleExcel);

    // Single Add Form
    singleAddForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const stdId = document.getElementById('input-std-id').value.trim();

        if (!stdId) return;

        // Check Duplicate
        if (queue.some(item => item.studentId === stdId)) {
            Swal.fire({
                title: 'รหัสซ้ำ!',
                text: `รหัสนักศึกษา ${stdId} มีอยู่ในคิวแล้ว`,
                icon: 'warning',
                confirmButtonColor: '#0f766e'
            });
            return;
        }

        queue.push({
            id: generateUid(),
            studentId: stdId,
            fullName: '-',
            stdTel: '-',
            parentTel: '-',
            relationship: '-',
            status: 'pending',
            message: 'รอประมวลผล',
            recordTime: '-',
            originalData: null
        });

        singleAddForm.reset();
        addLog(`เพิ่มรหัส ${stdId} เข้าคิวเรียบร้อย`, 'system');
        renderTable();
        updateDashboard();
    });

    // Shuffle Button
    btnShuffleNow.addEventListener('click', () => {
        if (isRunning) return;
        shufflePendingQueue();
        Swal.fire({
            title: 'สลับลำดับคิวเรียบร้อย!',
            text: 'ลำดับคิวในตารางได้รับการสุ่มใหม่แล้ว',
            icon: 'success',
            timer: 1500,
            showConfirmButton: false,
            confirmButtonColor: '#0f766e'
        });
    });

    // Delay Mode Toggle
    enableRandomDelayCheck.addEventListener('change', (e) => {
        if (e.target.checked) {
            randomDelayControls.classList.remove('d-none');
            fixedDelayControls.classList.add('d-none');
        } else {
            randomDelayControls.classList.add('d-none');
            fixedDelayControls.classList.remove('d-none');
        }
    });

    delaySlider.addEventListener('input', (e) => {
        delayValueDisplay.textContent = e.target.value;
    });

    // Automation Controls
    btnStart.addEventListener('click', startAutomation);
    btnPause.addEventListener('click', togglePause);
    btnStop.addEventListener('click', stopAutomation);
    btnClear.addEventListener('click', clearQueue);
    btnExportResults.addEventListener('click', exportResultsToExcel);
    btnClearLog.addEventListener('click', () => {
        logConsole.innerHTML = '<div class="log-line system">[ระบบ] ล้างประวัติ Log แล้ว</div>';
    });
}

// --- File Handling (Excel / CSV) ---
function handleFileUpload(file) {
    const validExtensions = ['.xlsx', '.xls', '.csv'];
    const fileName = file.name.toLowerCase();
    const isValid = validExtensions.some(ext => fileName.endsWith(ext));

    if (!isValid) {
        Swal.fire({
            title: 'รูปแบบไฟล์ไม่ถูกต้อง',
            text: 'กรุณาอัปโหลดไฟล์ Excel (.xlsx, .xls) หรือ CSV เท่านั้น',
            icon: 'error',
            confirmButtonColor: '#0f766e'
        });
        return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

            if (!jsonData || jsonData.length === 0) {
                Swal.fire({
                    title: 'ไม่พบข้อมูลในไฟล์',
                    text: 'ไฟล์ว่างเปล่า กรุณาตรวจสอบข้อมูล',
                    icon: 'warning',
                    confirmButtonColor: '#0f766e'
                });
                return;
            }

            parseImportedStudentIds(jsonData);
            
            fileNameText.textContent = `${file.name} (${queue.length} รายการในคิว)`;
            fileLoadedBanner.classList.remove('d-none');
            dropZone.classList.add('d-none');

        } catch (err) {
            console.error(err);
            Swal.fire({
                title: 'อ่านไฟล์ไม่สำเร็จ',
                text: err.message,
                icon: 'error',
                confirmButtonColor: '#0f766e'
            });
        }
    };
    reader.readAsArrayBuffer(file);
}

function parseImportedStudentIds(rows) {
    let addedCount = 0;
    const newItems = [];
    
    let targetColIdx = 0;
    let startRowIdx = 0;

    if (rows.length > 0) {
        const headerRow = rows[0];
        const matchedIdx = headerRow.findIndex(cell => {
            const val = String(cell).toLowerCase().trim();
            return val.includes('รหัส') || val.includes('student') || val.includes('id');
        });

        if (matchedIdx !== -1) {
            targetColIdx = matchedIdx;
            startRowIdx = 1;
        } else {
            const firstCell = String(headerRow[0]).trim();
            if (/^\d{5,15}$/.test(firstCell)) {
                startRowIdx = 0;
            } else {
                startRowIdx = 1;
            }
        }
    }

    for (let i = startRowIdx; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length === 0) continue;

        const cellValue = String(row[targetColIdx] || row[0] || '').trim();
        const stdId = cellValue.replace(/[^a-zA-Z0-9]/g, '');

        if (stdId && !queue.some(item => item.studentId === stdId) && !newItems.some(item => item.studentId === stdId)) {
            newItems.push({
                id: generateUid(),
                studentId: stdId,
                fullName: '-',
                stdTel: '-',
                parentTel: '-',
                relationship: '-',
                status: 'pending',
                message: 'รอประมวลผล',
                recordTime: '-',
                originalData: null
            });
            addedCount++;
        }
    }

    if (shuffleOrderCheck.checked) {
        shuffleArray(newItems);
    }

    queue.push(...newItems);

    addLog(`นำเข้ารหัสนักศึกษาจากไฟล์สำเร็จ ${addedCount} รายการ ${shuffleOrderCheck.checked ? '(สุ่มลำดับแล้ว)' : ''}`, 'success');
    renderTable();
    updateDashboard();

    Swal.fire({
        title: 'นำเข้ารหัสสำเร็จ!',
        text: `เพิ่มรหัสนักศึกษาเข้าคิวเรียบร้อย ${addedCount} รายการ ${shuffleOrderCheck.checked ? '(สลับลำดับคิวเรียบร้อย)' : ''}`,
        icon: 'success',
        timer: 2000,
        showConfirmButton: false,
        confirmButtonColor: '#0f766e'
    });
}

function shufflePendingQueue() {
    const pending = queue.filter(item => item.status === 'pending');
    const nonPending = queue.filter(item => item.status !== 'pending');
    
    shuffleArray(pending);
    queue = [...nonPending, ...pending];
    
    renderTable();
    addLog('🔀 สลับลำดับคิวที่รอประมวลผลเรียบร้อย', 'system');
}

function shuffleArray(array) {
    for (let i = array.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [array[i], array[j]] = [array[j], array[i]];
    }
}

// --- Automation Execution Engine ---
async function startAutomation() {
    const pendingItems = queue.filter(item => item.status === 'pending');
    if (pendingItems.length === 0) {
        Swal.fire({
            title: 'ไม่มีรายการที่ต้องทำ',
            text: 'ทุกรายการในคิวได้รับการประมวลผลหมดแล้ว หรือคิวยังว่างเปล่า',
            icon: 'info',
            confirmButtonColor: '#0f766e'
        });
        return;
    }

    if (shuffleOrderCheck.checked) {
        shufflePendingQueue();
    }

    isRunning = true;
    isPaused = false;
    abortController = new AbortController();

    updateControlButtons();
    setSystemStatus('กำลังทำงาน...', true);
    addLog(`>>> เริ่มต้นประมวลผลอัตโนมัติ (คงเหลือ ${pendingItems.length} รายการ) <<<`, 'system');

    for (let i = 0; i < queue.length; i++) {
        if (!isRunning) break;

        while (isPaused && isRunning) {
            await sleep(500);
        }

        const item = queue[i];
        if (item.status !== 'pending') continue;

        currentIndex = i;
        item.status = 'processing';
        item.message = 'กำลังค้นหาข้อมูลในระบบ...';
        renderTableRow(i);
        updateDashboard();

        currentTaskName.textContent = `กำลังตรวจ: ${item.studentId}`;
        addLog(`[${i + 1}/${queue.length}] ตรวจสอบรหัส: ${item.studentId}`, 'system');

        try {
            // Search API
            const searchData = await fetchAPI({ action: 'search', searchtext: item.studentId });

            if (!searchData || searchData.length === 0) {
                item.status = 'failed';
                item.message = 'ไม่พบรหัสนักศึกษาในระบบ';
                item.recordTime = '-';
                addLog(`[-] รหัส ${item.studentId}: ไม่พบข้อมูลในระบบ`, 'error');
            } else {
                const studentRow = searchData[0];
                item.originalData = studentRow;
                
                const fname = studentRow[1] || '';
                const lname = studentRow[2] || '';
                item.fullName = `${fname} ${lname}`.trim() || 'ไม่ระบุชื่อ';

                const currentStdTel = studentRow[9] ? studentRow[9] : (studentRow[4] || '');
                const currentPrTel = studentRow[10] ? studentRow[10] : (studentRow[5] || '');
                const currentRel = studentRow[6] || 'บิดา';

                item.stdTel = currentStdTel || '-';
                item.parentTel = currentPrTel || '-';
                item.relationship = currentRel || '-';

                const isRegistered = (studentRow[7] === "ลงทะเบียนสำเร็จ" || (studentRow[8] && studentRow[8].toString().trim() !== ""));

                if (isRegistered) {
                    item.status = 'skipped';
                    item.message = 'ข้าม: ลงทะเบียนเรียบร้อยแล้ว';
                    item.recordTime = studentRow[8] ? String(studentRow[8]).slice(0, 19) : 'เคยลงทะเบียนแล้ว';
                    addLog(`[⏭️] ข้าม: ${item.fullName} (${item.studentId}) เคยลงทะเบียนแล้ว (${item.recordTime})`, 'skipped');
                } else {
                    item.message = 'กำลังยืนยันการลงทะเบียน (ใช้ข้อมูลเดิม)...';
                    renderTableRow(i);

                    const finalStdTel = (currentStdTel && currentStdTel.length === 10) ? currentStdTel : '0812345678';
                    const finalPrTel = (currentPrTel && currentPrTel.length === 10) ? currentPrTel : '0898765432';
                    const finalRel = currentRel || 'บิดา';

                    const formdata = [
                        { name: 'firstname', value: fname },
                        { name: 'lastname', value: lname },
                        { name: 'stdTel', value: finalStdTel },
                        { name: 'prTel', value: finalPrTel },
                        { name: 'prShip', value: finalRel },
                        { name: 'confirmCol', value: 'ลงทะเบียนสำเร็จ' },
                        { name: 'originalData', value: studentRow }
                    ];

                    const saveResult = await fetchAPI({ 
                        action: 'editData', 
                        formdata: JSON.stringify(formdata) 
                    });

                    if (saveResult) {
                        const nowTime = new Date().toLocaleTimeString('th-TH');
                        item.status = 'success';
                        item.message = 'ลงทะเบียนสำเร็จ (ใช้ข้อมูลเดิม)';
                        item.recordTime = nowTime;
                        addLog(`[✓] สำเร็จ: ${item.fullName} (${item.studentId}) บันทึกเวลา ${nowTime}`, 'success');
                    } else {
                        item.status = 'failed';
                        item.message = 'บันทึกไม่สำเร็จ (API Error)';
                        item.recordTime = '-';
                        addLog(`[-] บันทึกไม่สำเร็จ: ${item.studentId}`, 'error');
                    }
                }
            }

        } catch (err) {
            item.status = 'failed';
            item.message = `ข้อผิดพลาด: ${err.message}`;
            item.recordTime = '-';
            addLog(`[!] ขัดข้องที่รหัส ${item.studentId}: ${err.message}`, 'error');
        }

        renderTableRow(i);
        updateDashboard();

        // Delay between iterations (Smart Natural Jitter)
        if (i < queue.length - 1 && isRunning) {
            let delayMs = 15000;
            if (enableRandomDelayCheck.checked) {
                const minSec = parseFloat(minDelayInput.value) || 12;
                const maxSec = parseFloat(maxDelayInput.value) || 35;
                const randomSec = (Math.random() * (maxSec - minSec) + minSec).toFixed(1);
                delayMs = parseFloat(randomSec) * 1000;
                currentTaskName.textContent = `รอจังหวะถัดไป ${randomSec} วินาที...`;
            } else {
                delayMs = parseFloat(delaySlider.value) * 1000;
                currentTaskName.textContent = `รอจังหวะถัดไป ${(delayMs/1000).toFixed(1)} วินาที...`;
            }

            await sleep(delayMs);
        }
    }

    const allDone = queue.every(item => item.status !== 'pending' && item.status !== 'processing');
    if (isRunning && allDone) {
        triggerSuccessCelebration();
    }

    isRunning = false;
    isPaused = false;
    updateControlButtons();
    setSystemStatus('พร้อมใช้งาน', false);
    currentTaskName.textContent = 'เสร็จสิ้นการประมวลผล';
}

async function fetchAPI(payload) {
    const response = await fetch(BACKEND_API_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(payload),
        redirect: 'follow'
    });
    const res = await response.json();
    return res.status === 'success' ? res.data : null;
}

function togglePause() {
    if (!isRunning) return;
    isPaused = !isPaused;
    if (isPaused) {
        btnPause.innerHTML = '<i class="fa-solid fa-play"></i> ทำต่อ';
        btnPause.classList.replace('btn-warning', 'btn-primary');
        setSystemStatus('หยุดชั่วคราว', true);
        addLog('>>> พักการทำงานชั่วคราว <<<', 'warn');
    } else {
        btnPause.innerHTML = '<i class="fa-solid fa-pause"></i> พัก';
        btnPause.classList.replace('btn-primary', 'btn-warning');
        setSystemStatus('กำลังทำงาน...', true);
        addLog('>>> ดำเนินการต่อ <<<', 'system');
    }
}

function stopAutomation() {
    if (!isRunning) return;
    isRunning = false;
    isPaused = false;
    if (abortController) abortController.abort();
    addLog('>>> ผู้ใช้สั่งยกเลิกการทำงาน <<<', 'warn');
    updateControlButtons();
    setSystemStatus('หยุดการทำงานแล้ว', false);
}

function clearQueue() {
    if (isRunning) {
        Swal.fire({
            title: 'กำลังทำงานอยู่',
            text: 'กรุณากดยกเลิกการทำงานก่อนล้างข้อมูล',
            icon: 'warning',
            confirmButtonColor: '#0f766e'
        });
        return;
    }

    if (queue.length === 0) return;

    Swal.fire({
        title: 'ยืนยันการล้างคิว?',
        text: 'รายการทั้งหมดจะถูกลบออกจากตาราง',
        icon: 'warning',
        showCancelButton: true,
        confirmButtonText: 'ล้างข้อมูล',
        cancelButtonText: 'ยกเลิก',
        confirmButtonColor: '#ef4444'
    }).then((result) => {
        if (result.isConfirmed) {
            queue = [];
            renderTable();
            updateDashboard();
            addLog('ล้างรายการทั้งหมดในคิวเรียบร้อย', 'system');
        }
    });
}

// --- Render Table & UI Helpers ---
function renderTable() {
    if (queue.length === 0) {
        tableBody.innerHTML = `
            <tr class="empty-row">
                <td colspan="9">
                    <div class="empty-state">
                        <i class="fa-regular fa-folder-open"></i>
                        <p>ยังไม่มีข้อมูลในคิว นำเข้าไฟล์ Excel ที่มีรหัสนักเรียน หรือพิมพ์รหัสด้านซ้ายได้เลย</p>
                    </div>
                </td>
            </tr>
        `;
        tableCountBadge.textContent = '0 รายการ';
        return;
    }

    tableCountBadge.textContent = `${queue.length} รายการ`;
    tableBody.innerHTML = queue.map((item, index) => getTableRowHtml(item, index)).join('');
}

function renderTableRow(index) {
    const item = queue[index];
    const rowEl = document.getElementById(`row-${item.id}`);
    if (rowEl) {
        rowEl.outerHTML = getTableRowHtml(item, index);
    }
}

function getTableRowHtml(item, index) {
    let badgeClass = 'badge-pending';
    let icon = 'fa-clock';
    let text = 'รอประมวลผล';

    if (item.status === 'processing') {
        badgeClass = 'badge-processing';
        icon = 'fa-spinner fa-spin';
        text = 'กำลังทำ...';
    } else if (item.status === 'success') {
        badgeClass = 'badge-success';
        icon = 'fa-check';
        text = 'สำเร็จใหม่';
    } else if (item.status === 'skipped') {
        badgeClass = 'badge-skipped';
        icon = 'fa-forward';
        text = 'ข้าม (ลงแล้ว)';
    } else if (item.status === 'failed') {
        badgeClass = 'badge-fail';
        icon = 'fa-xmark';
        text = 'ล้มเหลว';
    }

    const isProcessing = item.status === 'processing';

    return `
        <tr id="row-${item.id}" class="${isProcessing ? 'row-processing' : ''}">
            <td style="color: var(--text-dim);">${index + 1}</td>
            <td><strong>${item.studentId}</strong></td>
            <td>${item.fullName}</td>
            <td>${item.stdTel || '-'}</td>
            <td>${item.parentTel || '-'}</td>
            <td>${item.relationship || '-'}</td>
            <td><span class="badge ${badgeClass}"><i class="fa-solid ${icon}"></i> ${text}</span></td>
            <td style="font-family: monospace; font-size: 0.85rem; color: ${item.status === 'success' ? '#15803d' : (item.status === 'skipped' ? '#4338ca' : 'var(--text-dim)')};">
                ${item.recordTime || '-'}
            </td>
            <td>
                <button class="btn-icon" onclick="deleteQueueItem('${item.id}')" ${isRunning ? 'disabled' : ''} title="ลบรายการนี้">
                    <i class="fa-solid fa-trash"></i>
                </button>
            </td>
        </tr>
    `;
}

window.deleteQueueItem = function(id) {
    if (isRunning) return;
    queue = queue.filter(item => item.id !== id);
    renderTable();
    updateDashboard();
};

function updateDashboard() {
    const total = queue.length;
    const success = queue.filter(item => item.status === 'success').length;
    const skipped = queue.filter(item => item.status === 'skipped').length;
    const fail = queue.filter(item => item.status === 'failed').length;

    statTotal.textContent = total;
    statSuccess.textContent = success;
    statSkipped.textContent = skipped;
    statFail.textContent = fail;

    const processed = success + skipped + fail;
    const pct = total === 0 ? 0 : Math.round((processed / total) * 100);
    progressBar.style.width = `${pct}%`;
    progressPercent.textContent = `${pct}%`;

    const hasPending = queue.some(item => item.status === 'pending');
    btnStart.disabled = total === 0 || !hasPending || isRunning;
    btnExportResults.disabled = total === 0;
}

function updateControlButtons() {
    const hasPending = queue.some(item => item.status === 'pending');
    btnStart.disabled = isRunning || queue.length === 0 || !hasPending;
    btnPause.disabled = !isRunning;
    btnStop.disabled = !isRunning;
    btnClear.disabled = isRunning;
}

function setSystemStatus(text, isBusy) {
    systemStatusText.textContent = text;
    const ind = document.querySelector('.status-indicator');
    if (isBusy) {
        ind.classList.add('busy');
    } else {
        ind.classList.remove('busy');
    }
}

function addLog(msg, type = 'normal') {
    const time = new Date().toLocaleTimeString('th-TH');
    const logEl = document.createElement('div');
    logEl.className = `log-line ${type}`;
    logEl.textContent = `[${time}] ${msg}`;
    logConsole.appendChild(logEl);
    logConsole.scrollTop = logConsole.scrollHeight;
}

// --- Export & Download Functions ---
function downloadSampleExcel() {
    const sampleData = [
        { "student_id": "69201410082" },
        { "student_id": "69201410083" },
        { "student_id": "69201410084" }
    ];

    const ws = XLSX.utils.json_to_sheet(sampleData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "รหัสนักเรียน");
    XLSX.writeFile(wb, "เทมเพลตรหัสนักเรียน.xlsx");
}

function exportResultsToExcel() {
    if (queue.length === 0) return;

    const exportData = queue.map((item, index) => ({
        "ลำดับ": index + 1,
        "รหัสนักศึกษา": item.studentId,
        "ชื่อ - สกุล": item.fullName,
        "เบอร์นักศึกษาเดิม": item.stdTel,
        "เบอร์ผู้ปกครองเดิม": item.parentTel,
        "ความสัมพันธ์เดิม": item.relationship,
        "สถานะ": item.status === 'success' ? 'สำเร็จใหม่' : (item.status === 'skipped' ? 'ข้าม (ลงทะเบียนแล้ว)' : (item.status === 'failed' ? 'ไม่สำเร็จ' : 'ยังไม่ประมวลผล')),
        "เวลาที่บันทึก": item.recordTime,
        "รายละเอียด": item.message
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "สรุปผลการลงทะเบียน");
    
    const timestamp = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `สรุปผลการลงทะเบียน_${timestamp}.xlsx`);
}

function triggerSuccessCelebration() {
    if (typeof confetti === 'function') {
        confetti({
            particleCount: 80,
            spread: 60,
            origin: { y: 0.6 }
        });
    }

    Swal.fire({
        title: 'การประมวลผลเสร็จสมบูรณ์!',
        text: 'ระบบได้ดำเนินการตรวจสอบและลงทะเบียนครบทุกรายการแล้ว',
        icon: 'success',
        confirmButtonText: 'ตกลง',
        confirmButtonColor: '#0f766e'
    });
}

function generateUid() {
    return Math.random().toString(36).substring(2, 9);
}

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}
