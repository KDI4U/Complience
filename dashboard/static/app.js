// Compliance Automation Dashboard - Frontend Logic
let currentService = 'esic';
let accountsData = [];
let isProcessRunning = false;
let autoScroll = true;
let editingRow = null;
let eventSource = null;
let liveDownloadedCount = 0;
let liveFailedCount = 0;
let selectedLoginIds = new Set();
let currentModalStep = 1;
let currentModalTitle = '';

// DOM Elements
const tabEsic = document.getElementById('tab-esic');
const tabEpfo = document.getElementById('tab-epfo');
const totalAccountsEl = document.getElementById('stat-total');
const successAccountsEl = document.getElementById('stat-success');
const failedAccountsEl = document.getElementById('stat-failed');
const drivePathEl = document.getElementById('stat-drive');
const statLiveDownloadedEl = document.getElementById('stat-live-downloaded');
const headerLiveDownloadedEl = document.getElementById('header-live-downloaded-count');
const accountsTableBody = document.getElementById('accounts-tbody');
const searchInput = document.getElementById('account-search');
const terminalOutput = document.getElementById('terminal-output');
const serverStatusIndicator = document.getElementById('server-status-dot');
const serverStatusText = document.getElementById('server-status-text');

// Stage Tracking Elements
const headerStagePill = document.getElementById('header-stage-pill');
const headerStageText = document.getElementById('header-stage-text');
const headerStageIcon = document.getElementById('header-stage-icon');
const stageTrackerContainer = document.getElementById('stage-tracker-container');
const stageTrackerBoldText = document.getElementById('stage-tracker-bold-text');
const stageTrackerSubtext = document.getElementById('stage-tracker-subtext');

// Single Run Inputs
const singleUserIdInput = document.getElementById('single-userid');
const singlePasswordInput = document.getElementById('single-password');
const singleHeadlessCb = document.getElementById('single-headless');
const btnRunSingle = document.getElementById('btn-run-single');

// Pipeline Run Inputs
const pipelineStepSelect = document.getElementById('pipeline-step');
const pipelineHeadlessCb = document.getElementById('pipeline-headless');
const btnRunPipeline = document.getElementById('btn-run-pipeline');
const btnStopPipeline = document.getElementById('btn-stop-pipeline');

// Modals
const modalAccount = document.getElementById('modal-account');
const modalTitle = document.getElementById('modal-title');
const modalUserId = document.getElementById('modal-userid');
const modalPassword = document.getElementById('modal-password');
const btnSaveAccount = document.getElementById('btn-save-account');

const modalRecipients = document.getElementById('modal-recipients');
const recipientsToInput = document.getElementById('recipients-to');
const recipientsCcInput = document.getElementById('recipients-cc');
const btnSaveRecipients = document.getElementById('btn-save-recipients');

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    setupEventListeners();
    connectLogStream();
    loadAccounts();
    loadRecipients();
    checkProcessStatus();
    setInterval(checkProcessStatus, 3000);
});

function setupEventListeners() {
    tabEsic.addEventListener('click', () => switchService('esic'));
    tabEpfo.addEventListener('click', () => switchService('epfo'));

    searchInput.addEventListener('input', (e) => {
        renderAccountsTable(e.target.value.trim().toLowerCase());
    });

    const singleCodeSelect = document.getElementById('single-code-select');
    if (singleCodeSelect) {
        singleCodeSelect.addEventListener('change', () => {
            const val = singleCodeSelect.value;
            if (!val) return;
            const found = accountsData.find(a => String(a.loginid).trim() === String(val).trim());
            singleUserIdInput.value = val;
            singlePasswordInput.value = found ? (found.password || '') : '';
            showToast(`Selected account: ${val}`, 'info');
        });
    }

    btnRunSingle.addEventListener('click', handleRunSingle);
    btnRunPipeline.addEventListener('click', handleRunPipeline);
    btnStopPipeline.addEventListener('click', handleStopProcess);

    document.getElementById('btn-add-account').addEventListener('click', () => openAccountModal());
    document.getElementById('modal-close-btn').addEventListener('click', closeAccountModal);
    document.getElementById('modal-cancel-btn').addEventListener('click', closeAccountModal);
    btnSaveAccount.addEventListener('click', handleSaveAccount);

    // Email & Report Buttons
    document.getElementById('btn-quick-send-email').addEventListener('click', handleSendFailureEmail);
    document.getElementById('btn-trigger-email').addEventListener('click', handleSendFailureEmail);
    document.getElementById('btn-generate-report-only').addEventListener('click', handleGenerateReportOnly);
    document.getElementById('btn-edit-recipients').addEventListener('click', openRecipientsModal);
    document.getElementById('modal-recipients-close-btn').addEventListener('click', closeRecipientsModal);
    btnSaveRecipients.addEventListener('click', handleSaveRecipients);
    document.getElementById('modal-recipients-open-excel').addEventListener('click', () => {
        openPathOnHost('recipients_excel');
    });

    document.getElementById('btn-clear-logs').addEventListener('click', () => {
        terminalOutput.innerHTML = '';
    });

    document.getElementById('btn-copy-logs').addEventListener('click', () => {
        const text = terminalOutput.innerText;
        navigator.clipboard.writeText(text);
        showToast('Logs copied to clipboard', 'success');
    });

    document.getElementById('btn-autoscroll').addEventListener('click', (e) => {
        autoScroll = !autoScroll;
        e.target.innerText = autoScroll ? 'Auto-scroll: ON' : 'Auto-scroll: OFF';
    });

    document.getElementById('toggle-single-pwd').addEventListener('click', () => {
        const type = singlePasswordInput.type === 'password' ? 'text' : 'password';
        singlePasswordInput.type = type;
    });

    document.getElementById('btn-open-folder').addEventListener('click', () => {
        const target = currentService === 'esic' ? 'esic_folder' : 'epfo_folder';
        openPathOnHost(target);
    });

    document.getElementById('btn-open-excel').addEventListener('click', () => {
        const target = currentService === 'esic' ? 'esic_excel' : 'epfo_excel';
        openPathOnHost(target);
    });

    // Step Badges Click -> Open Step Modal
    document.querySelectorAll('.clickable-step').forEach(btn => {
        btn.addEventListener('click', () => {
            const stepNum = parseInt(btn.getAttribute('data-step') || '1');
            const stepTitle = btn.getAttribute('data-title') || `Step ${stepNum}`;
            openStepRunnerModal(stepNum, stepTitle);
        });
    });

    // Step Runner Modal Events
    const stepCloseBtn = document.getElementById('modal-step-close-btn');
    if (stepCloseBtn) stepCloseBtn.addEventListener('click', closeStepRunnerModal);
    const stepCancelBtn = document.getElementById('modal-step-cancel-btn');
    if (stepCancelBtn) stepCancelBtn.addEventListener('click', closeStepRunnerModal);
    const stepRunBtn = document.getElementById('modal-step-run-btn');
    if (stepRunBtn) stepRunBtn.addEventListener('click', executeStepModalAction);
    const stepUseSelected = document.getElementById('modal-step-use-selected');
    if (stepUseSelected) {
        stepUseSelected.addEventListener('click', () => {
            const input = document.getElementById('modal-step-codes');
            if (selectedLoginIds.size === 0) {
                showToast('No codes checked in table. Check 1 or 2 codes in the table below first.', 'info');
                return;
            }
            if (input) input.value = Array.from(selectedLoginIds).join(', ');
            showToast(`Filled ${selectedLoginIds.size} checked code(s)`, 'success');
        });
    }

    // Batch Actions Toolbar Events
    const btnRetryFailed = document.getElementById('btn-retry-failed-codes');
    if (btnRetryFailed) btnRetryFailed.addEventListener('click', handleRetryFailedCodes);
    const btnRunSelected = document.getElementById('btn-run-selected-codes');
    if (btnRunSelected) btnRunSelected.addEventListener('click', handleRunSelectedCodes);
    const btnClearSel = document.getElementById('btn-clear-selection');
    if (btnClearSel) btnClearSel.addEventListener('click', clearSelectedCodes);
    const selectAllCb = document.getElementById('select-all-accounts');
    if (selectAllCb) selectAllCb.addEventListener('change', handleSelectAllToggle);
}

function updateStageUI(stage, isRunning, serviceName) {
    const rawStage = (stage || (isRunning ? 'RUNNING...' : 'READY / IDLE')).toUpperCase();
    
    if (headerStageText) headerStageText.innerText = rawStage;
    if (stageTrackerBoldText) stageTrackerBoldText.innerText = rawStage;

    const svc = (serviceName || currentService).toLowerCase();

    // Reset badges
    for (let i = 1; i <= 7; i++) {
        const badge = document.getElementById(`step-badge-${i}`);
        if (badge) badge.classList.remove('active', 'completed');
    }
    const mailBadge = document.getElementById('step-badge-7') || document.getElementById('step-badge-mail');
    if (mailBadge) mailBadge.classList.remove('active', 'completed');

    if (!isRunning) {
        if (headerStagePill) {
            headerStagePill.className = 'active-stage-pill';
            if (rawStage.includes('COMPLETED')) {
                headerStagePill.classList.add('completed');
                for (let i = 1; i <= 7; i++) {
                    const b = document.getElementById(`step-badge-${i}`);
                    if (b) b.classList.add('completed');
                }
                if (mailBadge) mailBadge.classList.add('completed');
            }
        }
        if (stageTrackerContainer) stageTrackerContainer.classList.remove('running');
        if (headerStageIcon) headerStageIcon.innerText = rawStage.includes('COMPLETED') ? '✅' : '⚡';
        return;
    }

    if (stageTrackerContainer) stageTrackerContainer.classList.add('running');
    if (headerStagePill) {
        headerStagePill.className = 'active-stage-pill running';
        if (svc === 'epfo') headerStagePill.classList.add('epfo-stage');
    }

    // Determine active step index
    let activeStep = 0;
    if (rawStage.includes('STEP 1') || rawStage.includes('CLEAN')) activeStep = 1;
    else if (rawStage.includes('STEP 2') || rawStage.includes('PORTAL') || rawStage.includes('DOWNLOAD')) activeStep = 2;
    else if (rawStage.includes('STEP 3') || rawStage.includes('MERGE')) activeStep = 3;
    else if (rawStage.includes('STEP 4') || rawStage.includes('DB') || rawStage.includes('POSTGRESQL')) activeStep = 4;
    else if (rawStage.includes('STEP 5') || rawStage.includes('COMPARISON') || rawStage.includes('COMPARE')) activeStep = 5;
    else if (rawStage.includes('STEP 6') || rawStage.includes('HRMS') || rawStage.includes('UPLOAD')) activeStep = 6;
    else if (rawStage.includes('STEP 7') || rawStage.includes('EMAIL') || rawStage.includes('MAIL')) activeStep = 7;

    // Highlight badges
    if (activeStep >= 1 && activeStep <= 7) {
        for (let i = 1; i < activeStep; i++) {
            const b = document.getElementById(`step-badge-${i}`);
            if (b) b.classList.add('completed');
        }
        const curr = document.getElementById(`step-badge-${activeStep}`) || (activeStep === 7 ? mailBadge : null);
        if (curr) curr.classList.add('active');
        if (activeStep === 7) {
            if (headerStagePill) headerStagePill.classList.add('mailing');
            if (headerStageIcon) headerStageIcon.innerText = '📧';
        } else {
            if (headerStageIcon) headerStageIcon.innerText = '🔄';
        }
    }
}

function switchService(service) {
    if (currentService === service) return;
    currentService = service;

    const quickEmailBtn = document.getElementById('btn-quick-send-email');
    const triggerEmailBtn = document.getElementById('btn-trigger-email');
    const viewReportBtn = document.getElementById('btn-generate-report-only');
    const bannerDesc = document.getElementById('report-banner-recipients-desc');

    if (service === 'esic') {
        tabEsic.classList.add('active', 'esic-tab');
        tabEpfo.classList.remove('active', 'epfo-tab');
        btnRunSingle.className = 'btn-primary esic-btn';
        btnRunPipeline.className = 'btn-primary esic-btn';
        document.getElementById('service-title-text').innerText = 'ESIC Compliance Operations';
        updatePipelineStepsForEsic();
        if (quickEmailBtn) quickEmailBtn.innerHTML = `<span>📧</span> <span>Email ESIC Report</span>`;
        if (triggerEmailBtn) triggerEmailBtn.innerHTML = `<span>🚀</span> <span>Send ESIC Mail Now</span>`;
        if (viewReportBtn) viewReportBtn.innerHTML = `<span>📑</span> <span>View ESIC Failure Excel</span>`;
        if (bannerDesc) bannerDesc.innerHTML = `Fresh Excel of <strong>Failed & Not Downloaded ESIC</strong> accounts automatically generated and emailed to: <code id="banner-to-email">arun_mishra@uistech.in</code>`;
        const chipPortalEsic = document.getElementById('chip-portal-url');
        if (chipPortalEsic) {
            chipPortalEsic.href = 'https://portal.esic.gov.in/ESICInsurance1/RevenueOne/MonthlyContribution/EmployerLogin.aspx';
            chipPortalEsic.innerHTML = '<span>🌐 Step 2: ESIC Portal Login</span>';
        }
    } else {
        tabEpfo.classList.add('active', 'epfo-tab');
        tabEsic.classList.remove('active', 'esic-tab');
        btnRunSingle.className = 'btn-primary epfo-btn';
        btnRunPipeline.className = 'btn-primary epfo-btn';
        document.getElementById('service-title-text').innerText = 'EPFO Compliance Operations';
        updatePipelineStepsForEpfo();
        if (quickEmailBtn) quickEmailBtn.innerHTML = `<span>📧</span> <span>Email EPFO Report</span>`;
        if (triggerEmailBtn) triggerEmailBtn.innerHTML = `<span>🚀</span> <span>Send EPFO Mail Now</span>`;
        if (viewReportBtn) viewReportBtn.innerHTML = `<span>📑</span> <span>View EPFO Failure Excel</span>`;
        if (bannerDesc) bannerDesc.innerHTML = `Fresh Excel of <strong>Failed & Not Downloaded EPFO</strong> accounts automatically generated and emailed to: <code id="banner-to-email">arun_mishra@uistech.in</code>`;
        const chipPortalEpfo = document.getElementById('chip-portal-url');
        if (chipPortalEpfo) {
            chipPortalEpfo.href = 'https://unifiedportal-emp.epfindia.gov.in/epfo/';
            chipPortalEpfo.innerHTML = '<span>🌐 Step 2: EPFO Portal Login</span>';
        }
    }

    singleUserIdInput.value = '';
    singlePasswordInput.value = '';
    clearSelectedCodes();

    loadAccounts();
}

function updatePipelineStepsForEsic() {
    pipelineStepSelect.innerHTML = `
        <option value="">Run Full Pipeline (All 7 Steps)</option>
        <option value="1">Step 1: Clean Old ESIC Data</option>
        <option value="2">Step 2: ESIC Portal Download</option>
        <option value="3">Step 3: Merge & Column Rename</option>
        <option value="4">Step 4: PostgreSQL DB Download</option>
        <option value="5">Step 5: Data Comparison</option>
        <option value="6">Step 6: HRMS ESIC Master Upload</option>
        <option value="7">Step 7: Send Email Status Report</option>
    `;
}

function updatePipelineStepsForEpfo() {
    pipelineStepSelect.innerHTML = `
        <option value="">Run Full Pipeline (All 7 Steps)</option>
        <option value="1">Step 1: Old Data Cleanup</option>
        <option value="2">Step 2: EPFO Portal Download</option>
        <option value="3">Step 3: Excel Merge</option>
        <option value="4">Step 4: Database Download</option>
        <option value="5">Step 5: Data Comparison</option>
        <option value="6">Step 6: HRMS Upload</option>
        <option value="7">Step 7: Send Email Status Report</option>
    `;
}

// Fetch Accounts from Backend
async function loadAccounts() {
    try {
        const res = await fetch(`/api/accounts?service=${currentService}`);
        if (!res.ok) throw new Error('Failed to load accounts');
        const data = await res.json();

        accountsData = data.accounts || [];
        totalAccountsEl.innerText = data.total;
        successAccountsEl.innerText = data.success;
        failedAccountsEl.innerText = data.failed;
        drivePathEl.innerText = `D:\\Complience (${currentService.toUpperCase()})`;

        if (!isProcessRunning) {
            statLiveDownloadedEl.innerText = data.success;
            headerLiveDownloadedEl.innerText = data.success;
        }

        const failedCountEl = document.getElementById('failed-codes-count');
        if (failedCountEl) failedCountEl.innerText = data.failed || 0;

        renderAccountsTable(searchInput.value.trim().toLowerCase());
        updateSelectionUI();
        populateSingleCodeSelect();
    } catch (err) {
        showToast('Error loading accounts: ' + err.message, 'error');
    }
}

// Render Accounts Table
function renderAccountsTable(filterText = '') {
    accountsTableBody.innerHTML = '';

    const filtered = accountsData.filter(item => {
        if (!filterText) return true;
        return (item.loginid || '').toLowerCase().includes(filterText) ||
               (item.response || '').toLowerCase().includes(filterText);
    });

    if (filtered.length === 0) {
        accountsTableBody.innerHTML = `
            <tr>
                <td colspan="6" style="text-align: center; color: var(--text-muted); padding: 30px;">
                    No accounts found in Excel sheet. Click "+ Add Account" to create one.
                </td>
            </tr>
        `;
        return;
    }

    filtered.forEach((acc, index) => {
        const tr = document.createElement('tr');
        const isChecked = selectedLoginIds.has(acc.loginid);
        if (isChecked) tr.classList.add('row-selected');

        const resp = acc.response || 'PENDING';
        let badgeClass = 'status-pending';
        let badgeIcon = '⏳';
        if (resp.toUpperCase().includes('SUCCESS')) {
            badgeClass = 'status-success';
            badgeIcon = '✓';
        } else if (resp.toUpperCase().includes('FAIL') || resp.toUpperCase().includes('ERROR')) {
            badgeClass = 'status-error';
            badgeIcon = '✗';
        }

        tr.innerHTML = `
            <td style="text-align: center; width: 38px;">
                <input type="checkbox" class="acc-row-checkbox" data-loginid="${escapeHtml(acc.loginid)}" ${isChecked ? 'checked' : ''}>
            </td>
            <td style="color: var(--text-muted); width: 50px;">${index + 1}</td>
            <td class="td-mono" style="font-weight: 600; color: #93c5fd;">${escapeHtml(acc.loginid)}</td>
            <td>
                <div style="display: flex; align-items: center; gap: 8px;">
                    <span class="password-masked" id="pwd-text-${acc.row}">••••••••</span>
                    <button class="btn-icon-action" style="padding: 2px 6px;" title="Show/Hide Password" onclick="togglePasswordView(${acc.row}, '${escapeHtml(acc.password)}')">
                        👁
                    </button>
                </div>
            </td>
            <td>
                <span class="status-badge ${badgeClass}" title="${escapeHtml(resp)}">
                    ${badgeIcon} ${escapeHtml(resp)}
                </span>
            </td>
            <td>
                <div class="action-buttons">
                    <button class="btn-row-run" title="Run full pipeline (Steps 1 to 7) for this code" onclick="quickRunDirect('${escapeHtml(acc.loginid)}')">
                        ⚡ Run
                    </button>
                    <button class="btn-icon-action" title="Prefill in single form" onclick="prefillAndRunSingle('${escapeHtml(acc.loginid)}', '${escapeHtml(acc.password)}')">
                        📝
                    </button>
                    <button class="btn-icon-action" title="Edit credentials" onclick="openAccountModal(${acc.row}, '${escapeHtml(acc.loginid)}', '${escapeHtml(acc.password)}')">
                        ✏
                    </button>
                    <button class="btn-icon-action delete-btn" title="Delete account" onclick="handleDeleteAccount(${acc.row}, '${escapeHtml(acc.loginid)}')">
                        🗑
                    </button>
                </div>
            </td>
        `;

        const cb = tr.querySelector('.acc-row-checkbox');
        if (cb) {
            cb.addEventListener('change', (e) => {
                const code = e.target.getAttribute('data-loginid');
                if (e.target.checked) {
                    selectedLoginIds.add(code);
                    tr.classList.add('row-selected');
                } else {
                    selectedLoginIds.delete(code);
                    tr.classList.remove('row-selected');
                }
                updateSelectionUI();
            });
        }

        tr.addEventListener('click', (e) => {
            if (e.target.closest('button') || e.target.closest('input')) return;
            prefillAndRunSingle(acc.loginid, acc.password);
        });

        accountsTableBody.appendChild(tr);
    });
    updateSelectionUI();
}

function togglePasswordView(row, plainPassword) {
    const el = document.getElementById(`pwd-text-${row}`);
    if (el.innerText === '••••••••') {
        el.innerText = plainPassword;
    } else {
        el.innerText = '••••••••';
    }
}

// Populate Single Code Select Dropdown
function populateSingleCodeSelect() {
    const sel = document.getElementById('single-code-select');
    if (!sel) return;
    const currentVal = sel.value;
    sel.innerHTML = '<option value="">-- Choose Account from Excel --</option>';
    accountsData.forEach(acc => {
        const opt = document.createElement('option');
        opt.value = acc.loginid;
        const respTag = acc.response ? ` (${acc.response.substring(0, 24)})` : '';
        opt.textContent = `${acc.loginid}${respTag}`;
        sel.appendChild(opt);
    });
    if (currentVal) sel.value = currentVal;
}

// Prefill and Run Single Account
function prefillAndRunSingle(loginid, password) {
    const sel = document.getElementById('single-code-select');
    if (sel) sel.value = loginid;
    singleUserIdInput.value = loginid;
    singlePasswordInput.value = password;
    singleUserIdInput.scrollIntoView({ behavior: 'smooth', block: 'center' });
    showToast(`Loaded ${loginid} for single execution`, 'success');
}

// Single Run Execution
async function handleRunSingle() {
    let loginid = singleUserIdInput.value.trim();
    let password = singlePasswordInput.value.trim();
    const headless = singleHeadlessCb.checked;

    if (!loginid) {
        showToast('Please select or enter an Account / User ID', 'error');
        return;
    }

    // Auto-lookup password from Excel if empty
    if (!password) {
        const found = accountsData.find(a => String(a.loginid).trim() === loginid);
        if (found && found.password) {
            password = found.password;
            singlePasswordInput.value = password;
        }
    }

    if (isProcessRunning) {
        showToast('A process is already active. Please wait or stop it.', 'error');
        return;
    }

    try {
        btnRunSingle.disabled = true;
        btnRunSingle.innerHTML = `<span>⏳</span> Running Steps 1-7 for ${loginid}...`;

        const res = await fetch('/api/run-single', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                service: currentService,
                loginid: loginid,
                password: password,
                headless: headless
            })
        });

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || 'Failed to start execution');
        }

        showToast(`Full pipeline (Steps 1 to 7) started for ${loginid}`, 'success');
        checkProcessStatus();

        // Scroll to terminal to see live logs
        const term = document.querySelector('.terminal-section');
        if (term) term.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    } catch (err) {
        showToast(err.message, 'error');
        btnRunSingle.disabled = false;
        btnRunSingle.innerHTML = `<span>⚡</span> Run Pipeline for This Code (Steps 1 - 7)`;
    }
}

// Pipeline Run Execution
async function handleRunPipeline() {
    if (isProcessRunning) {
        showToast('A process is already active', 'error');
        return;
    }

    const stepVal = pipelineStepSelect.value ? parseInt(pipelineStepSelect.value) : null;
    const headless = pipelineHeadlessCb.checked;

    try {
        btnRunPipeline.disabled = true;
        btnRunPipeline.innerHTML = `<span>⏳</span> Pipeline Starting...`;

        const res = await fetch('/api/run-pipeline', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                service: currentService,
                step: stepVal,
                headless: headless
            })
        });

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || 'Failed to start pipeline');
        }

        showToast(`${currentService.toUpperCase()} Pipeline started`, 'success');
        checkProcessStatus();
    } catch (err) {
        showToast(err.message, 'error');
        btnRunPipeline.disabled = false;
        btnRunPipeline.innerHTML = `<span>🚀</span> Start Automation Pipeline`;
    }
}

// ============================================================
// BATCH ACTIONS, SELECTION & DIRECT QUICK RUN
// ============================================================
function updateSelectionUI() {
    const count = selectedLoginIds.size;
    const btnRunSelected = document.getElementById('btn-run-selected-codes');
    const btnClear = document.getElementById('btn-clear-selection');
    const selectedCountEl = document.getElementById('selected-codes-count');
    const selectAllCb = document.getElementById('select-all-accounts');

    if (selectedCountEl) selectedCountEl.innerText = count;

    if (count > 0) {
        if (btnRunSelected) btnRunSelected.style.display = 'inline-flex';
        if (btnClear) btnClear.style.display = 'inline-flex';
    } else {
        if (btnRunSelected) btnRunSelected.style.display = 'none';
        if (btnClear) btnClear.style.display = 'none';
    }

    if (selectAllCb) {
        const checkboxes = document.querySelectorAll('.acc-row-checkbox');
        if (checkboxes.length > 0) {
            const allChecked = Array.from(checkboxes).every(cb => cb.checked);
            const someChecked = Array.from(checkboxes).some(cb => cb.checked);
            selectAllCb.checked = allChecked;
            selectAllCb.indeterminate = someChecked && !allChecked;
        } else {
            selectAllCb.checked = false;
            selectAllCb.indeterminate = false;
        }
    }
}

function clearSelectedCodes() {
    selectedLoginIds.clear();
    document.querySelectorAll('.acc-row-checkbox').forEach(cb => {
        cb.checked = false;
    });
    document.querySelectorAll('#accounts-tbody tr').forEach(tr => {
        tr.classList.remove('row-selected');
    });
    const selectAllCb = document.getElementById('select-all-accounts');
    if (selectAllCb) {
        selectAllCb.checked = false;
        selectAllCb.indeterminate = false;
    }
    updateSelectionUI();
}

function handleSelectAllToggle(e) {
    const isChecked = e.target.checked;
    const checkboxes = document.querySelectorAll('.acc-row-checkbox');
    checkboxes.forEach(cb => {
        cb.checked = isChecked;
        const code = cb.getAttribute('data-loginid');
        const row = cb.closest('tr');
        if (isChecked) {
            if (code) selectedLoginIds.add(code);
            if (row) row.classList.add('row-selected');
        } else {
            if (code) selectedLoginIds.delete(code);
            if (row) row.classList.remove('row-selected');
        }
    });
    updateSelectionUI();
}

async function handleRetryFailedCodes() {
    if (isProcessRunning) {
        showToast('A process is already active. Please wait or stop it first.', 'error');
        return;
    }

    try {
        const res = await fetch(`/api/failed-accounts?service=${currentService}`);
        if (!res.ok) throw new Error('Could not fetch failed accounts');
        const data = await res.json();
        const count = data.count || 0;

        if (count === 0) {
            showToast('No failed accounts found in Excel sheet!', 'info');
            return;
        }

        const confirmMsg = `Found ${count} failed code(s) in ${currentService.toUpperCase()} Excel. Retry portal download for these ${count} code(s)?`;
        if (!confirm(confirmMsg)) return;

        showToast(`Running full pipeline (Steps 1 to 7) for ${count} failed account(s)...`, 'info');

        const runRes = await fetch('/api/run-pipeline', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                service: currentService,
                only_failed: true
            })
        });

        if (!runRes.ok) {
            const err = await runRes.json();
            throw new Error(err.detail || 'Failed to start retry run');
        }

        checkProcessStatus();
    } catch (err) {
        showToast(err.message, 'error');
    }
}

async function handleRunSelectedCodes() {
    if (isProcessRunning) {
        showToast('A process is already active. Please wait or stop it first.', 'error');
        return;
    }

    if (selectedLoginIds.size === 0) {
        showToast('Please check 1 or more codes in the table first', 'info');
        return;
    }

    const count = selectedLoginIds.size;
    const codesArr = Array.from(selectedLoginIds);

    // If exactly 1 code is selected, run via single-run endpoint with immediate output
    if (count === 1) {
        const singleCode = codesArr[0];
        const found = accountsData.find(a => String(a.loginid).trim() === String(singleCode).trim());
        const pwd = found ? (found.password || '') : '';

        // Sync with left inputs
        const sel = document.getElementById('single-code-select');
        if (sel) sel.value = singleCode;
        singleUserIdInput.value = singleCode;
        if (pwd) singlePasswordInput.value = pwd;

        try {
            showToast(`Starting full pipeline (Steps 1 to 7) for code ${singleCode}...`, 'info');
            const res = await fetch('/api/run-single', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    service: currentService,
                    loginid: singleCode,
                    password: pwd,
                    headless: false
                })
            });

            if (!res.ok) {
                const err = await res.json();
                throw new Error(err.detail || 'Failed to start execution');
            }

            showToast(`Running Steps 1 to 7 for: ${singleCode}`, 'success');
            checkProcessStatus();
            const term = document.querySelector('.terminal-section');
            if (term) term.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        } catch (err) {
            showToast(err.message, 'error');
        }
        return;
    }

    // Multiple codes: run full pipeline (Steps 1 to 7) with target codes
    const codesStr = codesArr.join(',');
    try {
        showToast(`Starting download for ${count} selected code(s)...`, 'info');
        const res = await fetch('/api/run-pipeline', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                service: currentService,
                codes: codesStr
            })
        });

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || 'Failed to start execution');
        }

        showToast(`Processing ${count} selected code(s)`, 'success');
        checkProcessStatus();
        const term = document.querySelector('.terminal-section');
        if (term) term.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    } catch (err) {
        showToast(err.message, 'error');
    }
}

async function quickRunDirect(loginid) {
    if (isProcessRunning) {
        showToast('A process is already active. Please wait or stop it first.', 'error');
        return;
    }

    try {
        const found = accountsData.find(a => String(a.loginid).trim() === String(loginid).trim());
        const pwd = found ? (found.password || '') : '';

        // Pre-fill inputs
        const sel = document.getElementById('single-code-select');
        if (sel) sel.value = loginid;
        singleUserIdInput.value = loginid;
        if (pwd) singlePasswordInput.value = pwd;

        showToast(`Starting full pipeline (Steps 1 to 7) for code ${loginid}...`, 'info');
        const res = await fetch('/api/run-single', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                service: currentService,
                loginid: loginid,
                password: pwd,
                headless: false
            })
        });

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || 'Failed to start single execution');
        }

        showToast(`Running Steps 1 to 7 for ${loginid}`, 'success');
        checkProcessStatus();
        const term = document.querySelector('.terminal-section');
        if (term) term.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    } catch (err) {
        showToast(err.message, 'error');
    }
}

// Open Specific Step Output File or Directory on Windows Host
async function openStepTargetFile(step) {
    try {
        const res = await fetch(`/api/open-step-target?service=${currentService}&step=${step}`, { method: 'POST' });
        const data = await res.json();
        if (!res.ok) throw new Error(data.detail || 'Could not open file');
        showToast(data.message || `Opened Step ${step} target`, 'success');
    } catch (err) {
        showToast(err.message, 'error');
    }
}

// ============================================================
// STEP BADGE RUNNER MODAL LOGIC
// ============================================================
function openStepRunnerModal(stepNum, stepTitle) {
    currentModalStep = stepNum;
    currentModalTitle = stepTitle;

    const modal = document.getElementById('modal-step-runner');
    const numEl = document.getElementById('modal-step-num');
    const headingEl = document.getElementById('modal-step-heading');
    const descEl = document.getElementById('modal-step-desc');
    const singleSub = document.getElementById('modal-step-single-sub');
    const fromSub = document.getElementById('modal-step-from-sub');
    const codesGroup = document.getElementById('modal-step-codes-group');
    const codesInput = document.getElementById('modal-step-codes');
    const btnText = document.getElementById('modal-step-btn-text');

    if (numEl) numEl.innerText = stepNum;
    if (headingEl) headingEl.innerText = stepTitle;
    if (descEl) descEl.innerText = `You can execute Step ${stepNum} individually or resume the full pipeline from here.`;
    if (singleSub) singleSub.innerText = `Execute ${stepTitle} only, then stop.`;
    if (fromSub) fromSub.innerText = `Execute sequentially from Step ${stepNum} through Step 7 (Send Email).`;
    if (btnText) btnText.innerText = `Execute Step ${stepNum}`;

    if (codesGroup) {
        if (stepNum === 2) {
            codesGroup.style.display = 'block';
            if (selectedLoginIds.size > 0 && !codesInput.value) {
                codesInput.value = Array.from(selectedLoginIds).join(', ');
            }
        } else {
            codesGroup.style.display = 'none';
        }
    }

    // Configure Manual Work & Direct Links inside modal
    const portalLink = document.getElementById('modal-step-portal-link');
    const portalBtnText = document.getElementById('modal-step-portal-btn-text');
    const fileBtn = document.getElementById('modal-step-file-btn');
    const fileBtnText = document.getElementById('modal-step-file-btn-text');
    const manualDesc = document.getElementById('modal-step-manual-desc');

    if (portalLink && fileBtn) {
        if (stepNum === 2) {
            portalLink.style.display = 'inline-flex';
            portalLink.href = currentService === 'esic'
                ? 'https://portal.esic.gov.in/ESICInsurance1/RevenueOne/MonthlyContribution/EmployerLogin.aspx'
                : 'https://unifiedportal-emp.epfindia.gov.in/epfo/';
            if (portalBtnText) portalBtnText.innerText = `Open ${currentService.toUpperCase()} Official Portal`;
            if (manualDesc) manualDesc.innerText = `Login to official portal in browser or open downloaded accounts folder.`;
        } else if (stepNum === 6) {
            portalLink.style.display = 'inline-flex';
            portalLink.href = 'https://hrms.uistech.in/master/epfo-esic-master';
            if (portalBtnText) portalBtnText.innerText = 'Open HRMS Master Upload Portal';
            if (manualDesc) manualDesc.innerText = `Open HRMS portal in browser to upload Master CSV manually, or open the CSV file directly.`;
        } else {
            portalLink.style.display = 'none';
            if (manualDesc) manualDesc.innerText = `View or edit the generated Excel/CSV file for this step.`;
        }

        const stepFileLabels = {
            1: 'Open Data Folder',
            2: 'Open Downloaded Folder',
            3: 'Open Merged Excel File',
            4: 'Open PostgreSQL DB Data File',
            5: 'Open Comparison Output File',
            6: 'Open HRMS_Upload.csv File',
            7: 'Open Email Reports Folder'
        };
        if (fileBtnText) fileBtnText.innerText = stepFileLabels[stepNum] || 'Open Step File / Folder';
        fileBtn.onclick = () => openStepTargetFile(stepNum);
    }

    if (modal) modal.classList.add('active');
}

function closeStepRunnerModal() {
    const modal = document.getElementById('modal-step-runner');
    if (modal) modal.classList.remove('active');
}

async function executeStepModalAction() {
    if (isProcessRunning) {
        showToast('A process is already active. Please wait or stop it first.', 'error');
        return;
    }

    const mode = document.querySelector('input[name="step-execution-mode"]:checked')?.value || 'single_step';
    const isSingle = mode === 'single_step';
    const codesInput = document.getElementById('modal-step-codes');
    const codes = codesInput ? codesInput.value.trim() : '';
    const headlessCb = document.getElementById('modal-step-headless');
    const headless = headlessCb ? headlessCb.checked : false;

    closeStepRunnerModal();

    // Clear failed highlight for this step if user retries it
    const badge = document.getElementById(`step-badge-${currentModalStep}`);
    if (badge) badge.classList.remove('failed');

    try {
        const payload = {
            service: currentService,
            headless: headless
        };

        if (isSingle) {
            payload.step = currentModalStep;
        } else {
            payload.from_step = currentModalStep;
        }

        if (currentModalStep === 2 && codes) {
            payload.codes = codes;
        }

        const res = await fetch('/api/run-pipeline', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || 'Failed to start execution');
        }

        const modeLabel = isSingle ? `Step ${currentModalStep} only` : `Pipeline from Step ${currentModalStep} to 7`;
        showToast(`Running ${modeLabel}...`, 'success');
        checkProcessStatus();
    } catch (err) {
        showToast(err.message, 'error');
    }
}

// Stop Process
async function handleStopProcess() {
    try {
        const res = await fetch('/api/stop', { method: 'POST' });
        const data = await res.json();
        showToast(data.message || 'Process terminated', 'success');
        checkProcessStatus();
    } catch (err) {
        showToast('Error stopping process: ' + err.message, 'error');
    }
}

// Check Process Status
async function checkProcessStatus() {
    try {
        const res = await fetch('/api/status');
        const data = await res.json();

        isProcessRunning = data.running;

        // Update live stage UI
        updateStageUI(data.stage, data.running, data.service);

        if (data.running) {
            serverStatusIndicator.classList.add('running');
            serverStatusText.innerText = `Running: ${data.mode || 'Task'} (${data.elapsed}s)`;
            btnStopPipeline.style.display = 'flex';
            btnRunPipeline.disabled = true;
            btnRunSingle.disabled = true;

            // Update live download counts
            if (data.downloaded_count !== undefined) {
                liveDownloadedCount = data.downloaded_count;
                statLiveDownloadedEl.innerText = liveDownloadedCount;
                headerLiveDownloadedEl.innerText = liveDownloadedCount;
            }
        } else {
            serverStatusIndicator.classList.remove('running');
            serverStatusText.innerText = 'Ready / Idle';
            btnStopPipeline.style.display = 'none';
            btnRunPipeline.disabled = false;
            btnRunSingle.disabled = false;
            btnRunSingle.innerHTML = `<span>⚡</span> Download Data for This Code`;
            btnRunPipeline.innerHTML = `<span>🚀</span> Start Automation Pipeline`;
        }
    } catch (err) {
        // ignore network glitches
    }
}

// Live SSE Log Stream with real-time download tracking & stage updates
function connectLogStream() {
    if (eventSource) {
        eventSource.close();
    }

    eventSource = new EventSource('/api/logs/stream');

    eventSource.onmessage = (event) => {
        try {
            const payload = JSON.parse(event.data);
            const line = payload.line;

            // Update live stage from SSE stream
            if (payload.stage !== undefined) {
                updateStageUI(payload.stage, payload.running, currentService);
            }

            // Real-time counter update from payload
            if (payload.downloaded !== undefined && payload.downloaded > 0) {
                liveDownloadedCount = payload.downloaded;
                statLiveDownloadedEl.innerText = liveDownloadedCount;
                headerLiveDownloadedEl.innerText = liveDownloadedCount;
            }

            appendLogLine(line);
        } catch (e) {
            appendLogLine(event.data);
        }
    };

    eventSource.onerror = () => {
        setTimeout(connectLogStream, 3000);
    };
}

function appendLogLine(line) {
    if (!line) return;

    const div = document.createElement('div');
    div.className = 'terminal-line';

    const lower = line.toLowerCase();
    if (line.includes('[SYSTEM]')) {
        div.classList.add('system');
    } else if (line.includes('[REAL-TIME]') || lower.includes('downloaded!')) {
        div.classList.add('success');
        div.style.fontWeight = '700';
    } else if (lower.includes('success') || lower.includes('[ ok ]')) {
        div.classList.add('success');
    } else if (lower.includes('failed') || lower.includes('error') || lower.includes('[ fail ]') || lower.includes('exception')) {
        div.classList.add('error');
    } else if (lower.includes('warn')) {
        div.classList.add('warn');
    } else if (lower.includes('step ') || lower.includes('opening ')) {
        div.classList.add('info');
    }

    div.innerText = line;
    terminalOutput.appendChild(div);

    if (autoScroll) {
        terminalOutput.scrollTop = terminalOutput.scrollHeight;
    }

    // Detect failed step from logs
    const failMatch = line.match(/Step\s*([1-7]):.*\[\s*FAIL\s*\]/i);
    if (failMatch) {
        const stepNo = failMatch[1];
        const badge = document.getElementById(`step-badge-${stepNo}`);
        if (badge) {
            badge.classList.remove('active');
            badge.classList.add('failed');
        }
        showToast(`⚠️ Step ${stepNo} Failed! Click the Step ${stepNo} badge above to retry or resume.`, 'error');
    }

    // Live update account list when an account finishes
    if (line.includes('Excel updated') || line.includes('SUCCESS -') || line.includes('FAILED -')) {
        loadAccounts();
    }
}

// Account Modal Operations
function openAccountModal(row = null, loginid = '', password = '') {
    editingRow = row;
    if (row) {
        modalTitle.innerText = `Edit Account (Row ${row})`;
        modalUserId.value = loginid;
        modalPassword.value = password;
    } else {
        modalTitle.innerText = `Add New Account to ${currentService.toUpperCase()}`;
        modalUserId.value = '';
        modalPassword.value = '';
    }
    modalAccount.classList.add('active');
    modalUserId.focus();
}

function closeAccountModal() {
    modalAccount.classList.remove('active');
    editingRow = null;
}

async function handleSaveAccount() {
    const loginid = modalUserId.value.trim();
    const password = modalPassword.value.trim();

    if (!loginid || !password) {
        showToast('User ID and Password are required', 'error');
        return;
    }

    try {
        btnSaveAccount.disabled = true;

        if (editingRow) {
            const res = await fetch(`/api/accounts/${editingRow}`, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    service: currentService,
                    loginid: loginid,
                    password: password
                })
            });
            if (!res.ok) throw new Error('Update failed');
            showToast('Account updated successfully in Excel', 'success');
        } else {
            const res = await fetch('/api/accounts', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    service: currentService,
                    loginid: loginid,
                    password: password
                })
            });
            if (!res.ok) throw new Error('Create failed');
            showToast('Account added successfully to Excel', 'success');
        }

        closeAccountModal();
        loadAccounts();
    } catch (err) {
        showToast(err.message, 'error');
    } finally {
        btnSaveAccount.disabled = false;
    }
}

async function handleDeleteAccount(row, loginid) {
    if (!confirm(`Are you sure you want to delete account "${loginid}" (Row ${row}) from Excel?`)) {
        return;
    }

    try {
        const res = await fetch(`/api/accounts/${row}?service=${currentService}`, {
            method: 'DELETE'
        });
        if (!res.ok) throw new Error('Failed to delete account');
        showToast(`Account ${loginid} removed from Excel`, 'success');
        loadAccounts();
    } catch (err) {
        showToast(err.message, 'error');
    }
}

// ============================================================
// OUTLOOK REPORT EMAIL & RECIPIENTS LOGIC (SERVICE-SPECIFIC)
// ============================================================
async function handleSendFailureEmail() {
    const btn = document.getElementById('btn-quick-send-email');
    const btn2 = document.getElementById('btn-trigger-email');
    const svcUpper = currentService.toUpperCase();

    try {
        btn.disabled = true;
        btn2.disabled = true;
        btn.innerHTML = `<span>⏳</span> Generating & Sending ${svcUpper}...`;
        btn2.innerHTML = `<span>⏳</span> Sending...`;

        const res = await fetch(`/api/reports/send-email?service=${currentService}`, { method: 'POST' });
        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || 'Email send failed');
        }
        const data = await res.json();

        showToast(`${svcUpper} Email successfully sent! Attached: ${data.report_name}`, 'success');
    } catch (err) {
        showToast('Error sending email: ' + err.message, 'error');
    } finally {
        btn.disabled = false;
        btn2.disabled = false;
        btn.innerHTML = `<span>📧</span> <span>Email ${svcUpper} Report</span>`;
        btn2.innerHTML = `<span>🚀</span> <span>Send ${svcUpper} Mail Now</span>`;
    }
}

async function handleGenerateReportOnly() {
    const svcUpper = currentService.toUpperCase();
    try {
        const res = await fetch(`/api/reports/generate-excel?service=${currentService}`, { method: 'POST' });
        if (!res.ok) throw new Error('Report generation failed');
        const data = await res.json();

        showToast(`Generated: ${data.report_name} (${data.failed_count || data.total_count || 0} issues)`, 'success');
        // Open reports folder so user can see the file
        openPathOnHost('reports_dir');
    } catch (err) {
        showToast('Error generating report: ' + err.message, 'error');
    }
}

async function loadRecipients() {
    try {
        const res = await fetch('/api/reports/recipients');
        if (!res.ok) return;
        const data = await res.json();
        recipientsToInput.value = (data.to || []).join(', ');
        recipientsCcInput.value = (data.cc || []).join(', ');

        const bannerTo = document.getElementById('banner-to-email');
        const bannerCc = document.getElementById('banner-cc-emails');
        if (bannerTo && data.to && data.to.length) {
            bannerTo.innerText = data.to.join(', ');
        }
        if (bannerCc && data.cc && data.cc.length) {
            bannerCc.innerText = data.cc.join(', ');
        }
    } catch (err) {
        // ignore
    }
}

function openRecipientsModal() {
    loadRecipients();
    modalRecipients.classList.add('active');
}

function closeRecipientsModal() {
    modalRecipients.classList.remove('active');
}

async function handleSaveRecipients() {
    const toList = recipientsToInput.value.split(',').map(s => s.trim()).filter(Boolean);
    const ccList = recipientsCcInput.value.split(',').map(s => s.trim()).filter(Boolean);

    if (toList.length === 0) {
        showToast('At least one TO recipient is required', 'error');
        return;
    }

    try {
        btnSaveRecipients.disabled = true;
        const res = await fetch('/api/reports/recipients', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ to: toList, cc: ccList })
        });
        if (!res.ok) throw new Error('Save failed');

        showToast('Email recipients updated and saved to Excel!', 'success');
        closeRecipientsModal();
        await loadRecipients();
    } catch (err) {
        showToast(err.message, 'error');
    } finally {
        btnSaveRecipients.disabled = false;
    }
}

// Open Path on host machine in Windows Explorer
async function openPathOnHost(target) {
    try {
        const res = await fetch('/api/open-path', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ target: target })
        });
        const data = await res.json();
        showToast(data.message || 'Opened in Windows Explorer', 'success');
    } catch (err) {
        showToast('Could not open path: ' + err.message, 'error');
    }
}

// Toast Notification
function showToast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.innerText = message;
    container.appendChild(toast);

    setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transform = 'translateY(10px)';
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}

function escapeHtml(text) {
    if (text === null || text === undefined) return '';
    return String(text)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}
