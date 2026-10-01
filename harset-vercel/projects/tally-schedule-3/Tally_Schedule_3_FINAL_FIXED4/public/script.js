document.addEventListener('DOMContentLoaded', () => {
    // State to hold data for merging
    let leftData = null;
    let rightData = null;
    let leftName = '';
    let rightName = '';
    let leftCompanyType = '';
    let rightCompanyType = '';
    // Per-side Pvt Ltd share capital info
    let leftPvtInfo = null;
    let rightPvtInfo = null;
    // Pending modal resolution
    let modalResolve = null;
    // Flag to prevent concurrent generation
    let isGenerating = false;

    // ── Helpers ─────────────────────────────────────────────────────────────
    function isPvtLtd(companyType) {
        return companyType === 'Private Limited';
    }

    // ── Company dropdown helpers ──────────────────────────────────────────
    function populateDropdowns(companies) {
        ['left', 'right'].forEach(side => {
            let el = document.getElementById(`companyName-${side}`);
            if (el && !el.setCompanies) {
                const newEl = document.createElement('company-select');
                newEl.id = `companyName-${side}`;
                newEl.setAttribute('side', side);
                newEl.setAttribute('label', side === 'left' ? 'Company Name (Current Year)' : 'Company Name (Previous Year)');
                newEl.setAttribute('placeholder', '-- Select Company --');
                el.replaceWith(newEl);
                el = newEl;
            }
            if (el && el.setCompanies) el.setCompanies(companies);
        });
    }

    async function loadCompanies() {
        try {
            const tallyHost = localStorage.getItem('tallyHost') || '172.16.1.4';
            const tallyPort = localStorage.getItem('tallyPort') || '9000';
            const res = await fetch('/api/companies', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ tallyHost, tallyPort })
            });
            const data = await res.json();
            const statusEl = document.getElementById('companies-status');
            
            if (data.companies && data.companies.length > 0) {
                populateDropdowns(data.companies);
                if (statusEl) statusEl.textContent = `✓ ${data.companies.length} companies loaded from Tally`;
                console.log(`[Dropdown] ${data.companies.length} companies loaded.`);
            } else {
                if (statusEl) statusEl.textContent = '⚠ No companies found in Tally. Type name manually.';
                // Fallback: convert selects to text inputs
                ['left', 'right'].forEach(side => {
                    const sel = document.getElementById(`companyName-${side}`);
                    if (!sel) return;
                    const inp = document.createElement('input');
                    inp.type = 'text'; inp.id = `companyName-${side}`;
                    inp.name = 'companyName'; inp.required = true;
                    inp.placeholder = 'e.g. ECS'; inp.className = 'company-select';
                    sel.replaceWith(inp);
                });
            }
        } catch (e) {
            const statusEl = document.getElementById('companies-status');
            if (statusEl) statusEl.textContent = '⚠ Error loading companies. Tally server may be unreachable.';
            console.error('[Dropdown] Error:', e.message);
        }
    }
    loadCompanies();

    // ── Pvt Ltd Modal ─────────────────────────────────────────────────────
    function showPvtModal(companyName) {
        return new Promise((resolve) => {
            modalResolve = resolve;
            document.getElementById('modal-company-name').textContent = companyName;
            document.getElementById('modal-auth-capital').value = '';
            document.getElementById('modal-no-shares').value = '';
            document.getElementById('modal-face-value').value = '';
            document.getElementById('pvtltd-modal').classList.remove('hidden');
        });
    }

    document.getElementById('modal-confirm').addEventListener('click', () => {
        const authCap = parseFloat(document.getElementById('modal-auth-capital').value) || 0;
        const noShares = parseFloat(document.getElementById('modal-no-shares').value) || 0;
        const faceVal = parseFloat(document.getElementById('modal-face-value').value) || 0;
        document.getElementById('pvtltd-modal').classList.add('hidden');
        if (modalResolve) modalResolve({ authCap, noShares, faceVal });
    });

    document.getElementById('modal-skip').addEventListener('click', () => {
        document.getElementById('pvtltd-modal').classList.add('hidden');
        if (modalResolve) modalResolve(null);
    });

    // ── Fiscal Year Dropdowns ─────────────────────────────────────────────
    function buildYearDropdowns() {
        const currentYear = new Date().getFullYear();
        const startYear = 2020;

        ['left', 'right'].forEach(side => {
            const fromSel = document.getElementById(`fromYear-${side}`);
            const toSel   = document.getElementById(`toYear-${side}`);

            // Populate options 2020 → current year
            for (let y = startYear; y <= currentYear; y++) {
                const optFrom = document.createElement('option');
                optFrom.value = y;
                optFrom.textContent = `${y}`;
                fromSel.appendChild(optFrom);

                const optTo = document.createElement('option');
                optTo.value = y;
                optTo.textContent = `${y}`;
                toSel.appendChild(optTo);
            }

            // Auto-fill hidden fromDate when From Year changes
            fromSel.addEventListener('change', () => {
                const y = fromSel.value;
                if (!y) {
                    document.getElementById(`fromDate-${side}`).value = '';
                    document.getElementById(`fromHint-${side}`).textContent = '';
                    return;
                }
                // Start: 1 April of selected year  →  YYYYMMDD for the server
                document.getElementById(`fromDate-${side}`).value = `${y}0401`;
                document.getElementById(`fromHint-${side}`).textContent = `→ 01 Apr ${y}`;

                // Auto-suggest matching end year if not set
                if (!toSel.value) {
                    const nextY = parseInt(y) + 1;
                    if (nextY <= currentYear) toSel.value = nextY;
                    if (toSel.value) toSel.dispatchEvent(new Event('change'));
                }
            });

            // Auto-fill hidden toDate when To Year changes
            toSel.addEventListener('change', () => {
                const y = toSel.value;
                if (!y) {
                    document.getElementById(`toDate-${side}`).value = '';
                    document.getElementById(`toHint-${side}`).textContent = '';
                    return;
                }
                // End: 31 March of selected year  →  YYYYMMDD
                document.getElementById(`toDate-${side}`).value = `${y}0331`;
                document.getElementById(`toHint-${side}`).textContent = `→ 31 Mar ${y}`;
            });
        });
    }
    buildYearDropdowns();

    document.getElementById('btn-view-bs-table').addEventListener('click', () => {
        openTabularReport('bs');
    });

    document.getElementById('btn-view-pl-table').addEventListener('click', () => {
        openTabularReport('pl');
    });

    // ── Edit Panel Logic ──────────────────────────────────────────────────
    ['left', 'right'].forEach(side => {
        const editBtn    = document.getElementById(`edit-btn-${side}`);
        const editPanel  = document.getElementById(`edit-panel-${side}`);

        // Toggle panel open/close
        editBtn.addEventListener('click', () => {
            const isOpen = !editPanel.classList.contains('hidden');
            if (isOpen) {
                editPanel.classList.add('hidden');
                editBtn.classList.remove('active');
                editBtn.textContent = '✎ Edit';
            } else {
                // Pre-fill with existing pvtInfo if already applied
                const existing = side === 'left' ? leftPvtInfo : rightPvtInfo;
                if (existing) {
                    document.getElementById(`edit-auth-cap-${side}`).value  = existing.authCap  || '';
                    document.getElementById(`edit-face-val-${side}`).value  = existing.faceVal  || '';
                    document.getElementById(`edit-no-shares-${side}`).value = existing.noShares || '';
                }
                editPanel.classList.remove('hidden');
                editBtn.classList.add('active');
                editBtn.textContent = '✕ Close';
            }
        });

        // Cancel
        document.querySelector(`.edit-cancel-btn[data-side="${side}"]`).addEventListener('click', () => {
            editPanel.classList.add('hidden');
            editBtn.classList.remove('active');
            editBtn.textContent = '✎ Edit';
        });

        // Apply — re-render Note 1 section in the report
        document.querySelector(`.edit-apply-btn[data-side="${side}"]`).addEventListener('click', () => {
            const authCap  = parseFloat(document.getElementById(`edit-auth-cap-${side}`).value)  || 0;
            const faceVal  = parseFloat(document.getElementById(`edit-face-val-${side}`).value)  || 0;
            const noShares = parseFloat(document.getElementById(`edit-no-shares-${side}`).value) || 0;

            const pvtInfo = { authCap, faceVal, noShares };
            if (side === 'left') leftPvtInfo = pvtInfo;
            else rightPvtInfo = pvtInfo;

            const data = side === 'left' ? leftData : rightData;
            if (!data) return;

            // Re-render the full report with updated pvtInfo
            const container = document.getElementById(`output-${side}`);
            renderReport(data, container, pvtInfo);

            // Close panel
            editPanel.classList.add('hidden');
            editBtn.classList.remove('active');
            editBtn.textContent = '✎ Edit';
        });
    });

    function getVal(data, noteNum) {
        if (!data || !data.mapping) return 0;
        const map = data.mapping;
        
        // Standardize finding the note by matching the "Note X -" or "Note X " prefix
        const key = Object.keys(map).find(k => k.startsWith(`Note ${noteNum} -`) || k.startsWith(`Note ${noteNum} `));
        
        if (key && map[key] !== undefined) {
            return map[key];
        }

        // Specific fallbacks if the key naming deviates slightly
        if (noteNum === 1 && map['Note 1 - Share Capital'] !== undefined) return map['Note 1 - Share Capital'];
        if (noteNum === 2 && map['Note 2 - Profit/Loss'] !== undefined) return map['Note 2 - Profit/Loss'];
        
        return 0;
    }

    function fmtTab(num) {
        if (num === '' || num === null || num === undefined) return '-';
        if (num === 0) return '-';
        return Number(num).toLocaleString('en-IN', { maximumFractionDigits: 2, minimumFractionDigits: 2 });
    }

    // Format YYYY-MM-DD to YYYYMMDD
    function formatDate(dateString) {
        return dateString.replace(/-/g, '');
    }

    // Attach listener to both forms
    document.querySelectorAll('.mapping-form').forEach(form => {
        form.addEventListener('submit', async (e) => {
            e.preventDefault();
            
            const side = form.getAttribute('data-side');
            const submitBtn = document.getElementById(`btn-${side}`);
            const btnText = submitBtn.querySelector('.btn-text');
            const spinner = submitBtn.querySelector('.spinner');
            const resultContainer = document.getElementById(`result-${side}`);
            const reportOutput = document.getElementById(`output-${side}`);
            const errorContainer = document.getElementById(`error-${side}`);
            
            // Reset state
            errorContainer.classList.add('hidden');
            resultContainer.classList.add('hidden');
            
            // Get values — use web component API for company name
            const compEl = document.getElementById(`companyName-${side}`);
            const cName = (compEl.getValue ? compEl.getValue() : compEl.value || '').trim();
            const companyType = document.getElementById(`companyType-${side}`).value;
            // Dates are pre-filled by the year dropdowns into hidden inputs (YYYYMMDD)
            const fromDate = document.getElementById(`fromDate-${side}`).value;
            const toDate   = document.getElementById(`toDate-${side}`).value;
            const fromYear = document.getElementById(`fromYear-${side}`).value;
            const toYear   = document.getElementById(`toYear-${side}`).value;

            if (!companyType) {
                showError(side, 'Please select a type of company.');
                return;
            }
            if (!cName) {
                showError(side, 'Please select a company.');
                return;
            }
            if (!fromYear || !toYear) {
                showError(side, 'Please select both the start and end financial year.');
                return;
            }

            if (parseInt(toYear) <= parseInt(fromYear)) {
                showError(side, 'The end financial year must be strictly greater than the start year.');
                return;
            }

            if (isGenerating) {
                showError(side, 'Another generation is in progress. Please wait.');
                return;
            }
            isGenerating = true;

            // Check if Pvt Ltd → show modal first (before loading spinner)
            let pvtInfo = null;
            if (isPvtLtd(companyType)) {
                pvtInfo = await showPvtModal(cName);
            }
            if (side === 'left') { leftPvtInfo = pvtInfo; leftCompanyType = companyType; }
            else { rightPvtInfo = pvtInfo; rightCompanyType = companyType; }

            // Loading state
            submitBtn.disabled = true;
            btnText.textContent = 'Processing...';
            spinner.style.display = 'block';

            try {
                const tallyHost = localStorage.getItem('tallyHost') || '172.16.1.4';
                const tallyPort = localStorage.getItem('tallyPort') || '9000';

                const response = await fetch('/api/run-mapping', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ companyName: cName, fromDate, toDate, pvtLtdInfo: pvtInfo, companyType, tallyHost, tallyPort })
                });

                const data = await response.json();

                if (response.ok && data.success) {
                    renderReport(data.result, reportOutput, pvtInfo);
                    
                    // Store state for merge and view toggles
                    if (side === 'left') {
                        leftData = data.result;
                        leftName = cName;
                    } else {
                        rightData = data.result;
                        rightName = cName;
                    }

                    resultContainer.classList.remove('hidden');
                    checkMergeAvailability();
                } else {
                    throw new Error(data.error || 'Failed to generate mapping');
                }
            } catch (error) {
                showError(side, error.message);
                if (side === 'left') leftData = null;
                else rightData = null;
                checkMergeAvailability();
            } finally {
                isGenerating = false;
                submitBtn.disabled = false;
                btnText.textContent = 'Generate Mapping';
                spinner.style.display = 'none';
            }
        });
    });

    // Copy Text logic for both sides
    document.querySelectorAll('.copy-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const side = e.target.getAttribute('data-side');
            const reportOutput = document.getElementById(`output-${side}`);
            const text = reportOutput.innerText;
            
            navigator.clipboard.writeText(text).then(() => {
                const originalText = e.target.textContent;
                e.target.textContent = 'Copied!';
                setTimeout(() => { e.target.textContent = originalText; }, 2000);
            }).catch(err => console.error('Failed to copy text: ', err));
        });
    });

    // Merge Logic
    const mergeBtn = document.getElementById('merge-btn');
    mergeBtn.addEventListener('click', async () => {
        const btnText = mergeBtn.querySelector('.btn-text');
        const spinner = mergeBtn.querySelector('.spinner');
        const mergeError = document.getElementById('merge-error');
        const downloadLink = document.getElementById('merged-download-link');

        mergeError.classList.add('hidden');
        downloadLink.classList.add('hidden');
        mergeBtn.disabled = true;
        btnText.textContent = 'Merging...';
        spinner.style.display = 'block';

        try {
            const leftYear = document.getElementById('toYear-left').value;
            const rightYear = document.getElementById('toYear-right').value;

            const response = await fetch('/api/merge-mapping', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ 
                    leftName, leftData, rightName, rightData, leftYear, rightYear,
                    leftSessionId: leftData?.sessionId,
                    rightSessionId: rightData?.sessionId,
                    companyType: leftCompanyType || rightCompanyType,
                    leftCompanyType, rightCompanyType,
                    leftPvtInfo, rightPvtInfo
                })
            });

            const data = await response.json();
            if (response.ok && data.success) {
                downloadLink.href = `/api/download/${data.excelFileName}`;
                downloadLink.classList.remove('hidden');
                
                // Keep the pulse effect or change it to indicate success
                mergeBtn.classList.remove('pulse-btn');
                btnText.textContent = 'Merged Successfully!';
            } else {
                throw new Error(data.error || 'Failed to merge');
            }
        } catch(error) {
            mergeError.textContent = error.message;
            mergeError.classList.remove('hidden');
            btnText.textContent = 'Merge & Download Excel';
            mergeBtn.classList.add('pulse-btn');
        } finally {
            mergeBtn.disabled = false;
            spinner.style.display = 'none';
        }
    });

    function checkMergeAvailability() {
        const mergeSection = document.getElementById('merge-section');
        const mergeBtn = document.getElementById('merge-btn');
        const downloadLink = document.getElementById('merged-download-link');
        const viewToggles = document.getElementById('view-toggles');

        if (leftData || rightData) {
            viewToggles.classList.remove('hidden');
            const compType = leftCompanyType || rightCompanyType;
            const isRwaTrust = compType === 'RWA' || compType === 'Trust' || ((leftData && leftData.mapping && leftData.mapping["Note 21 - Other Expenses"] !== undefined && leftData.mapping["Note 25 - Other Expenses"] === undefined) || (rightData && rightData.mapping && rightData.mapping["Note 21 - Other Expenses"] !== undefined && rightData.mapping["Note 25 - Other Expenses"] === undefined));
            const plBtn = document.getElementById('btn-view-pl-table');
            if (plBtn) {
                plBtn.textContent = isRwaTrust ? 'Income & Expenditure' : 'P&L';
            }
        } else {
            viewToggles.classList.add('hidden');
        }

        if (leftData && rightData) {
            mergeSection.classList.remove('hidden');
            mergeSection.scrollIntoView({ behavior: 'smooth', block: 'end' });
            
            // Reset merge button state if it was clicked previously
            mergeBtn.querySelector('.btn-text').textContent = 'Merge & Download Excel';
            mergeBtn.classList.add('pulse-btn');
            downloadLink.classList.add('hidden');
        } else {
            mergeSection.classList.add('hidden');
        }
    }

    function showError(side, message) {
        const errorContainer = document.getElementById(`error-${side}`);
        errorContainer.textContent = message;
        errorContainer.classList.remove('hidden');
    }

    function fmt(num) {
        if (num === null || num === undefined) return '-';
        if (typeof num !== 'number') return num;
        return '₹ ' + num.toLocaleString('en-IN', { maximumFractionDigits: 2, minimumFractionDigits: 2 });
    }

    function createRow(label, value, type = '') {
        return `
            <div class="report-row ${type}">
                <span>${label}</span>
                <span class="money-val">${fmt(value)}</span>
            </div>
        `;
    }

    function renderReport(data, container, pvtInfo = null) {
        if (!data || data.error) {
            container.innerHTML = `<div class="error">${data?.error || 'Invalid data received.'}</div>`;
            return;
        }

        let html = '';
        
        function getMappingVal(mapObj, n) {
            const key = Object.keys(mapObj).find(k => k.startsWith(`Note ${n} -`));
            return key ? mapObj[key] : 0;
        }

        // Determine if this is the dynamic 25-note layout (Partnership/Proprietorship/HUF)
        if (data.companyType === 'Partnership Firm' || data.companyType === 'Proprietorship' || data.companyType === 'HUF' || (data.mapping && data.mapping["Note 25 - Other Expenses"] !== undefined)) {
            html += `<h3 style="margin-top:20px; border-bottom:2px solid #ccc; padding-bottom:5px;">Balance Sheet</h3>`;
            html += `<div class="report-section">`;
            html += `<div class="report-title" style="color:var(--primary-color);"><span>I. EQUITY AND LIABILITIES</span></div>`;
            
            const liab = [
                {n: 3, label: "Owners' Capital Account"},
                {n: 4, label: "Reserves and Surplus"},
                {n: 5, label: "Borrowings"},
                {n: 6, label: "Deferred tax Liabilities"},
                {n: 7, label: "Other long term Liabilities"},
                {n: 8, label: "Provisions"},
                {n: 9, label: "Trade Payables"},
                {n: 10, label: "Other current Liabilities"}
            ];
            let totalLiab = 0;
            liab.forEach(item => {
                const val = getMappingVal(data.mapping, item.n);
                totalLiab += val;
                html += createRow(`Note ${item.n} - ${item.label}`, val);
            });
            html += createRow('Total Liabilities', totalLiab, 'report-total');
            html += `</div>`;
            
            html += `<div class="report-section">`;
            html += `<div class="report-title" style="color:var(--primary-color);"><span>II. ASSETS</span></div>`;
            
            const assets = [
                {n: 11, label: "Property, Plant & Equipment"},
                {n: 12, label: "Investments"},
                {n: 13, label: "Loans and Advances"},
                {n: 14, label: "Non current assets"},
                {n: 15, label: "Inventory"},
                {n: 16, label: "Trade Receivables"},
                {n: 17, label: "Cash and Bank Balance"},
                {n: 18, label: "Other current assets"}
            ];
            let totalAsset = 0;
            assets.forEach(item => {
                const val = getMappingVal(data.mapping, item.n);
                totalAsset += val;
                html += createRow(`Note ${item.n} - ${item.label}`, val);
            });
            html += createRow('Total Assets', totalAsset, 'report-total');
            html += `</div>`;
            
            html += `<h3 style="margin-top:30px; border-bottom:2px solid #ccc; padding-bottom:5px;">Statement of Profit and Loss</h3>`;
            html += `<div class="report-section">`;
            html += `<div class="report-title" style="color:var(--primary-color);"><span>I. Revenue</span></div>`;
            
            const rev1 = getMappingVal(data.mapping, 19);
            const rev2 = getMappingVal(data.mapping, 20);
            html += createRow('Note 19 - Revenue from Operations', rev1);
            html += createRow('Note 20 - Other income', rev2);
            const totalRev = rev1 + rev2;
            html += createRow('Total Revenue', totalRev, 'report-total');
            html += `</div>`;
            
            html += `<div class="report-section">`;
            html += `<div class="report-title" style="color:var(--primary-color);"><span>II. Expenses</span></div>`;
            
            const exp = [
                {n: 21, label: "Cost of goods sold"},
                {n: 22, label: "Employee benefits expense"},
                {n: 23, label: "Finance cost"},
                {n: 24, label: "Depreciation and amortization expense"},
                {n: 25, label: "Other Expenses"}
            ];
            let totalExp = 0;
            exp.forEach(item => {
                const val = getMappingVal(data.mapping, item.n);
                totalExp += val;
                html += createRow(`Note ${item.n} - ${item.label}`, val);
            });
            html += createRow('Total Expenses', totalExp, 'report-total');
            html += `</div>`;
            
            html += `<div class="report-section">`;
            html += createRow('Net Profit / (Loss)', totalRev - totalExp, 'report-total');
            html += `</div>`;
            
            container.innerHTML = html;
            return;
        }

        // Determine if this is the dynamic 21-note layout (RWA/Trust)
        if (data.companyType === 'RWA' || data.companyType === 'Trust' || (data.mapping && data.mapping["Note 25 - Other Expenses"] === undefined && data.mapping["Note 21 - Other Expenses"] !== undefined)) {
            html += `<h3 style="margin-top:20px; border-bottom:2px solid #ccc; padding-bottom:5px;">Balance Sheet</h3>`;
            html += `<div class="report-section">`;
            html += `<div class="report-title" style="color:var(--primary-color);"><span>I. SOURCES OF FUNDS (LIABILITIES)</span></div>`;
            
            const liab = [
                {n: 3, label: "NPOs funds"},
                {n: 4, label: "Borrowings"},
                {n: 5, label: "Other long term Liabilities"},
                {n: 6, label: "Provisions"},
                {n: 7, label: "Payables"},
                {n: 8, label: "Other current Liabilities"}
            ];
            let totalLiab = 0;
            liab.forEach(item => {
                const val = getMappingVal(data.mapping, item.n);
                totalLiab += val;
                html += createRow(`Note ${item.n} - ${item.label}`, val);
            });
            html += createRow('Total Sources of Funds', totalLiab, 'report-total');
            html += `</div>`;
            
            html += `<div class="report-section">`;
            html += `<div class="report-title" style="color:var(--primary-color);"><span>II. APPLICATION OF FUNDS (ASSETS)</span></div>`;
            
            const assets = [
                {n: 9, label: "PPE"},
                {n: 10, label: "Investments - Non current and current"},
                {n: 11, label: "Loans and Advances"},
                {n: 12, label: "Other Non current assets"},
                {n: 13, label: "Receivables"},
                {n: 14, label: "Cash and Bank Balance"},
                {n: 15, label: "Other current assets"}
            ];
            let totalAsset = 0;
            assets.forEach(item => {
                const val = getMappingVal(data.mapping, item.n);
                totalAsset += val;
                html += createRow(`Note ${item.n} - ${item.label}`, val);
            });
            html += createRow('Total Application of Funds', totalAsset, 'report-total');
            html += `</div>`;
            
            html += `<h3 style="margin-top:30px; border-bottom:2px solid #ccc; padding-bottom:5px;">Statement of Income and Expenditure</h3>`;
            html += `<div class="report-section">`;
            html += `<div class="report-title" style="color:var(--primary-color);"><span>I. Income</span></div>`;
            
            const rev = getMappingVal(data.mapping, 16);
            html += createRow('Note 16 - Other income', rev);
            const totalRev = rev;
            html += createRow('Total Income', totalRev, 'report-total');
            html += `</div>`;
            
            html += `<div class="report-section">`;
            html += `<div class="report-title" style="color:var(--primary-color);"><span>II. Expenses</span></div>`;
            
            const exp = [
                {n: 17, label: "Cost of goods sold"},
                {n: 18, label: "Employee benefits expense"},
                {n: 19, label: "Finance Cost"},
                {n: 20, label: "Depreciation and amortization expense"},
                {n: 21, label: "Other Expenses"}
            ];
            let totalExp = 0;
            exp.forEach(item => {
                const val = getMappingVal(data.mapping, item.n);
                totalExp += val;
                html += createRow(`Note ${item.n} - ${item.label}`, val);
            });
            html += createRow('Total Expenses', totalExp, 'report-total');
            html += `</div>`;
            
            html += `<div class="report-section">`;
            html += createRow('Net Surplus / (Deficit)', totalRev - totalExp, 'report-total');
            html += `</div>`;
            
            container.innerHTML = html;
            return;
        }


        // Note 1 – with optional Pvt Ltd breakdown
        const shareCapital = data.mapping['Note 1 - Share Capital'] || 0;
        html += `<div class="report-section">
            <div class="report-title"><span>Note 1 - Share Capital</span> <span class="money-val">${fmt(shareCapital)}</span></div>`;
        if (pvtInfo) {
            html += createRow('Authorized Capital', pvtInfo.authCap, 'sub-item');
            html += createRow(`No. of Shares × ₹${pvtInfo.faceVal} (Face Value)`, pvtInfo.noShares, 'sub-item');
            html += createRow('Paid-up Capital (Issued & Subscribed)', shareCapital, 'sub-item report-total');
        }
        html += `</div>`;

        // Note 2
        html += `<div class="report-section">
            <div class="report-title"><span>Note 2 - Profit/Loss</span> <span class="money-val">${fmt(data.mapping['Note 2 - Profit/Loss'])}</span></div>
        </div>`;

        // Note 3
        html += `<div class="report-section">
            <div class="report-title"><span>Note 3 - Long-term Borrowings</span> <span class="money-val">${fmt(data.loansInfo ? data.loansInfo.total : 0)}</span></div>`;
        if (data.loansInfo) {
            html += createRow('(A) Secured Loans', data.loansInfo.secTotal, 'sub-item');
            for (const [k, v] of Object.entries(data.loansInfo.secured || {})) {
                html += createRow(`- ${k}`, v, 'sub-sub-item');
            }
            html += createRow('(B) Unsecured Loans', data.loansInfo.unsecTotal, 'sub-item');
            for (const [k, v] of Object.entries(data.loansInfo.unsecured || {})) {
                html += createRow(`- ${k}`, v, 'sub-sub-item');
            }
        }
        html += `</div>`;

        // Note 4
        html += `<div class="report-section">
            <div class="report-title"><span>Note 4 - Deferred Tax Liabilities</span> <span class="money-val">${fmt(data.mapping['Note 4 - '])}</span></div>
        </div>`;

        // Note 5
        html += `<div class="report-section">
            <div class="report-title"><span>Note 5 - Provisions</span> <span class="money-val">${fmt(data.provisionsInfo ? data.provisionsInfo.total : 0)}</span></div>`;
        if (data.provisionsInfo) {
            let charCode = 65;
            for (const [k, v] of Object.entries(data.provisionsInfo.shortTermItems || {})) {
                html += createRow(`(${String.fromCharCode(charCode++)}) ${k}`, v, 'sub-item');
            }
            if (data.provisionsInfo.provisionForTaxation) {
                html += createRow(`(${String.fromCharCode(charCode++)}) Provision for Taxation`, data.provisionsInfo.provisionForTaxation, 'sub-item');
            }
        }
        html += `</div>`;

        // Note 6 - Trade Payables (with Sundry Creditors breakdown)
        html += `<div class="report-section">
            <div class="report-title"><span>Note 6 - Trade Payables</span> <span class="money-val">${fmt(data.sundryCredInfo ? data.sundryCredInfo.total : data.mapping['Note 6 - Trade Payables'])}</span></div>`;
        if (data.sundryCredInfo && Object.keys(data.sundryCredInfo.ledgers).length > 0) {
            for (const [k, v] of Object.entries(data.sundryCredInfo.ledgers)) {
                html += createRow(`- ${k}`, v, 'sub-item');
            }
        }
        html += `</div>`;

        // Note 7
        html += `<div class="report-section">
            <div class="report-title"><span>Note 7 - Other Current Liabilities</span> <span class="money-val">${fmt(data.otherCurrLiabInfo ? data.otherCurrLiabInfo.total : data.mapping['Note 7 - Other Current Liabilities'])}</span></div>`;
        if (data.otherCurrLiabInfo) {
            let charCode = 65;
            for (const [k, v] of Object.entries(data.otherCurrLiabInfo.items || {})) {
                html += createRow(`(${String.fromCharCode(charCode++)}) ${k}`, v, 'sub-item');
            }
        }
        html += `</div>`;

        // Note 8
        html += `<div class="report-section">
            <div class="report-title"><span>Note 8 - PPE (Fixed Assets)</span> <span class="money-val">${fmt(data.mapping['Note 8 - PPE (Fixed Assets)'])}</span></div>
        </div>`;

        // Note 9
        html += `<div class="report-section">
            <div class="report-title"><span>Note 9 - Long Term Loans & Advances</span> <span class="money-val">${fmt(data.advancesInfo ? data.advancesInfo.total : data.mapping['Note 9 - Long Term Loans & Advances'])}</span></div>`;
        if (data.advancesInfo) {
            html += createRow('Unsecured advances : Considered Good', '', 'sub-item');
            let charCode = 97;
            for (const [k, v] of Object.entries(data.advancesInfo.items || {})) {
                html += createRow(`(${String.fromCharCode(charCode++)}) ${k}`, v === 0 ? '-' : v, 'sub-sub-item');
            }
        }
        html += `</div>`;

        // Note 10 - Investments
        html += `<div class="report-section">
            <div class="report-title"><span>Note 10 - Investments</span> <span class="money-val">${fmt(data.investmentsInfo ? data.investmentsInfo.total : data.mapping['Note 10 - Investments'])}</span></div>`;
        if (data.investmentsInfo && Object.keys(data.investmentsInfo.ledgers).length > 0) {
            for (const [k, v] of Object.entries(data.investmentsInfo.ledgers)) {
                html += createRow(`- ${k}`, Math.abs(v), 'sub-item');
            }
        }
        html += `</div>`;

        html += `<div class="report-section">
            <div class="report-title"><span>Note 11 - Sundry Debtors</span> <span class="money-val">${fmt(data.sundryDebtorsInfo ? data.sundryDebtorsInfo.total : data.mapping['Note 11 - Sundry Debtors'])}</span></div>`;
        if (data.sundryDebtorsInfo && Object.keys(data.sundryDebtorsInfo.ledgers).length > 0) {
            for (const [k, v] of Object.entries(data.sundryDebtorsInfo.ledgers))
                html += createRow(`- ${k}`, Math.abs(v), 'sub-item');
        }
        html += `</div>`;

        // Note 12
        html += `<div class="report-section">
            <div class="report-title"><span>Note 12 - Cash & Cash Equivalents</span> <span class="money-val">${fmt(data.cashInfo ? data.cashInfo.total : data.mapping['Note 12 - Cash & Cash Equivalents'])}</span></div>`;
        if (data.cashInfo) {
            html += createRow('Balances with Banks', data.cashInfo.bankAccounts, 'sub-item');
            html += createRow('Cheques, drafts on hand', '-', 'sub-item');
            html += createRow('Cash on Hand', data.cashInfo.cashInHand, 'sub-item');
            html += createRow('Others', '-', 'sub-item');
        }
        html += `</div>`;

        // Note 13 & 14
        html += `<div class="report-section">
            <div class="report-title"><span>Note 13 - Other Current Assets</span> <span class="money-val">${fmt(data.mapping['Note 13 - Other Current Assets'])}</span></div>
            <div class="report-title" style="margin-top: 1rem;"><span>Note 14 - Revenue from Operations</span> <span class="money-val">${fmt(data.mapping['Note 14 - Revenue from Operations'])}</span></div>
        </div>`;

        // Note 15
        html += `<div class="report-section">
            <div class="report-title"><span>Note 15 - Other Income</span> <span class="money-val">${fmt(data.indirectIncomesInfo ? data.indirectIncomesInfo.total : data.mapping['Note 15 - Other Income'])}</span></div>`;
        if (data.indirectIncomesInfo) {
            for (const [k, v] of Object.entries(data.indirectIncomesInfo.items || {})) {
                html += createRow(k, v === 0 ? '-' : v, 'sub-item');
            }
        }
        html += `</div>`;

        // Note 16
        html += `<div class="report-section">
            <div class="report-title"><span>Note 16 - Employee Benefit Expenses</span> <span class="money-val">${fmt(data.note16Info ? data.note16Info.total : data.mapping['Note 16 - Employee Benefit Expenses'])}</span></div>`;
        if (data.note16Info) {
            for (const [k, v] of Object.entries(data.note16Info.items || {})) {
                html += createRow(k, v, 'sub-item');
            }
        }
        html += `</div>`;

        // Note 17 & 18
        html += `<div class="report-section">
            <div class="report-title"><span>Note 17 - Depreciation</span> <span class="money-val">${fmt(data.mapping['Note 17 - Depreciation'])}</span></div>
            <div class="report-title" style="margin-top: 1rem;"><span>Note 18 - Other Expenses</span> <span class="money-val">${fmt(data.mapping['Note 18 - Other Expenses'])}</span></div>
        </div>`;

        container.innerHTML = html;
    }

    function openTabularReport(type) {
        if (!leftData && !rightData) return alert('Please generate at least one mapping first.');

        const reportPage = document.getElementById('report-page');
        const container = document.getElementById('report-page-table-container');
        const title = document.getElementById('report-page-title');
        const compType = leftCompanyType || rightCompanyType;
        const isRwaTrust = compType === 'RWA' || compType === 'Trust' || (leftData && leftData.mapping && leftData.mapping["Note 21 - Other Expenses"] !== undefined && leftData.mapping["Note 25 - Other Expenses"] === undefined);
        title.textContent = type === 'bs' ? 'Balance Sheet' : (isRwaTrust ? 'Statement of Income and Expenditure' : 'Statement of Profit and Loss');

        let thHtml = '';
        if (leftData) thHtml += `<th>${leftName || 'Left Report'}</th>`;
        if (rightData) thHtml += `<th>${rightName || 'Right Report'}</th>`;

        function r(label, noteText, noteNum, cls = '') {
            let leftVal = noteNum ? getVal(leftData, noteNum) : '';
            let rightVal = noteNum ? getVal(rightData, noteNum) : '';
            let html = `<tr><td class="${cls}">${label}</td><td style="text-align:center;">${noteText}</td>`;
            if (leftData) html += `<td class="num">${noteNum ? fmtTab(leftVal) : ''}</td>`;
            if (rightData) html += `<td class="num">${noteNum ? fmtTab(rightVal) : ''}</td>`;
            html += `</tr>`;
            return html;
        }

        function sum(start, end, d) {
            let tot = 0;
            for (let i = start; i <= end; i++) tot += Number(getVal(d, i)) || 0;
            return tot;
        }

        let bodyHtml = '';

        if (type === 'bs') {
            if (leftCompanyType === 'Partnership Firm' || leftCompanyType === 'Proprietorship' || leftCompanyType === 'HUF' || (leftData && leftData.mapping && leftData.mapping["Note 25 - Other Expenses"] !== undefined)) {
                bodyHtml += r('I. SOURCES OF FUNDS (LIABILITIES)', '', null, 'bold');
                bodyHtml += r('Owner\'s Capital Account', '3', 3, 'item');
                bodyHtml += r('Reserves and Surplus', '4', 4, 'item');
                bodyHtml += r('Borrowings', '5', 5, 'item');
                bodyHtml += r('Deferred Tax Liabilities', '6', 6, 'item');
                bodyHtml += r('Other Long-term Liabilities', '7', 7, 'item');
                bodyHtml += r('Provisions', '8', 8, 'item');
                bodyHtml += r('Trade Payables', '9', 9, 'item');
                bodyHtml += r('Other Current Liabilities', '10', 10, 'item');
                let lT1 = leftData ? sum(3, 10, leftData) : 0;
                let rT1 = rightData ? sum(3, 10, rightData) : 0;
                bodyHtml += `<tr class="total-row"><td colspan="2" class="bold" style="text-align:right;">Total Sources of Funds</td>`;
                if (leftData) bodyHtml += `<td class="num bold">${fmtTab(lT1)}</td>`;
                if (rightData) bodyHtml += `<td class="num bold">${fmtTab(rT1)}</td>`;
                bodyHtml += `</tr>`;
                
                bodyHtml += r('II. APPLICATION OF FUNDS (ASSETS)', '', null, 'bold');
                bodyHtml += r('PPE (Fixed Assets)', '11', 11, 'item');
                bodyHtml += r('Investments', '12', 12, 'item');
                bodyHtml += r('Loans and Advances', '13', 13, 'item');
                bodyHtml += r('Non-current Assets', '14', 14, 'item');
                bodyHtml += r('Inventory', '15', 15, 'item');
                bodyHtml += r('Trade Receivables', '16', 16, 'item');
                bodyHtml += r('Cash and Bank Balance', '17', 17, 'item');
                bodyHtml += r('Other Current Assets', '18', 18, 'item');
                let lT2 = leftData ? sum(11, 18, leftData) : 0;
                let rT2 = rightData ? sum(11, 18, rightData) : 0;
                bodyHtml += `<tr class="total-row"><td colspan="2" class="bold" style="text-align:right;">Total Application of Funds</td>`;
                if (leftData) bodyHtml += `<td class="num bold">${fmtTab(lT2)}</td>`;
                if (rightData) bodyHtml += `<td class="num bold">${fmtTab(rT2)}</td>`;
                bodyHtml += `</tr>`;

            } else if (leftCompanyType === 'RWA' || leftCompanyType === 'Trust' || (leftData && leftData.mapping && leftData.mapping["Note 21 - Other Expenses"] !== undefined && leftData.mapping["Note 25 - Other Expenses"] === undefined)) {
                bodyHtml += r('I. SOURCES OF FUNDS (LIABILITIES)', '', null, 'bold');
                bodyHtml += r('NPOs funds', '3', 3, 'item');
                bodyHtml += r('Borrowings', '4', 4, 'item');
                bodyHtml += r('Other long term Liabilities', '5', 5, 'item');
                bodyHtml += r('Provisions', '6', 6, 'item');
                bodyHtml += r('Payables', '7', 7, 'item');
                bodyHtml += r('Other current Liabilities', '8', 8, 'item');
                let lT1 = leftData ? sum(3, 8, leftData) : 0;
                let rT1 = rightData ? sum(3, 8, rightData) : 0;
                bodyHtml += `<tr class="total-row"><td colspan="2" class="bold" style="text-align:right;">Total Sources of Funds</td>`;
                if (leftData) bodyHtml += `<td class="num bold">${fmtTab(lT1)}</td>`;
                if (rightData) bodyHtml += `<td class="num bold">${fmtTab(rT1)}</td>`;
                bodyHtml += `</tr>`;
                
                bodyHtml += r('II. APPLICATION OF FUNDS (ASSETS)', '', null, 'bold');
                bodyHtml += r('PPE (Fixed Assets)', '9', 9, 'item');
                bodyHtml += r('Investments - Non current and current', '10', 10, 'item');
                bodyHtml += r('Loans and Advances', '11', 11, 'item');
                bodyHtml += r('Other Non current assets', '12', 12, 'item');
                bodyHtml += r('Receivables', '13', 13, 'item');
                bodyHtml += r('Cash and Bank Balance', '14', 14, 'item');
                bodyHtml += r('Other current assets', '15', 15, 'item');
                let lT2 = leftData ? sum(9, 15, leftData) : 0;
                let rT2 = rightData ? sum(9, 15, rightData) : 0;
                bodyHtml += `<tr class="total-row"><td colspan="2" class="bold" style="text-align:right;">Total Application of Funds</td>`;
                if (leftData) bodyHtml += `<td class="num bold">${fmtTab(lT2)}</td>`;
                if (rightData) bodyHtml += `<td class="num bold">${fmtTab(rT2)}</td>`;
                bodyHtml += `</tr>`;

            } else {
                bodyHtml += r('I. EQUITY AND LIABILITIES', '', null, 'bold');
                bodyHtml += r("1. Shareholders' funds", '', null, 'bold');
                bodyHtml += r('(a) Share capital', '1', 1, 'item');
                bodyHtml += r('(b) Reserves and surplus', '2', 2, 'item');
                bodyHtml += r('(c) Money received against share warrants', '', null, 'item');
                bodyHtml += r('2. Share application money pending allotment', '', null, 'bold');
                bodyHtml += r('3. Non-current liabilities', '', null, 'bold');
                bodyHtml += r('(a) Long-term borrowings', '3', 3, 'item');
                bodyHtml += r('(b) Deferred tax liabilities (net)', '4', 4, 'item');
                bodyHtml += r('(c) Other Long-term liabilities', '', null, 'item');
                bodyHtml += r('(d) Long-term provisions', '', null, 'item');
                bodyHtml += r('4. Current liabilities', '', null, 'bold');
                bodyHtml += r('(a) Short-term borrowings', '', null, 'item');
                bodyHtml += r('(b) Trade payables', '', null, 'item');
                bodyHtml += r('&nbsp;&nbsp;&nbsp;&nbsp;(i) Total outstanding dues of micro and small enterprises', '', null, 'sub-item');
                bodyHtml += r('&nbsp;&nbsp;&nbsp;&nbsp;(ii) Total outstanding dues of creditors other than micro and small enterprises', '6', 6, 'sub-item');
                bodyHtml += r('(c) Other current liabilities', '7', 7, 'item');
                bodyHtml += r('(d) Short-term provisions', '5', 5, 'item');

                let lT1 = leftData ? sum(1, 7, leftData) : 0;
                let rT1 = rightData ? sum(1, 7, rightData) : 0;
                bodyHtml += `<tr class="total-row"><td colspan="2" class="bold" style="text-align:right;">TOTAL</td>`;
                if (leftData) bodyHtml += `<td class="num bold">${fmtTab(lT1)}</td>`;
                if (rightData) bodyHtml += `<td class="num bold">${fmtTab(rT1)}</td>`;
                bodyHtml += `</tr>`;

                bodyHtml += r('II. ASSETS', '', null, 'bold');
                bodyHtml += r('1. Non-current assets', '', null, 'bold');
                bodyHtml += r('(a) Property Plant and Equipment and Intangible assets', '', null, 'item');
                bodyHtml += r('&nbsp;&nbsp;&nbsp;&nbsp;(i) Property, Plant and Equipment', '8', 8, 'sub-item');
                bodyHtml += r('&nbsp;&nbsp;&nbsp;&nbsp;(ii) Intangible assets', '', null, 'sub-item');
                bodyHtml += r('&nbsp;&nbsp;&nbsp;&nbsp;(iii) Capital Work In Progress', '', null, 'sub-item');
                bodyHtml += r('&nbsp;&nbsp;&nbsp;&nbsp;(iv) Intangible Assets Under Developments', '', null, 'sub-item');
                bodyHtml += r('(b) Non-current investments', '10', 10, 'item');
                bodyHtml += r('(c) Deferred tax assets (net)', '', null, 'item');
                bodyHtml += r('(d) Long-term loans and advances', '9', 9, 'item');
                bodyHtml += r('(e) Other non-current assets', '', null, 'item');
                bodyHtml += r('2. Current assets', '', null, 'bold');
                bodyHtml += r('(a) Current investments', '', null, 'item');
                bodyHtml += r('(b) Inventories', '', null, 'item');
                bodyHtml += r('(c) Trade receivables', '11', 11, 'item');
                bodyHtml += r('(d) Cash and cash equivalents', '12', 12, 'item');
                bodyHtml += r('(e) Short-term loans and advances', '', null, 'item');
                bodyHtml += r('(f) Other current assets', '13', 13, 'item');

                let lT2 = leftData ? sum(8, 13, leftData) : 0;
                let rT2 = rightData ? sum(8, 13, rightData) : 0;
                bodyHtml += `<tr class="total-row"><td colspan="2" class="bold" style="text-align:right;">TOTAL</td>`;
                if (leftData) bodyHtml += `<td class="num bold">${fmtTab(lT2)}</td>`;
                if (rightData) bodyHtml += `<td class="num bold">${fmtTab(rT2)}</td>`;
                bodyHtml += `</tr>`;

            }

        } else if (type === 'pl') {
            if (leftCompanyType === 'Partnership Firm' || leftCompanyType === 'Proprietorship' || leftCompanyType === 'HUF' || (leftData && leftData.mapping && leftData.mapping["Note 25 - Other Expenses"] !== undefined)) {
                bodyHtml += r('I. INCOME', '', null, 'bold');
                bodyHtml += r('Revenue from Operations', '19', 19, 'item');
                bodyHtml += r('Other Income', '20', 20, 'item');
                let lInc = leftData ? sum(19, 20, leftData) : 0;
                let rInc = rightData ? sum(19, 20, rightData) : 0;
                bodyHtml += `<tr class="total-row"><td colspan="2" class="bold" style="text-align:right;">Total Income</td>`;
                if (leftData) bodyHtml += `<td class="num bold">${fmtTab(lInc)}</td>`;
                if (rightData) bodyHtml += `<td class="num bold">${fmtTab(rInc)}</td>`;
                bodyHtml += `</tr>`;
                
                bodyHtml += r('II. EXPENDITURE', '', null, 'bold');
                bodyHtml += r('Cost of Goods Sold', '21', 21, 'item');
                bodyHtml += r('Employee Benefits Expense', '22', 22, 'item');
                bodyHtml += r('Finance Cost', '23', 23, 'item');
                bodyHtml += r('Depreciation and Amortisation', '24', 24, 'item');
                bodyHtml += r('Other Expenses', '25', 25, 'item');
                let lExp = leftData ? sum(21, 25, leftData) : 0;
                let rExp = rightData ? sum(21, 25, rightData) : 0;
                bodyHtml += `<tr class="total-row"><td colspan="2" class="bold" style="text-align:right;">Total Expenditure</td>`;
                if (leftData) bodyHtml += `<td class="num bold">${fmtTab(lExp)}</td>`;
                if (rightData) bodyHtml += `<td class="num bold">${fmtTab(rExp)}</td>`;
                bodyHtml += `</tr>`;
                
                let lProf = lInc - lExp;
                let rProf = rInc - rExp;
                bodyHtml += `<tr class="total-row"><td colspan="2" class="bold" style="text-align:right;">III. Net Profit / (Loss)</td>`;
                if (leftData) bodyHtml += `<td class="num bold">${fmtTab(lProf)}</td>`;
                if (rightData) bodyHtml += `<td class="num bold">${fmtTab(rProf)}</td>`;
                bodyHtml += `</tr>`;


            } else if (leftCompanyType === 'RWA' || leftCompanyType === 'Trust' || (leftData && leftData.mapping && leftData.mapping["Note 21 - Other Expenses"] !== undefined && leftData.mapping["Note 25 - Other Expenses"] === undefined)) {
                bodyHtml += r('I. INCOME', '', null, 'bold');
                bodyHtml += r('Other Income', '16', 16, 'item');
                let lInc = leftData ? sum(16, 16, leftData) : 0;
                let rInc = rightData ? sum(16, 16, rightData) : 0;
                bodyHtml += `<tr class="total-row"><td colspan="2" class="bold" style="text-align:right;">Total Income</td>`;
                if (leftData) bodyHtml += `<td class="num bold">${fmtTab(lInc)}</td>`;
                if (rightData) bodyHtml += `<td class="num bold">${fmtTab(rInc)}</td>`;
                bodyHtml += `</tr>`;
                
                bodyHtml += r('II. EXPENDITURE', '', null, 'bold');
                bodyHtml += r('Cost of Goods Sold', '17', 17, 'item');
                bodyHtml += r('Employee Benefits Expense', '18', 18, 'item');
                bodyHtml += r('Finance Cost', '19', 19, 'item');
                bodyHtml += r('Depreciation and Amortisation Expense', '20', 20, 'item');
                bodyHtml += r('Other Expenses', '21', 21, 'item');
                let lExp = leftData ? sum(17, 21, leftData) : 0;
                let rExp = rightData ? sum(17, 21, rightData) : 0;
                bodyHtml += `<tr class="total-row"><td colspan="2" class="bold" style="text-align:right;">Total Expenditure</td>`;
                if (leftData) bodyHtml += `<td class="num bold">${fmtTab(lExp)}</td>`;
                if (rightData) bodyHtml += `<td class="num bold">${fmtTab(rExp)}</td>`;
                bodyHtml += `</tr>`;
                
                let lProf = lInc - lExp;
                let rProf = rInc - rExp;
                bodyHtml += `<tr class="total-row"><td colspan="2" class="bold" style="text-align:right;">III. Net Surplus / (Deficit)</td>`;
                if (leftData) bodyHtml += `<td class="num bold">${fmtTab(lProf)}</td>`;
                if (rightData) bodyHtml += `<td class="num bold">${fmtTab(rProf)}</td>`;
                bodyHtml += `</tr>`;


            } else {
                bodyHtml += r('I. Revenue from operations', '14', 14, 'bold');
                bodyHtml += r('II. Other Income', '15', 15, 'bold');

                let lInc = leftData ? sum(14, 15, leftData) : 0;
                let rInc = rightData ? sum(14, 15, rightData) : 0;
                bodyHtml += `<tr class="total-row"><td colspan="2" class="bold" style="text-align:right;">III. TOTAL INCOME ( I + II )</td>`;
                if (leftData) bodyHtml += `<td class="num bold">${fmtTab(lInc)}</td>`;
                if (rightData) bodyHtml += `<td class="num bold">${fmtTab(rInc)}</td>`;
                bodyHtml += `</tr>`;

                bodyHtml += r('IV. EXPENSES', '', null, 'bold');
                bodyHtml += r('(a) Cost of materials consumed', '', null, 'item');
                bodyHtml += r('(b) Purchases of Stock In Trade', '', null, 'item');
                bodyHtml += r('(c) Changes in inventories of finished goods', '', null, 'item');
                bodyHtml += r('(d) Changes in work-in-progress and stock-in-trade', '', null, 'item');
                bodyHtml += r('(e) Employee benefits expenses', '16', 16, 'item');
                bodyHtml += r('(f) Depreciation and amortisation expenses', '17', 17, 'item');
                bodyHtml += r('(g) Finance costs', '', null, 'item');
                bodyHtml += r('(h) Other expenses', '18', 18, 'item');

                let lExp = leftData ? sum(16, 18, leftData) : 0;
                let rExp = rightData ? sum(16, 18, rightData) : 0;
                bodyHtml += `<tr class="total-row"><td colspan="2" class="bold" style="text-align:right;">TOTAL EXPENSES</td>`;
                if (leftData) bodyHtml += `<td class="num bold">${fmtTab(lExp)}</td>`;
                if (rightData) bodyHtml += `<td class="num bold">${fmtTab(rExp)}</td>`;
                bodyHtml += `</tr>`;

                let lProf = lInc - lExp;
                let rProf = rInc - rExp;
                bodyHtml += `<tr class="total-row"><td colspan="2" class="bold" style="text-align:right;">V. Profit before exceptional and extraordinary items and tax (III - IV)</td>`;
                if (leftData) bodyHtml += `<td class="num bold">${fmtTab(lProf)}</td>`;
                if (rightData) bodyHtml += `<td class="num bold">${fmtTab(rProf)}</td>`;
                bodyHtml += `</tr>`;
                
                bodyHtml += r('VI. Exceptional items', '', null, 'bold');
                bodyHtml += `<tr class="total-row"><td colspan="2" class="bold" style="text-align:right;">VII. Profit before extraordinary items and tax ( V- VI)</td>`;
                if (leftData) bodyHtml += `<td class="num bold">${fmtTab(lProf)}</td>`;
                if (rightData) bodyHtml += `<td class="num bold">${fmtTab(rProf)}</td>`;
                bodyHtml += `</tr>`;
                
                bodyHtml += r('VIII. Extraordinary Items', '', null, 'bold');
                bodyHtml += `<tr class="total-row"><td colspan="2" class="bold" style="text-align:right;">IX. Profit before tax (VII-VIII)</td>`;
                if (leftData) bodyHtml += `<td class="num bold">${fmtTab(lProf)}</td>`;
                if (rightData) bodyHtml += `<td class="num bold">${fmtTab(rProf)}</td>`;
                bodyHtml += `</tr>`;
                
                bodyHtml += r('X. Tax Expense:', '', null, 'bold');
                bodyHtml += r('(a) Current tax expense', '', null, 'item');
                bodyHtml += r('(b) Deferred tax', '', null, 'item');
                bodyHtml += `<tr class="total-row"><td colspan="2" class="bold" style="text-align:right;">XI. Profit / (Loss) from continuing operations (VII-VIII)</td>`;
                if (leftData) bodyHtml += `<td class="num bold">${fmtTab(lProf)}</td>`;
                if (rightData) bodyHtml += `<td class="num bold">${fmtTab(rProf)}</td>`;
                bodyHtml += `</tr>`;
            }

        }

        container.innerHTML = `
            <table>
                <thead>
                    <tr>
                        <th style="width:50%;">Particulars</th>
                        <th style="width:10%;">Note No.</th>
                        ${thHtml}
                    </tr>
                </thead>
                <tbody>${bodyHtml}</tbody>
            </table>`;

        // Show the report page overlay
        reportPage.classList.remove('hidden');
        reportPage.scrollTop = 0;
    }

    // Back button
    document.getElementById('btn-back').addEventListener('click', () => {
        document.getElementById('report-page').classList.add('hidden');
    });

    // Print button
    document.getElementById('btn-print').addEventListener('click', () => {
        window.print();
    });

    // ── Checker and Settings Logic ──────────────────────────────────────────
    const typeSelect = document.getElementById('checker-type-select');
    const compSelect = document.getElementById('checker-company-select');
    const checkResult = document.getElementById('checker-result');
    let fullCompanyList = [];

    async function initChecker() {
        try {
            const res = await fetch('/api/local-company-list');
            fullCompanyList = await res.json();
            
            // Populate types
            const types = [...new Set(fullCompanyList.map(c => c.type).filter(Boolean))];
            types.sort().forEach(type => {
                const opt = document.createElement('option');
                opt.value = type;
                opt.textContent = type;
                typeSelect.appendChild(opt);
            });
            
            updateCheckerCompanies('');
        } catch(e) {
            console.error('Error fetching local company list:', e);
        }
    }

    function updateCheckerCompanies(selectedType) {
        compSelect.innerHTML = '<option value="">-- Select Company --</option>';
        const filtered = selectedType 
            ? fullCompanyList.filter(c => c.type === selectedType) 
            : fullCompanyList;
            
        filtered.sort((a,b) => a.name.localeCompare(b.name)).forEach(comp => {
            const opt = document.createElement('option');
            opt.value = comp.name;
            opt.textContent = comp.name;
            compSelect.appendChild(opt);
        });
    }

    if (typeSelect) {
        typeSelect.addEventListener('change', (e) => {
            updateCheckerCompanies(e.target.value);
            checkResult.classList.add('hidden');
        });
    }

    if (compSelect) {
        compSelect.addEventListener('change', async (e) => {
            const selectedName = e.target.value;
            if(!selectedName) {
                checkResult.classList.add('hidden');
                return;
            }
            
            try {
                const tallyHost = localStorage.getItem('tallyHost') || '172.16.1.4';
                const tallyPort = localStorage.getItem('tallyPort') || '9000';
                const res = await fetch('/api/companies', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ tallyHost, tallyPort })
                });
                const data = await res.json();
                const tallyCompanies = data.companies || [];
                
                checkResult.classList.remove('hidden', 'success', 'error');
                // Partial match or exact match check
                const found = tallyCompanies.some(tc => tc.toLowerCase().includes(selectedName.toLowerCase()) || selectedName.toLowerCase().includes(tc.toLowerCase()));
                
                if(found) {
                    checkResult.textContent = 'The company you have selected is found';
                    checkResult.classList.add('success');
                } else {
                    checkResult.textContent = 'The company you have selected is not avaliable in this IP';
                    checkResult.classList.add('error');
                }
            } catch(err) {
                console.error('Error checking company:', err);
            }
        });
    }

    // ── Server Settings Logic ──────────────────────────────────────────
    const settingsBtn = document.getElementById('server-settings-btn');
    const settingsDialog = document.getElementById('server-settings-dialog');
    const refreshServerBtn = document.getElementById('refresh-server-btn');

    if (settingsBtn) {
        settingsBtn.addEventListener('click', () => {
            settingsDialog.classList.toggle('hidden');
        });
    }

    if (refreshServerBtn) {
        refreshServerBtn.addEventListener('click', async () => {
            const ipInput = document.getElementById('tally-ip-input');
            const portInput = document.getElementById('tally-port-input');
            const tallyHost = ipInput ? ipInput.value.trim() : '';
            const tallyPort = portInput ? portInput.value.trim() : '';
            
            if (tallyHost) localStorage.setItem('tallyHost', tallyHost);
            if (tallyPort) localStorage.setItem('tallyPort', tallyPort);
            
            refreshServerBtn.textContent = 'Refreshing...';
            refreshServerBtn.disabled = true;
            
            try {
                alert('Session refreshed with new local settings!');
                settingsDialog.classList.add('hidden');
                loadCompanies(); // reload the dropdowns
                updateConfigDisplay({ host: localStorage.getItem('tallyHost') || '172.16.1.4', port: localStorage.getItem('tallyPort') || '9000' });
            } catch(e) {
                console.error(e);
            } finally {
                refreshServerBtn.textContent = 'Refresh Session';
                refreshServerBtn.disabled = false;
            }
        });
    }

    async function loadConfig() {
        const host = localStorage.getItem('tallyHost') || '172.16.1.4';
        const port = localStorage.getItem('tallyPort') || '9000';
        updateConfigDisplay({ host, port });
        
        const ipInput = document.getElementById('tally-ip-input');
        const portInput = document.getElementById('tally-port-input');
        if (ipInput) ipInput.value = host;
        if (portInput) portInput.value = port;
    }

    function updateConfigDisplay(config) {
        if(config && config.host && config.port) {
            const ipEl = document.getElementById('display-ip');
            const portEl = document.getElementById('display-port');
            if (ipEl) ipEl.textContent = config.host;
            if (portEl) portEl.textContent = config.port;
        }
    }

    initChecker();
    loadConfig();
});