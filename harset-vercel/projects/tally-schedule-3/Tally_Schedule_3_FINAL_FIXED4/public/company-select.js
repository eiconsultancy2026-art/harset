/**
 * <company-select> Web Component
 * A styled dropdown with a built-in "Pvt. Ltd only" checkbox filter.
 *
 * Attributes:
 *   side        – "left" | "right"
 *   label       – text shown as the dropdown label
 *   placeholder – default option text
 *
 * Methods:
 *   setCompanies(list)  – populate with a string array
 *   getValue()          – returns the currently selected value
 *
 * Events:
 *   change – fired whenever the selected company changes
 */
class CompanySelect extends HTMLElement {
    constructor() {
        super();
        this._companies = [];
        this._pvtOnly   = false;
        this.attachShadow({ mode: 'open' });
    }

    connectedCallback() {
        this._render();
    }

    get side()        { return this.getAttribute('side')        || ''; }
    get label()       { return this.getAttribute('label')       || 'Company'; }
    get placeholder() { return this.getAttribute('placeholder') || '-- Select Company --'; }

    _isPvtLtd(name) {
        return /pvt\.?\s*ltd|private\s+limited/i.test(name);
    }

    _render() {
        this.shadowRoot.innerHTML = `
            <style>
                :host { display: block; width: 100%; }

                .wrapper {
                    position: relative;
                    width: 100%;
                }

                /* Custom dropdown trigger */
                .trigger {
                    display: flex;
                    align-items: center;
                    justify-content: space-between;
                    background: #1e293b;
                    border: 1px solid #334155;
                    border-radius: 0.5rem;
                    padding: 0.75rem 1rem;
                    color: white;
                    font-size: 1rem;
                    font-family: 'Inter', sans-serif;
                    cursor: pointer;
                    transition: border-color 0.2s, box-shadow 0.2s;
                    user-select: none;
                    gap: 0.5rem;
                }
                .trigger:hover, .trigger.open {
                    border-color: #3b82f6;
                    box-shadow: 0 0 0 2px rgba(59,130,246,0.25);
                }
                .trigger-text { flex: 1; text-align: left; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; min-width: 0; }
                .trigger-arrow { transition: transform 0.2s; color: #94a3b8; font-size:0.85rem; flex-shrink:0; }
                .trigger.open .trigger-arrow { transform: rotate(180deg); }

                /* Dropdown panel */
                .panel {
                    position: absolute;
                    top: calc(100% + 4px);
                    left: 0; right: 0;
                    background: #1e293b;
                    border: 1px solid #3b82f6;
                    border-radius: 0.6rem;
                    box-shadow: 0 12px 30px rgba(0,0,0,0.5);
                    z-index: 9999;
                    overflow: hidden;
                    display: none;
                    animation: fadeIn 0.15s ease;
                }
                .panel.open { display: block; }

                @keyframes fadeIn {
                    from { opacity: 0; transform: translateY(-6px); }
                    to   { opacity: 1; transform: translateY(0); }
                }

                /* Filter bar inside panel */
                .filter-bar {
                    display: flex;
                    align-items: center;
                    gap: 0.5rem;
                    padding: 0.6rem 0.9rem;
                    border-bottom: 1px solid #334155;
                    background: rgba(59,130,246,0.07);
                }
                .filter-bar label {
                    display: flex;
                    align-items: center;
                    gap: 0.4rem;
                    font-size: 0.82rem;
                    color: #94a3b8;
                    cursor: pointer;
                    font-family: 'Inter', sans-serif;
                    user-select: none;
                }
                .filter-bar input[type=checkbox] {
                    width: 15px; height: 15px;
                    accent-color: #3b82f6;
                    cursor: pointer;
                }
                .filter-bar input:checked + span {
                    color: #60a5fa;
                    font-weight: 600;
                }

                /* Options list */
                .options {
                    max-height: 220px;
                    overflow-y: auto;
                    scrollbar-width: thin;
                    scrollbar-color: #334155 transparent;
                }
                .option {
                    padding: 0.65rem 1rem;
                    font-size: 0.95rem;
                    font-family: 'Inter', sans-serif;
                    color: #e2e8f0;
                    cursor: pointer;
                    transition: background 0.12s;
                }
                .option:hover { background: rgba(59,130,246,0.18); color: white; }
                .option.selected { background: rgba(59,130,246,0.28); color: #93c5fd; font-weight: 600; }
                .option.placeholder { color: #64748b; cursor: default; font-style: italic; }
                .option.empty-msg { color: #ef4444; font-size: 0.85rem; font-style: italic; cursor: default; }
            </style>

            <div class="wrapper">
                <div class="trigger" id="trigger">
                    <span class="trigger-text" id="trigger-text">${this.placeholder}</span>
                    <span class="trigger-arrow">▼</span>
                </div>
                <div class="panel" id="panel">
                    <div class="filter-bar">
                        <label>
                            <input type="checkbox" id="pvt-check">
                            <span>Pvt. Ltd only</span>
                        </label>
                    </div>
                    <div class="options" id="options"></div>
                </div>
            </div>
        `;

        this._trigger     = this.shadowRoot.getElementById('trigger');
        this._triggerText = this.shadowRoot.getElementById('trigger-text');
        this._panel       = this.shadowRoot.getElementById('panel');
        this._optionsList = this.shadowRoot.getElementById('options');
        this._pvtCheck    = this.shadowRoot.getElementById('pvt-check');

        this._selected = '';   // current value

        // Toggle open/close
        this._trigger.addEventListener('click', (e) => {
            e.stopPropagation();
            this._togglePanel();
        });

        // Pvt Ltd checkbox
        this._pvtCheck.addEventListener('change', () => {
            this._pvtOnly = this._pvtCheck.checked;
            this._renderOptions();
        });

        // Stop propagation on panel clicks so it doesn't close when clicking checkbox/options
        this._panel.addEventListener('click', (e) => {
            e.stopPropagation();
        });

        // Close on outside click
        document.addEventListener('click', () => this._closePanel());

        this._renderOptions();
    }

    _togglePanel() {
        const open = this._panel.classList.contains('open');
        open ? this._closePanel() : this._openPanel();
    }

    _openPanel() {
        this._panel.classList.add('open');
        this._trigger.classList.add('open');
    }

    _closePanel() {
        this._panel.classList.remove('open');
        this._trigger.classList.remove('open');
    }

    _renderOptions() {
        const list = this._pvtOnly
            ? this._companies.filter(c => this._isPvtLtd(c))
            : this._companies;

        this._optionsList.innerHTML = '';

        if (this._companies.length === 0) {
            this._addOption('', 'Loading...', 'placeholder');
            return;
        }

        if (list.length === 0 && this._pvtOnly) {
            this._addOption('', 'No Pvt Ltd companies found', 'empty-msg');
            return;
        }

        // Placeholder row
        this._addOption('', this.placeholder, 'placeholder');

        list.forEach(c => {
            const div = this._addOption(c, c, c === this._selected ? 'selected' : '');
            div.addEventListener('click', (e) => {
                e.stopPropagation();
                this._selectValue(c);
            });
        });
    }

    _addOption(value, text, cls = '') {
        const div = document.createElement('div');
        div.className = `option ${cls}`.trim();
        div.dataset.value = value;
        div.textContent = text;
        this._optionsList.appendChild(div);
        return div;
    }

    _selectValue(val) {
        this._selected = val;
        this._triggerText.textContent = val || this.placeholder;
        this._closePanel();
        this._renderOptions();   // refresh selected highlight
        this.dispatchEvent(new CustomEvent('change', {
            detail: { value: val },
            bubbles: true,
            composed: true
        }));
    }

    /** Public API **/
    setCompanies(companies) {
        this._companies = companies || [];
        // Reset selected if it's no longer in the list
        if (!this._companies.includes(this._selected)) {
            this._selected = '';
            this._triggerText.textContent = this.placeholder;
        }
        this._renderOptions();
    }

    getValue() {
        return this._selected;
    }

    reset() {
        this._selectValue('');
    }
}

customElements.define('company-select', CompanySelect);
