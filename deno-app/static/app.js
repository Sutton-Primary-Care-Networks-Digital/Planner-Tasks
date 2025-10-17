/**
 * Microsoft Planner Task Creator - Frontend Application
 */

class PlannerApp {
    constructor() {
        this.sessionId = null;
        this.parsedData = null;
        this.processedTasks = null;
        this.planners = [];
        this.buckets = [];
        this.selectedPlannerId = null;
        this.selectedPlannerGroupId = null;
        this.selectedBucketId = null;
        this.plannerMembers = [];
        this.assigneeOverrides = {};
        this.statusOverrides = {};
        this.bucketSelectionSet = new Set(); // Selected missing bucket names
        
        this.init();
    }

    normalizeStatusKey(status) {
        if (!status) return '';
        const s = String(status).toLowerCase().trim();
        if (s === 'done' || s === 'completed' || s === 'complete') return 'complete';
        if (s === 'in progress' || s === 'in_progress') return 'in_progress';
        if (s === 'not started' || s === 'not_started') return 'not_started';
        return s.replace(/\s+/g, '_');
    }

    handleCreateAllBucketsToggle(event) {
        const checked = event.target.checked;
        // Treat as select-all/none for inline bucket selection
        const list = document.getElementById('bucket-selection-list-inline');
        const boxes = list ? list.querySelectorAll('.bucket-sel-check-inline') : [];
        if (!(this.bucketSelectionSet instanceof Set)) this.bucketSelectionSet = new Set();
        this.bucketSelectionSet.clear();
        boxes.forEach(cb => {
            cb.checked = checked;
            const name = cb.getAttribute('data-name');
            if (checked) this.bucketSelectionSet.add(name);
        });
        // Update both previews
        this.showTaskPreview();
        // If Final Preview is open, refresh
        this.buildFinalPreview();
    }
    
    init() {
        this.setupEventListeners();
        this.checkExistingAuth();
    }
    
    async checkExistingAuth() {
        // Check if we have a stored session ID
        const storedSessionId = localStorage.getItem('planner-session-id');
        if (storedSessionId) {
            this.sessionId = storedSessionId;
            try {
                const response = await this.apiRequest('/api/auth/status', { suppressAuthToast: true });
                if (response.authenticated) {
                    await this.onAuthenticated();
                    return;
                }
                // Stored session but not authenticated
                this.sessionId = null;
                localStorage.removeItem('planner-session-id');
                this.showStep('auth');
                return;
            } catch (error) {
                // Auth check failed, clear stored session
                this.sessionId = null;
                localStorage.removeItem('planner-session-id');
                this.showStep('auth');
                return;
            }
        }
        
        // No valid existing auth, show auth step
        this.showStep('auth');
    }

    async verifyPlannerAccess() {
        // Perform a silent preflight call to confirm we can list planners; if not, auto sign-out without toast
        try {
            const data = await this.apiRequest('/api/planners', { suppressAuthToast: true });
            const planners = Array.isArray(data.planners) ? data.planners : [];
            if (planners.length === 0) {
                this.sessionId = null;
                localStorage.removeItem('planner-session-id');
                this.showStep('auth');
                return false;
            }
            return true;
        } catch (_e) {
            // Treat any failure as invalid auth; sign out silently
            this.sessionId = null;
            localStorage.removeItem('planner-session-id');
            this.showStep('auth');
            return false;
        }
    }
    
    setupEventListeners() {
        // Authentication
        document.getElementById('browser-auth-btn').addEventListener('click', () => this.startBrowserAuth());
        document.getElementById('sign-out-btn').addEventListener('click', () => this.signOut());
        
        // File upload
        document.getElementById('file-input').addEventListener('change', (e) => this.handleFileSelect(e));
        document.getElementById('process-file-btn').addEventListener('click', () => this.processFile());
        
        // Column mapping
        document.getElementById('map-columns-btn').addEventListener('click', () => this.processDataMapping());
        
        // Planner selection
        document.getElementById('planner-select').addEventListener('change', (e) => this.handlePlannerChange(e));
        document.getElementById('bucket-select').addEventListener('change', (e) => this.handleBucketChange(e));
        document.getElementById('confirm-creation').addEventListener('change', (e) => this.handleConfirmationChange(e));
        document.getElementById('create-tasks-btn').addEventListener('click', () => this.createTasks());
        // Bucket preview controls (button may not exist anymore)
        const bpBtn = document.getElementById('bucket-preview-btn');
        if (bpBtn) bpBtn.addEventListener('click', () => this.openBucketPreview());
        const bpClose = document.getElementById('bucket-preview-close');
        if (bpClose) bpClose.addEventListener('click', () => this.closeBucketPreview());
        const bpOk = document.getElementById('bucket-preview-ok');
        if (bpOk) bpOk.addEventListener('click', () => this.closeBucketPreview());
        const bpCreate = document.getElementById('bucket-create-btn');
        if (bpCreate) bpCreate.addEventListener('click', () => this.createSelectedBuckets());
        // Create all missing buckets toggle
        document.getElementById('create-missing-buckets-toggle').addEventListener('change', (e) => this.handleCreateAllBucketsToggle(e));
        // Close any open assignee dropdowns when clicking outside
        document.addEventListener('click', (e) => {
            const menus = document.querySelectorAll('.assignee-menu');
            menus.forEach(menu => {
                if (!menu.contains(e.target) && !menu.previousElementSibling?.contains(e.target)) {
                    menu.classList.add('hidden');
                }
            });
        });

        // Delegated handlers to make assignee checkboxes resilient across rerenders
        document.addEventListener('change', (e) => {
            const t = e.target;
            if (!(t instanceof HTMLElement)) return;
            if (t instanceof HTMLInputElement && t.classList && t.classList.contains('assignee-check')) {
                e.stopPropagation();
                const idx = parseInt(t.getAttribute('data-task-index') || '-1', 10);
                const id = t.getAttribute('value') || '';
                if (isNaN(idx) || !id) return;
                if (!this.assigneeOverrides[idx]) this.assigneeOverrides[idx] = Array.isArray(this.assigneeOverrides[idx]) ? this.assigneeOverrides[idx] : [];
                const list = new Set(this.assigneeOverrides[idx]);
                if (t.checked) list.add(id); else list.delete(id);
                this.assigneeOverrides[idx] = Array.from(list);
                const menu = t.closest('.assignee-menu');
                const btn = menu?.previousElementSibling;
                if (btn && btn.classList.contains('assignee-toggle')) {
                    const count = this.assigneeOverrides[idx].length;
                    const label = count === 0 ? 'Assign...' : (count === 1 ? ((this.plannerMembers.find(m=>m.id===this.assigneeOverrides[idx][0])||{}).displayName || '1 selected') : `${count} selected`);
                    btn.textContent = label;
                }
                const summary = document.querySelector(`.assignee-summary[data-task-index="${idx}"]`);
                if (summary) {
                    summary.textContent = this.renderAssigneeSummary(this.assigneeOverrides[idx]);
                }
                const info = document.getElementById(`assignee-info-${idx}`);
                if (info) {
                    const ids = this.assigneeOverrides[idx] || [];
                    if (ids.length > 0) {
                        const names = ids.map(id2 => (this.plannerMembers.find(m => m.id === id2) || {}).displayName).filter(Boolean).join(', ');
                        info.innerHTML = `<span class="text-green-700">👤 ${names}</span>`;
                    } else {
                        info.innerHTML = '';
                    }
                }
            }
        }, true);
        document.addEventListener('mousedown', (e) => {
            const t = e.target;
            if (!(t instanceof HTMLElement)) return;
            if (t.classList && (t.classList.contains('assignee-check') || t.closest('.assignee-list label'))) {
                e.stopPropagation();
            }
        }, true);
        document.addEventListener('click', (e) => {
            const t = e.target;
            if (!(t instanceof HTMLElement)) return;
            if (t.classList && (t.classList.contains('assignee-check') || t.closest('.assignee-list label'))) {
                e.stopPropagation();
            }
        }, true);
        
        // Results
        document.getElementById('create-more-btn').addEventListener('click', () => this.reset());

        // Top nav: Workflow & Prompt page
        const workflowLink = document.getElementById('workflow-link');
        if (workflowLink) {
            workflowLink.addEventListener('click', (e) => {
                e.preventDefault();
                this.showWorkflow();
            });
        }
        const homeLink = document.getElementById('home-link');
        if (homeLink) {
            homeLink.addEventListener('click', (e) => {
                e.preventDefault();
                this.goHome();
            });
        }
        const backToTool = document.getElementById('back-to-tool');
        if (backToTool) {
            backToTool.addEventListener('click', (e) => {
                e.preventDefault();
                this.goHome();
            });
        }

        // Final Preview modal controls
        const fpClose = document.getElementById('final-preview-close');
        const fpCancel = document.getElementById('final-preview-cancel');
        const fpConfirm = document.getElementById('final-preview-confirm');
        if (fpClose) fpClose.addEventListener('click', () => this.closeFinalPreview());
        if (fpCancel) fpCancel.addEventListener('click', () => this.closeFinalPreview());
        if (fpConfirm) fpConfirm.addEventListener('click', () => this.createTasksConfirmed());

        // Inline Bucket selection list wiring
        this.buildBucketSelectionListInline();

        // Copilot prompt copy handler
        const copyBtn = document.getElementById('copy-copilot-prompt');
        if (copyBtn) {
            copyBtn.addEventListener('click', async () => {
                try {
                    const ta = document.getElementById('copilot-prompt');
                    if (ta) {
                        await navigator.clipboard.writeText(ta.value);
                        const ok = document.getElementById('copilot-prompt-copied');
                        if (ok) {
                            ok.classList.remove('hidden');
                            setTimeout(() => ok.classList.add('hidden'), 1500);
                        }
                    }
                } catch (e) {
                    this.showToast('Copy failed. Select and copy manually.', 'warning');
                }
            });
        }
    }
    
    showStep(step) {
        // Hide all sections
        const sections = ['auth-section', 'upload-section', 'column-mapping-section', 'planner-section', 'results-section', 'workflow-section'];
        sections.forEach(sectionId => {
            document.getElementById(sectionId).classList.add('hidden');
        });
        
        // Show target section
        document.getElementById(`${step}-section`).classList.remove('hidden');
        
        // Update step indicator if authenticated
        if (this.sessionId) {
            document.getElementById('step-indicator').classList.remove('hidden');
            document.getElementById('user-info').classList.remove('hidden');
            this.updateStepIndicator(step);
        }
    }

    showWorkflow() {
        // Hide step indicator while in docs page
        document.getElementById('step-indicator').classList.add('hidden');
        // Show workflow page only
        this.showStep('workflow');
        // Ensure workflow-section is visible even though not a numbered step
        document.getElementById('workflow-section')?.classList.remove('hidden');
    }

    goHome() {
        // If authenticated, go to current logical step (upload) else auth
        if (this.sessionId) {
            this.showStep('upload');
        } else {
            this.showStep('auth');
        }
    }

    openBucketSelection() {
        const modal = document.getElementById('bucket-selection-modal');
        this.buildBucketSelectionList();
        modal.classList.remove('hidden');
    }

    closeBucketSelection() {
        document.getElementById('bucket-selection-modal').classList.add('hidden');
    }

    buildBucketSelectionList() {
        const list = document.getElementById('bucket-selection-list');
        if (!list) return;
        const tasks = (this.processedTasks && this.processedTasks.tasks) ? this.processedTasks.tasks : [];
        const { createMap } = this.collectBucketPreviewData(tasks);
        // Initialize selection set if not present
        if (!(this.bucketSelectionSet instanceof Set)) {
            this.bucketSelectionSet = new Set(Object.keys(createMap));
        }
        list.innerHTML = '';
        const names = Object.keys(createMap).sort();
        names.forEach(name => {
            const id = `bucket-sel-${name.replace(/[^a-z0-9]/gi, '_')}`;
            const li = document.createElement('li');
            li.className = 'flex items-center justify-between';
            li.innerHTML = `<label for="${id}" class="text-sm text-gray-700">${name} <span class="text-xs text-gray-500">(${createMap[name]} task(s))</span></label>`+
                           `<input id="${id}" type="checkbox" class="bucket-sel-check" data-name="${name}" ${this.bucketSelectionSet.has(name)?'checked':''}>`;
            list.appendChild(li);
        });
        // Wire individual checks
        list.querySelectorAll('.bucket-sel-check').forEach(cb => {
            cb.addEventListener('change', (e) => {
                const name = e.target.getAttribute('data-name');
                if (e.target.checked) this.bucketSelectionSet.add(name); else this.bucketSelectionSet.delete(name);
                // Keep Select All in sync
                const bsSelectAll = document.getElementById('bucket-select-all');
                if (bsSelectAll) {
                    bsSelectAll.checked = this.bucketSelectionSet.size === names.length;
                }
            });
        });
        // Initialize Select All state
        const bsSelectAll = document.getElementById('bucket-select-all');
        if (bsSelectAll) {
            bsSelectAll.checked = this.bucketSelectionSet.size === names.length && names.length > 0;
        }
    }

    toggleBucketSelectAll(event) {
        const checked = event.target.checked;
        const list = document.getElementById('bucket-selection-list');
        const boxes = list ? list.querySelectorAll('.bucket-sel-check') : [];
        if (!(this.bucketSelectionSet instanceof Set)) this.bucketSelectionSet = new Set();
        boxes.forEach(cb => {
            cb.checked = checked;
            const name = cb.getAttribute('data-name');
            if (checked) this.bucketSelectionSet.add(name); else this.bucketSelectionSet.delete(name);
        });
    }

    applyBucketSelection() {
        this.closeBucketSelection();
        // Rebuild Final Preview to reflect new selection
        this.buildFinalPreview();
    }
    
    updateStepIndicator(currentStep) {
        const steps = {
            'auth': 1,
            'upload': 2,
            'column-mapping': 3,
            'planner': 4,
            'results': 5
        };
        
        const currentStepNum = steps[currentStep] || 1;
        
        for (let i = 1; i <= 5; i++) {
            const stepElement = document.getElementById(`step-${i}`);
            const circle = stepElement.querySelector('div');
            const text = stepElement.querySelector('span');
            
            if (i <= currentStepNum) {
                circle.classList.remove('bg-gray-300', 'text-gray-600');
                circle.classList.add('bg-teal-600', 'text-white');
                text.classList.remove('text-gray-500');
                text.classList.add('text-gray-700');
            } else {
                circle.classList.remove('bg-teal-600', 'text-white');
                circle.classList.add('bg-gray-300', 'text-gray-600');
                text.classList.remove('text-gray-700');
                text.classList.add('text-gray-500');
            }
        }
    }
    
    showLoading(message = 'Processing...') {
        document.getElementById('loading-message').textContent = message;
        document.getElementById('loading-overlay').classList.remove('hidden');
    }
    
    hideLoading() {
        document.getElementById('loading-overlay').classList.add('hidden');
    }
    
    showError(message) {
        alert(`Error: ${message}`);
    }

    showToast(message, type = 'info', durationMs = 3500) {
        const cont = document.getElementById('toast-container');
        if (!cont) return;
        const el = document.createElement('div');
        const base = 'fade-in px-4 py-2 rounded shadow text-sm flex items-center gap-2';
        const style = type === 'error' ? 'bg-red-600 text-white' : type === 'warning' ? 'bg-yellow-500 text-black' : 'bg-gray-900 text-white';
        el.className = `${base} ${style}`;
        el.innerHTML = `${type === 'error' ? '⚠️' : type === 'warning' ? '⚠️' : 'ℹ️'} <span>${message}</span>`;
        cont.appendChild(el);
        setTimeout(() => { try { cont.removeChild(el); } catch {} }, durationMs);
    }
    
    async apiRequest(endpoint, options = {}) {
        const headers = {
            'Content-Type': 'application/json',
            ...options.headers
        };
        
        if (this.sessionId) {
            headers['X-Session-ID'] = this.sessionId;
        }
        
        const response = await fetch(endpoint, {
            ...options,
            headers
        });
        
        if (!response.ok) {
            // Handle token expiration or invalid session
            if (response.status === 401 || response.status === 403) {
                try { await response.clone().json(); } catch {}
                this.sessionId = null;
                localStorage.removeItem('planner-session-id');
                if (!options.suppressAuthToast) {
                    this.showToast('Session expired. Please sign in again.', 'warning');
                }
                this.showStep('auth');
                throw new Error('Session expired. Please sign in again.');
            }
            const errorData = await response.json().catch(() => ({ error: 'Unknown error' }));
            throw new Error(errorData.error || `HTTP ${response.status}`);
        }
        
        return response.json();
    }

    async ensureAuthenticatedOrSignOut() {
        try {
            const status = await this.apiRequest('/api/auth/status', { suppressAuthToast: true });
            if (status && status.authenticated) return true;
        } catch (_e) {
            // apiRequest already handled UI sign-out and redirect on 401/403
            return false;
        }
        // Fallback: if status returned but not authenticated
        this.sessionId = null;
        localStorage.removeItem('planner-session-id');
        this.showStep('auth');
        return false;
    }
    
    async startBrowserAuth() {
        try {
            this.showLoading('Opening browser for authentication...');
            
            // Call the interactive auth endpoint (this will open browser and handle everything)
            const response = await this.apiRequest('/api/auth/interactive', {
                method: 'POST'
            });
            
            if (response.success) {
                this.sessionId = response.sessionId;
                localStorage.setItem('planner-session-id', response.sessionId);
                this.hideLoading();
                await this.onAuthenticated();
            } else {
                throw new Error(response.error || 'Authentication failed');
            }
            
        } catch (error) {
            this.hideLoading();
            this.showError(error.message);
        }
    }
    
    
    async onAuthenticated() {
        // Proactively verify that Planner APIs are accessible; if not, we will be redirected to sign-in
        const ok = await this.verifyPlannerAccess();
        if (!ok) return;
        this.showStep('upload');
    }
    
    async signOut() {
        try {
            await this.apiRequest('/api/auth/signout', { method: 'POST' });
        } catch (error) {
            // Ignore errors
        }
        
        this.sessionId = null;
        localStorage.removeItem('planner-session-id');
        document.getElementById('user-info').classList.add('hidden');
        document.getElementById('step-indicator').classList.add('hidden');
        this.reset();
        this.showStep('auth');
    }
    
    handleFileSelect(event) {
        const file = event.target.files[0];
        if (file) {
            document.getElementById('file-name').textContent = `${file.name} (${Math.round(file.size / 1024)} KB)`;
            document.getElementById('file-info').classList.remove('hidden');
        }
    }
    
    async processFile() {
        const fileInput = document.getElementById('file-input');
        const file = fileInput.files[0];
        
        if (!file) {
            this.showError('Please select a file first');
            return;
        }
        
        try {
            this.showLoading('Processing file...');
            
            const formData = new FormData();
            formData.append('file', file);
            
            const response = await fetch('/api/parse-file', {
                method: 'POST',
                headers: {
                    'X-Session-ID': this.sessionId
                },
                body: formData
            });
            
            if (!response.ok) {
                throw new Error('Failed to process file');
            }
            
            const data = await response.json();
            this.parsedData = data;
            
            this.populateColumnMapping(data.availableColumns);
            this.showStep('column-mapping');
            
        } catch (error) {
            this.showError(error.message);
        } finally {
            this.hideLoading();
        }
    }
    
    populateColumnMapping(columns) {
        const selects = [
            'title-column', 'description-column', 'start-date-column',
            'due-date-column', 'assignee-column', 'bucket-column', 'status-column'
        ];
        
        selects.forEach(selectId => {
            const select = document.getElementById(selectId);
            select.innerHTML = '<option value="">Select column...</option>';
            columns.forEach(column => {
                const option = document.createElement('option');
                option.value = column;
                option.textContent = column;
                select.appendChild(option);
            });
        });
        
        // Show available columns
        const columnsList = document.getElementById('available-columns');
        columnsList.innerHTML = '';
        columns.forEach((column, index) => {
            const li = document.createElement('li');
            li.textContent = `${index + 1}. ${column}`;
            li.className = 'p-2 bg-gray-100 rounded';
            columnsList.appendChild(li);
        });
        
        // Auto-select common column names
        this.autoSelectColumns(columns);
    }
    
    autoSelectColumns(columns) {
        const mappings = [
            { select: 'title-column', patterns: ['title', 'task', 'name', 'subject'] },
            { select: 'description-column', patterns: ['description', 'desc', 'details', 'notes'] },
            { select: 'start-date-column', patterns: ['start', 'begin', 'start_date', 'startdate'] },
            { select: 'due-date-column', patterns: ['due', 'end', 'deadline', 'due_date', 'duedate'] },
            { select: 'assignee-column', patterns: ['assign', 'user', 'owner', 'responsible', 'assignee'] },
            { select: 'bucket-column', patterns: ['bucket', 'category', 'group', 'section'] },
            { select: 'status-column', patterns: ['status', 'state', 'progress', 'complete'] }
        ];
        
        mappings.forEach(mapping => {
            const select = document.getElementById(mapping.select);
            const matchedColumn = columns.find(column =>
                mapping.patterns.some(pattern =>
                    column.toLowerCase().includes(pattern.toLowerCase())
                )
            );
            
            if (matchedColumn) {
                select.value = matchedColumn;
            }
        });
    }
    
    async processDataMapping() {
        const titleColumn = document.getElementById('title-column').value;
        if (!titleColumn) {
            this.showError('Please select a title column');
            return;
        }
        
        const columnMapping = {
            title: titleColumn,
            description: document.getElementById('description-column').value || undefined,
            startDate: document.getElementById('start-date-column').value || undefined,
            dueDate: document.getElementById('due-date-column').value || undefined,
            assignee: document.getElementById('assignee-column').value || undefined,
            bucketName: document.getElementById('bucket-column').value || undefined,
            status: document.getElementById('status-column').value || undefined
        };
        
        try {
            this.showLoading('Processing data mapping...');
            
            const data = await this.apiRequest('/api/process-data', {
                method: 'POST',
                body: JSON.stringify({ columnMapping })
            });
            
            this.processedTasks = data;
            
            // Load planners
            await this.loadPlanners();
            
            this.showStep('planner');
            
        } catch (error) {
            this.showError(error.message);
        } finally {
            this.hideLoading();
        }
    }
    
    async loadPlanners() {
        try {
            const ok = await this.ensureAuthenticatedOrSignOut();
            if (!ok) return; // Redirected to sign-in already
            const data = await this.apiRequest('/api/planners', { suppressAuthToast: true });
            this.planners = Array.isArray(data.planners) ? data.planners : [];
            if (this.planners.length === 0) {
                // Treat as invalid/expired access; sign out silently and redirect
                this.sessionId = null;
                localStorage.removeItem('planner-session-id');
                this.showStep('auth');
                return;
            }
            
            const select = document.getElementById('planner-select');
            select.innerHTML = '<option value="">Select a planner...</option>';
            
            this.planners.forEach(planner => {
                const option = document.createElement('option');
                option.value = planner.id;
                option.textContent = `${planner.title} (${planner.groupName})`;
                select.appendChild(option);
            });
            
        } catch (error) {
            // apiRequest handles 401/403; for other errors, redirect silently
            this.sessionId = null;
            localStorage.removeItem('planner-session-id');
            this.showStep('auth');
        }
    }
    
    async handlePlannerChange(event) {
        const plannerId = event.target.value;
        this.selectedPlannerId = plannerId;
        const planner = this.planners.find(p => p.id === plannerId);
        this.selectedPlannerGroupId = planner ? planner.groupId : null;
        
        if (!plannerId) {
            document.getElementById('bucket-section').classList.add('hidden');
            document.getElementById('task-preview').classList.add('hidden');
            return;
        }
        
        try {
            this.showLoading('Loading buckets...');
            
            const data = await this.apiRequest(`/api/planners/${plannerId}/buckets`);
            this.buckets = data.buckets;
            
            const select = document.getElementById('bucket-select');
            select.innerHTML = '<option value="">Select default bucket...</option>';
            
            this.buckets.forEach(bucket => {
                const option = document.createElement('option');
                option.value = bucket.id;
                option.textContent = bucket.name;
                select.appendChild(option);
            });
            
            document.getElementById('bucket-section').classList.remove('hidden');
            
            // Look up assignees if we haven't already
            if (this.processedTasks && !this.processedTasks.assigneesLoaded) {
                await this.lookupAssignees();
            }
            
            // Load planner members for manual assignment dropdowns
            try {
                const membersResp = await this.apiRequest(`/api/planners/${plannerId}/members${this.selectedPlannerGroupId ? `?groupId=${encodeURIComponent(this.selectedPlannerGroupId)}` : ''}`);
                this.plannerMembers = membersResp.members || [];
            } catch (_e) {
                this.plannerMembers = [];
            }

            // Refresh bucketInfo for tasks against live plan buckets
            try {
                const bucketsResp = await this.apiRequest('/api/lookup-buckets', {
                    method: 'POST',
                    body: JSON.stringify({ planId: plannerId })
                });
                if (this.processedTasks) {
                    this.processedTasks.tasks = bucketsResp.tasks;
                }
            } catch (_e) {
                // leave tasks as-is
            }
            
            // Show preview
            document.getElementById('task-preview').classList.remove('hidden');
            // Rebuild inline bucket selection based on tasks
            this.buildBucketSelectionListInline();
            this.showTaskPreview();
        } catch (error) {
            this.showError(error.message);
        } finally {
            this.hideLoading();
        }
    }

    async lookupAssignees() {
        try {
            this.showLoading('Looking up assignees...');
            const data = await this.apiRequest('/api/lookup-assignees', {
                method: 'POST',
                body: JSON.stringify({
                    planId: this.selectedPlannerId,
                    groupId: this.selectedPlannerGroupId,
                }),
            });
            if (!this.processedTasks) {
                this.processedTasks = {};
            }
            this.processedTasks.assigneesLoaded = true;
            this.processedTasks.tasks = data.tasks;
        } catch (error) {
            console.warn('Assignee lookup failed:', error.message);
        } finally {
            this.hideLoading();
        }
    }

    showTaskPreview() {
        const listEl = document.getElementById('task-list');
        if (!listEl) return;
        listEl.innerHTML = '';

        const tasks = (this.processedTasks && this.processedTasks.tasks) ? this.processedTasks.tasks : [];
        this.updateTaskStats(tasks);

        const preview = tasks; // show all tasks in preview
        preview.forEach((task, idx) => {
            const div = document.createElement('div');
            div.className = 'p-3 bg-white rounded border';

            // Assignee info display prioritizes overrides
            const overrideIds = Array.isArray(this.assigneeOverrides[idx]) ? this.assigneeOverrides[idx] : [];
            let assigneeInfo = '';
            if (overrideIds.length > 0) {
                const names = overrideIds
                    .map(id => (this.plannerMembers.find(m => m.id === id) || {}).displayName)
                    .filter(Boolean)
                    .join(', ');
                assigneeInfo = `<span class="text-green-700">👤 ${names}</span>`;
            } else if (Array.isArray(task.assigneeUsers) && task.assigneeUsers.length > 0) {
                const names = task.assigneeUsers.map(u => u.displayName || u.originalName || 'Unknown').join(', ');
                assigneeInfo = `<span class="text-green-700">👤 ${names}</span>`;
            } else if (task.assigneeNeedsReview) {
                assigneeInfo = `<span class="text-yellow-700">👤 Needs review</span>`;
            } else if (task.assigneeUser) {
                assigneeInfo = `<span class="text-gray-700">👤 ${task.assigneeUser.originalName}</span>`;
            } else {
                assigneeInfo = `<span class="text-gray-500">👤 Not found</span>`;
            }

            let bucketInfo = '';
            if (task.bucketInfo) {
                bucketInfo = `<span class="text-blue-600">📂 ${task.bucketInfo.name}</span>`;
            } else if (task.bucketName) {
                bucketInfo = `<span class="text-yellow-600">📂 ${task.bucketName} (not found)</span>`;
            } else {
                const def = (this.buckets || []).find(b => b.id === this.selectedBucketId);
                bucketInfo = `<span class="text-gray-600">📂 ${def ? def.name : 'Default'}</span>`;
            }

            const startDate = task.startDate ? task.startDate.substring(0, 10) : '';
            const dueDate = task.dueDate ? task.dueDate.substring(0, 10) : '';
            const desc = task.description ? task.description : '';

            // Assignee dropdown with checkboxes + search + summary
            if (!Array.isArray(this.assigneeOverrides[idx])) {
                const detected = Array.isArray(task.assigneeUsers) ? task.assigneeUsers.map(u => u.id).filter(Boolean) : [];
                if (detected.length > 0) this.assigneeOverrides[idx] = detected;
            }
            const selectedIds = Array.isArray(this.assigneeOverrides[idx]) ? this.assigneeOverrides[idx] : [];
            const candidateChecks = (task.assigneeCandidates || []).map(c => `
                <label class=\"flex items-center justify-between py-1 text-sm\">\n                  <span class=\"mr-3\">${c.displayName}</span>\n                  <input type=\"checkbox\" class=\"assignee-check\" data-task-index=\"${idx}\" value=\"${c.id}\" ${selectedIds.includes(c.id)?'checked':''}>\n                </label>`).join('');
            const memberChecks = (this.plannerMembers || []).map(m => `
                <label class=\"flex items-center justify-between py-1 text-sm\">\n                  <span class=\"mr-3\">${m.displayName}</span>\n                  <input type=\"checkbox\" class=\"assignee-check\" data-task-index=\"${idx}\" value=\"${m.id}\" ${selectedIds.includes(m.id)?'checked':''}>\n                </label>`).join('');
            const selectedSummary = selectedIds.length === 0
                ? 'Assign...'
                : (selectedIds.length === 1
                    ? ((this.plannerMembers.find(m=>m.id===selectedIds[0])||{}).displayName || '1 selected')
                    : `${selectedIds.length} selected`);
            const manualAssignHtml = `
              <div class="mt-2 relative inline-block">
                <button type="button" class="assignee-toggle border border-gray-300 rounded px-2 py-1 bg-white text-sm" data-task-index="${idx}">${selectedSummary}</button>
                <div class="assignee-menu hidden absolute z-50 mt-1 w-72 max-h-64 overflow-y-auto bg-white border border-gray-200 rounded shadow">
                  <div class="px-3 py-2 border-b bg-gray-50">
                    <input type="text" class="assignee-search w-full border border-gray-300 rounded px-2 py-1 text-sm" placeholder="Search..." data-task-index="${idx}">
                  </div>
                  ${candidateChecks ? `<div class="px-3 py-1 text-xs text-gray-500 bg-gray-50">Suggested</div>` : ''}
                  <div class="px-3 assignee-list">${candidateChecks}</div>
                  ${memberChecks ? `<div class="px-3 py-1 text-xs text-gray-500 bg-gray-50 border-t">All members</div>` : ''}
                  <div class="px-3 assignee-list">${memberChecks}</div>
                </div>
              </div>
              <div class="assignee-summary text-xs text-gray-600 mt-1" data-task-index="${idx}">${this.renderAssigneeSummary(selectedIds)}</div>`;

            // Status dropdown
            const statusOptions = [
                { key: 'not_started', label: 'Not Started' },
                { key: 'in_progress', label: 'In Progress' },
                { key: 'complete', label: 'Complete' }
            ];
            const currentStatus = (this.statusOverrides[idx]) || (task.status ? task.status : '');
            const normalizedKey = this.normalizeStatusKey(currentStatus);
            const statusSelectHtml = `
              <div class=\"mt-2\">\n                <label class=\"text-sm text-gray-700 mr-2\">Status:</label>\n                <select data-task-index=\"${idx}\" class=\"status-select border border-gray-300 rounded px-2 py-1\">\n                  <option value=\"\">Select status...</option>\n                  ${statusOptions.map(opt => `<option value=\"${opt.key}\" ${normalizedKey===opt.key?'selected':''}>${opt.label}</option>`).join('')}\n                </select>\n              </div>`;

            // Bucket create info text based on inline selection
            const missing = !task.bucketInfo && task.bucketName;
            let bucketCreateNote = '';
            if (missing) {
                const willCreate = this.bucketSelectionSet.has(task.bucketName);
                bucketCreateNote = willCreate ? `<span class="ml-2 text-xs text-green-700">(will be created)</span>`
                                              : `<span class="ml-2 text-xs text-gray-500">(will use default)</span>`;
            }

            div.innerHTML = `
                <div class="font-semibold text-gray-900 mb-1">Title: <span class="font-medium">${task.title || ''}</span></div>
                ${desc ? `<div class="text-sm text-gray-700 mb-1">Description: <span class="text-gray-600">${desc}</span></div>` : ''}
                <div class="text-sm text-gray-600 space-x-4 items-center">
                    <span id="assignee-info-${idx}">${assigneeInfo}</span>
                    <span>${bucketInfo}${bucketCreateNote}</span>
                    ${startDate ? `<span class="text-teal-700">🟢 Start: ${startDate}</span>` : ''}
                    ${dueDate ? `<span class="text-purple-700">📅 Due: ${dueDate}</span>` : ''}
                </div>
                ${manualAssignHtml}
                ${statusSelectHtml}
            `;

            listEl.appendChild(div);
        });

        // Wire assignee dropdown toggle
        document.querySelectorAll('.assignee-toggle').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const container = e.target.closest('div');
                const menu = container?.querySelector('.assignee-menu');
                if (menu) menu.classList.toggle('hidden');
            });
        });

        // Prevent clicks inside the assignee menu from bubbling to document (which would close it)
        document.querySelectorAll('.assignee-menu').forEach(menu => {
            menu.addEventListener('click', (e) => {
                e.stopPropagation();
            });
        });

        // Wire assignee checkbox lists
        document.querySelectorAll('.assignee-check').forEach(cb => {
            // Prevent outside click handler from firing before change
            cb.addEventListener('mousedown', (e) => e.stopPropagation());
            cb.addEventListener('change', (e) => {
                // Keep menu open and prevent outside-click handler from triggering
                e.stopPropagation();
                const idx = parseInt(e.target.getAttribute('data-task-index'), 10);
                const id = e.target.value;
                if (!this.assigneeOverrides[idx]) this.assigneeOverrides[idx] = Array.isArray(this.assigneeOverrides[idx]) ? this.assigneeOverrides[idx] : [];
                const list = new Set(this.assigneeOverrides[idx]);
                if (e.target.checked) {
                    list.add(id);
                } else {
                    list.delete(id);
                }
                this.assigneeOverrides[idx] = Array.from(list);
                // Update button label summary
                const menu = e.target.closest('.assignee-menu');
                const container = menu?.previousElementSibling;
                const btn = container;
                if (btn && btn.classList.contains('assignee-toggle')) {
                    const count = this.assigneeOverrides[idx].length;
                    const label = count === 0 ? 'Assign...' : (count === 1 ? ((this.plannerMembers.find(m=>m.id===this.assigneeOverrides[idx][0])||{}).displayName || '1 selected') : `${count} selected`);
                    btn.textContent = label;
                }
                // Update textual summary under the dropdown
                const summary = document.querySelector(`.assignee-summary[data-task-index="${idx}"]`);
                if (summary) {
                    summary.textContent = this.renderAssigneeSummary(this.assigneeOverrides[idx]);
                }
                // Update the main assignee info display in the task card
                const info = document.getElementById(`assignee-info-${idx}`);
                if (info) {
                    const ids = this.assigneeOverrides[idx] || [];
                    if (ids.length > 0) {
                        const names = ids.map(id => (this.plannerMembers.find(m => m.id === id) || {}).displayName).filter(Boolean).join(', ');
                        info.innerHTML = `<span class="text-green-700">👤 ${names}</span>`;
                    } else {
                        info.innerHTML = '';
                    }
                }
            });
            // Also intercept direct clicks on the checkbox element
            cb.addEventListener('click', (e) => e.stopPropagation());
        });

        // Also guard label interactions to avoid bubbling to document
        document.querySelectorAll('.assignee-list label').forEach(lbl => {
            lbl.addEventListener('mousedown', (e) => e.stopPropagation());
            lbl.addEventListener('click', (e) => e.stopPropagation());
        });

        // Wire status dropdowns
        document.querySelectorAll('.status-select').forEach(sel => {
            sel.addEventListener('change', (e) => {
                const idx = parseInt(e.target.getAttribute('data-task-index'), 10);
                const v = e.target.value || '';
                if (v) {
                    this.statusOverrides[idx] = v;
                } else if (this.statusOverrides[idx]) {
                    delete this.statusOverrides[idx];
                }
            });
        });

        // Wire assignee search filters
        document.querySelectorAll('.assignee-search').forEach(input => {
            // Prevent clicks/typing from closing the menu
            input.addEventListener('click', (e) => e.stopPropagation());
            input.addEventListener('keydown', (e) => e.stopPropagation());
            input.addEventListener('input', (e) => {
                const q = e.target.value.toLowerCase();
                const menu = e.target.closest('.assignee-menu');
                if (!menu) return;
                menu.querySelectorAll('.assignee-list label').forEach(label => {
                    const text = label.innerText.toLowerCase();
                    label.style.display = text.includes(q) ? '' : 'none';
                });
            });
        });
    }

    buildBucketSelectionListInline() {
        const container = document.getElementById('bucket-selection-list-inline');
        if (!container) return;
        container.innerHTML = '';
        const tasks = (this.processedTasks && this.processedTasks.tasks) ? this.processedTasks.tasks : [];
        // Compute missing buckets and counts
        const missingMap = {};
        tasks.forEach(t => {
            if (!t.bucketInfo && t.bucketName) {
                missingMap[t.bucketName] = (missingMap[t.bucketName] || 0) + 1;
            }
        });
        const names = Object.keys(missingMap).sort();
        // Initialize selection to all when first built
        if (this.bucketSelectionSet.size === 0) {
            names.forEach(n => this.bucketSelectionSet.add(n));
        }
        names.forEach(name => {
            const id = `bucket-sel-inline-${name.replace(/[^a-z0-9]/gi,'_')}`;
            const item = document.createElement('li');
            item.className = 'flex items-center';
            item.innerHTML = `<input id=\"${id}\" type=\"checkbox\" class=\"bucket-sel-check-inline mr-2\" data-name=\"${name}\" ${this.bucketSelectionSet.has(name)?'checked':''}>`+
                             `<label for=\"${id}\" class=\"text-sm text-gray-700\">${name} <span class=\"text-xs text-gray-500\">(${missingMap[name]} task(s))</span></label>`;
            container.appendChild(item);
        });
        // Wire inline checks (debounced rerender to not clobber assignee selections)
        container.querySelectorAll('.bucket-sel-check-inline').forEach(cb => {
            cb.addEventListener('change', (e) => {
                e.stopPropagation();
                const name = e.target.getAttribute('data-name');
                if (e.target.checked) this.bucketSelectionSet.add(name); else this.bucketSelectionSet.delete(name);
                try { if (this._bucketRerenderTimer) clearTimeout(this._bucketRerenderTimer); } catch {}
                this._bucketRerenderTimer = setTimeout(() => {
                    this.showTaskPreview();
                    this.buildFinalPreview();
                    this._bucketRerenderTimer = null;
                }, 150);
            });
            cb.addEventListener('click', (e) => e.stopPropagation());
        });
    }

    renderAssigneeSummary(selectedIds) {
        if (!Array.isArray(selectedIds) || selectedIds.length === 0) return '';
        const names = selectedIds
            .map(id => (this.plannerMembers.find(m => m.id === id) || {}).displayName)
            .filter(Boolean);
        return `Assigned: ${names.join(', ')}`;
    }

    updateTaskStats(tasks) {
        document.getElementById('total-tasks').textContent = String(tasks.length || 0);
        const withAssignees = tasks.filter(t => Array.isArray(t.assigneeUsers) && t.assigneeUsers.length > 0).length;
        const bucketsFound = tasks.filter(t => t.bucketInfo).length;
        const withDueDates = tasks.filter(t => !!t.dueDate).length;
        document.getElementById('assigned-tasks').textContent = String(withAssignees);
        document.getElementById('buckets-found').textContent = String(bucketsFound);
        document.getElementById('due-dates').textContent = String(withDueDates);
    }

    handleBucketChange(event) {
        this.selectedBucketId = event.target.value || null;
    }

    // Bucket preview logic
    openBucketPreview() {
        const tasks = (this.processedTasks && this.processedTasks.tasks) ? this.processedTasks.tasks : [];
        const { createMap, existingMap, defaultCount } = this.collectBucketPreviewData(tasks);

        const createList = document.getElementById('bucket-create-list');
        const existList = document.getElementById('bucket-existing-list');
        createList.innerHTML = '';
        existList.innerHTML = '';

        Object.keys(createMap).sort().forEach(name => {
            const li = document.createElement('li');
            li.textContent = `${name}: ${createMap[name]} task(s)`;
            createList.appendChild(li);
        });

        Object.keys(existingMap).sort().forEach(name => {
            const li = document.createElement('li');
            li.textContent = `${name}: ${existingMap[name]} task(s)`;
            existList.appendChild(li);
        });

        const liDefault = document.createElement('li');
        const defaultBucketName = (this.buckets.find(b => b.id === this.selectedBucketId) || {}).name || 'Default bucket';
        liDefault.textContent = `${defaultBucketName}: ${defaultCount} task(s)`;
        existList.appendChild(liDefault);

        document.getElementById('bucket-preview-modal').classList.remove('hidden');
    }

    closeBucketPreview() {
        document.getElementById('bucket-preview-modal').classList.add('hidden');
    }

    collectBucketPreviewData(tasks) {
        const createMap = {}; // name -> count
        const existingMap = {}; // name -> count
        let defaultCount = 0;

        const createAllToggled = !!(document.getElementById('create-missing-buckets-toggle')?.checked);

        tasks.forEach((t, idx) => {
            if (t.bucketInfo) {
                const name = t.bucketInfo.name;
                existingMap[name] = (existingMap[name] || 0) + 1;
                return;
            }

            if (t.bucketName) {
                const offerCreate = createAllToggled || (this.bucketCreateByTask && this.bucketCreateByTask[idx]);
                if (offerCreate) {
                    createMap[t.bucketName] = (createMap[t.bucketName] || 0) + 1;
                } else {
                    defaultCount += 1;
                }
            } else {
                defaultCount += 1;
            }
        });

        return { createMap, existingMap, defaultCount };
    }

    async createSelectedBuckets() {
        try {
            const tasks = (this.processedTasks && this.processedTasks.tasks) ? this.processedTasks.tasks : [];
            const { createMap } = this.collectBucketPreviewData(tasks);
            const bucketNames = Object.keys(createMap);
            if (bucketNames.length === 0) {
                this.closeBucketPreview();
                return;
            }
            if (!this.selectedPlannerId) {
                this.showError('Please select a planner first.');
                return;
            }
            this.showLoading('Creating buckets...');
            const resp = await this.apiRequest('/api/buckets/create', {
                method: 'POST',
                body: JSON.stringify({ planId: this.selectedPlannerId, bucketNames })
            });
            // Refresh local buckets list
            if (resp && Array.isArray(resp.buckets)) {
                this.buckets = resp.buckets;
                const select = document.getElementById('bucket-select');
                if (select) {
                    const current = this.selectedBucketId;
                    select.innerHTML = '<option value="">Select default bucket...</option>';
                    this.buckets.forEach(bucket => {
                        const option = document.createElement('option');
                        option.value = bucket.id;
                        option.textContent = bucket.name;
                        select.appendChild(option);
                    });
                    if (current && this.buckets.find(b => b.id === current)) {
                        select.value = current;
                    }
                }
            }
            // Re-enrich tasks with bucketInfo
            try {
                const bucketsResp = await this.apiRequest('/api/lookup-buckets', {
                    method: 'POST',
                    body: JSON.stringify({ planId: this.selectedPlannerId })
                });
                if (this.processedTasks && bucketsResp && Array.isArray(bucketsResp.tasks)) {
                    this.processedTasks.tasks = bucketsResp.tasks;
                }
            } catch (_e) {}
            this.closeBucketPreview();
            this.showTaskPreview();
        } catch (error) {
            this.showError(error.message);
        } finally {
            this.hideLoading();
        }
    }

    handleFileSelect(event) {
        const file = event.target.files[0];
        if (file) {
            document.getElementById('file-name').textContent = `${file.name} (${Math.round(file.size / 1024)} KB)`;
            document.getElementById('file-info').classList.remove('hidden');
            // Do not show preview yet; wait until a planner is selected
        }
    }

    handleConfirmationChange(event) {
        const createBtn = document.getElementById('create-tasks-btn');
        createBtn.disabled = !event.target.checked;
    }
    
    async createTasks() {
        if (!this.selectedPlannerId || !this.selectedBucketId) {
            this.showError('Please select a planner and bucket');
            return;
        }
        // Open Final Preview modal instead of immediate creation
        this.openFinalPreview();
    }

    openFinalPreview() {
        this.buildFinalPreview();
        document.getElementById('final-preview-modal').classList.remove('hidden');
    }

    closeFinalPreview() {
        document.getElementById('final-preview-modal').classList.add('hidden');
    }

    buildFinalPreview() {
        const tasks = (this.processedTasks && this.processedTasks.tasks) ? this.processedTasks.tasks : [];
        const { createMap, existingMap, defaultCount } = this.collectBucketPreviewData(tasks);
        // Apply bucket-level selection if any
        let selectedCreateMap = {};
        let unselectedCount = 0;
        Object.entries(createMap).forEach(([name, count]) => {
            if (this.bucketSelectionSet.has(name)) {
                selectedCreateMap[name] = count;
            } else {
                unselectedCount += count;
            }
        });
        // Buckets to be created
        const ul = document.getElementById('final-bucket-create-list');
        if (ul) {
            ul.innerHTML = '';
            const names = Object.keys(selectedCreateMap).sort();
            if (names.length === 0) {
                const li = document.createElement('li');
                li.textContent = 'No new buckets to create';
                ul.appendChild(li);
            } else {
                names.forEach(name => {
                    const li = document.createElement('li');
                    li.textContent = `${name}: ${selectedCreateMap[name]} task(s)`;
                    ul.appendChild(li);
                });
            }
        }
        // Existing buckets & default
        const exist = document.getElementById('final-bucket-existing-list');
        if (exist) {
            exist.innerHTML = '';
            Object.keys(existingMap).sort().forEach(name => {
                const li = document.createElement('li');
                li.textContent = `${name}: ${existingMap[name]} task(s)`;
                exist.appendChild(li);
            });
            const defLi = document.createElement('li');
            const defaultBucketName = (this.buckets.find(b => b.id === this.selectedBucketId) || {}).name || 'Default bucket';
            defLi.textContent = `${defaultBucketName}: ${defaultCount + unselectedCount} task(s)`;
            exist.appendChild(defLi);
        }
        // Compact tasks with more detail
        const tl = document.getElementById('final-task-list');
        if (tl) {
            tl.innerHTML = '';
            const preview = tasks.slice(0, 20);
            preview.forEach((t, idx) => {
                const row = document.createElement('div');
                row.className = 'flex items-center justify-between border rounded px-2 py-2';
                const bucket = t.bucketInfo ? t.bucketInfo.name : (t.bucketName ? `${t.bucketName} (not found)` : 'Default');
                const ids = Array.isArray(this.assigneeOverrides[idx]) ? this.assigneeOverrides[idx] : [];
                const names = ids.map(id => (this.plannerMembers.find(m => m.id === id) || {}).displayName).filter(Boolean).join(', ');
                const statusKey = this.normalizeStatusKey(this.statusOverrides[idx] || t.status || '');
                const statusLabel = statusKey==='complete'?'Complete':statusKey==='in_progress'?'In Progress':statusKey==='not_started'?'Not Started':'';
                row.innerHTML = `<div class="truncate"><div class="font-medium">${t.title || ''}</div><div class="text-xs text-gray-600">${t.description ? t.description : ''}</div></div><div class="ml-2 text-xs text-gray-600 text-right">${bucket}${names ? ` • ${names}` : ''}${statusLabel?` • ${statusLabel}`:''}</div>`;
                tl.appendChild(row);
            });
            if (preview.length === 0) {
                const row = document.createElement('div');
                row.className = 'text-sm text-gray-600';
                row.textContent = 'No tasks to show';
                tl.appendChild(row);
            }
        }
    }

    async createTasksConfirmed() {
        try {
            this.showLoading('Preparing to create tasks...');
            // Decide if we need to create buckets: either global toggle on (select all),
            // or there are any selected missing buckets in the inline list
            const createAllOn = !!(document.getElementById('create-missing-buckets-toggle')?.checked);
            const tasks = (this.processedTasks && this.processedTasks.tasks) ? this.processedTasks.tasks : [];
            const { createMap } = this.collectBucketPreviewData(tasks);
            let bucketNames = Object.keys(createMap);
            if (this.bucketSelectionSet instanceof Set) {
                bucketNames = bucketNames.filter(n => this.bucketSelectionSet.has(n));
            }
            let needCreate = createAllOn || bucketNames.length > 0;
            if (needCreate) {
                try {
                    // bucketNames already filtered above
                    if (bucketNames.length > 0) {
                        await this.apiRequest('/api/buckets/create', {
                            method: 'POST',
                            body: JSON.stringify({ planId: this.selectedPlannerId, bucketNames })
                        });
                        // Refresh bucket list dropdown
                        const data = await this.apiRequest(`/api/planners/${this.selectedPlannerId}/buckets`);
                        this.buckets = data.buckets;
                        const select = document.getElementById('bucket-select');
                        if (select) {
                            const current = this.selectedBucketId;
                            select.innerHTML = '<option value="">Select default bucket...</option>';
                            this.buckets.forEach(bucket => {
                                const option = document.createElement('option');
                                option.value = bucket.id;
                                option.textContent = bucket.name;
                                select.appendChild(option);
                            });
                            if (current && this.buckets.find(b => b.id === current)) {
                                select.value = current;
                            }
                        }
                        // Re-enrich tasks against new buckets
                        const bucketsResp = await this.apiRequest('/api/lookup-buckets', {
                            method: 'POST',
                            body: JSON.stringify({ planId: this.selectedPlannerId })
                        });
                        if (this.processedTasks && Array.isArray(bucketsResp.tasks)) {
                            this.processedTasks.tasks = bucketsResp.tasks;
                        }
                    }
                } catch (e) {
                    console.warn('Bucket create pre-step failed:', e?.message || e);
                }
            }

            this.showLoading('Creating tasks...');
            const data = await this.apiRequest('/api/create-tasks', {
                method: 'POST',
                body: JSON.stringify({
                    planId: this.selectedPlannerId,
                    bucketId: this.selectedBucketId,
                    overrides: this.assigneeOverrides,
                    statusOverrides: this.statusOverrides
                })
            });
            this.closeFinalPreview();
            this.showResults(data);
        } catch (error) {
            this.showError(error.message);
        } finally {
            this.hideLoading();
        }
    }
    
    showResults(data) {
        const summary = data.summary;
        
        // Update summary stats
        document.getElementById('created-count').textContent = summary.created;
        document.getElementById('assigned-count').textContent = summary.assigned;
        document.getElementById('warning-count').textContent = summary.total - summary.created - summary.failed;
        document.getElementById('failed-count').textContent = summary.failed;
        
        // Show detailed results
        const resultsDetails = document.getElementById('results-details');
        resultsDetails.innerHTML = '';
        
        data.results.forEach(result => {
            const div = document.createElement('div');
            div.className = `p-3 rounded border-l-4 ${result.success ? 'border-green-500 bg-green-50' : 'border-red-500 bg-red-50'}`;
            
            if (result.success) {
                let assigneeText = '';
                if (result.assignedUsers && result.assignedUsers.length > 0) {
                    const names = result.assignedUsers.map(u => u.displayName).join(', ');
                    assigneeText = ` → ${names}`;
                }
                
                div.innerHTML = `
                    <div class="flex items-center">
                        <i class="fas fa-check-circle text-green-600 mr-2"></i>
                        <span class="font-medium text-green-800">${result.task}</span>
                        <span class="text-green-600 text-sm">${assigneeText}</span>
                    </div>
                `;
            } else {
                div.innerHTML = `
                    <div class="flex items-center">
                        <i class="fas fa-times-circle text-red-600 mr-2"></i>
                        <span class="font-medium text-red-800">${result.task}</span>
                        <span class="text-red-600 text-sm ml-2">(${result.error})</span>
                    </div>
                `;
            }
            
            resultsDetails.appendChild(div);
        });
        
        this.showStep('results');
    }
    
    reset() {
        this.parsedData = null;
        this.processedTasks = null;
        this.planners = [];
        this.buckets = [];
        this.selectedPlannerId = null;
        this.selectedBucketId = null;
        
        // Reset form elements
        document.getElementById('file-input').value = '';
        document.getElementById('file-info').classList.add('hidden');
        document.getElementById('confirm-creation').checked = false;
        
        this.showStep('upload');
    }
}

// Initialize the application when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    new PlannerApp();
});