<script>
    // GLOBAL STATE
    let currentUser = null;
    let candidatesData = [];
    let stagesData = [];
    let departmentsData = [];
    let recruitersData = [];
    let usersData = [];
    let emailTemplatesData = [];
    let aliasData = {}; // Store aliases from backend
    let initialData = {}; // Store initial load data
    let newsData = [];
    let projectsData = [];
    let ticketsData = [];
    let recruitmentChartInstance = null;
    let sourceChartInstance = null;
    let funnelChartInstance = null;
    let rejectionChartInstance = null;
    let timeToHireChartInstance = null;

    function formatDateForInput(dStr) {
        if (!dStr) return '';
        const d = new Date(dStr);
        if (isNaN(d.getTime())) return '';
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    }

    // UTILITY FUNCTIONS
    function hashCode(str) {
        let hash = 0;
        for (let i = 0; i < str.length; i++) {
            const char = str.charCodeAt(i);
            hash = ((hash << 5) - hash) + char;
            hash = hash & hash;
        }
        return hash;
    }

    // Helper to get value using Aliases
    function getVal(c, key) {
        if (!c) return '';
        if (c[key] !== undefined) return c[key];
        const aliases = (aliasData && aliasData[key]) ? aliasData[key] : [];
        for (let a of aliases) {
            if (c[a] !== undefined) return c[a];
        }
        return '';
    }

    function showLoadingTable(selector) {
        const tbody = document.querySelector(selector);
        if (tbody) {
            tbody.innerHTML = `
                <tr>
                    <td colspan="10" class="text-center py-5">
                        <div class="spinner-border text-primary" role="status">
                            <span class="visually-hidden">Loading...</span>
                        </div>
                        <div class="mt-2 text-muted">Đang tải dữ liệu...</div>
                    </td>
                </tr>
            `;
        }
    }

    // 0. INITIALIZATION
    document.addEventListener('DOMContentLoaded', function () {
        // PERMANENT LOGIN BYPASS REMOVED
        checkSession();

        // Initialize Chatbot logic
        const chatCircle = document.getElementById('chat-circle');
        if (chatCircle) chatCircle.onclick = toggleChat;
    });

    function checkSession() {
        const storedUser = localStorage.getItem('ats_user');
        if (storedUser) {
            currentUser = JSON.parse(storedUser);
            document.getElementById('current-user-display').innerText = currentUser.username;
            document.getElementById('login-screen').style.display = 'none';
            document.getElementById('app-container').style.display = 'block';
            updateUIForRole();
            loadDashboardData();
        }
    }




    // ============================================
    // AI CHATBOT LOGIC
    // ============================================    // AI CHATBOT LOGIC
    function toggleChat() {
        const box = document.getElementById('chat-box');
        if (!box) return;

        const isHidden = box.style.display === 'none';
        box.style.display = isHidden ? 'flex' : 'none';

        if (isHidden) {
            const input = document.getElementById('chat-input');
            if (input) setTimeout(() => input.focus(), 300);
            const badge = document.getElementById('chat-badge');
            if (badge) badge.style.display = 'none';

            // Add initial animation if not already present
            box.classList.add('animate__fadeInUp');
        }
    }

    function sendMessage() {
        const input = document.getElementById('chat-input');
        const content = document.getElementById('chat-content');
        if (!input || !content || !input.value.trim()) return;

        const msg = input.value.trim();
        input.value = '';

        // Append User Message
        content.innerHTML += `
            <div class="d-flex flex-column align-items-end mb-4">
                <div class="msg-user shadow-sm p-3 rounded-4 bg-primary text-white" style="border-bottom-right-radius: 5px !important;">
                    ${msg}
                </div>
                <small class="text-muted mt-1 me-2" style="font-size: 10px;">Bạn • Vừa xong</small>
            </div>
        `;
        content.scrollTop = content.scrollHeight;

        // Loading bubble
        content.innerHTML += `
            <div id="ai-typing" class="d-flex flex-column align-items-start mb-4">
                <div class="msg-ai shadow-sm p-3 rounded-4 bg-white" style="border-bottom-left-radius: 5px !important;">
                    <span class="spinner-grow spinner-grow-sm text-primary" role="status"></span>
                    <span class="ms-1">AI đang suy nghĩ...</span>
                </div>
            </div>
        `;
        content.scrollTop = content.scrollHeight;

        google.script.run
            .withSuccessHandler(function (response) {
                const typing = document.getElementById('ai-typing');
                if (typing) typing.remove();

                content.innerHTML += `
                    <div class="d-flex flex-column align-items-start mb-4">
                        <div class="msg-ai shadow-sm p-3 rounded-4 bg-white" style="border-bottom-left-radius: 5px !important;">
                            ${response.replace(/\n/g, '<br>')}
                        </div>
                        <small class="text-muted mt-1 ms-2" style="font-size: 10px;">AI • Vừa xong</small>
                    </div>
                `;
                content.scrollTop = content.scrollHeight;
            })
            .withFailureHandler(function (err) {
                const typing = document.getElementById('ai-typing');
                if (typing) typing.remove();
                Swal.fire('Lỗi Chat', err.toString(), 'error');
            })
            .apiChatWithGemini(msg);
    }

    // ============================================
    // TICKET & JOB LINKING LOGIC
    // ============================================
    function viewLinkedJD() {
        let ticketId = null;
        const tickIdInput = document.getElementById('tick-id');
        if (tickIdInput) {
            ticketId = tickIdInput.value;
        } else {
            const modal = document.getElementById('viewTicketModal');
            if (modal) {
                const idHeader = modal.querySelector('.modal-title')?.innerText?.match(/Ticket ID: ([^\s]+)/);
                if (idHeader) ticketId = idHeader[1];
            }
        }

        if (!ticketId) {
            Swal.fire('Lỗi', 'Không xác định được mã Ticket hiện tại.', 'error');
            return;
        }

        google.script.run.withSuccessHandler(function (jobs) {
            const linkedJob = jobs.find(j => String(j.ticketId || j.TicketID) === String(ticketId));
            if (linkedJob) {
                Swal.fire({
                    title: 'Mô tả công việc (JD)',
                    html: `<div class="text-start p-3 bg-light rounded" style="max-height: 400px; overflow-y: auto; white-space: pre-wrap;">${linkedJob.description || linkedJob['Mô tả']}</div>`,
                    width: '700px',
                    confirmButtonText: 'Đã hiểu'
                });
            } else {
                Swal.fire('Chú ý', 'Ticket này chưa được liên kết với bản Tin tuyển dụng (JD) nào trong danh sách Jobs.', 'info');
            }
        }).apiGetTableData('Jobs');
    }

    function populateJobTicketDropdown() {
        const ticketSelect = document.getElementById('job-ticket-id');
        if (!ticketSelect) return;

        // Use ticketsData from global state instead of separate API call for speed
        if (ticketsData && ticketsData.length > 0) {
            ticketSelect.innerHTML = '<option value="">-- Chọn Ticket đã tạo --</option>';
            ticketsData.forEach(t => {
                const id = t.TicketID || t['Mã Ticket'] || t.ID || 'N/A';
                const title = t.Position || t['Vị trí'] || t.Title || '';
                ticketSelect.innerHTML += `<option value="${id}">${id} - ${title}</option>`;
            });
        }
    }

    function populateJobDepartmentDropdown() {
        const deptSelect = document.getElementById('job-dept');
        if (!deptSelect) return;
        deptSelect.innerHTML = '<option value="">-- Chọn Phòng ban --</option>';
        if (departmentsData && departmentsData.length > 0) {
            departmentsData.forEach(d => {
                const name = d.name || d['Phòng ban'] || d['Department'];
                if (name) deptSelect.innerHTML += `<option value="${name}">${name}</option>`;
            });
        }
    }

    function onJobDepartmentChange() {
        const deptName = document.getElementById('job-dept').value;
        const posSelect = document.getElementById('job-position');
        if (!posSelect) return;
        posSelect.innerHTML = '<option value="">-- Chọn Vị trí --</option>';
        if (!deptName) return;

        const deptRecord = departmentsData.find(d => (d.name || d['Phòng ban'] || d['Department']) === deptName);
        if (deptRecord && deptRecord.positions) {
            deptRecord.positions.forEach(pos => {
                posSelect.innerHTML += `<option value="${pos}">${pos}</option>`;
            });
        }
    }

    function populateJobLocationDropdown() {
        const locSelect = document.getElementById('job-location');
        if (!locSelect) return;
        locSelect.innerHTML = '<option value="">-- Chọn Địa điểm --</option>';

        // Try office locations from companyInfo addresses list
        const offices = initialData.companyInfo?.addresses || initialData.companyInfo?.['addresses'] || [];
        if (Array.isArray(offices) && offices.length > 0) {
            offices.forEach(off => {
                locSelect.innerHTML += `<option value="${off}">${off}</option>`;
            });
        } else {
            // Fallback to the old format or hardcoded
            const legacyOffices = initialData.companyInfo?.offices || [];
            if (Array.isArray(legacyOffices) && legacyOffices.length > 0) {
                legacyOffices.forEach(off => locSelect.innerHTML += `<option value="${off}">${off}</option>`);
            } else {
                ['Hà Nội', 'Hồ Chí Minh', 'Đà Nẵng', 'Remote'].forEach(off => {
                    locSelect.innerHTML += `<option value="${off}">${off}</option>`;
                });
            }
        }
    }

    document.addEventListener('shown.bs.modal', function (event) {
        if (event.target.id === 'addJobModal') {
            populateJobTicketDropdown();
            populateJobDepartmentDropdown();
            populateJobLocationDropdown();
        }
    });

    // 1. SIMPLE NAVIGATION
    function showSection(sectionId, element) {
        // Consolidated RBAC Check
        const adminOnlySections = ['settings', 'jobs'];
        if (adminOnlySections.includes(sectionId) && (!currentUser || currentUser.role !== 'Admin')) {
            Swal.fire('Truy cập bị từ chối', 'Bạn không có quyền truy cập mục này.', 'error');
            return;
        }

        const evalRoles = ['Admin', 'Manager', 'Recruiter'];
        if (sectionId === 'evaluations' && (!currentUser || !evalRoles.includes(currentUser.role))) {
            Swal.fire('Truy cập bị từ chối', 'Bạn không có quyền truy cập mục này.', 'error');
            return;
        }

        document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
        const section = document.getElementById(sectionId);
        if (section) section.classList.add('active');

        if (element) {
            document.querySelectorAll('.sidebar .nav-link').forEach(l => l.classList.remove('active'));
            element.classList.add('active');
            const titleMap = {
                'dashboard': 'Tổng Quan Hoạt Động',
                'kanban': 'Bảng Theo Dõi Ứng Viên (Kanban)',
                'candidates': 'Danh Sách Ứng Viên',
                'jobs': 'Quản Lý Vị Trí Tuyển Dụng',
                'settings': 'Cài Đặt Hệ Thống',
                'evaluations': 'Quản lý Đánh giá Phỏng vấn',
                'recruitment-hub': 'Quản lý Tuyển dụng (Dự án & Tickets)'
            };
            document.getElementById('page-title').innerText = titleMap[sectionId] || 'Trang Quản Trị';

            if (sectionId === 'jobs') loadJobs();
            if (sectionId === 'settings') loadSettings();
            if (sectionId === 'candidates') renderCandidatesTable();
            if (sectionId === 'kanban') renderKanbanBoard();
            if (sectionId === 'evaluations') loadEvaluations();
            if (sectionId === 'recruitment-hub') {
                loadProjects();
                loadTickets();
            }
            if (sectionId === 'dashboard') {
                if (typeof updateDashboardStats === 'function') updateDashboardStats();
                loadActivityLogs();
            }
        }
    }

    // 2. LOGIN LOGIC
    function handleLogin() {
        let u = document.getElementById('login-username').value;
        let p = document.getElementById('login-password').value;
        if (u) u = u.trim();
        if (p) p = p.trim();
        const msg = document.getElementById('login-message');
        const loader = document.getElementById('login-loader');

        if (!u || !p) {
            msg.innerText = 'Vui lòng nhập đầy đủ thông tin.';
            return;
        }
        msg.innerText = '';
        loader.style.display = 'inline-block';

        google.script.run.withSuccessHandler(function (response) {
            loader.style.display = 'none';
            if (response.success) {
                currentUser = response.user;
                localStorage.setItem('ats_user', JSON.stringify(currentUser));
                document.getElementById('current-user-display').innerText = currentUser.username;
                document.getElementById('login-screen').style.display = 'none';
                document.getElementById('app-container').style.display = 'block';
                updateUIForRole();
                loadDashboardData();
            } else {
                msg.innerText = response.message;
            }
        }).withFailureHandler(function (err) {
            loader.style.display = 'none';
            msg.innerText = 'Lỗi kết nối: ' + err.message;
        }).apiLogin(u, p);
    }

    function logout() {
        localStorage.removeItem('ats_user');
        currentUser = null;

        // Hide App, Show Login
        document.getElementById('app-container').style.display = 'none';
        document.getElementById('login-screen').style.display = 'block';

        // Clear forms
        document.getElementById('login-username').value = '';
        document.getElementById('login-password').value = '';

        // Optional: Reload to clear memory states
        // location.reload(); 
    }

    // LISTENER FOR ADD CANDIDATE MODAL TO POPULATE RECRUITERS
    document.addEventListener('DOMContentLoaded', function () {

        // Handle dynamically changing status dropdown based on selected Ticket
        const detailTicketId = document.getElementById('detail-ticket-id');
        if (detailTicketId) {
            detailTicketId.addEventListener('change', function () {
                const currentStatus = document.getElementById('detail-status')?.value;
                populateStatusDropdown(currentStatus);
            });
        }
    });

    // GENERIC RECRUITER POPULATOR
    function populateRecruiterSelect(elementId, selectedValue = null) {
        const select = document.getElementById(elementId);
        if (!select) return;

        select.innerHTML = '<option value="">Chọn nhân viên</option>';

        // Combine recruitersData and usersData (if they are recruiters)
        // Or just use recruitersData if it's the source of truth
        let list = recruitersData;
        if (!list || list.length === 0) list = usersData; // Fallback

        if (list && list.length > 0) {
            list.forEach(r => {
                const opt = document.createElement('option');
                const val = r.name || r.Full_Name || r.Username;
                opt.value = val;
                opt.innerText = val;
                if (val === selectedValue) opt.selected = true;
                select.appendChild(opt);
            });
        }
    }

    // 3. LOAD DATA
    function loadDashboardData() {
        console.log('🔄 loadDashboardData() called');
        google.script.run.withSuccessHandler(function (data) {
            console.log('✅ SUCCESS HANDLER CALLED');

            if (!data) {
                console.warn('⚠️ Backend returned null! Using safe defaults...');
                data = {
                    candidates: [], stages: [], departments: [], recruiters: [], users: [],
                    emailTemplates: [], projects: [], tickets: [], aliases: {}
                };
            }

            candidatesData = data.candidates || [];
            stagesData = data.stages || [];
            departmentsData = data.departments || [];
            recruitersData = data.recruiters || [];
            usersData = data.users || [];
            console.log('✅ Users Data Loaded: ' + usersData.length);
            emailTemplatesData = data.emailTemplates || [];
            aliasData = data.aliases || {};
            projectsData = data.projects || [];
            ticketsData = data.tickets || [];
            window.rejectionReasonsData = data.rejectionReasons || [];
            initialData = data || {};

            updateDashboardStats();
            populateFilterDropdowns();
            renderKanbanBoard();
            renderCandidatesTable();
            if (typeof renderRecruiters === 'function') renderRecruiters();

            // Check for evaluations (if Manager)
            checkForPendingEvaluations();

            // Load Notifications
            if (typeof loadNotifications === 'function') loadNotifications();
        })
            .withFailureHandler(function (error) {
                console.error('❌ FAILURE:', error);
                alert('Lỗi: ' + error.message);
            })
            .apiGetInitialData(currentUser ? currentUser.username : '');
    }

    // RBAC: Update UI based on Role
    function updateUIForRole() {
        if (!currentUser) return;
        const role = currentUser.role || 'Viewer';

        console.log('Applying RBAC for Role:', role);

        // 1. Settings Access
        const navSettings = document.getElementById('nav-settings');
        if (role !== 'Admin') {
            if (navSettings) navSettings.style.display = 'none';
        } else {
            if (navSettings) navSettings.style.display = 'block';
        }

        // 2. Jobs Access (Admin Only)
        const navJobs = document.getElementById('nav-jobs');
        if (role !== 'Admin') {
            if (navJobs) navJobs.style.display = 'none';
        } else {
            if (navJobs) navJobs.style.display = 'block';
        }

        // 3. Evaluations Access (Admin, Manager, Recruiter)
        const navEvals = document.getElementById('nav-evaluations');
        if (navEvals) {
            if (role === 'Admin' || role === 'Manager' || role === 'Recruiter') {
                navEvals.style.display = 'block';
            } else {
                navEvals.style.display = 'none';
            }
        }

        // 2. Viewer Restrictions
        if (role === 'Viewer') {
            // Hide "Add Candidate" buttons
            document.querySelectorAll('.btn-add-candidate').forEach(el => el.style.display = 'none');

            // Disable specific interactions if needed (handled in renderKanbanBoard / renderCandidatesTable)
            // We'll add a global class to body to help CSS/JS
            document.body.classList.add('role-viewer');
        } else {
            document.querySelectorAll('.btn-add-candidate').forEach(el => el.style.display = 'inline-block');
            document.body.classList.remove('role-viewer');
        }

        // 3. Manager/User Logic
        // Manager uses same UI as User but data is filtered by Backend.
        // User cannot access Settings (handled above).

        // 4. Update Profile Display
        const disp = document.getElementById('current-user-display');
        if (disp) {
            disp.innerText = `${currentUser.name} (${role})`;
        }
    }

    function updateDashboardStats() {
        document.getElementById('stat-total-candidates').innerText = candidatesData.length;

        // Calculate stats based on Status/Stage
        const hiredCount = candidatesData.filter(c => c.Status === 'Offer' || c.Status === 'Hired').length;
        const interviewCount = candidatesData.filter(c => c.Status === 'Interview' || c.Status === 'Phỏng vấn').length;
        const rejectedCount = candidatesData.filter(c => c.Status === 'Rejected' || c.Status === 'Loại').length;

        document.getElementById('stat-hired').innerText = hiredCount;
        document.getElementById('stat-interviewing').innerText = interviewCount;
        document.getElementById('stat-rejected').innerText = rejectedCount;

        // Update Chart if exists
        updateCharts();
    }

    function updateCharts() {
        // Group by Month (Created Date)
        const monthCounts = Array(12).fill(0);
        candidatesData.forEach(c => {
            if (c.Applied_Date) {
                const d = new Date(c.Applied_Date);
                if (!isNaN(d)) monthCounts[d.getMonth()]++;
            }
        });

        // We need to access the chart instance. 
        // Since we created it in DOMContentLoaded, let's attach it to window or re-create.
        // For simplicity, let's just trigger a re-render if the canvas exists and we have a global ref, 
        // or easier: just destroy and recreate if we can access the context.
        // Actually, let's look at the init code.
    }

    // 4. RENDER KANBAN
    function renderKanbanBoard() {
        const container = document.getElementById('kanban-container');
        if (!container) return;

        // Populate Project Filter if empty
        const projFilter = document.getElementById('kanban-filter-project');
        if (projFilter && projFilter.options.length <= 1 && projectsData.length > 0) {
            projectsData.forEach(p => {
                const opt = document.createElement('option');
                opt.value = p['Mã Dự án'];
                opt.innerText = p['Tên Dự án'];
                projFilter.appendChild(opt);
            });
        }

        // Show loading indicator
        const loadingIndicator = document.getElementById('kanban-loading');
        if (loadingIndicator) loadingIndicator.style.display = 'block';

        // Clear container (keeping loading indicator)
        Array.from(container.children).forEach(child => {
            if (child.id !== 'kanban-loading') child.remove();
        });

        // Determine which stages to use
        const projectCode = projFilter ? projFilter.value : '';
        const selectedProject = projectsData.find(p => p['Mã Dự án'] === projectCode);

        let dynamicStages = [];
        if (selectedProject && selectedProject['Quy trình (Workflow)']) {
            const workflowParts = selectedProject['Quy trình (Workflow)'].split(',').map(s => s.trim());
            dynamicStages = workflowParts.map((name, index) => ({
                Stage_Name: name,
                Order: index + 1,
                Color: '#FFC107'
            }));
        } else {
            dynamicStages = (stagesData && stagesData.length > 0) ? [...stagesData] : [
                { Stage_Name: 'Apply', Order: 1, Color: '#0d6efd' },
                { Stage_Name: 'Interview', Order: 2, Color: '#fd7e14' },
                { Stage_Name: 'Offer', Order: 3, Color: '#198754' },
                { Stage_Name: 'Rejected', Order: 4, Color: '#dc3545' }
            ];
        }
        dynamicStages.sort((a, b) => a.Order - b.Order);

        // DEDUPLICATE and FILTER candidates
        const seen = new Set();
        let uniqueCandidates = candidatesData.filter(c => {
            if (seen.has(c.ID)) return false;
            seen.add(c.ID);
            return true;
        });

        // Filter by Project if selected
        if (projectCode) {
            uniqueCandidates = uniqueCandidates.filter(c => {
                const tID = getVal(c, 'TicketID');
                if (!tID) return false;
                const t = ticketsData.find(x => x['Mã Ticket'] == tID);
                return t && t['Mã Dự án'] === projectCode;
            });
        }

        // Search Filter
        const searchQ = document.getElementById('kanban-search')?.value.toLowerCase() || '';
        if (searchQ) {
            uniqueCandidates = uniqueCandidates.filter(c =>
                (c.Name || '').toLowerCase().includes(searchQ) ||
                (c.ID || '').toString().toLowerCase().includes(searchQ)
            );
        }

        const fragment = document.createDocumentFragment();

        dynamicStages.forEach(stage => {
            const col = document.createElement('div');
            col.className = 'kanban-column';
            col.innerHTML = `
         <div class="kanban-header" style="border-bottom-color: ${stage.Color || '#ccc'}">
             <span>${stage.Stage_Name}</span>
             <span class="badge bg-secondary rounded-pill count-badge">0</span>
         </div>
         <div class="kanban-items" data-stage="${stage.Stage_Name}">
           <!-- Items go here -->
         </div>
       `;

            // Append to fragment instead of container (FASTER!)
            fragment.appendChild(col);

            // Filter candidates for this stage
            // Try both Status and Stage fields, case-insensitive
            const candidatesInStage = uniqueCandidates.filter(c => {
                const candidateStatus = (c.Status || '').toString().trim();
                const candidateStage = (c.Stage || '').toString().trim();
                const stageName = stage.Stage_Name.toString().trim();

                return candidateStatus.toLowerCase() === stageName.toLowerCase() ||
                    candidateStage.toLowerCase() === stageName.toLowerCase();
            });

            const itemsContainer = col.querySelector('.kanban-items');
            col.querySelector('.count-badge').innerText = candidatesInStage.length;

            // Debug logging
            if (candidatesInStage.length > 0) {
                console.log(`Stage "${stage.Stage_Name}" has ${candidatesInStage.length} candidates:`, candidatesInStage);
            }

            candidatesInStage.forEach(c => {
                const card = document.createElement('div');
                card.className = 'kanban-card';
                card.setAttribute('data-id', c.ID);
                card.style.cursor = 'pointer';

                // Generate avatar initials
                const initials = (c.Name || 'U').split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase();
                const avatarColors = ['#007bff', '#28a745', '#dc3545', '#ffc107', '#17a2b8', '#6f42c1', '#e83e8c'];
                const avatarColor = avatarColors[Math.abs(hashCode(c.ID || '')) % avatarColors.length];

                // Determine status badge
                let statusBadge = '';
                if (c.Status) {
                    const badgeColors = {
                        'Mới tiếp nhận': 'primary',
                        'Đã liên lạc': 'success',
                        'Chờ phản hồi': 'warning',
                        'Đang trao đổi': 'info'
                    };
                    const badgeClass = badgeColors[c.Status] || 'secondary';
                    statusBadge = `<span class="badge bg-${badgeClass} ms-2" style="font-size: 0.7rem;">${c.Status}</span>`;
                }

                card.innerHTML = `
                    <div class="d-flex align-items-start mb-2">
                        <div class="candidate-avatar me-2" style="background: ${avatarColor}; width: 40px; height: 40px; border-radius: 50%; display: flex; align-items: center; justify-content: center; color: white; font-weight: bold; font-size: 14px;">
                            ${initials}
                        </div>
                        <div class="flex-grow-1">
                            <h6 class="mb-0 fw-bold" style="font-size: 0.95rem;">${c.Name || 'N/A'}</h6>
                            ${statusBadge}
                        </div>
                    </div>
                    <p class="mb-1 small text-muted"><i class="fas fa-briefcase me-1"></i>${c.Position || 'N/A'}</p>
                    <p class="mb-1 small text-muted"><i class="fas fa-envelope me-1"></i>${c.Email || 'N/A'}</p>
                    <p class="mb-1 small text-muted"><i class="fas fa-phone me-1"></i>${c.Phone || 'N/A'}</p>
                    ${c.Recruiter ? `<p class="mb-2 small text-muted"><i class="fas fa-user-tie me-1"></i>${c.Recruiter}</p>` : ''}
                    <div class="d-flex justify-content-between align-items-center mt-2 pt-2" style="border-top: 1px solid #eee;">
                        <small class="text-muted">${c.Applied_Date ? new Date(c.Applied_Date).toLocaleDateString('vi-VN') : ''}</small>
                        <div class="btn-group btn-group-sm" role="group">
                            <!-- Helper for Viewer Check -->
                            ${(!currentUser || currentUser.role !== 'Viewer') ? `
                            <button class="btn btn-outline-secondary btn-sm p-1" title="Gửi email" onclick="event.stopPropagation(); openSendEmailModal('${c.ID}');">
                                <i class="fas fa-envelope" style="font-size: 0.75rem;"></i>
                            </button>
                            ` : ''}
                            
                            <!-- VIEW CV BUTTON -->
                            ${c.CV_Link ? `
                            <button class="btn btn-outline-info btn-sm p-1" title="Xem CV" onclick="event.stopPropagation(); viewCandidateCV('${c.CV_Link}');">
                                <i class="fas fa-file-alt" style="font-size: 0.75rem;"></i>
                            </button>
                            ` : ''}

                            <!-- EVALUATION BUTTON (Only for Interview Stage) -->
                            ${(stage.Stage_Name.toLowerCase().includes('phỏng vấn') || stage.Stage_Name.toLowerCase().includes('interview')) ? `
                            <button class="btn btn-outline-success btn-sm p-1" title="Tạo Đánh giá" onclick="event.stopPropagation(); openCreateEvaluationModal('${c.ID}');">
                                <i class="fas fa-clipboard-check" style="font-size: 0.75rem;"></i>
                            </button>
                            ` : ''}

                            <button class="btn btn-outline-primary btn-sm p-1" title="Xem chi tiết" onclick="event.stopPropagation(); openCandidateDetail('${c.ID}');">
                                <i class="fas fa-eye" style="font-size: 0.75rem;"></i>
                            </button>
                            ${(!currentUser || currentUser.role !== 'Viewer') ? `
                            <button class="btn btn-outline-warning btn-sm p-1" title="Sửa" onclick="event.stopPropagation(); editCandidate('${c.ID}');">
                                <i class="fas fa-edit" style="font-size: 0.75rem;"></i>
                            </button>
                            <button class="btn btn-outline-danger btn-sm p-1" title="Xóa" onclick="event.stopPropagation(); deleteCandidate('${c.ID}');">
                                <i class="fas fa-trash" style="font-size: 0.75rem;"></i>
                            </button>
                            ` : ''}
                        </div>
                    </div>
                        </div>
                    </div>
                `;

                // Prevent click when dragging
                let isDragging = false;
                card.addEventListener('mousedown', () => { isDragging = false; });
                card.addEventListener('mousemove', () => { isDragging = true; });

                // Click on card to open detail (only if not dragging AND not clicking on buttons)
                card.addEventListener('mouseup', function (e) {
                    // Check if click is on a button or inside button group
                    const clickedOnButton = e.target.closest('button') || e.target.closest('.btn-group');

                    if (!isDragging && !clickedOnButton) {
                        setTimeout(() => openCandidateDetail(c.ID), 100);
                    }
                    isDragging = false;
                });

                itemsContainer.appendChild(card);
            });

            // Init Sortable for this column
            if (typeof Sortable === 'undefined') {
                console.error('❌ Sortable library not loaded!');
            } else {
                if (!currentUser || currentUser.role !== 'Viewer') {
                    console.log(`✅ Initializing Sortable for stage: ${stage.Stage_Name}`);
                    new Sortable(itemsContainer, {
                        group: 'kanban-shared',
                        animation: 200,
                        ghostClass: 'sortable-ghost',
                        chosenClass: 'sortable-chosen',
                        dragClass: 'sortable-drag',
                        forceFallback: true,
                        fallbackClass: 'sortable-fallback',
                        fallbackOnBody: true,
                        swapThreshold: 0.65,
                        onStart: function (evt) {
                            console.log('🚀 Drag started');
                            evt.item.style.opacity = '0.5';
                        },
                        onEnd: function (evt) {
                            evt.item.style.opacity = '1';

                            const itemEl = evt.item;
                            const newStage = evt.to.getAttribute('data-stage');
                            const candidateId = itemEl.getAttribute('data-id');
                            const oldStage = evt.from.getAttribute('data-stage');

                            console.log('🔄 Drag ended - From:', oldStage, 'To:', newStage, 'ID:', candidateId);

                            // Only update if stage actually changed
                            if (newStage && candidateId && newStage !== oldStage) {
                                if (newStage.toLowerCase().includes('loại') || newStage.toLowerCase().includes('reject')) {
                                    promptRejectionReason(function (rejectionData) {
                                        executeStatusUpdate(candidateId, newStage, rejectionData);
                                    }, function () {
                                        // On cancel, revert
                                        loadDashboardData();
                                    });
                                } else {
                                    executeStatusUpdate(candidateId, newStage, null);
                                }
                            } else {
                                console.log('⏭️ Stage unchanged, no update needed');
                            }
                        }
                    });
                }
            }
        });

        function executeStatusUpdate(candidateId, newStage, rejectionData) {
            console.log('💾 Saving change to backend...');
            google.script.run.withSuccessHandler(function (res) {
                if (!res.success) {
                    Swal.fire('Lỗi', res.message, 'error');
                    loadDashboardData();
                } else {
                    const c = candidatesData.find(x => x.ID == candidateId);
                    if (c) {
                        c.Stage = newStage;
                        updateDashboardStats();
                    }

                    Swal.fire({
                        icon: 'success',
                        title: 'Đã cập nhật!',
                        text: `Đã chuyển sang ${newStage}`,
                        timer: 1500,
                        showConfirmButton: false
                    }).then(() => {
                        const s = newStage.toLowerCase();
                        if (s.includes('phỏng vấn') || s.includes('pv') || s.includes('interview') ||
                            s.includes('offer') || s.includes('nhận việc') ||
                            s.includes('loại') || s.includes('reject')) {

                            Swal.fire({
                                title: 'Gửi Email?',
                                text: `Bạn có muốn gửi email cho ứng viên không?`,
                                icon: 'question',
                                showCancelButton: true,
                                confirmButtonText: 'Soạn Email',
                                cancelButtonText: 'Không'
                            }).then((emailResult) => {
                                if (emailResult.isConfirmed) {
                                    openSendEmailModal(candidateId);
                                }
                            });
                        }
                    });
                }
            }).withFailureHandler(function (error) {
                console.error('Error updating status:', error);
                Swal.fire('Lỗi', 'Không thể cập nhật: ' + error.message, 'error');
                loadDashboardData();
            }).apiUpdateCandidateStatus(candidateId, newStage, rejectionData);
        }

        // Append entire fragment to DOM at once (SINGLE REFLOW - MUCH FASTER!)
        container.appendChild(fragment);

        // Hide loading indicator
        if (loadingIndicator) loadingIndicator.style.display = 'none';
    }

    function renderCandidatesTable() {
        const tbody = document.querySelector('#candidates-table tbody');
        if (!tbody) return;
        tbody.innerHTML = '';

        // Use document fragment for batch DOM update
        const fragment = document.createDocumentFragment();

        candidatesData.forEach(c => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
           <td>${c.ID}</td>
           <td class="fw-bold cursor-pointer text-primary" onclick="openCandidateDetail('${c.ID}')">${c.Name}</td>
           <td>${c.Position}</td>
           <td>${c.Applied_Date ? new Date(c.Applied_Date).toLocaleDateString() : ''}</td>
           <td><span class="badge bg-info text-dark">${c.Status}</span></td>
           <td>
             <div class="btn-group btn-group-sm">
                ${(!currentUser || currentUser.role !== 'Viewer') ? `
                <button class="btn btn-outline-secondary btn-sm" title="Gửi email" onclick="openSendEmailModal('${c.ID}')">
                    <i class="fas fa-envelope"></i>
                </button>
                ` : ''}

                ${c.CV_Link ? `
                <button class="btn btn-outline-info btn-sm" title="Xem CV" onclick="viewCandidateCV('${c.CV_Link}')">
                    <i class="fas fa-file-alt"></i>
                </button>
                ` : ''}

                 <button class="btn btn-outline-info btn-sm" onclick="openCandidateDetail('${c.ID}', 'view')" title="Xem chi tiết"><i class="fas fa-eye"></i></button>

                 ${(!currentUser || currentUser.role !== 'Viewer') ? `
                 <button class="btn btn-outline-primary btn-sm" onclick="openCandidateDetail('${c.ID}', 'edit')" title="Chỉnh sửa"><i class="fas fa-edit"></i></button>
                 <button class="btn btn-outline-danger btn-sm" title="Xóa" onclick="deleteCandidate('${c.ID}')">
                     <i class="fas fa-trash"></i>
                 </button>
                 ` : ''}
              </div>
            </td>
        `;
            fragment.appendChild(tr);
        });

        // Append all at once
        tbody.appendChild(fragment);
    }

    function viewCandidateCV(link) {
        if (!link) {
            Swal.fire('Lỗi', 'Ứng viên này chưa có link CV', 'warning');
            return;
        }

        const modal = new bootstrap.Modal(document.getElementById('cvPreviewModal'));
        const iframe = document.getElementById('cv-iframe');
        const downloadBtn = document.getElementById('cv-download-btn');

        // Set download/open new tab link
        downloadBtn.href = link;

        // Process Link for Embedding
        let embedLink = link;

        // Handle Google Drive Links
        if (link.includes('drive.google.com')) {
            // Replace /view or /edit with /preview
            if (link.includes('/view')) {
                embedLink = link.replace('/view', '/preview');
            } else if (link.includes('/edit')) {
                embedLink = link.replace('/edit', '/preview');
            } else if (!link.includes('/preview')) {
                // Try appending /preview if it ends with nothing or just ID
                // Simple heuristic: if it doesn't have an action, append /preview
                if (!link.endsWith('/')) embedLink += '/preview';
                else embedLink += 'preview';
            }
        }
        // Handle Dropbox, etc. if needed (future)

        iframe.src = embedLink;

        modal.show();

        // Clear src on close to stop audio/video if any
        document.getElementById('cvPreviewModal').addEventListener('hidden.bs.modal', function () {
            iframe.src = '';
        });
    }

    // 9. CANDIDATE DETAILS
    // -------------------------------------------------------------------------
    // 4. CANDIDATE DETAIL MODAL LOGIC (OPEN/SAVE)
    // -------------------------------------------------------------------------

    // OPEN CANDIDATE DETAIL MODAL
    function openCandidateDetail(candidateId = null, mode = 'edit') {
        const modal = document.getElementById('candidateDetailModal');
        const form = document.getElementById('edit-candidate-form');
        const modalTitle = modal.querySelector('.modal-title');
        const saveBtn = modal.querySelector('.modal-footer .btn-primary');
        const notesHistoryDiv = document.getElementById('detail-notes-history');

        // Reset Tabs - always start with the first tab
        const firstTabEl = document.querySelector('#candidateTabs .nav-link:first-child');
        if (firstTabEl) {
            const firstTab = bootstrap.Tab.getOrCreateInstance(firstTabEl);
            firstTab.show();
        }

        modal.setAttribute('data-mode', mode);
        window._currentCandidateStage = '';

        if (!candidateId) {
            // NEW MODE
            if (form) form.reset();
            const idInput = document.getElementById('current-candidate-id');
            if (idInput) idInput.value = '';
            modalTitle.innerHTML = '<i class="fas fa-user-plus me-2"></i>Thêm ứng viên mới';

            // RESET AI MATCH UI
            resetAiMatchUI();

            if (saveBtn) {
                saveBtn.innerHTML = '<i class="fas fa-plus me-2"></i>Tạo Hồ sơ';
                saveBtn.style.display = 'block';
                saveBtn.onclick = saveCandidateDetail;
            }

            if (notesHistoryDiv) notesHistoryDiv.innerHTML = '<div class="text-center text-muted py-3">Chưa có ghi chú.</div>';

            // Populate Dropdowns with defaults
            if (typeof populateDepartmentDropdown === 'function') {
                populateDepartmentDropdown('', () => {
                    if (typeof populatePositionDropdown === 'function') populatePositionDropdown('');
                });
            }
            if (typeof populateTicketDropdown === 'function') populateTicketDropdown('');
            if (typeof populateRecruiterSelect === 'function') populateRecruiterSelect('detail-recruiter', '');
            if (typeof populateStatusDropdown === 'function') populateStatusDropdown('Apply');

            const contactStatus = document.getElementById('detail-contact-status');
            if (contactStatus) contactStatus.value = 'Mới tiếp nhận';

            new bootstrap.Modal(modal).show();
            return;
        }

        // EDIT/VIEW MODE
        const c = candidatesData.find(x => x.ID == candidateId);
        if (!c) {
            Swal.fire('Lỗi', 'Không tìm thấy thông tin ứng viên', 'error');
            return;
        }

        document.getElementById('current-candidate-id').value = c.ID;
        modal.setAttribute('data-candidate-id', c.ID);
        modalTitle.innerHTML = `<i class="fas fa-user-edit me-2"></i>Chi tiết: ${c.Name || 'Ứng viên'}`;

        // RESET AI MATCH UI
        resetAiMatchUI();
        window._currentCandidateStage = c.Stage || '';

        // Map data to fields
        document.getElementById('detail-name').value = c.Name || '';
        document.getElementById('detail-gender').value = c.Gender || '';
        document.getElementById('detail-dob').value = c.Birth_Year || '';
        document.getElementById('detail-phone').value = c.Phone || '';
        document.getElementById('detail-email').value = c.Email || '';
        document.getElementById('detail-department').value = c.Department || '';
        document.getElementById('detail-position').value = c.Position || '';
        document.getElementById('detail-experience').value = c.Experience || '';
        document.getElementById('detail-school').value = c.School || '';
        document.getElementById('detail-education-level').value = c.Education_Level || '';
        document.getElementById('detail-major').value = c.Major || '';
        document.getElementById('detail-salary').value = c.Salary_Expectation || '';
        document.getElementById('detail-source').value = c.Source || '';
        document.getElementById('detail-recruiter').value = c.Recruiter || '';
        document.getElementById('detail-status').value = c.Stage || '';
        document.getElementById('detail-contact-status').value = c.Status || '';
        document.getElementById('detail-cv-link').value = c.CV_Link || '';

        // Rejection Data
        if (document.getElementById('detail-rejection-source')) document.getElementById('detail-rejection-source').value = c.Rejection_Source || '';
        if (document.getElementById('detail-rejection-type')) document.getElementById('detail-rejection-type').value = c.Rejection_Type || '';
        if (document.getElementById('detail-rejection-reason')) document.getElementById('detail-rejection-reason').value = c.Rejection_Reason || '';

        // Populate dynamic dropdowns
        if (typeof populateTicketDropdown === 'function') populateTicketDropdown(c.TicketID);

        // Handle Visibility of Rejection tracking
        const rejectionSection = document.getElementById('rejection-detail-section');
        const isRejected = (c.Stage || '').toLowerCase().includes('loại') || (c.Stage || '').toLowerCase().includes('reject') || (c.Stage || '').toLowerCase().includes('từ chối');
        if (rejectionSection) rejectionSection.style.display = isRejected ? 'block' : 'none';

        // Note history
        renderNoteHistory(c.Notes, notesHistoryDiv);
        document.getElementById('detail-new-note').value = '';

        // Configure Save/Close Button
        if (saveBtn) {
            if (mode === 'view') {
                saveBtn.innerHTML = '<i class="fas fa-times me-2"></i>Đóng';
                saveBtn.classList.remove('btn-primary');
                saveBtn.classList.add('btn-secondary');
                saveBtn.onclick = () => bootstrap.Modal.getInstance(modal).hide();
            } else {
                saveBtn.innerHTML = '<i class="fas fa-save me-2"></i>Cập nhật Hồ sơ';
                saveBtn.classList.add('btn-primary');
                saveBtn.classList.remove('btn-secondary');
                saveBtn.style.display = 'block';
                saveBtn.onclick = saveCandidateDetail;
            }
        }

        // Configure Related Action Buttons (Connect to Hub/Email/PDF)
        const hubBtn = modal.querySelector('button[onclick="openDocumentHub()"]');
        if (hubBtn) hubBtn.onclick = () => openDocumentHub(c.ID);

        const pdfBtn = modal.querySelector('button[onclick="exportCandidateToPDF()"]');
        if (pdfBtn) pdfBtn.onclick = () => exportCandidatePDF(c.ID);

        const emailBtn = modal.querySelector('button[onclick="openSendEmail()"]');
        if (emailBtn) emailBtn.onclick = () => openSendEmailModal(c.ID);

        // RESTORE VIEW CV BUTTON
        const viewCvBtn = document.getElementById('btn-view-cv-detail');
        if (viewCvBtn) {
            viewCvBtn.onclick = () => {
                const link = document.getElementById('detail-cv-link').value;
                if (link) window.open(link, '_blank');
                else Swal.fire('Thông báo', 'Chưa có link CV để xem.', 'info');
            };
        }

        // CV FILE UPLOAD FEEDBACK
        const fileInput = document.getElementById('detail-cv-file');
        if (fileInput) {
            fileInput.onchange = (e) => {
                if (e.target.files && e.target.files[0]) {
                    const fileName = e.target.files[0].name;
                    const linkInput = document.getElementById('detail-cv-link');
                    if (linkInput) {
                        linkInput.value = "[File sẵn sàng: " + fileName + "]";
                        linkInput.classList.add('text-success', 'fw-bold');
                    }
                }
            };
        }

        new bootstrap.Modal(modal).show();
    }

    // FUNCTION TO POPULATE REJECTION TYPE BY SOURCE
    function populateRejectionType(source, selectedValue = '') {
        const select = document.getElementById('detail-rejection-type');
        if (!select) return;

        select.innerHTML = '<option value="">-- Chọn loại --</option>';
        if (!source) return;

        const filtered = (window.rejectionReasonsData || []).filter(r => r.Type === source);
        filtered.sort((a, b) => (a.Order || 0) - (b.Order || 0)).forEach(r => {
            const opt = document.createElement('option');
            opt.value = r.Reason;
            opt.innerText = r.Reason;
            if (r.Reason === selectedValue) opt.selected = true;
            select.appendChild(opt);
        });
    }




    // 5. INIT CHARTS
    // (Already declared globally at the top)

    document.addEventListener('DOMContentLoaded', function () {
        initCharts();
    });

    // ACTIVITY LOGS
    function loadActivityLogs() {
        const container = document.getElementById('activity-log-container');
        if (container) container.innerHTML = '<div class="text-center text-muted py-3"><div class="spinner-border spinner-border-sm text-primary" role="status"></div> Đang tải...</div>';

        google.script.run.withSuccessHandler(renderActivityLogs).apiGetActivityLogs(20);
    }

    function renderActivityLogs(logs) {
        console.log('Activity Logs received:', logs); // DEBUG
        const container = document.getElementById('activity-log-container');
        if (!container) return;

        if (!logs || logs.length === 0) {
            container.innerHTML = '<div class="text-center text-muted py-3">Chưa có hoạt động nào.</div>';
            return;
        }

        let html = '<ul class="list-group list-group-flush">';
        logs.forEach(log => {
            // Handle different date formats or raw ISO string
            let timeStr = '';
            try {
                const dateObj = new Date(log.timestamp);
                if (!isNaN(dateObj)) {
                    timeStr = dateObj.toLocaleString('vi-VN');
                } else {
                    timeStr = log.timestamp; // Fallback to raw string
                }
            } catch (e) {
                timeStr = 'N/A';
            }

            let icon = 'fas fa-info-circle';
            let color = 'text-primary';
            let bgColor = 'bg-light';

            const lowAction = (log.action || '').toLowerCase();
            if (lowAction.includes('email')) { icon = 'fas fa-envelope'; color = 'text-warning'; bgColor = 'bg-warning-subtle'; }
            else if (lowAction.includes('thêm ứng viên')) { icon = 'fas fa-user-plus'; color = 'text-success'; bgColor = 'bg-success-subtle'; }
            else if (lowAction.includes('chuyển trạng thái')) { icon = 'fas fa-exchange-alt'; color = 'text-info'; bgColor = 'bg-info-subtle'; }
            else if (lowAction.includes('ghi chú')) { icon = 'fas fa-sticky-note'; color = 'text-secondary'; bgColor = 'bg-secondary-subtle'; }
            else if (lowAction.includes('cập nhật')) { icon = 'fas fa-edit'; color = 'text-primary'; }

            html += `
                <li class="list-group-item d-flex align-items-start border-0 border-bottom py-3">
                    <div class="me-3 mt-1 ${color} p-2 rounded-circle ${bgColor}" style="width: 40px; height: 40px; display: flex; align-items: center; justify-content: center;">
                        <i class="${icon}"></i>
                    </div>
                    <div class="flex-grow-1">
                        <div class="d-flex justify-content-between">
                            <strong>${log.user}</strong>
                            <small class="text-muted" style="font-size: 0.8rem;">${timeStr}</small>
                        </div>
                        <div class="text-dark mt-1">${log.details}</div>
                        ${log.action ? `<small class="text-muted fst-italic">${log.action}</small>` : ''}
                    </div>
                </li>
            `;
        });
        html += '</ul>';
        container.innerHTML = html;
    }

    function initCharts() {
        if (document.getElementById('recruitmentChart')) {
            const ctx1 = document.getElementById('recruitmentChart').getContext('2d');
            recruitmentChartInstance = new Chart(ctx1, {
                type: 'line',
                data: {
                    labels: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
                    datasets: [{
                        label: 'Ứng tuyển theo tháng',
                        data: Array(12).fill(0), // Init empty
                        borderColor: '#FFC107',
                        tension: 0.4,
                        fill: true,
                        backgroundColor: 'rgba(255, 193, 7, 0.1)'
                    }]
                },
                options: { responsive: true }
            });
        }

        // Ensure Source Chart also exists (if added in HTML)
        if (document.getElementById('sourceChart')) {
            const ctx2 = document.getElementById('sourceChart').getContext('2d');
            sourceChartInstance = new Chart(ctx2, {
                type: 'doughnut',
                data: {
                    labels: ['Website', 'LinkedIn', 'Facebook', 'Referral', 'Other'],
                    datasets: [{
                        data: [0, 0, 0, 0, 0],
                        backgroundColor: ['#FFC107', '#0D6EFD', '#198754', '#DC3545', '#6C757D']
                    }]
                }
            });
        }

        // 3. Funnel Chart
        if (document.getElementById('funnelChart')) {
            const ctx3 = document.getElementById('funnelChart').getContext('2d');
            funnelChartInstance = new Chart(ctx3, {
                type: 'bar',
                data: {
                    labels: [],
                    datasets: [{
                        label: 'Số lượng ứng viên',
                        data: [],
                        backgroundColor: '#0dcaf0',
                        borderRadius: 5
                    }]
                },
                options: {
                    indexAxis: 'y',
                    responsive: true,
                    scales: {
                        x: { beginAtZero: true }
                    }
                }
            });
        }
    }

    // Update data for chart
    function updateCharts() {
        if (!recruitmentChartInstance) return;


        // 1. Line Chart Data
        const monthCounts = Array(12).fill(0);
        candidatesData.forEach(c => {
            if (c.Applied_Date) {
                const dateParts = c.Applied_Date.split('T')[0].split('-'); // YYYY-MM-DD
                if (dateParts.length === 3) {
                    const month = parseInt(dateParts[1], 10) - 1; // 0-indexed
                    if (month >= 0 && month <= 11) monthCounts[month]++;
                }
            }
        });

        recruitmentChartInstance.data.datasets[0].data = monthCounts;
        recruitmentChartInstance.update();

        // 2. Source Chart Data (if implemented)
        if (sourceChartInstance) {
            const sources = { 'Website': 0, 'LinkedIn': 0, 'Facebook': 0, 'Referral': 0, 'Other': 0 };
            candidatesData.forEach(c => {
                let s = c.Source || 'Other';
                if (sources.hasOwnProperty(s)) sources[s]++;
                else sources['Other']++;
            });
            sourceChartInstance.data.datasets[0].data = Object.values(sources);
            sourceChartInstance.update();
        }

        // 3. Funnel & Analytics
        calculateAnalytics();
    }

    function calculateAnalytics() {
        // A. Funnel Data
        if (funnelChartInstance) {
            const stageCounts = {};
            if (stagesData.length > 0) {
                stagesData.sort((a, b) => a.Order - b.Order).forEach(s => stageCounts[s.Stage_Name] = 0);
            } else {
                ['Apply', 'Screening', 'Interview', 'Offer', 'Hired'].forEach(s => stageCounts[s] = 0);
            }

            candidatesData.forEach(c => {
                const s = getVal(c, 'Stage');
                if (s && stageCounts.hasOwnProperty(s)) {
                    stageCounts[s]++;
                }
            });

            funnelChartInstance.data.labels = Object.keys(stageCounts);
            funnelChartInstance.data.datasets[0].data = Object.values(stageCounts);
            funnelChartInstance.update();

            // Calculate Conversion Rate
            const total = candidatesData.length;
            const hiredCount = candidatesData.filter(c => {
                const s = (getVal(c, 'Stage') || '').toLowerCase();
                return s.includes('hired') || s.includes('đã tuyển') || s.includes('nhận việc') || s.includes('official');
            }).length;

            if (total > 0 && document.getElementById('stat-conversion-rate')) {
                const rate = ((hiredCount / total) * 100).toFixed(1);
                document.getElementById('stat-conversion-rate').innerText = rate + '%';
            }
        }

        // B. Advanced Analytics: Rejection & Time to Hire
        const hiredCandidates = candidatesData.filter(c => {
            const s = (getVal(c, 'Stage') || '').toLowerCase();
            return s.includes('hired') || s.includes('đã tuyển') || s.includes('nhận việc') || s.includes('official');
        });

        // 1. Time to Hire
        let totalDays = 0;
        let hireCount = 0;
        const deptTime = {}; // For chart

        hiredCandidates.forEach(c => {
            const applied = getVal(c, 'Applied_Date');
            const hired = getVal(c, 'Hire_Date');
            if (applied && hired) {
                const d1 = new Date(applied);
                const d2 = new Date(hired);
                if (!isNaN(d1) && !isNaN(d2)) {
                    const diffDays = Math.ceil(Math.abs(d2 - d1) / (1000 * 60 * 60 * 24));
                    totalDays += diffDays;
                    hireCount++;

                    const dept = getVal(c, 'Department') || 'Khác';
                    if (!deptTime[dept]) deptTime[dept] = { total: 0, count: 0 };
                    deptTime[dept].total += diffDays;
                    deptTime[dept].count++;
                }
            }
        });

        if (hireCount > 0 && document.getElementById('stat-time-to-hire')) {
            document.getElementById('stat-time-to-hire').innerText = (totalDays / hireCount).toFixed(1);
        }

        if (timeToHireChartInstance) {
            const labels = Object.keys(deptTime);
            const data = labels.map(l => (deptTime[l].total / deptTime[l].count).toFixed(1));
            timeToHireChartInstance.data.labels = labels;
            timeToHireChartInstance.data.datasets[0].data = data;
            timeToHireChartInstance.update();
        }

        // 2. Rejection Analysis
        if (rejectionChartInstance) {
            const rejectReasons = {};
            candidatesData.forEach(c => {
                const s = (getVal(c, 'Stage') || '').toLowerCase();
                if (s.includes('loại') || s.includes('reject')) {
                    const reason = getVal(c, 'Rejection_Reason') || 'Không rõ lý do';
                    rejectReasons[reason] = (rejectReasons[reason] || 0) + 1;
                }
            });

            rejectionChartInstance.data.labels = Object.keys(rejectReasons);
            rejectionChartInstance.data.datasets[0].data = Object.values(rejectReasons);
            rejectionChartInstance.update();
        }
    }





    // 6.1 SAVE CANDIDATE DETAIL (UNIFIED)
    function saveCandidateDetail() {
        if (!currentUser) return;

        const id = document.getElementById('current-candidate-id').value;
        const isNew = !id;
        const modal = document.getElementById('candidateDetailModal');
        const mode = modal.getAttribute('data-mode');

        if (mode === 'view') {
            Swal.fire('Thông báo', 'Bạn đang ở chế độ chỉ xem.', 'info');
            return;
        }

        // Collect Data
        const data = {
            ID: id,
            Name: document.getElementById('detail-name').value,
            Gender: document.getElementById('detail-gender').value,
            Birth_Year: document.getElementById('detail-dob').value,
            Phone: document.getElementById('detail-phone').value,
            Email: document.getElementById('detail-email').value,
            Department: document.getElementById('detail-department').value,
            Position: document.getElementById('detail-position').value,
            Experience: document.getElementById('detail-experience').value,
            School: document.getElementById('detail-school').value,
            Education_Level: document.getElementById('detail-education-level').value,
            Major: document.getElementById('detail-major').value,
            Salary_Expectation: document.getElementById('detail-salary').value,
            Source: document.getElementById('detail-source').value,
            Recruiter: document.getElementById('detail-recruiter').value,
            TicketID: document.getElementById('detail-ticket-id').value,
            Stage: document.getElementById('detail-status').value,
            Status: document.getElementById('detail-contact-status').value,
            CV_Link: document.getElementById('detail-cv-link').value,
            Rejection_Source: document.getElementById('detail-rejection-source') ? document.getElementById('detail-rejection-source').value : '',
            Rejection_Type: document.getElementById('detail-rejection-type') ? document.getElementById('detail-rejection-type').value : '',
            Rejection_Reason: document.getElementById('detail-rejection-reason') ? document.getElementById('detail-rejection-reason').value : '',

            NewNote: document.getElementById('detail-new-note').value,
            User: currentUser.username || currentUser.email
        };

        // Validation
        if (!data.Name || !data.Phone || !data.Email) {
            Swal.fire('Lỗi', 'Vui lòng điền Họ tên, SĐT và Email.', 'warning');
            return;
        }

        const btn = document.querySelector('#candidateDetailModal .modal-footer .btn-primary');
        const originalText = btn.innerText;
        btn.innerText = 'Đang lưu...';
        btn.disabled = true;

        const newStage = data.Stage;
        const currentStage = window._currentCandidateStage || '';

        // Handle Rejection Reason popup or visibility
        if (newStage !== currentStage && (newStage.toLowerCase().includes('loại') || newStage.toLowerCase().includes('reject') || newStage.toLowerCase().includes('từ chối'))) {
            // If the fields are visible and filled in the modal, use those.
            // Otherwise, prompt using the legacy popup for compatibility if needed.
            const typeInModal = document.getElementById('detail-rejection-type').value;
            const reasonInModal = document.getElementById('detail-rejection-reason').value;

            if (typeInModal || reasonInModal) {
                // If fields are already filled (e.g. user updated them in modal), 
                // data.Rejection_Type/Reason are already set from collection step.
                continueSaving();
            } else {
                promptRejectionReason(function (rejectionData) {
                    data.Rejection_Type = rejectionData.type;
                    data.Rejection_Reason = rejectionData.reason;
                    continueSaving();
                }, function () {
                    // Cancel
                    btn.innerText = originalText;
                    btn.disabled = false;
                });
            }
        } else {
            continueSaving();
        }

        function continueSaving() {
            const file = document.getElementById('detail-cv-file').files[0];
            if (file) {
                const reader = new FileReader();
                reader.onload = function (e) {
                    const base64 = e.target.result.split(',')[1];
                    const fileData = { name: file.name, type: file.type, data: base64 };
                    sendToBackend(data, fileData);
                };
                reader.readAsDataURL(file);
            } else {
                sendToBackend(data, null);
            }
        }


        function sendToBackend(formData, fileData) {
            const run = google.script.run.withSuccessHandler(function (res) {
                btn.innerText = originalText;
                btn.disabled = false;

                if (res.success) {
                    Swal.fire('Thành công', res.message || 'Hành động hoàn tất!', 'success').then(() => {
                        if (isNew) {
                            Swal.fire({
                                title: 'Gửi Email?',
                                text: "Gửi email xác nhận đã nhận hồ sơ cho ứng viên?",
                                icon: 'question',
                                showCancelButton: true,
                                confirmButtonText: 'Gửi ngay',
                                cancelButtonText: 'Không'
                            }).then((result) => {
                                if (result.isConfirmed) {
                                    google.script.run.apiSendApplicationReceivedEmail(res.candidateId || res.ID);
                                }
                            });
                        }
                    });

                    bootstrap.Modal.getInstance(document.getElementById('candidateDetailModal')).hide();
                    loadDashboardData();
                } else {
                    Swal.fire('Lỗi', res.message, 'error');
                }
            }).withFailureHandler(function (err) {
                btn.innerText = originalText;
                btn.disabled = false;
                Swal.fire('Lỗi hệ thống', err.message, 'error');
            });

            if (isNew) {
                run.apiCreateCandidate(formData, fileData);
            } else {
                run.apiUpdateCandidate(formData, fileData);
            }
        }
    }

    // 7. JOB MANAGEMENT (UPDATED)
    let jobsData = [];

    function loadJobs() {
        const tbody = document.querySelector('#jobs-table tbody');
        if (tbody) tbody.innerHTML = '<tr><td colspan="7" class="text-center"><div class="spinner-border text-primary"></div></td></tr>';

        google.script.run.withSuccessHandler(function (data) {
            jobsData = data || [];

            // Check for Debug Error
            if (jobsData.length > 0 && jobsData[0].ID === 'CRITICAL') {
                Swal.fire('Lỗi Backend V3', jobsData[0].Title, 'error');
            }
            if (jobsData.length > 0 && jobsData[0].ID === 'WARN') {
                console.warn(jobsData[0].Title);
            }

            renderJobs();
        }).withFailureHandler(function (err) {
            Swal.fire('Lỗi', 'Không thể tải danh sách tin tuyển dụng (V3): ' + err.message, 'error');
            if (tbody) tbody.innerHTML = '<tr><td colspan="7" class="text-center text-danger">Lỗi tải dữ liệu (V3)</td></tr>';
        }).apiGetJobsV3();
    }

    function renderJobs() {
        const tbody = document.querySelector('#jobs-table tbody');
        if (!tbody) return;
        tbody.innerHTML = '';

        if (jobsData.length === 0) {
            tbody.innerHTML = '<tr><td colspan="7" class="text-center py-4">Chưa có tin tuyển dụng nào.</td></tr>';
            return;
        }

        jobsData.forEach(job => {
            const tr = document.createElement('tr');

            // Status Badge
            let statusBadge = '';
            if (job.Status === 'Open' || job.Status === 'Mở' || job.Status === 'Đang tuyển') {
                statusBadge = '<span class="badge bg-success">Đang tuyển</span>';
            } else {
                statusBadge = '<span class="badge bg-secondary">Đã đóng</span>';
            }

            // Safe property access
            const id = job.ID || '';
            const title = job.Title || '';
            const dept = job.Department || '';
            const loc = job.Location || '';
            const date = job.Created_Date ? job.Created_Date.toString().slice(0, 10) : '';

            tr.innerHTML = `
                <td><small class="text-muted">${id}</small></td>
                <td class="fw-bold"><a href="javascript:void(0)" onclick="viewJob('${id}')" class="text-decoration-none text-dark">${title}</a></td>
                <td><span class="badge bg-light text-dark border">${dept}</span></td>
                <td>${loc}</td>
                <td>${date}</td>
                <td>${statusBadge}</td>
                <td>
                    ${(!currentUser || currentUser.role !== 'Viewer') ? `
                    <button class="btn btn-sm btn-info text-white" onclick="editJob('${id}')" title="Chỉnh sửa">
                        <i class="fas fa-edit"></i>
                    </button>
                    <button class="btn btn-sm btn-outline-primary" onclick="toggleJobStatus('${id}', '${job.Status}')" title="Đổi trạng thái">
                        <i class="fas fa-sync-alt"></i>
                    </button>
                    <button class="btn btn-sm btn-danger" onclick="deleteJob('${id}')" title="Xóa">
                        <i class="fas fa-trash"></i>
                    </button>
                    ` : '<span class="text-muted small">Chỉ xem</span>'}
                </td>
            `;
            tbody.appendChild(tr);
        });
    }

    let isEditing = false;
    let editingId = null;

    function resetJobForm() {
        isEditing = false;
        editingId = null;
        document.getElementById('add-job-form').reset();
        document.querySelector('#addJobModal .modal-title').innerText = 'Tạo Tin Tuyển Dụng Mới';
        document.querySelector('#addJobModal .btn-primary-custom').innerText = 'Tạo Tin';
    }

    // Call this when clicking "Tạo Tin Mới"
    // We can attach it globally or via listener
    document.addEventListener('click', function (e) {
        if (e.target && e.target.innerText && e.target.innerText.includes('Tạo Tin Mới')) {
            resetJobForm();
        }
    });

    function saveJob() {
        const form = document.getElementById('add-job-form');
        if (!form) return;

        const title = form.querySelector('[name="title"]').value;
        const department = form.querySelector('[name="department"]').value;
        const position = form.querySelector('[name="position"]').value; // Added
        const location = form.querySelector('[name="location"]').value;
        const type = form.querySelector('[name="type"]').value;
        const description = form.querySelector('[name="description"]').value;
        const ticketId = form.querySelector('[name="ticketId"]').value; // Added

        if (!title) {
            Swal.fire('Lỗi', 'Vui lòng nhập tiêu đề', 'warning');
            return;
        }

        const jobData = {
            id: editingId, // for update
            title: title,
            department: department,
            position: position,
            location: location,
            type: type,
            description: description,
            ticketId: ticketId
        };

        const btn = document.querySelector('#addJobModal .btn-primary-custom');
        const originalText = btn.innerHTML;
        btn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Đang lưu...';
        btn.disabled = true;

        const handler = function (res) {
            btn.innerHTML = originalText;
            btn.disabled = false;
            if (res.success) {
                Swal.fire('Thành công', isEditing ? 'Đã cập nhật tin!' : 'Đã tạo tin tuyển dụng!', 'success');
                const modalEl = document.getElementById('addJobModal');
                const modal = bootstrap.Modal.getInstance(modalEl);
                if (modal) modal.hide();
                form.reset();
                loadJobs();
            } else {
                Swal.fire('Lỗi', res.message, 'error');
            }
        };

        const failure = function (error) {
            btn.innerHTML = originalText;
            btn.disabled = false;
            Swal.fire('Lỗi', error.message, 'error');
        };

        if (isEditing) {
            google.script.run.withSuccessHandler(handler).withFailureHandler(failure).apiUpdateJobV3(jobData);
        } else {
            google.script.run.withSuccessHandler(handler).withFailureHandler(failure).apiCreateJob(jobData);
        }
    }

    function editJob(id) {
        const job = jobsData.find(j => j.ID == id);
        if (!job) return;

        isEditing = true;
        editingId = id;

        // Populate Form
        const form = document.getElementById('add-job-form');
        form.querySelector('[name="title"]').value = job.Title || '';
        form.querySelector('[name="department"]').value = job.Department || '';
        form.querySelector('[name="location"]').value = job.Location || '';
        form.querySelector('[name="type"]').value = job.Type || '';
        form.querySelector('[name="description"]').value = job.Description || '';

        // UI Updates
        document.querySelector('#addJobModal .modal-title').innerText = 'Cập nhật Tin Tuyển Dụng';
        document.querySelector('#addJobModal .btn-primary-custom').innerText = 'Cập nhật';

        // Show Modal
        const modal = new bootstrap.Modal(document.getElementById('addJobModal'));
        modal.show();
    }

    function toggleJobStatus(id, currentStatus) {
        // Toggle logic: If Open/Mở -> Closed. Else -> Open.
        let newStatus = 'Closed';
        if (currentStatus !== 'Open' && currentStatus !== 'Mở' && currentStatus !== 'Đang tuyển') {
            newStatus = 'Open';
        }

        google.script.run.withSuccessHandler(function (res) {
            if (res.success) {
                // Update local data for speed
                const j = jobsData.find(x => x.ID == id);
                if (j) j.Status = newStatus;
                renderJobs();

                const toast = Swal.mixin({
                    toast: true,
                    position: 'top-end',
                    showConfirmButton: false,
                    timer: 3000
                });
                toast.fire({ icon: 'success', title: 'Đã cập nhật trạng thái' });
            } else {
                Swal.fire('Lỗi', res.message, 'error');
            }
        }).apiUpdateJobStatus(id, newStatus);
    }

    function viewJob(id) {
        const job = jobsData.find(j => j.ID == id);
        if (!job) return;

        document.getElementById('view-job-title').innerText = job.Title;
        document.getElementById('view-job-dept').innerText = job.Department;
        document.getElementById('view-job-loc').innerText = job.Location;
        document.getElementById('view-job-type').innerText = job.Type;
        document.getElementById('view-job-status').innerText = job.Status;
        document.getElementById('view-job-date').innerText = job.Created_Date ? job.Created_Date.toString().slice(0, 10) : '';
        document.getElementById('view-job-desc').innerText = job.Description || 'Không có mô tả.';

        // Wire up the "Edit" button inside the view modal
        const editBtn = document.getElementById('btn-edit-from-view');
        editBtn.onclick = function () {
            // Hide view modal
            const viewModalEl = document.getElementById('viewJobModal');
            const viewModal = bootstrap.Modal.getInstance(viewModalEl);
            if (viewModal) viewModal.hide();

            // Open edit modal
            editJob(id);
        };

        const modal = new bootstrap.Modal(document.getElementById('viewJobModal'));
        modal.show();
    }

    function deleteJob(id) {
        Swal.fire({
            title: 'Xóa tin tuyển dụng?',
            text: "Hành động này không thể hoàn tác!",
            icon: 'warning',
            showCancelButton: true,
            confirmButtonColor: '#d33',
            cancelButtonColor: '#3085d6',
            confirmButtonText: 'Xóa',
            cancelButtonText: 'Hủy'
        }).then((result) => {
            if (result.isConfirmed) {
                google.script.run.withSuccessHandler(function (res) {
                    if (res.success) {
                        Swal.fire('Đã xóa!', 'Tin tuyển dụng đã bị xóa.', 'success');
                        loadJobs();
                    } else {
                        Swal.fire('Lỗi', res.message, 'error');
                    }
                }).apiDeleteJob(id);
            }
        })
    }

    // Initialize listener for Jobs section
    document.addEventListener('DOMContentLoaded', function () {
        // Check for existing listener approach (onclick in HTML)
        // We can also attach to link click
        const jobsLink = document.querySelector('a[onclick*="showSection(\'jobs\'"]');
        if (jobsLink) {
            jobsLink.addEventListener('click', function () {
                // Delay slightly to let UI switch
                setTimeout(loadJobs, 100);
            });
        }
    });

    // 8. SETTINGS MANAGEMENT
    function loadSettings() {
        google.script.run.withSuccessHandler(function (data) {
            usersData = data.users || [];
            stagesData = data.stages || [];
            renderSettings(data);
        }).apiGetSettings();
        loadEmailTemplates();
        loadNews();
    }

    // ... (renderSettings exists) ...

    // 8. SETTINGS MANAGEMENT
    function loadEmailTemplates() {
        google.script.run.withSuccessHandler(function (data) {
            emailTemplatesData = data;
            const list = document.getElementById('email-template-list');
            if (!list) return;
            list.innerHTML = '';

            emailTemplatesData.forEach(t => {
                const item = document.createElement('a');
                item.href = '#';
                item.className = 'list-group-item list-group-item-action';
                item.innerText = t.Name;
                item.onclick = (e) => {
                    e.preventDefault();
                    document.querySelectorAll('#email-template-list a').forEach(a => a.classList.remove('active'));
                    item.classList.add('active');
                    selectTemplate(t.ID);
                };
                list.appendChild(item);
            });
        }).apiGetEmailTemplates();
    }

    function selectTemplate(id) {
        const t = emailTemplatesData.find(x => x.ID == id);
        if (!t) return;

        document.getElementById('template-editor-title').innerText = 'Đang sửa: ' + t.Name;
        document.getElementById('email-template-form').style.display = 'block';
        document.getElementById('tpl-id').value = t.ID;
        document.getElementById('tpl-name').value = t.Name;
        document.getElementById('tpl-subject').value = t.Subject;
        document.getElementById('tpl-body').value = t.Body;
    }

    function addEmailTemplate() {
        document.getElementById('template-editor-title').innerText = 'Thêm Mẫu Mới';
        document.getElementById('email-template-form').style.display = 'block';
        document.getElementById('tpl-id').value = '';
        document.getElementById('tpl-name').value = '';
        document.getElementById('tpl-subject').value = '';
        document.getElementById('tpl-body').value = '';

        // Remove active class from list
        document.querySelectorAll('#email-template-list a').forEach(a => a.classList.remove('active'));
    }

    function deleteEmailTemplate() {
        const id = document.getElementById('tpl-id').value;
        if (!id) {
            // If creating new, just hide form
            document.getElementById('email-template-form').style.display = 'none';
            return;
        }

        Swal.fire({
            title: 'Chắc chắn xóa?',
            text: "Không thể hoàn tác hành động này!",
            icon: 'warning',
            showCancelButton: true,
            confirmButtonColor: '#d33',
            cancelButtonColor: '#3085d6',
            confirmButtonText: 'Xóa ngay'
        }).then((result) => {
            if (result.isConfirmed) {
                google.script.run.withSuccessHandler(function (res) {
                    if (res.success) {
                        Swal.fire('Đã xóa!', 'Mẫu email đã bị xóa.', 'success');
                        document.getElementById('email-template-form').style.display = 'none';
                        loadEmailTemplates();
                    } else {
                        Swal.fire('Lỗi', res.message, 'error');
                    }
                }).apiDeleteEmailTemplate(id);
            }
        });
    }

    function saveEmailTemplate() {
        const data = {
            id: document.getElementById('tpl-id').value,
            name: document.getElementById('tpl-name').value,
            subject: document.getElementById('tpl-subject').value,
            body: document.getElementById('tpl-body').value
        };

        if (!data.name || !data.subject) {
            Swal.fire('Lỗi', 'Vui lòng nhập Tên mẫu và Tiêu đề', 'error');
            return;
        }

        const btn = document.querySelector('#email-template-form .btn-primary'); // Save button
        const text = btn.innerText;
        btn.innerText = 'Đang lưu...';
        btn.disabled = true;

        google.script.run.withSuccessHandler(function (res) {
            btn.innerText = text;
            btn.disabled = false;

            if (res.success) {
                Swal.fire('Thành công', 'Đã lưu mẫu email', 'success');
                // Reload list to reflect changes
                loadEmailTemplates();
                // Optionally hide form if it was new, or keep open. 
                // Keeping open is fine, but we might want to update the title if Name changed.
                // Simpler to just refresh list.
            } else {
                Swal.fire('Lỗi', res.message, 'error');
            }
        }).apiSaveEmailTemplate(data);
    }

    function renderSettings(data) {
        renderCompanyInfo(data.companyInfo);
        // Users Table
        const userTbody = document.querySelector('#users-table tbody');
        if (userTbody) {
            userTbody.innerHTML = '';
            // Use global usersData if available, fallback to data.users
            const users = usersData.length ? usersData : (data.users || []);

            users.forEach(u => {
                userTbody.innerHTML += `
                    <tr>
                        <td>${u.Username}</td>
                        <td>${u.Full_Name || ''}</td>
                        <td>${u.Email || ''}</td>
                        <td>${u.Phone || ''}</td>
                        <td><span class="badge bg-secondary">${u.Role}</span></td>
                        <td>${u.Department || ''}</td>
                        <td>
                             <button class="btn btn-sm btn-outline-primary me-1" onclick="openUserModal('${u.Username}')">Sửa</button>
                             <button class="btn btn-sm btn-outline-danger" onclick="deleteUser('${u.Username}')">Xóa</button>
                        </td>
                    </tr>
                `;
            });
        }

        // Rejection Reasons
        if (typeof renderRejectionReasons === 'function') {
            renderRejectionReasons(window.rejectionReasonsData || []);
        }

        // Rejection Reasons
        if (typeof renderRejectionReasons === 'function') {
            renderRejectionReasons(window.rejectionReasonsData || []);
        }
    }

    function renderCompanyInfo(info) {
        if (!info) return;

        const nameEl = document.getElementById('comp-name');
        const taxEl = document.getElementById('comp-taxcode');
        const emailEl = document.getElementById('comp-email');
        const phoneEl = document.getElementById('comp-phone');
        const worktimeEl = document.getElementById('comp-worktime');

        if (nameEl) nameEl.value = info.name || '';
        if (taxEl) taxEl.value = info.taxcode || '';
        if (emailEl) emailEl.value = info.email || '';
        if (phoneEl) phoneEl.value = info.phone || '';
        if (worktimeEl) worktimeEl.value = info.worktime || '';

        const addrContainer = document.getElementById('comp-addresses-container');
        if (addrContainer) {
            addrContainer.innerHTML = '';
            if (info.addresses && Array.isArray(info.addresses) && info.addresses.length > 0) {
                info.addresses.forEach(addr => addCompanyAddress(addr));
            } else {
                addCompanyAddress();
            }
        }

        const signerContainer = document.getElementById('comp-signers-container');
        if (signerContainer) {
            signerContainer.innerHTML = '';
            if (info.signers && Array.isArray(info.signers) && info.signers.length > 0) {
                info.signers.forEach(s => {
                    // Handle both old string format and new object format
                    if (typeof s === 'string') {
                        addCompanySigner({ name: s, position: '' });
                    } else {
                        addCompanySigner(s);
                    }
                });
            } else {
                addCompanySigner();
            }
        }
    }

    function addCompanyAddress(value = '') {
        const container = document.getElementById('comp-addresses-container');
        if (!container) return;
        const div = document.createElement('div');
        div.className = 'input-group';
        div.innerHTML = `
            <input type="text" class="form-control" placeholder="Nhập địa chỉ..." value="${value}">
            <button class="btn btn-outline-danger" onclick="this.parentElement.remove()"><i class="fas fa-times"></i></button>
        `;
        container.appendChild(div);
    }

    function addCompanySigner(signer = { name: '', position: '' }) {
        const container = document.getElementById('comp-signers-container');
        if (!container) return;
        const div = document.createElement('div');
        div.className = 'input-group';
        div.innerHTML = `
            <input type="text" class="form-control w-50" placeholder="Họ tên..." value="${signer.name || ''}">
            <input type="text" class="form-control" placeholder="Chức vụ..." value="${signer.position || ''}">
            <button class="btn btn-outline-danger" onclick="this.parentElement.remove()"><i class="fas fa-times"></i></button>
        `;
        container.appendChild(div);
    }

    function saveCompanyInfo() {
        const nameEl = document.getElementById('comp-name');
        const taxEl = document.getElementById('comp-taxcode');
        const emailEl = document.getElementById('comp-email');
        const phoneEl = document.getElementById('comp-phone');
        const worktimeEl = document.getElementById('comp-worktime');

        const addrInputs = document.querySelectorAll('#comp-addresses-container input');
        const addresses = Array.from(addrInputs).map(i => i.value.trim()).filter(v => v);

        const signerRows = document.querySelectorAll('#comp-signers-container .input-group');
        const signers = Array.from(signerRows).map(row => {
            const inputs = row.querySelectorAll('input');
            return {
                name: inputs[0].value.trim(),
                position: inputs[1].value.trim()
            };
        }).filter(s => s.name);

        const info = {
            name: nameEl ? nameEl.value.trim() : '',
            taxcode: taxEl ? taxEl.value.trim() : '',
            email: emailEl ? emailEl.value.trim() : '',
            phone: phoneEl ? phoneEl.value.trim() : '',
            worktime: worktimeEl ? worktimeEl.value.trim() : '',
            addresses: addresses,
            signers: signers
        };

        const btn = document.querySelector('#tab-company .btn-primary');
        const originalText = btn ? btn.innerHTML : 'Lưu thông tin';
        if (btn) {
            btn.innerHTML = '<span class="spinner-border spinner-border-sm me-1"></span> Đang lưu...';
            btn.disabled = true;
        }

        google.script.run.withSuccessHandler(function (res) {
            if (btn) {
                btn.innerHTML = originalText;
                btn.disabled = false;
            }
            if (res.success) {
                Swal.fire('Thành công', res.message, 'success');
            } else {
                Swal.fire('Lỗi', res.message, 'error');
            }
        }).apiSaveCompanyInfo(info);
    }


    function renderRejectionReasons(reasons) {
        const companyList = document.getElementById('rejection-company-list');
        const candidateList = document.getElementById('rejection-candidate-list');
        if (companyList) companyList.innerHTML = '';
        if (candidateList) candidateList.innerHTML = '';

        reasons.sort((a, b) => (a.Order || 0) - (b.Order || 0)).forEach(r => {
            addRejectionReasonRow(r.Type, r.Reason, r.ID);
        });
    }

    function addRejectionReasonRow(type, reason = '', id = '') {
        const container = document.getElementById(type === 'Company' ? 'rejection-company-list' : 'rejection-candidate-list');
        if (!container) return;

        const div = document.createElement('div');
        div.className = 'input-group input-group-sm rejection-reason-row mb-2';
        div.setAttribute('data-type', type);
        div.setAttribute('data-id', id || 'R' + new Date().getTime());
        div.innerHTML = `
            <input type="text" class="form-control" value="${reason}" placeholder="Nhập lý do...">
            <button class="btn btn-outline-danger" onclick="this.parentElement.remove()"><i class="fas fa-times"></i></button>
        `;
        container.appendChild(div);
    }

    function saveRejectionReasons() {
        const reasons = [];
        document.querySelectorAll('.rejection-reason-row').forEach((row, index) => {
            const input = row.querySelector('input');
            const type = row.getAttribute('data-type');
            const rid = row.getAttribute('data-id');
            if (input.value.trim()) {
                reasons.push({
                    ID: rid,
                    Type: type,
                    Reason: input.value.trim(),
                    Order: index + 1
                });
            }
        });

        const btn = document.querySelector('#tab-rejection .btn-primary');
        const originalText = btn.innerHTML;
        btn.innerHTML = '<span class="spinner-border spinner-border-sm me-1"></span> Đang lưu...';
        btn.disabled = true;

        google.script.run.withSuccessHandler(function (res) {
            btn.innerHTML = originalText;
            btn.disabled = false;
            if (res.success) {
                Swal.fire('Thành công', 'Đã lưu danh sách lý do từ chối', 'success');
                window.rejectionReasonsData = reasons;
            } else {
                Swal.fire('Lỗi', res.message, 'error');
            }
        }).apiSaveRejectionReasons(reasons);
    }

    function promptRejectionReason(onConfirm, onCancel) {
        const reasons = window.rejectionReasonsData || [];
        const companyReasons = reasons.filter(r => r.Type === 'Company').sort((a, b) => a.Order - b.Order);
        const candidateReasons = reasons.filter(r => r.Type === 'Candidate').sort((a, b) => a.Order - b.Order);

        let html = `
            <div class="text-start">
                <label class="form-label small fw-bold">Loại từ chối</label>
                <select id="swal-reject-type" class="form-select mb-3" onchange="updateSwalReasons()">
                    <option value="Company">Công ty từ chối</option>
                    <option value="Candidate">Ứng viên từ chối</option>
                </select>
                <label class="form-label small fw-bold">Lý do cụ thể</label>
                <select id="swal-reject-reason" class="form-select">
                    ${companyReasons.map(r => `<option value="${r.Reason}">${r.Reason}</option>`).join('')}
                    <option value="Khác">Khác...</option>
                </select>
            </div>
        `;

        // We need a way to update the second dropdown when the first one changes
        // Since SweetAlert's HTML is injected, we'll use a global helper or inline script
        window.updateSwalReasons = function () {
            const type = document.getElementById('swal-reject-type').value;
            const reasonSelect = document.getElementById('swal-reject-reason');
            const list = type === 'Company' ? companyReasons : candidateReasons;
            reasonSelect.innerHTML = list.map(r => `<option value="${r.Reason}">${r.Reason}</option>`).join('') + '<option value="Khác">Khác...</option>';
        };

        Swal.fire({
            title: 'Thông tin từ chối',
            html: html,
            icon: 'warning',
            showCancelButton: true,
            confirmButtonText: 'Xác nhận',
            cancelButtonText: 'Hủy',
            preConfirm: () => {
                const type = document.getElementById('swal-reject-type').value;
                let reason = document.getElementById('swal-reject-reason').value;
                if (reason === 'Khác') {
                    return Swal.showValidationMessage('Vui lòng chọn hoặc nhập lý do khác (Tính năng nhập text đang phát triển)');
                }
                return { type: type, reason: reason };
            }
        }).then((result) => {
            if (result.isConfirmed) {
                onConfirm(result.value);
            } else {
                if (onCancel) onCancel();
            }
        });
    }



    function openUserModal(username = null) {
        console.log('openUserModal called with username:', username);
        const modal = document.getElementById('addUserModal');
        const form = document.getElementById('user-form');
        const title = document.getElementById('userModalLabel');
        const modeInput = document.getElementById('u-mode');

        // Populate Departments (ensure Departments loaded)
        const deptSelect = document.getElementById('u-department');

        if (deptSelect) {
            const populateDepts = () => {
                const selectEl = document.getElementById('u-department');
                if (!selectEl) {
                    console.error('Dept Select Element not found on populate!');
                    return;
                }
                console.log('Populating department dropdown... Data:', departmentsData);
                selectEl.innerHTML = '<option value="All">All (Admin)</option>';

                if (departmentsData && Array.isArray(departmentsData) && departmentsData.length > 0) {
                    departmentsData.forEach(d => {
                        const opt = document.createElement('option');
                        opt.value = d.name;
                        opt.innerText = d.name;
                        selectEl.appendChild(opt);
                    });
                } else {
                    console.warn('departmentsData is empty or invalid in openUserModal');
                }
            };

            // ALWAYS fetch settings to ensure we have the very latest departments
            // This fixes the issue where user adds a dept in Settings but it doesn't show up here immediately
            console.log('Fetching departments for User Modal...');
            google.script.run.withSuccessHandler(function (res) {
                console.log('Departments fetch result:', res);
                if (res.debug) console.log('Server Debug:', res.debug);

                if (res.success) {
                    departmentsData = res.departments || [];
                    console.log('Updated departmentsData:', departmentsData);
                    populateDepts();
                } else {
                    console.error('Failed to fetch departments:', res.message);
                    populateDepts(); // fallback
                }
            }).withFailureHandler(function (err) {
                console.error('API Call Failed:', err);
                populateDepts(); // use what we have
            }).apiGetDepartments();

        }

        // Show the modal explicitly since we removed data-bs-toggle
        new bootstrap.Modal(modal).show();


        if (username) {
            // EDIT MODE
            const user = usersData.find(u => u.Username === username);
            if (!user) return;

            title.innerText = 'Sửa Người Dùng';
            modeInput.value = 'edit';

            document.getElementById('u-username').value = user.Username;
            document.getElementById('u-username').disabled = true; // Cannot change username

            document.getElementById('u-password').value = ''; // Don't show password
            document.getElementById('u-password-req').style.display = 'none';
            document.getElementById('u-password-hint').style.display = 'block';
            document.getElementById('u-password').required = false;

            document.getElementById('u-fullname').value = user.Full_Name || '';
            document.getElementById('u-email').value = user.Email || '';
            document.getElementById('u-phone').value = user.Phone || '';
            document.getElementById('u-role').value = user.Role || 'User';
            document.getElementById('u-department').value = user.Department || 'All';

        } else {
            // ADD MODE
            title.innerText = 'Thêm Người Dùng';
            modeInput.value = 'add';
            form.reset();
            document.getElementById('u-username').disabled = false;
            document.getElementById('u-password-req').style.display = 'inline';
            document.getElementById('u-password-hint').style.display = 'none';
            document.getElementById('u-password').required = true;
        }

        new bootstrap.Modal(modal).show();
    }

    function saveUser() {
        const mode = document.getElementById('u-mode').value;
        const user = {
            username: document.getElementById('u-username').value,
            password: document.getElementById('u-password').value,
            fullname: document.getElementById('u-fullname').value,
            email: document.getElementById('u-email').value,
            phone: document.getElementById('u-phone').value,
            role: document.getElementById('u-role').value,
            department: document.getElementById('u-department').value
        };

        if (!user.username) {
            Swal.fire('Lỗi', 'Vui lòng nhập Username', 'error');
            return;
        }

        // For Add mode, password is required
        if (mode === 'add' && !user.password) {
            Swal.fire('Lỗi', 'Vui lòng nhập Password', 'error');
            return;
        }

        const btn = document.querySelector('#addUserModal .btn-primary');
        const originalText = btn.innerText;
        btn.innerText = 'Đang lưu...';
        btn.disabled = true;

        const handler = function (res) {
            btn.innerText = originalText;
            btn.disabled = false;

            if (res.success) {
                Swal.fire('Thành công', mode === 'add' ? 'Đã thêm người dùng' : 'Đã cập nhật thông tin', 'success');

                // Properly hide modal and remove backdrop
                const modalEl = document.getElementById('addUserModal');
                const modal = bootstrap.Modal.getInstance(modalEl);
                if (modal) {
                    modal.hide();
                } else {
                    // Fallback if instance not found
                    new bootstrap.Modal(modalEl).hide();
                }

                // Manually remove backdrop if it sticks
                const backdrop = document.querySelector('.modal-backdrop');
                if (backdrop) backdrop.remove();
                document.body.classList.remove('modal-open');
                document.body.style.overflow = '';
                document.body.style.paddingRight = '';

                loadSettings(); // Reload table
            } else {
                Swal.fire('Lỗi', res.message, 'error');
            }
        };

        if (mode === 'add') {
            google.script.run.withSuccessHandler(handler).withFailureHandler(function (err) {
                btn.innerText = originalText;
                btn.disabled = false;
                Swal.fire('Lỗi', 'Lỗi kết nối: ' + err.message, 'error');
            }).apiCreateUser(user);
        } else {
            google.script.run.withSuccessHandler(handler).withFailureHandler(function (err) {
                btn.innerText = originalText;
                btn.disabled = false;
                Swal.fire('Lỗi', 'Lỗi kết nối: ' + err.message, 'error');
            }).apiEditUser(user);
        }
    }

    function deleteUser(username) {
        if (!confirm('Bạn có chắc muốn xóa user này?')) return;
        google.script.run.withSuccessHandler(function (res) {
            if (res.success) loadSettings();
            else Swal.fire('Lỗi', res.message, 'error');
        }).apiDeleteUser(username);
    }

    // CANDIDATE DETAIL & MANAGEMENT FUNCTIONS (openCandidateDetail is defined earlier at line ~510)


    function editCandidate(candidateId) {
        openCandidateDetail(candidateId);
    }

    function deleteCandidate(candidateId) {
        Swal.fire({
            title: 'Xác nhận xóa?',
            text: 'Bạn có chắc muốn xóa ứng viên này?',
            icon: 'warning',
            showCancelButton: true,
            confirmButtonColor: '#dc3545',
            cancelButtonColor: '#6c757d',
            confirmButtonText: 'Xóa',
            cancelButtonText: 'Hủy'
        }).then((result) => {
            if (result.isConfirmed) {
                // Call API to delete
                google.script.run.withSuccessHandler(function (res) {
                    if (res.success) {
                        Swal.fire('Đã xóa!', 'Ứng viên đã được xóa.', 'success');
                        loadDashboardData(); // Reload data
                    } else {
                        Swal.fire('Lỗi', res.message, 'error');
                    }
                }).apiDeleteCandidate(candidateId);
            }
        });
    }




    // ============================================
    // DEPARTMENT & POSITION MANAGEMENT
    // ============================================



    function loadDepartments() {
        google.script.run
            .withSuccessHandler(function (res) {
                if (res.success) {
                    departmentsData = res.departments || [];
                    renderDepartments();
                } else {
                    console.error('Failed to load departments:', res.message);
                }
            })
            .withFailureHandler(function (error) {
                console.error('Failed to load departments:', error);
            })
            .apiGetDepartments();
    }

    function renderDepartments() {
        const container = document.getElementById('departments-container');
        if (!container) return;

        container.innerHTML = '';

        departmentsData.forEach(dept => {
            const card = document.createElement('div');
            card.className = 'col-md-4';

            const positionsList = dept.positions.map(pos => `
                <li class="list-group-item d-flex justify-content-between align-items-center py-2">
                    <span>${pos}</span>
                    <div>
                        <button class="btn btn-sm btn-outline-secondary me-1" onclick="promptEditPosition('${dept.name}', '${pos}')" title="Sửa">
                            <i class="fas fa-edit" style="font-size: 0.7rem;"></i>
                        </button>
                        <button class="btn btn-sm btn-outline-danger" onclick="deletePosition('${dept.name}', '${pos}')" title="Xóa">
                            <i class="fas fa-trash" style="font-size: 0.7rem;"></i>
                        </button>
                    </div>
                </li>
            `).join('');

            card.innerHTML = `
                <div class="card h-100">
                    <div class="card-header bg-primary text-white d-flex justify-content-between align-items-center">
                        <h6 class="mb-0">${dept.name}</h6>
                        <div>
                            <button class="btn btn-sm btn-outline-light me-1" onclick="promptEditDepartment('${dept.name}')" title="Sửa">
                                <i class="fas fa-edit"></i>
                            </button>
                            <button class="btn btn-sm btn-outline-light" onclick="deleteDepartment('${dept.name}')" title="Xóa">
                                <i class="fas fa-trash"></i>
                            </button>
                        </div>
                    </div>
                    <div class="card-body p-0">
                        <ul class="list-group list-group-flush">
                            ${positionsList || '<li class="list-group-item text-muted">Chưa có vị trí</li>'}
                        </ul>
                    </div>
                    <div class="card-footer bg-white">
                        <button class="btn btn-sm btn-outline-primary w-100" onclick="promptAddPosition('${dept.name}')">
                            <i class="fas fa-plus"></i> Thêm vị trí
                        </button>
                    </div>
                </div>
            `;

            container.appendChild(card);
        });
    }

    function promptAddDepartment() {
        Swal.fire({
            title: 'Thêm Phòng ban mới',
            input: 'text',
            inputPlaceholder: 'Nhập tên phòng ban...',
            showCancelButton: true,
            confirmButtonText: 'Thêm',
            cancelButtonText: 'Hủy',
            inputValidator: (value) => {
                if (!value) {
                    return 'Vui lòng nhập tên phòng ban!';
                }
            }
        }).then((result) => {
            if (result.isConfirmed) {
                addDepartment(result.value);
            }
        });
    }

    function addDepartment(deptName) {
        google.script.run
            .withSuccessHandler(function (res) {
                if (res.success) {
                    Swal.fire('Thành công!', 'Đã thêm phòng ban', 'success');
                    loadDepartments(); // Reload
                    refreshAllDropdowns();  // Auto-refresh dropdowns
                } else {
                    Swal.fire('Lỗi', res.message, 'error');
                }
            })
            .withFailureHandler(function (error) {
                Swal.fire('Lỗi', error.message, 'error');
            })
            .apiAddDepartment(deptName);
    }

    function promptAddPosition(deptName) {
        Swal.fire({
            title: 'Thêm Vị trí mới',
            text: `Phòng ban: ${deptName}`,
            input: 'text',
            inputPlaceholder: 'Nhập tên vị trí...',
            showCancelButton: true,
            confirmButtonText: 'Thêm',
            cancelButtonText: 'Hủy',
            inputValidator: (value) => {
                if (!value) {
                    return 'Vui lòng nhập tên vị trí!';
                }
            }
        }).then((result) => {
            if (result.isConfirmed) {
                addPosition(deptName, result.value);
            }
        });
    }

    function addPosition(deptName, position) {
        google.script.run
            .withSuccessHandler(function (res) {
                if (res.success) {
                    Swal.fire('Thành công!', 'Đã thêm vị trí', 'success');
                    loadDepartments(); // Reload
                    refreshAllDropdowns();  // Auto-refresh dropdowns
                } else {
                    Swal.fire('Lỗi', res.message, 'error');
                }
            })
            .withFailureHandler(function (error) {
                Swal.fire('Lỗi', error.message, 'error');
            })
            .apiAddPosition(deptName, position);
    }

    function deleteDepartment(deptName) {
        Swal.fire({
            title: 'Xác nhận xóa phòng ban?',
            text: `Bạn có chắc muốn xóa "${deptName}"?`,
            icon: 'warning',
            showCancelButton: true,
            confirmButtonColor: '#dc3545',
            cancelButtonColor: '#6c757d',
            confirmButtonText: 'Xóa',
            cancelButtonText: 'Hủy'
        }).then((result) => {
            if (result.isConfirmed) {
                google.script.run
                    .withSuccessHandler(function (res) {
                        if (res.success) {
                            Swal.fire('Đã xóa!', 'Phòng ban đã được xóa', 'success');
                            loadDepartments();
                            refreshAllDropdowns();  // Auto-refresh dropdowns
                        } else {
                            Swal.fire('Lỗi', res.message, 'error');
                        }
                    })
                    .withFailureHandler(function (error) {
                        Swal.fire('Lỗi', error.message, 'error');
                    })
                    .apiDeleteDepartment(deptName);
            }
        });
    }

    function deletePosition(deptName, position) {
        Swal.fire({
            title: 'Xác nhận xóa vị trí?',
            text: `Bạn có chắc muốn xóa "${position}"?`,
            icon: 'warning',
            showCancelButton: true,
            confirmButtonColor: '#dc3545',
            cancelButtonColor: '#6c757d',
            confirmButtonText: 'Xóa',
            cancelButtonText: 'Hủy'
        }).then((result) => {
            if (result.isConfirmed) {
                google.script.run
                    .withSuccessHandler(function (res) {
                        if (res.success) {
                            Swal.fire('Đã xóa!', 'Vị trí đã được xóa', 'success');
                            loadDepartments();
                            refreshAllDropdowns();  // Auto-refresh dropdowns
                        } else {
                            Swal.fire('Lỗi', res.message, 'error');
                        }
                    })
                    .withFailureHandler(function (error) {
                        Swal.fire('Lỗi', error.message, 'error');
                    })
                    .apiDeletePosition(deptName, position);
            }
        });
    }

    // EDIT DEPARTMENT
    function promptEditDepartment(oldName) {
        Swal.fire({
            title: 'Sửa tên phòng ban',
            input: 'text',
            inputValue: oldName,
            inputPlaceholder: 'Nhập tên mới',
            showCancelButton: true,
            confirmButtonText: 'Lưu',
            cancelButtonText: 'Hủy',
            preConfirm: (newName) => {
                if (!newName) {
                    Swal.showValidationMessage('Vui lòng nhập tên phòng ban');
                }
                return newName;
            }
        }).then((result) => {
            if (result.isConfirmed && result.value) {
                editDepartment(oldName, result.value);
            }
        });
    }

    function editDepartment(oldName, newName) {
        google.script.run
            .withSuccessHandler(function (res) {
                if (res.success) {
                    Swal.fire('Đã cập nhật!', 'Tên phòng ban đã được cập nhật', 'success');
                    loadDepartments();
                    refreshAllDropdowns();  // Auto-refresh dropdowns
                } else {
                    Swal.fire('Lỗi', res.message, 'error');
                }
            })
            .withFailureHandler(function (error) {
                Swal.fire('Lỗi', error.message, 'error');
            })
            .apiEditDepartment(oldName, newName);
    }

    // EDIT POSITION
    function promptEditPosition(deptName, oldPosition) {
        Swal.fire({
            title: 'Sửa tên vị trí',
            text: `Phòng ban: ${deptName}`,
            input: 'text',
            inputValue: oldPosition,
            inputPlaceholder: 'Nhập tên mới',
            showCancelButton: true,
            confirmButtonText: 'Lưu',
            cancelButtonText: 'Hủy',
            preConfirm: (newPosition) => {
                if (!newPosition) {
                    Swal.showValidationMessage('Vui lòng nhập tên vị trí');
                }
                return newPosition;
            }
        }).then((result) => {
            if (result.isConfirmed && result.value) {
                editPosition(deptName, oldPosition, result.value);
            }
        });
    }

    function editPosition(deptName, oldPosition, newPosition) {
        google.script.run
            .withSuccessHandler(function (res) {
                if (res.success) {
                    Swal.fire('Đã cập nhật!', 'Tên vị trí đã được cập nhật', 'success');
                    loadDepartments();
                    refreshAllDropdowns();  // Auto-refresh dropdowns
                } else {
                    Swal.fire('Lỗi', res.message, 'error');
                }
            })
            .withFailureHandler(function (error) {
                Swal.fire('Lỗi', error.message, 'error');
            })
            .apiEditPosition(deptName, oldPosition, newPosition);
    }

    // ============================================
    // EMAIL WORKFLOW
    // ============================================

    let quill = null;
    let emailAttachments = [];
    let draftTimer = null;

    function initQuill() {
        if (quill) return;
        quill = new Quill('#email-body-quill', {
            theme: 'snow',
            modules: {
                toolbar: [
                    [{ 'font': [] }, { 'size': [] }],
                    ['bold', 'italic', 'underline', 'strike'],
                    [{ 'color': [] }, { 'background': [] }],
                    [{ 'list': 'ordered' }, { 'list': 'bullet' }],
                    ['link', 'image'],
                    ['clean']
                ]
            }
        });

        // Auto-save draft on change
        quill.on('text-change', () => {
            clearTimeout(draftTimer);
            draftTimer = setTimeout(saveDraft, 5000); // Save after 5s of inactivity
        });
    }

    function toggleEmailField(type) {
        const container = document.getElementById(`email-${type}-container`);
        container.style.display = container.style.display === 'none' ? 'flex' : 'none';
        if (container.style.display === 'flex') {
            document.getElementById(`email-${type}-input`).focus();
        }
    }

    function handleRecipientInput(event, type) {
        if (event.key === 'Enter' || event.key === ',') {
            event.preventDefault();
            const input = event.target;
            const email = input.value.trim();
            if (email) {
                if (validateEmail(email)) {
                    addRecipientTag(email, type);
                    input.value = '';
                } else {
                    Swal.showValidationMessage('Email không hợp lệ');
                }
            }
        }
    }

    function validateEmail(email) {
        return /\S+@\S+\.\S+/.test(email);
    }

    function addRecipientTag(email, type) {
        const tagsContainer = document.getElementById(`email-${type}-tags`);
        const input = document.getElementById(`email-${type}-input`);

        const tag = document.createElement('div');
        tag.className = 'email-tag';
        tag.innerHTML = `${email} <i class="fas fa-times" onclick="removeRecipientTag(this)"></i>`;
        tag.setAttribute('data-email', email);

        tagsContainer.insertBefore(tag, input);
    }

    function removeRecipientTag(element) {
        element.parentElement.remove();
    }

    function handleFileAttachments(input) {
        const files = input.files;
        if (!files.length) return;

        const list = document.getElementById('email-attachments-list');
        list.style.display = 'block';

        Array.from(files).forEach(file => {
            if (file.size > 10 * 1024 * 1024) { // 10MB limit
                Swal.fire('Lỗi', `File ${file.name} quá lớn (tối đa 10MB)`, 'warning');
                return;
            }

            const reader = new FileReader();
            reader.onload = function (e) {
                const base64 = e.target.result.split(',')[1];
                emailAttachments.push({
                    name: file.name,
                    content: base64,
                    contentType: file.type
                });

                const badge = document.createElement('div');
                badge.className = 'attachment-badge';
                badge.innerHTML = `
                    <i class="far fa-file-alt me-2"></i>
                    <span>${file.name} (${(file.size / 1024).toFixed(1)} KB)</span>
                    <i class="fas fa-times attachment-remove" onclick="removeAttachment(this, '${file.name}')"></i>
                `;
                list.appendChild(badge);
            };
            reader.readAsDataURL(file);
        });
        input.value = ''; // Reset input
    }

    function removeAttachment(element, name) {
        emailAttachments = emailAttachments.filter(a => a.name !== name);
        element.parentElement.remove();
        if (emailAttachments.length === 0) {
            document.getElementById('email-attachments-list').style.display = 'none';
        }
    }

    function openSendEmailModal(candidateId) {
        const c = candidatesData.find(x => x.ID == candidateId);
        if (!c) return;

        initQuill();

        // Reset form tags and attachments
        document.querySelectorAll('.email-tag').forEach(t => t.remove());
        emailAttachments = [];
        document.getElementById('email-attachments-list').innerHTML = '';
        document.getElementById('email-attachments-list').style.display = 'none';

        document.getElementById('email-candidate-id').value = candidateId;
        document.getElementById('sendEmailModal').setAttribute('data-candidate-id', candidateId);
        document.getElementById('email-to').value = c.Email || '';
        document.getElementById('email-subject').value = '';
        quill.setContents([]);

        // Populate Templates Dropdown
        if (emailTemplatesData.length === 0) {
            google.script.run.withSuccessHandler(function (data) {
                emailTemplatesData = data;
                populateTemplateDropdown();
            }).apiGetEmailTemplates();
        } else {
            populateTemplateDropdown();
        }

        const modal = new bootstrap.Modal(document.getElementById('sendEmailModal'));
        modal.show();

        // Load draft if exists
        const draft = localStorage.getItem(`draft_email_${candidateId}`);
        if (draft) {
            const d = JSON.parse(draft);
            document.getElementById('email-subject').value = d.subject || '';
            quill.root.innerHTML = d.body || '';
            document.getElementById('email-draft-status').innerText = `Tải lại bản nháp lúc ${d.time}`;
        } else {
            document.getElementById('email-draft-status').innerText = '';
        }
    }

    function populateTemplateDropdown() {
        const tplSelect = document.getElementById('email-template-select');
        tplSelect.innerHTML = '<option value="">-- Mẫu --</option>';
        emailTemplatesData.forEach(t => {
            const opt = document.createElement('option');
            opt.value = t.ID;
            opt.innerText = t.Name;
            tplSelect.appendChild(opt);
        });
    }

    function loadTemplateContent(tplId) {
        if (!tplId) return;
        const t = emailTemplatesData.find(x => x.ID == tplId);
        if (!t) return;

        const candidateId = document.getElementById('email-candidate-id').value;
        const c = candidatesData.find(x => x.ID == candidateId);

        let subject = t.Subject;
        let body = t.Body;

        if (c) {
            const name = c.Name || '';
            const pos = c.Position || '';
            const date = new Date().toLocaleDateString('vi-VN');

            subject = subject.replace(/\[Name\]|{{name}}/gi, name)
                .replace(/\[Position\]|{{position}}/gi, pos);

            body = body.replace(/\[Name\]|{{name}}/gi, name)
                .replace(/\[Position\]|{{position}}/gi, pos)
                .replace(/\[Date\]|{{date}}/gi, date);
        }

        document.getElementById('email-subject').value = subject;
        quill.root.innerHTML = body.replace(/\n/g, '<br>');
    }

    function saveDraft() {
        const candidateId = document.getElementById('email-candidate-id').value;
        if (!candidateId) return;

        const draft = {
            subject: document.getElementById('email-subject').value,
            body: quill.root.innerHTML,
            time: new Date().toLocaleTimeString('vi-VN')
        };
        localStorage.setItem(`draft_email_${candidateId}`, JSON.stringify(draft));
        document.getElementById('email-draft-status').innerText = `Đã lưu tự động lúc ${draft.time}`;
    }

    function clearDraft() {
        const candidateId = document.getElementById('email-candidate-id').value;
        localStorage.removeItem(`draft_email_${candidateId}`);
        document.getElementById('email-subject').value = '';
        quill.setContents([]);
        document.getElementById('email-draft-status').innerText = 'Đã xóa bản nháp';
    }

    function sendEmail() {
        const to = document.getElementById('email-to').value;
        const subject = document.getElementById('email-subject').value;
        const body = quill.root.innerHTML;

        // Collect CC and BCC tags
        const getEmailsFromTags = (type) => {
            return Array.from(document.querySelectorAll(`#email-${type}-tags .email-tag`))
                .map(t => t.getAttribute('data-email')).join(',');
        };
        const cc = getEmailsFromTags('cc');
        const bcc = getEmailsFromTags('bcc');

        if (!to) return Swal.fire('Lỗi', 'Chưa có người nhận', 'error');
        if (!subject || quill.getText().trim() === '') return Swal.fire('Lỗi', 'Thiếu tiêu đề hoặc nội dung', 'error');

        const submitBtn = document.getElementById('btn-send-email');
        const originalText = submitBtn.innerHTML;
        submitBtn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Đang gửi...';
        submitBtn.disabled = true;

        const candidateId = document.getElementById('email-candidate-id').value;

        google.script.run.withSuccessHandler(function (response) {
            submitBtn.innerHTML = originalText;
            submitBtn.disabled = false;
            if (response.success) {
                Swal.fire('Thành công', response.message, 'success');
                const modal = bootstrap.Modal.getInstance(document.getElementById('sendEmailModal'));
                modal.hide();
                localStorage.removeItem(`draft_email_${candidateId}`);
                if (document.getElementById('activity-log-container')) loadActivityLogs();
            } else {
                Swal.fire('Lỗi', response.message, 'error');
            }
        }).withFailureHandler(function (error) {
            submitBtn.innerHTML = originalText;
            submitBtn.disabled = false;
            Swal.fire('Lỗi', error.message, 'error');
        }).apiSendSmartEmail(to, cc, bcc, subject, body, candidateId, emailAttachments);
    }

    function openScheduleSendModal() {
        Swal.fire({
            title: 'Hẹn giờ gửi',
            html: '<input type="datetime-local" id="schedule-time" class="form-control">',
            showCancelButton: true,
            confirmButtonText: 'Xác nhận',
            preConfirm: () => {
                const time = document.getElementById('schedule-time').value;
                if (!time) return Swal.showValidationMessage('Vui lòng chọn thời gian');
                return time;
            }
        }).then((result) => {
            if (result.isConfirmed) {
                Swal.fire('Thông báo', 'Tính năng hẹn giờ đang được thiết lập backend. Mail sẽ được gửi lúc ' + result.value, 'info');
            }
        });
    }


    // REFRESH ALL DROPDOWNS (called after any Settings change)
    function refreshAllDropdowns() {
        console.log('Refreshing all dropdowns...');

        // Refresh Kanban filter dropdowns
        populateFilterDropdowns();

        // Refresh candidate detail modal dropdowns (if modal is open)
        const detailModal = document.getElementById('candidateDetailModal');
        if (detailModal && detailModal.classList.contains('show')) {
            const currentDept = document.getElementById('detail-department')?.value;
            const currentPos = document.getElementById('detail-position')?.value;
            const currentStatus = document.getElementById('detail-status')?.value;

            populateDepartmentDropdown(currentDept, () => {
                populatePositionDropdown(currentPos);
            });
            populateStatusDropdown(currentStatus);  // NEW - refresh status dropdown
        }

        // Refresh add candidate modal dropdowns (if modal is open)
        const addModal = document.getElementById('addCandidateModal');
        if (addModal && addModal.classList.contains('show')) {
            const currentDept = document.getElementById('add-department')?.value;
            const currentStatus = document.getElementById('add-status')?.value;

            // Populate add-department dropdown directly (populateDepartmentDropdown only works for detail-department)
            const addDeptDropdown = document.getElementById('add-department');
            if (addDeptDropdown && departmentsData.length > 0) {
                addDeptDropdown.innerHTML = '<option value="">Chọn phòng ban</option>';
                departmentsData.forEach(dept => {
                    const option = document.createElement('option');
                    option.value = dept.name;
                    option.textContent = dept.name;
                    if (dept.name === currentDept) option.selected = true;
                    addDeptDropdown.appendChild(option);
                });
                if (currentDept) {
                    populateAddPositionDropdown();
                }
            }
        }
    }

    // POPULATE STATUS DROPDOWN (Detail Modal)
    function populateStatusDropdown(selectedStatus) {
        const statusDropdown = document.getElementById('detail-status');
        if (!statusDropdown) return;

        const ticketId = document.getElementById('detail-ticket-id')?.value;
        let customStages = null;

        if (ticketId && ticketsData && projectsData) {
            const ticket = ticketsData.find(t => String(t['Mã Ticket']) === String(ticketId));
            if (ticket) {
                const project = projectsData.find(p => p['Mã Dự án'] === ticket['Mã Dự án']);
                if (project && project['Quy trình (Workflow)']) {
                    customStages = project['Quy trình (Workflow)'].split(',').map(s => s.trim());
                }
            }
        }

        statusDropdown.innerHTML = '<option value="">Chọn giai đoạn</option>';

        if (customStages && customStages.length > 0) {
            customStages.forEach(stageName => {
                const option = document.createElement('option');
                option.value = stageName;
                option.textContent = stageName;
                if (stageName === selectedStatus) option.selected = true;
                statusDropdown.appendChild(option);
            });
        } else if (stagesData && stagesData.length > 0) {
            // Sort by order
            const sorted = [...stagesData].sort((a, b) => (a.Order || 0) - (b.Order || 0));
            sorted.forEach(stage => {
                const option = document.createElement('option');
                option.value = stage.Stage_Name;
                option.textContent = stage.Stage_Name;
                if (stage.Stage_Name === selectedStatus) option.selected = true;
                statusDropdown.appendChild(option);
            });
        } else {
            // Fallback to default stages if no custom stages defined
            const defaultStages = ['Apply', 'Call Interview', 'Interview', 'Offer', 'Hired', 'Rejected'];
            defaultStages.forEach(stageName => {
                const option = document.createElement('option');
                option.value = stageName;
                option.textContent = stageName;
                if (stageName === selectedStatus) option.selected = true;
                statusDropdown.appendChild(option);
            });
        }
    }

    // POPULATE STATUS DROPDOWN (Add Candidate Modal)

    // Load departments when Settings tab is shown
    document.addEventListener('shown.bs.tab', function (e) {
        if (e.target.getAttribute('href') === '#tab-departments') {
            loadDepartments();
        }
    });

    // Populate department dropdown for candidate form
    function populateDepartmentDropdown(selectedValue, callback) {
        const deptDropdown = document.getElementById('detail-department');
        if (!deptDropdown) return;

        // Load departments if not already loaded
        if (departmentsData.length === 0) {
            google.script.run
                .withSuccessHandler(function (res) {
                    if (res.success) {
                        departmentsData = res.departments || [];
                        fillDepartmentOptions(deptDropdown, selectedValue);
                        if (callback) callback();
                    }
                })
                .apiGetDepartments();
        } else {
            fillDepartmentOptions(deptDropdown, selectedValue);
            if (callback) callback();
        }
    }

    function fillDepartmentOptions(dropdown, selectedValue) {
        dropdown.innerHTML = '<option value="">Chọn phòng ban</option>';
        departmentsData.forEach(dept => {
            const option = document.createElement('option');
            option.value = dept.name;
            option.textContent = dept.name;
            if (dept.name === selectedValue) {
                option.selected = true;
            }
            dropdown.appendChild(option);
        });
    }

    // Populate position dropdown based on selected department
    function populatePositionDropdown(selectedPosition) {
        const deptDropdown = document.getElementById('detail-department');
        const posDropdown = document.getElementById('detail-position');

        if (!deptDropdown || !posDropdown) return;

        const selectedDept = deptDropdown.value;
        posDropdown.innerHTML = '<option value="">Chọn vị trí</option>';

        if (!selectedDept) return;

        const dept = departmentsData.find(d => d.name === selectedDept);
        if (dept && dept.positions) {
            dept.positions.forEach(pos => {
                const option = document.createElement('option');
                option.value = pos;
                option.textContent = pos;
                if (pos === selectedPosition) {
                    option.selected = true;
                }
                posDropdown.appendChild(option);
            });
        }
    }

    // ============================================
    // KANBAN FILTER LOGIC
    // ============================================

    let kanbanFilters = {
        search: '',
        department: '',
        position: '',
        recruiter: '',
        dateFrom: '',
        dateTo: ''
    };

    // Populate filter dropdowns

    // RECRUITER MANAGEMENT FUNCTIONS
    function renderRecruiters() {
        const tbody = document.querySelector('#recruiters-table tbody');
        if (!tbody) return;
        tbody.innerHTML = '';

        if (recruitersData.length === 0) {
            tbody.innerHTML = '<tr><td colspan="7" class="text-center">Chưa có dữ liệu</td></tr>';
            return;
        }

        recruitersData.forEach(r => {
            const tr = document.createElement('tr');
            // ID fallback if old data
            const displayId = r.id || '';
            const displayJoinDate = r.joinDate ? new Date(r.joinDate).toLocaleDateString('vi-VN') : '';

            tr.innerHTML = `
                <td>${displayId}</td>
                <td>${r.name}</td>
                <td>${r.email || ''}</td>
                <td>${r.position || ''}</td>
                <td>${displayJoinDate}</td>
                <td class="text-center"><span class="badge bg-info text-white">${r.totalCandidates || 0}</span></td>
                <td class="text-center">
                    <button class="btn btn-sm btn-primary me-1" onclick="openRecruiterModal('${r.id}')" title="Sửa">
                        <i class="fas fa-edit"></i>
                    </button>
                    <button class="btn btn-sm btn-danger" onclick="deleteRecruiter('${r.id}')" title="Xóa">
                        <i class="fas fa-trash"></i>
                    </button>
                </td>
            `;
            tbody.appendChild(tr);
        });
    }

    // NEW: Open Recruiter Modal (Add or Edit)
    // Replaces promptAddRecruiter
    function openRecruiterModal(id = null) {
        const modalEl = document.getElementById('recruiterModal');
        const modal = new bootstrap.Modal(modalEl);

        // Reset form
        document.getElementById('recruiter-form').reset();
        document.getElementById('rec-id').value = '';

        if (id && id !== 'undefined' && id !== 'null') {
            // EDIT MODE
            document.getElementById('recruiterModalLabel').innerText = 'Cập nhật Người phụ trách';
            const rec = recruitersData.find(r => String(r.id) === String(id));
            if (rec) {
                document.getElementById('rec-id').value = rec.id || '';
                document.getElementById('rec-name').value = rec.name;
                document.getElementById('rec-email').value = rec.email || '';
                document.getElementById('rec-position').value = rec.position || '';
                document.getElementById('rec-joinDate').value = rec.joinDate || '';
            }
        } else {
            // ADD MODE
            document.getElementById('recruiterModalLabel').innerText = 'Thêm Người phụ trách';
            document.getElementById('rec-joinDate').value = new Date().toISOString().slice(0, 10);
        }

        modal.show();
    }

    // Legacy mapping
    function promptAddRecruiter() {
        openRecruiterModal();
    }

    // NEW: Save Recruiter
    function saveRecruiter() {
        const id = document.getElementById('rec-id').value;
        const name = document.getElementById('rec-name').value.trim();
        const email = document.getElementById('rec-email').value.trim();
        const position = document.getElementById('rec-position').value.trim();
        const joinDate = document.getElementById('rec-joinDate').value;

        if (!name) {
            Swal.fire('Lỗi', 'Vui lòng nhập tên', 'warning');
            return;
        }

        const data = {
            id: id,
            name: name,
            email: email,
            position: position,
            joinDate: joinDate
        };

        const handler = function (res) {
            if (res.success) {
                Swal.fire('Thành công', res.message, 'success');
                // Close modal
                const modalEl = document.getElementById('recruiterModal');
                const modal = bootstrap.Modal.getInstance(modalEl);
                if (modal) modal.hide();

                // Reload
                loadDashboardData();
            } else {
                Swal.fire('Lỗi', res.message, 'error');
            }
        };

        if (id) {
            // Edit
            google.script.run.withSuccessHandler(handler).apiEditRecruiter(data);
        } else {
            // Add
            google.script.run.withSuccessHandler(handler).apiAddRecruiter(data);
        }
    }

    // UPDATED: Delete Recruiter
    function deleteRecruiter(id) {
        if (!confirm('Bạn có chắc muốn xóa người nảy?')) return;
        google.script.run.withSuccessHandler(function (res) {
            if (res.success) {
                loadDashboardData();
            } else {
                Swal.fire('Lỗi', res.message, 'error');
            }
        }).apiDeleteRecruiter(id);
    }

    // Updated populateFilterDropdowns to use recruitersData
    function populateFilterDropdowns() {
        // Populate department filter
        const deptFilter = document.getElementById('filter-department');
        if (deptFilter && departmentsData.length > 0) {
            deptFilter.innerHTML = '<option value="">Tất cả phòng ban</option>';
            const uniqueDepts = [...new Set(departmentsData.map(d => d.name))];
            uniqueDepts.forEach(dept => {
                const option = document.createElement('option');
                option.value = dept;
                option.textContent = dept;
                if (kanbanFilters.department === dept) option.selected = true;
                deptFilter.appendChild(option);
            });
        }

        // Populate position filter (all positions from all departments)
        const posFilter = document.getElementById('filter-position');
        if (posFilter && departmentsData.length > 0) {
            posFilter.innerHTML = '<option value="">Tất cả vị trí</option>';
            const allPositions = [];
            departmentsData.forEach(dept => {
                dept.positions.forEach(pos => {
                    if (!allPositions.includes(pos)) {
                        allPositions.push(pos);
                    }
                });
            });
            allPositions.forEach(pos => {
                const option = document.createElement('option');
                option.value = pos;
                option.textContent = pos;
                if (kanbanFilters.position === pos) option.selected = true;
                posFilter.appendChild(option);
            });
        }

        // Populate recruiter filter (from recruitersData)
        const recruiterFilter = document.getElementById('filter-recruiter');
        if (recruiterFilter && recruitersData.length > 0) {
            recruiterFilter.innerHTML = '<option value="">Người phụ trách</option>';
            recruitersData.forEach(r => {
                const option = document.createElement('option');
                option.value = r.name;
                option.textContent = r.name;
                if (kanbanFilters.recruiter === r.name) option.selected = true;
                recruiterFilter.appendChild(option);
            });
        }
    }

    // Helper to populate recruiter dropdown in Modals
    function populateRecruiterSelect(selectId, selectedValue) {
        const select = document.getElementById(selectId);
        if (!select) return;

        select.innerHTML = '<option value="">Chọn người phụ trách</option>';
        if (recruitersData.length > 0) {
            recruitersData.forEach(r => {
                const option = document.createElement('option');
                option.value = r.name;
                option.textContent = r.name;
                if (r.name === selectedValue) {
                    option.selected = true;
                }
                select.appendChild(option);
            });
        }
    }

    function populateTicketDropdown(selectedTicketId) {
        const select = document.getElementById('detail-ticket-id');
        if (!select) return;
        select.innerHTML = '<option value="">-- Không có --</option>';

        if (ticketsData && ticketsData.length > 0) {
            ticketsData.forEach(t => {
                const option = document.createElement('option');
                option.value = t['Mã Ticket'];
                option.innerText = `[${t['Mã Ticket']}] ${t['Vị trí']} (${t['Mã Dự án']})`;
                if (String(t['Mã Ticket']) === String(selectedTicketId)) {
                    option.selected = true;
                }
                select.appendChild(option);
            });
        }
    }

    // Update refreshAllDropdowns to include recruiters AND SOURCES
    function refreshAllDropdowns() {
        console.log('🔄 Refreshing all dropdowns...');

        // 1. Refresh Department/Position filters
        populateFilterDropdowns();

        // 2. Refresh Sources
        populateCandidateSources();


        // 4. Refresh Detail Modal Dropdowns
        const detailModal = document.getElementById('candidateDetailModal');
        if (detailModal && detailModal.classList.contains('show')) {
            const currentRecruiter = document.getElementById('detail-recruiter')?.value;
            populateRecruiterSelect('detail-recruiter', currentRecruiter);
        }
    }

    // ============================================
    // PROJECT & TICKET MANAGEMENT
    // ============================================
    let projectStages = [];

    function addProjectStage(name = '') {
        projectStages.push(name);
        renderProjectStages();
    }

    function removeProjectStage(index) {
        projectStages.splice(index, 1);
        renderProjectStages();
    }

    function renderProjectStages() {
        const container = document.getElementById('project-stages-container');
        if (!container) return;
        container.innerHTML = '';
        projectStages.forEach((stage, index) => {
            const div = document.createElement('div');
            div.className = 'd-flex align-items-center bg-white p-2 border rounded';
            div.innerHTML = `
                <span class="me-2 text-muted"><i class="fas fa-grip-vertical"></i></span>
                <input type="text" class="form-control form-control-sm border-0 shadow-none" value="${stage}" 
                    placeholder="Tên bước (VD: Sơ vấn)" onchange="projectStages[${index}] = this.value">
                <button type="button" class="btn btn-sm btn-link text-danger ms-auto" onclick="removeProjectStage(${index})">
                    <i class="fas fa-times"></i>
                </button>
            `;
            container.appendChild(div);
        });
    }

    function loadProjects() {
        showLoadingTable('#projects-table tbody');
        google.script.run
            .withSuccessHandler(function (res) {
                console.log('✅ loadProjects success:', res);
                projectsData = (res && res.data) ? res.data : [];
                renderProjects();
            })
            .withFailureHandler(function (err) {
                console.error('❌ loadProjects error:', err);
                const tbody = document.querySelector('#projects-table tbody');
                if (tbody) tbody.innerHTML = '<tr><td colspan="7" class="text-center text-danger py-4">Lỗi tải dữ liệu dự án: ' + err.message + '</td></tr>';
            })
            .apiGetProjects();
    }

    function renderProjects() {
        const tbody = document.querySelector('#projects-table tbody');
        if (!tbody) return;
        tbody.innerHTML = '';

        if (!projectsData || projectsData.length === 0) {
            tbody.innerHTML = '<tr><td colspan="10" class="text-center py-4">Chưa có dự án nào.</td></tr>';
            return;
        }

        projectsData.forEach(p => {
            const id = p['Mã Dự án'];
            const projectTickets = ticketsData.filter(t => t['Mã Dự án'] === id);

            let totalHired = 0;
            let totalRequested = 0;
            let actualCost = 0;
            const hiredStatusList = ['Hired', 'Đã tuyển', 'Nhận việc', 'Đã nhận việc', 'Official'];

            projectTickets.forEach(t => {
                totalRequested += parseInt(t['Số lượng'] || 0);

                // Hired count for this ticket
                totalHired += candidatesData.filter(c => {
                    const tID = getVal(c, 'TicketID');
                    const s = getVal(c, 'Stage');
                    return String(tID).trim() === String(t['Mã Ticket']).trim() && hiredStatusList.includes(s);
                }).length;

                // Cost for this ticket
                try {
                    const costs = JSON.parse(t['Chi phí tuyển dụng'] || '[]');
                    if (Array.isArray(costs)) {
                        actualCost += costs.reduce((sum, item) => sum + (parseFloat(item.cost) || 0), 0);
                    }
                } catch (e) { }
            });

            const budget = parseFloat(p['Ngân sách']) || 0;
            const formatDate = (dStr) => {
                if (!dStr) return '...';
                const d = new Date(dStr);
                return isNaN(d.getTime()) ? dStr : d.toLocaleDateString('vi-VN');
            };

            const tr = document.createElement('tr');
            tr.style.cursor = 'pointer';
            tr.onclick = (e) => {
                if (e.target.closest('button')) return;
                openProjectModal(id);
            };

            tr.innerHTML = `
            <td class="ps-3"><small class="text-muted">${id || ''}</small></td>
            <td class="fw-bold text-primary">${p['Tên Dự án'] || ''}</td>
            <td class="text-center">
                <span class="text-success fw-bold">${totalHired}</span> / 
                <span class="fw-bold">${totalRequested}</span>
            </td>
            <td>
                <span class="fw-bold text-success">${new Intl.NumberFormat('vi-VN').format(actualCost)}</span> / 
                <span class="text-muted small">${new Intl.NumberFormat('vi-VN').format(budget)} đ</span>
            </td>
            <td>${p['Người quản lý'] || ''}</td>
            <td>${formatDate(p['Ngày bắt đầu'])}</td>
            <td>${formatDate(p['Ngày kết thúc'])}</td>
            <td><span class="badge ${p['Trạng thái'] === 'Active' ? 'bg-success' : 'bg-secondary'}">${p['Trạng thái'] || 'Draft'}</span></td>
            <td class="pe-3 text-center">
                <div class="btn-group">
                    <button class="btn btn-sm btn-outline-info" title="Xem chi tiết" onclick="openProjectModal('${id}')"><i class="fas fa-eye"></i></button>
                    ${(currentUser && currentUser.role === 'Admin') ? `
                        <button class="btn btn-sm btn-outline-primary" title="Sửa Dự án" onclick="openProjectModal('${id}')"><i class="fas fa-edit"></i></button>
                    ` : ''}
                </div>
            </td>
        `;
            tbody.appendChild(tr);
        });
    }

    function openProjectModal(id) {
        // Use standard bootstrap trigger if instance not found
        const modalEl = document.getElementById('projectModal');
        const modal = bootstrap.Modal.getOrCreateInstance(modalEl);
        const form = document.getElementById('project-form');
        form.reset();
        projectStages = [];

        // UI containers for stats and list
        const statsContainer = document.getElementById('proj-stats-container');
        const ticketsContainer = document.getElementById('proj-tickets-list-container');
        const tbody = document.getElementById('proj-tickets-tbody');

        if (id) {
            const proj = projectsData.find(p => p['Mã Dự án'] === id);
            if (proj) {
                document.getElementById('proj-code').value = proj['Mã Dự án'];
                document.getElementById('proj-name').value = proj['Tên Dự án'];
                document.getElementById('proj-manager').value = proj['Người quản lý'];
                document.getElementById('proj-start').value = formatDateForInput(proj['Ngày bắt đầu']);
                document.getElementById('proj-end').value = formatDateForInput(proj['Ngày kết thúc']);
                document.getElementById('proj-quota').value = proj['Chỉ tiêu'] || 0;
                document.getElementById('proj-budget').value = proj['Ngân sách'] || 0;

                const workflowStr = proj['Quy trình (Workflow)'] || '';
                projectStages = workflowStr.split(',').map(s => s.trim()).filter(s => s !== '');

                // --- Calculate Project Stats ---
                const projectTickets = ticketsData.filter(t => t['Mã Dự án'] === id);
                let totalTickets = projectTickets.length;
                let totalRequested = 0;
                let totalHired = 0; // New
                let totalProjectCost = 0;

                tbody.innerHTML = '';
                projectTickets.forEach(t => {
                    const ticketCode = t['Mã Ticket'];
                    const requested = parseInt(t['Số lượng'] || 0);
                    totalRequested += requested;

                    // Hired Count for this ticket (Refined to use Stage)
                    const hiredStatusList = ['Hired', 'Đã tuyển', 'Nhận việc', 'Đã nhận việc', 'Official'];
                    const hiredCount = candidatesData.filter(c => {
                        const tID = getVal(c, 'TicketID');
                        const s = getVal(c, 'Stage'); // Refined
                        return String(tID).trim() === String(ticketCode).trim() && hiredStatusList.includes(s);
                    }).length;
                    totalHired += hiredCount;

                    // Cost for this ticket
                    let ticketCost = 0;
                    try {
                        const costs = JSON.parse(t['Chi phí tuyển dụng'] || '[]');
                        if (Array.isArray(costs)) {
                            ticketCost = costs.reduce((sum, item) => sum + (parseFloat(item.cost) || 0), 0);
                        }
                    } catch (e) { }
                    totalProjectCost += ticketCost;

                    // Render Row
                    const deadline = t['Deadline'] || t['Hạn định tuyển dụng'] || '...';
                    const status = t['Trạng thái Phê duyệt'] || 'Pending';
                    const tr = document.createElement('tr');
                    tr.style.cursor = 'pointer';
                    tr.onclick = () => {
                        // Using a small delay to ensure modal transitions smoothly if needed
                        const modalEl = document.getElementById('projectModal');
                        const modal = bootstrap.Modal.getInstance(modalEl);
                        if (modal) modal.hide();
                        setTimeout(() => openTicketModal(ticketCode), 300);
                    };
                    tr.innerHTML = `
                    <td class="fw-bold">${t['Vị trí cần tuyển'] || t['Vị trí'] || 'N/A'}</td>
                    <td class="text-center">${requested}</td>
                    <td class="text-center text-success fw-bold">${hiredCount}</td>
                    <td>${deadline}</td>
                    <td>${new Intl.NumberFormat('vi-VN').format(ticketCost)} đ</td>
                    <td><span class="badge ${status === 'Approved' ? 'bg-success' : (status === 'Rejected' ? 'bg-danger' : 'bg-warning text-dark')}">${status}</span></td>
                `;
                    tbody.appendChild(tr);
                });

                // Update Stats UI
                document.getElementById('stat-proj-tickets').innerText = totalTickets;
                document.getElementById('stat-proj-requested').innerText = totalRequested;
                document.getElementById('stat-proj-hired').innerText = totalHired; // New
                document.getElementById('stat-proj-cost').innerText = new Intl.NumberFormat('vi-VN').format(totalProjectCost) + ' đ';

                statsContainer.style.display = 'block';
                ticketsContainer.style.display = 'block';
            }
        } else {
            // Default stages for new project
            projectStages = ['Ứng tuyển', 'Xét duyệt hồ sơ', 'Sơ vấn', 'Phỏng vấn', 'Phê duyệt nhận việc', 'Mời nhận việc', 'Đã nhận việc', 'Từ chối'];
            statsContainer.style.display = 'none';
            ticketsContainer.style.display = 'none';
        }

        renderProjectStages();
        modal.show();
    }

    function editProject(id) {
        openProjectModal(id);
    }

    function saveProject() {
        const data = {
            code: document.getElementById('proj-code').value,
            name: document.getElementById('proj-name').value,
            workflow: projectStages.join(', '),
            manager: document.getElementById('proj-manager').value,
            startDate: document.getElementById('proj-start').value,
            endDate: document.getElementById('proj-end').value,
            quota: parseFloat(document.getElementById('proj-quota').value) || 0,
            budget: parseFloat(document.getElementById('proj-budget').value) || 0
        };
        console.log('Saving Project - Payload:', data);

        if (!data.name || projectStages.length === 0) {
            Swal.fire('Lỗi', 'Vui lòng nhập tên dự án và ít nhất một bước quy trình.', 'warning');
            return;
        }

        google.script.run.withSuccessHandler(function (res) {
            if (res.success) {
                Swal.fire('Thành công', res.message, 'success');
                bootstrap.Modal.getInstance(document.getElementById('projectModal')).hide();
                loadProjects();
            } else {
                Swal.fire('Lỗi', res.message, 'error');
            }
        }).apiSaveProject(data);
    }

    function loadTickets() {
        showLoadingTable('#tickets-table tbody');
        google.script.run
            .withSuccessHandler(function (res) {
                console.log('✅ loadTickets success:', res);
                ticketsData = (res && res.data) ? res.data : [];
                renderTickets();
            })
            .withFailureHandler(function (err) {
                console.error('❌ loadTickets error:', err);
                const tbody = document.querySelector('#tickets-table tbody');
                if (tbody) tbody.innerHTML = '<tr><td colspan="8" class="text-center text-danger py-4">Lỗi tải dữ liệu phiếu: ' + err.message + '</td></tr>';
            })
            .apiGetTickets();
    }

    function renderTickets() {
        const tbody = document.querySelector('#tickets-table tbody');
        if (!tbody) return;
        tbody.innerHTML = '';

        if (!ticketsData || ticketsData.length === 0) {
            tbody.innerHTML = '<tr><td colspan="9" class="text-center py-4">Chưa có phiếu yêu cầu nào.</td></tr>';
            return;
        }

        ticketsData.forEach(t => {
            const tr = document.createElement('tr');
            tr.style.cursor = 'pointer';
            tr.onclick = (e) => {
                if (e.target.closest('button')) return;
                openTicketModal(t['Mã Ticket']);
            };

            const ticketCode = t['Mã Ticket'];

            // 1. Calculate Hired Count (Based on Stage as requested)
            const hiredCount = candidatesData.filter(c => {
                const tID = getVal(c, 'TicketID');
                const s = getVal(c, 'Stage'); // Refined to use Stage
                const hiredStatusList = ['Hired', 'Đã tuyển', 'Nhận việc', 'Đã nhận việc', 'Official'];
                return String(tID).trim() === String(ticketCode).trim() && hiredStatusList.includes(s);
            }).length;

            // 2. Calculate Total Cost
            let totalCost = 0;
            try {
                const costs = JSON.parse(t['Chi phí tuyển dụng'] || '[]');
                if (Array.isArray(costs)) {
                    totalCost = costs.reduce((sum, item) => sum + (parseFloat(item.cost) || 0), 0);
                }
            } catch (e) { }

            // 3. Project Name - Code
            const project = projectsData.find(p => p['Mã Dự án'] === t['Mã Dự án']);
            const projDisplay = project ? `${project['Tên Dự án']} - ${t['Mã Dự án']}` : (t['Mã Dự án'] || '');

            // 4. Date Formatting (DD/MM/YYYY)
            const formatDate = (dStr) => {
                if (!dStr) return '...';
                const d = new Date(dStr);
                if (isNaN(d.getTime())) return dStr;
                return d.toLocaleDateString('vi-VN');
            };

            let statusClass = 'bg-secondary';
            const status = t['Trạng thái Phê duyệt'] || 'Pending';
            if (status === 'Approved') statusClass = 'bg-success';
            if (status === 'Rejected') statusClass = 'bg-danger';
            if (status === 'Pending') statusClass = 'bg-warning text-dark';

            tr.innerHTML = `
            <td class="ps-3"><small class="text-muted">${ticketCode || ''}</small></td>
            <td class="fw-bold">${t['Vị trí cần tuyển'] || t['Vị trí'] || ''}</td>
            <td class="text-center">
                <span class="text-success fw-bold">${hiredCount}</span> / 
                <span class="fw-bold">${t['Số lượng'] || 1}</span>
            </td>
            <td class="fw-bold text-success">${new Intl.NumberFormat('vi-VN').format(totalCost)} đ</td>
                <td><span class="badge bg-light text-dark border">${projDisplay}</span></td>
                <td>${formatDate(t['Ngày bắt đầu'])}</td>
                <td><span class="text-danger fw-bold">${formatDate(t['Deadline'] || t['Hạn định tuyển dụng'])}</span></td>
                <td><span class="badge ${statusClass}">${status}</span></td>
                <td class="pe-3 text-center">
                    <div class="btn-group">
                        <button class="btn btn-sm btn-outline-info" title="Xem chi tiết" onclick="openTicketModal('${ticketCode}')"><i class="fas fa-eye"></i></button>
                        ${(currentUser && currentUser.role === 'Admin' && status === 'Pending') ? `
                            <button class="btn btn-sm btn-success" title="Duyệt Ticket" onclick="openTicketModal('${ticketCode}')"><i class="fas fa-check"></i></button>
                        ` : ''}
                    </div>
                </td>
            `;
            tbody.appendChild(tr);
        });
    }

    function openTicketModal(ticketId = null) {
        const modal = bootstrap.Modal.getOrCreateInstance(document.getElementById('ticketModal'));
        document.getElementById('ticket-form').reset();
        document.getElementById('ticket-stats-container').style.display = 'none';
        document.getElementById('admin-approval-section').style.display = 'none';
        document.getElementById('tick-costs-container').innerHTML = '';

        // Populate Projects Dropdown
        const projSelect = document.getElementById('tick-project');
        projSelect.innerHTML = '<option value="">-- Chọn Dự án --</option>';
        projectsData.forEach(p => {
            const opt = document.createElement('option');
            opt.value = p['Mã Dự án'];
            opt.textContent = p['Tên Dự án'];
            projSelect.appendChild(opt);
        });

        // Populate Dept Dropdown
        const deptSelect = document.getElementById('tick-department');
        deptSelect.innerHTML = '<option value="">-- Chọn Phòng ban --</option>';
        departmentsData.forEach(d => {
            const opt = document.createElement('option');
            opt.value = d.name;
            opt.textContent = d.name;
            deptSelect.appendChild(opt);
        });

        const posSelect = document.getElementById('tick-position');
        posSelect.innerHTML = '<option value="">-- Chọn Vị trí --</option>';

        deptSelect.onchange = () => {
            posSelect.innerHTML = '<option value="">-- Chọn Vị trí --</option>';
            const d = departmentsData.find(x => x.name === deptSelect.value);
            if (d && d.positions) {
                d.positions.forEach(p => {
                    const opt = document.createElement('option');
                    opt.value = p;
                    opt.textContent = p;
                    posSelect.appendChild(opt);
                });
            }
        };

        // Populate Offices
        const officeSelect = document.getElementById('tick-office');
        officeSelect.innerHTML = '<option value="">-- Chọn văn phòng --</option>';
        if (initialData && initialData.companyInfo && initialData.companyInfo.addresses) {
            initialData.companyInfo.addresses.forEach(addr => {
                const opt = document.createElement('option');
                const addrLabel = typeof addr === 'string' ? addr : (addr.city || addr.detail || '');
                opt.value = addrLabel;
                opt.textContent = addrLabel;
                officeSelect.appendChild(opt);
            });
        }

        // Populate Admin Recruiter Select
        const adminRecSelect = document.getElementById('tick-admin-recruiter');
        if (adminRecSelect) {
            adminRecSelect.innerHTML = '<option value="">-- Chọn Recruiter --</option>';
            recruitersData.forEach(r => {
                const opt = document.createElement('option');
                opt.value = r.email || r.Username;
                opt.textContent = r.name || r.Full_Name || r.Username;
                adminRecSelect.appendChild(opt);
            });
        }

        if (ticketId) {
            document.getElementById('ticketModal').setAttribute('data-current-code', ticketId);
            const t = ticketsData.find(x => String(x['Mã Ticket'] || '').trim() === String(ticketId).trim());
            if (t) {
                document.getElementById('tick-project').value = t['Mã Dự án'] || '';
                document.getElementById('tick-quantity').value = t['Số lượng'] || 1;
                const deptVal = String(t['Phòng ban'] || '').trim();
                document.getElementById('tick-department').value = deptVal;

                // Trigger dept change for positions
                const d = departmentsData.find(x => String(x.name || '').trim() === deptVal);
                if (d && d.positions) {
                    posSelect.innerHTML = '<option value="">-- Chọn Vị trí --</option>';
                    d.positions.forEach(p => {
                        const opt = document.createElement('option');
                        opt.value = p;
                        opt.textContent = p;
                        const savedPos = String(t['Vị trí cần tuyển'] || t['Vị trí'] || '').trim();
                        if (String(p).trim() === savedPos) opt.selected = true;
                        posSelect.appendChild(opt);
                    });
                }

                document.getElementById('tick-start').value = formatDateForInput(getVal(t, 'Ngày bắt đầu'));
                document.getElementById('tick-deadline').value = formatDateForInput(getVal(t, 'Deadline') || getVal(t, 'Hạn định tuyển dụng') || getVal(t, 'Hạn định tuyển'));
                document.getElementById('tick-work-type').value = getVal(t, 'Loại hình làm việc') || 'Fulltime';
                document.getElementById('tick-education').value = getVal(t, 'Học vấn') || getVal(t, 'Học Vấn') || 'Không yêu cầu';
                document.getElementById('tick-gender').value = getVal(t, 'Giới tính') || 'Không yêu cầu';
                document.getElementById('tick-age').value = getVal(t, 'Tuổi') || '';
                document.getElementById('tick-major').value = getVal(t, 'Chuyên môn') || '';
                document.getElementById('tick-experience').value = getVal(t, 'Kinh nghiệm') || 'Không yêu cầu';
                document.getElementById('tick-office').value = getVal(t, 'Văn phòng làm việc') || getVal(t, 'Văn phòng') || '';
                document.getElementById('tick-manager').value = getVal(t, 'Quản lý trực tiếp') || '';

                // SHOW STATS
                showTicketStats(ticketId, t);

                // Role-based visibility and buttons
                const status = String(t['Trạng thái Phê duyệt'] || 'Pending').trim();
                const submitBtn = document.getElementById('btn-save-ticket');
                const proposeBtn = document.getElementById('btn-propose-edit');

                const userRole = String(currentUser ? (currentUser.role || currentUser.Role || '') : '').trim();
                if (userRole === 'Admin') {
                    // Admin always sees everything and can Save
                    document.getElementById('admin-approval-section').style.display = 'block';
                    document.getElementById('tick-admin-recruiter').value = String(t['Recruiter phụ trách'] || '').trim();
                    submitBtn.style.display = 'inline-block';
                    submitBtn.innerText = 'Lưu thay đổi';
                    if (proposeBtn) proposeBtn.style.display = 'none';

                    // Load existing costs
                    try {
                        const savedCosts = JSON.parse(t['Chi phí tuyển dụng'] || '[]');
                        if (Array.isArray(savedCosts)) {
                            savedCosts.forEach(c => addCostRow(c.name, c.cost));
                        }
                    } catch (e) { }
                } else if (userRole === 'Manager') {
                    if (status === 'Approved') {
                        // Manager on Approved ticket -> Propose Edit
                        submitBtn.style.display = 'none';
                        if (proposeBtn) {
                            proposeBtn.style.display = 'inline-block';
                            proposeBtn.innerText = 'Đề xuất chỉnh sửa';
                        }
                        document.getElementById('admin-approval-section').style.display = 'none';
                    } else {
                        // Manager on Pending/Rejected -> Normal Save
                        submitBtn.style.display = 'inline-block';
                        submitBtn.innerText = 'Gửi Yêu cầu';
                        if (proposeBtn) proposeBtn.style.display = 'none';
                        document.getElementById('admin-approval-section').style.display = 'none';
                    }
                }
            }
        } else {
            // NEW TICKET
            document.getElementById('ticketModal').removeAttribute('data-current-code');
            const submitBtn = document.getElementById('btn-save-ticket');
            const proposeBtn = document.getElementById('btn-propose-edit');
            submitBtn.style.display = 'inline-block';
            submitBtn.innerText = 'Gửi Yêu cầu';
            if (proposeBtn) proposeBtn.style.display = 'none';

            const userRole = String(currentUser ? (currentUser.role || currentUser.Role || '') : '').trim();
            if (userRole === 'Admin') {
                document.getElementById('admin-approval-section').style.display = 'block';
            } else {
                document.getElementById('admin-approval-section').style.display = 'none';
            }
        }

        modal.show();
    }

    function showTicketStats(ticketCode, t) {
        if (!t) t = ticketsData.find(x => x['Mã Ticket'] === ticketCode);
        if (!t) return;

        document.getElementById('ticket-stats-container').style.display = 'block';

        const ticketCandidates = candidatesData.filter(c => String(getVal(c, 'TicketID')).trim() === String(ticketCode).trim());
        const hiredStatusList = ['Hired', 'Đã tuyển', 'Nhận việc', 'Đã nhận việc', 'Official'];
        const requested = parseInt(t['Số lượng'] || 0);
        const hiredCandidates = ticketCandidates.filter(c => hiredStatusList.includes(getVal(c, 'Stage')));

        // --- 1. Efficiency Calculation (%) ---
        const ticketDeadline = new Date(t['Deadline'] || t['Hạn định tuyển dụng']);
        let onTimeCount = 0;

        hiredCandidates.forEach(c => {
            const hireDateStr = getVal(c, 'Hire_Date');
            if (hireDateStr) {
                const hDate = new Date(hireDateStr);
                if (!isNaN(hDate.getTime()) && !isNaN(ticketDeadline.getTime())) {
                    if (hDate <= ticketDeadline) onTimeCount++;
                } else { onTimeCount++; }
            } else { onTimeCount++; }
        });

        const efficiencyPerc = requested > 0 ? ((onTimeCount / requested) * 100).toFixed(1) : 0;

        // --- 2. Update Stats UI ---
        document.getElementById('stat-hired-ratio').innerText = `${hiredCandidates.length}/${requested}`;

        let totalCost = 0;
        try {
            const costs = JSON.parse(t['Chi phí tuyển dụng'] || '[]');
            if (Array.isArray(costs)) {
                totalCost = costs.reduce((sum, item) => sum + (parseFloat(item.cost) || 0), 0);
            }
        } catch (e) { }
        document.getElementById('stat-total-costs').innerText = new Intl.NumberFormat('vi-VN').format(totalCost) + ' đ';

        const dStr = t['Deadline'] || t['Hạn định tuyển dụng'];
        if (dStr) {
            const dDate = new Date(dStr);
            const now = new Date();
            const days = Math.ceil((dDate - now) / (1000 * 60 * 60 * 24));
            document.getElementById('stat-days-left').innerText = days;
        } else { document.getElementById('stat-days-left').innerText = '--'; }

        document.getElementById('stat-efficiency-perc').innerText = efficiencyPerc + '%';

        // --- 3. Hired Candidates Table ---
        const hiredTbody = document.querySelector('#hired-candidates-table tbody');
        hiredTbody.innerHTML = '';

        if (hiredCandidates.length === 0) {
            hiredTbody.innerHTML = '<tr><td colspan="5" class="text-center py-2 text-muted italic">Chưa có ứng viên nào nhận việc</td></tr>';
        } else {
            hiredCandidates.forEach((c, index) => {
                const hireDateStr = getVal(c, 'Hire_Date') || '...';
                const hDate = new Date(hireDateStr);
                let statusBadge = '<span class="badge bg-success">Đúng hạn</span>';
                if (!isNaN(hDate.getTime()) && !isNaN(ticketDeadline.getTime()) && hDate > ticketDeadline) {
                    statusBadge = '<span class="badge bg-danger">Trễ hạn</span>';
                }

                const tr = document.createElement('tr');
                tr.innerHTML = `
                <td class="ps-3">${index + 1}</td>
                <td><div class="fw-bold">${getVal(c, 'Name') || 'N/A'}</div></td>
                <td>${hireDateStr}</td>
                <td>${statusBadge}</td>
                <td class="pe-3 text-center">
                    <div class="btn-group">
                        <button class="btn btn-xs btn-outline-primary" title="Hồ sơ" onclick="openCandidateDetail('${getVal(c, 'ID')}')"><i class="fas fa-user"></i></button>
                        <button class="btn btn-xs btn-outline-info" title="Xem CV" onclick="viewCandidateCV('${getVal(c, 'CV_Link')}')"><i class="fas fa-file-pdf"></i></button>
                        <button class="btn btn-xs btn-outline-secondary" title="Đánh giá" onclick="viewEvaluationDetail('${getVal(c, 'ID')}')"><i class="fas fa-clipboard-check"></i></button>
                    </div>
                </td>
            `;
                hiredTbody.appendChild(tr);
            });
        }

        // --- 4. Recruiter Performance (New Breakdown) ---
        const recPerformanceCont = document.getElementById('stat-recruiter-performance');
        recPerformanceCont.innerHTML = '';

        const recruiterStats = {};
        ticketCandidates.forEach(c => {
            const s = getVal(c, 'Stage');
            const rec = getVal(c, 'Recruiter') || 'Unknown';
            if (!recruiterStats[rec]) recruiterStats[rec] = { hired: 0, ontime: 0 };

            if (hiredStatusList.includes(s)) {
                recruiterStats[rec].hired++;
                const hireDateStr = getVal(c, 'Hire_Date');
                if (hireDateStr) {
                    const hDate = new Date(hireDateStr);
                    if (!isNaN(hDate.getTime()) && !isNaN(ticketDeadline.getTime())) {
                        if (hDate <= ticketDeadline) recruiterStats[rec].ontime++;
                    } else { recruiterStats[rec].ontime++; }
                } else { recruiterStats[rec].ontime++; }
            }
        });

        Object.keys(recruiterStats).forEach(recName => {
            const stats = recruiterStats[recName];
            if (stats.hired === 0) return;
            const percHired = requested > 0 ? ((stats.hired / requested) * 100).toFixed(0) : 0;
            const percOntime = stats.hired > 0 ? ((stats.ontime / stats.hired) * 100).toFixed(0) : 0;

            const div = document.createElement('div');
            div.className = 'd-flex justify-content-between align-items-center bg-white p-2 rounded border-start border-4 border-primary shadow-xs';
            div.style.fontSize = '0.8rem';
            div.innerHTML = `
            <span class="fw-bold text-dark"><i class="fas fa-user-tie me-1"></i> ${recName}</span>
            <div class="text-end">
                <span class="badge bg-light text-primary me-1 border">${stats.hired} / ${requested} (${percHired}%)</span>
                <span class="badge bg-light text-success border">${percOntime}% Đúng hạn</span>
            </div>
        `;
            recPerformanceCont.appendChild(div);
        });

        if (Object.keys(recruiterStats).length === 0) {
            recPerformanceCont.innerHTML = '<div class="text-center py-2 text-muted small">Chưa có kết quả từ Recruiter</div>';
        }
    }

    function addCostRow(name = '', cost = '') {
        const container = document.getElementById('tick-costs-container');
        if (!container) return;
        const div = document.createElement('div');
        div.className = 'input-group input-group-sm mb-1';
        div.innerHTML = `
            <input type="text" class="form-control" placeholder="Tên chi phí..." value="${name}">
            <input type="number" class="form-control" placeholder="Số tiền..." value="${cost}">
            <button class="btn btn-outline-danger" onclick="this.parentElement.remove()"><i class="fas fa-times"></i></button>
        `;
        container.appendChild(div);
    }

    function approveTicketAction(status) {
        const code = document.getElementById('ticketModal').getAttribute('data-current-code');
        if (!code) return;

        const recruiterEmail = document.getElementById('tick-admin-recruiter').value;
        const costRows = document.querySelectorAll('#tick-costs-container .input-group');
        const costs = Array.from(costRows).map(row => {
            const inputs = row.querySelectorAll('input');
            return {
                name: inputs[0].value.trim(),
                cost: inputs[1].value.trim()
            };
        }).filter(c => c.name);

        const approvalData = {
            status: status,
            recruiterEmail: recruiterEmail,
            costs: costs
        };

        Swal.fire({
            title: 'Xác nhận xử lý Ticket?',
            text: `Bạn đang thực hiện ${status === 'Approved' ? 'Phê duyệt' : 'Từ chối'} ticket này.`,
            icon: 'question',
            showCancelButton: true
        }).then(res => {
            if (res.isConfirmed) {
                google.script.run.withSuccessHandler(function (response) {
                    if (response.success) {
                        Swal.fire('Thành công', response.message, 'success');
                        bootstrap.Modal.getInstance(document.getElementById('ticketModal')).hide();
                        loadDashboardData();
                    } else {
                        Swal.fire('Lỗi', response.message, 'error');
                    }
                }).apiApproveTicket(code, approvalData, currentUser.username);
            }
        });
    }

    function saveTicket() {
        const form = document.getElementById('ticket-form');
        if (!form.checkValidity()) {
            form.reportValidity();
            return;
        }

        const currentCode = document.getElementById('ticketModal').getAttribute('data-current-code');
        const t = currentCode ? ticketsData.find(x => x['Mã Ticket'] === currentCode) : null;

        const ticketData = {
            code: currentCode || '',
            projectCode: document.getElementById('tick-project').value,
            quantity: parseInt(document.getElementById('tick-quantity').value),
            department: document.getElementById('tick-department').value,
            position: document.getElementById('tick-position').value,
            startDate: document.getElementById('tick-start').value,
            deadline: document.getElementById('tick-deadline').value,
            workType: document.getElementById('tick-work-type').value,
            education: document.getElementById('tick-education').value,
            gender: document.getElementById('tick-gender').value,
            age: document.getElementById('tick-age').value,
            major: document.getElementById('tick-major').value,
            experience: document.getElementById('tick-experience').value,
            office: document.getElementById('tick-office').value,
            directManager: document.getElementById('tick-manager').value,
            approvalStatus: t ? t['Trạng thái Phê duyệt'] : 'Pending'
        };

        // For Admin: Collect Recruiter and Costs
        if (currentUser && currentUser.role === 'Admin') {
            ticketData.recruiterEmail = document.getElementById('tick-admin-recruiter').value;
            const costRows = document.querySelectorAll('#tick-costs-container .input-group');
            ticketData.costs = Array.from(costRows).map(row => {
                const inputs = row.querySelectorAll('input');
                return {
                    name: inputs[0].value.trim(),
                    cost: inputs[1].value.trim()
                };
            }).filter(c => c.name);
        }

        const btn = document.activeElement && document.activeElement.id.includes('btn') ? document.activeElement : document.getElementById('btn-save-ticket');
        const originalText = btn.innerHTML;
        btn.innerHTML = '<span class="spinner-border spinner-border-sm me-1"></span> Đang lưu...';
        btn.disabled = true;

        google.script.run.withSuccessHandler(function (res) {
            btn.innerHTML = originalText;
            btn.disabled = false;
            if (res.success) {
                Swal.fire('Thành công', res.message, 'success');
                bootstrap.Modal.getInstance(document.getElementById('ticketModal')).hide();
                loadDashboardData();
            } else {
                Swal.fire('Lỗi', res.message, 'error');
            }
        })
            .withFailureHandler(function (err) {
                btn.innerHTML = originalText;
                btn.disabled = false;
                Swal.fire('Lỗi hệ thống', err.toString(), 'error');
            })
            .apiSaveTicket(ticketData, currentUser);
    }

    function showLoadingTable(selector) {
        const tbody = document.querySelector(selector);
        if (tbody) {
            const table = tbody.closest('table');
            const colCount = table ? table.querySelectorAll('thead th').length : 5;
            tbody.innerHTML = `<tr><td colspan="${colCount}" class="text-center py-4"><div class="spinner-border text-primary" role="status"></div><br><small class="text-muted">Đang tải dữ liệu...</small></td></tr>`;
        }
    }


    // Apply filters to candidate data
    function applyKanbanFilters() {
        let filtered = [...candidatesData];

        // Filter by search text (name, email, phone)
        if (kanbanFilters.search) {
            const searchLower = kanbanFilters.search.toLowerCase();
            filtered = filtered.filter(c => {
                const name = (c.Name || '').toLowerCase();
                const email = (c.Email || '').toLowerCase();
                const phone = (c.Phone || '').toLowerCase();
                return name.includes(searchLower) || email.includes(searchLower) || phone.includes(searchLower);
            });
        }

        // Filter by department
        if (kanbanFilters.department) {
            filtered = filtered.filter(c => c.Department === kanbanFilters.department);
        }

        // Filter by position
        if (kanbanFilters.position) {
            filtered = filtered.filter(c => c.Position === kanbanFilters.position);
        }

        // Filter by recruiter
        if (kanbanFilters.recruiter) {
            filtered = filtered.filter(c => c.Recruiter === kanbanFilters.recruiter);
        }

        // Filter by date range
        if (kanbanFilters.dateFrom) {
            const fromDate = new Date(kanbanFilters.dateFrom);
            filtered = filtered.filter(c => {
                if (!c.Applied_Date) return false;
                const appliedDate = new Date(c.Applied_Date);
                return appliedDate >= fromDate;
            });
        }

        if (kanbanFilters.dateTo) {
            const toDate = new Date(kanbanFilters.dateTo);
            toDate.setHours(23, 59, 59); // End of day
            filtered = filtered.filter(c => {
                if (!c.Applied_Date) return false;
                const appliedDate = new Date(c.Applied_Date);
                return appliedDate <= toDate;
            });
        }

        return filtered;
    }

    // Attach filter event listeners
    function attachFilterEventListeners() {
        const searchInput = document.getElementById('kanban-search');
        const deptFilter = document.getElementById('filter-department');
        const posFilter = document.getElementById('filter-position');
        const recruiterFilter = document.getElementById('filter-recruiter');
        const dateFromFilter = document.getElementById('filter-date-from');
        const dateToFilter = document.getElementById('filter-date-to');

        if (searchInput) {
            searchInput.addEventListener('input', function () {
                kanbanFilters.search = this.value;
                renderKanbanBoardWithFilters();
            });
        }

        if (deptFilter) {
            deptFilter.addEventListener('change', function () {
                kanbanFilters.department = this.value;
                renderKanbanBoardWithFilters();
            });
        }

        if (posFilter) {
            posFilter.addEventListener('change', function () {
                kanbanFilters.position = this.value;
                renderKanbanBoardWithFilters();
            });
        }

        if (recruiterFilter) {
            recruiterFilter.addEventListener('change', function () {
                kanbanFilters.recruiter = this.value;
                renderKanbanBoardWithFilters();
            });
        }

        if (dateFromFilter) {
            dateFromFilter.addEventListener('change', function () {
                kanbanFilters.dateFrom = this.value;
                renderKanbanBoardWithFilters();
            });
        }

        if (dateToFilter) {
            dateToFilter.addEventListener('change', function () {
                kanbanFilters.dateTo = this.value;
                renderKanbanBoardWithFilters();
            });
        }
    }

    // Render Kanban board with filters applied
    function renderKanbanBoardWithFilters() {
        const originalData = candidatesData;
        candidatesData = applyKanbanFilters();
        renderKanbanBoard();
        candidatesData = originalData; // Restore original data
    }

    // Populate position dropdown for Add Candidate modal

    // Open Add Candidate modal and populate department dropdown

    // Initialize filters when Kanban section loads
    document.addEventListener('DOMContentLoaded', function () {
        attachFilterEventListeners();

        // Populate filters when switching to Kanban section
        const kanbanNavLink = document.querySelector('[href="#kanban"]');
        if (kanbanNavLink) {
            kanbanNavLink.addEventListener('click', function () {
                setTimeout(() => {
                    populateFilterDropdowns();
                }, 300);
            });
        }

        // Prepare Add Candidate modal when opened
        const addCandidateModal = document.getElementById('addCandidateModal');
        if (addCandidateModal) {
            addCandidateModal.addEventListener('show.bs.modal', function () {
                prepareAddCandidateModal();
            });
        }
    });
    // Check Duplicate Candidate
    function checkDuplicateCandidate() {
        const phone = document.getElementById('add-phone').value;
        const email = document.getElementById('add-email').value;
        const warningMsg = document.getElementById('duplicate-warning-msg');
        const container = document.getElementById('duplicate-warning-container');

        if (!container || !warningMsg) return;

        if (!phone && !email) {
            container.style.display = 'none';
            return;
        }

        google.script.run.withSuccessHandler(function (res) {
            if (res.success && res.found) {
                warningMsg.innerHTML = `<strong>⚠️ Cảnh báo trùng lặp (${res.matchType}):</strong> Ứng viên <b>${res.name}</b> đã nộp hồ sơ ngày <b>${res.date}</b> cho vị trí <b>${res.position}</b>. <a href="${res.link}" target="_blank">Xem hồ sơ</a>`;
                container.style.display = 'block';
            } else {
                container.style.display = 'none';
            }
        }).apiCheckDuplicateCandidate(phone, email);
    }
    // NEWS MANAGEMENT
    function loadNews() {
        google.script.run.withSuccessHandler(function (data) {
            newsData = data;
            renderNews(data);
        }).apiGetNews();
    }

    function renderNews(data) {
        const tbody = document.querySelector('#news-table tbody');
        if (!tbody) return;
        tbody.innerHTML = '';
        data.forEach(item => {
            const dateStr = item.Date ? new Date(item.Date).toLocaleDateString() : '';
            // Get first image
            const imgs = item.Image ? item.Image.split(/[\n,;]+/).map(s => s.trim()).filter(s => s) : [];
            const firstImg = imgs.length > 0 ? imgs[0] : '';

            tbody.innerHTML += `
                <tr>
                    <td><small>${item.ID}</small></td>
                    <td><img src="${firstImg}" alt="img" style="height:40px; width:40px; object-fit:cover;" onerror="this.src='https://via.placeholder.com/40'"></td>
                    <td>${item.Title}</td>
                    <td>${dateStr}</td>
                    <td><span class="badge ${item.Status === 'Published' ? 'bg-success' : 'bg-secondary'}">${item.Status}</span></td>
                    <td class="text-center">
                        <button class="btn btn-sm btn-outline-primary" onclick="openNewsModal('${item.ID}')">Sửa</button>
                        <button class="btn btn-sm btn-outline-danger" onclick="deleteNews('${item.ID}')">Xóa</button>
                    </td>
                </tr>
            `;
        });
    }

    function openNewsModal(id) {
        const modal = new bootstrap.Modal(document.getElementById('newsModal'));
        if (id) {
            const item = newsData.find(x => x.ID == id);
            if (!item) return;
            document.getElementById('news-id').value = item.ID;
            document.getElementById('news-title').value = item.Title;
            document.getElementById('news-image').value = item.Image;
            document.getElementById('news-content').value = item.Content;
            document.getElementById('news-status').value = item.Status;
            document.getElementById('newsModalLabel').innerText = 'Cập nhật Tin Tức';
        } else {
            document.getElementById('news-form').reset();
            document.getElementById('news-id').value = '';
            document.getElementById('news-image').value = '';
            document.getElementById('newsModalLabel').innerText = 'Viết Bài Mới';
        }
        // Reset File Input
        const fileInput = document.getElementById('news-image-db');
        if (fileInput) fileInput.value = '';

        // Refresh Preview
        previewNewsFlagNodes();

        modal.show();
    }

    // IMAGE PREVIEW
    function previewNewsFlagNodes() {
        const fileInput = document.getElementById('news-image-db');
        const previewContainer = document.getElementById('news-image-preview');
        const hiddenInput = document.getElementById('news-image');

        // Clear previous previews of NEW files (keep existing URLs in hidden input?)
        // For simplicity, we'll rebuild preview from existing hidden input + new files
        previewContainer.innerHTML = '';

        // 1. Show existing URLs
        const existingUrls = hiddenInput.value ? hiddenInput.value.split(/[\n,;]+/).filter(s => s.trim()) : [];
        existingUrls.forEach(url => {
            const div = document.createElement('div');
            div.className = 'position-relative';
            div.innerHTML = `
                <img src="${url}" style="width: 80px; height: 80px; object-fit: cover; border-radius: 4px; border: 1px solid #ddd;">
                <button type="button" class="btn btn-sm btn-danger position-absolute top-0 end-0 p-0" style="width: 20px; height: 20px; line-height: 1;" onclick="removeImage('${url}')">&times;</button>
             `;
            previewContainer.appendChild(div);
        });

        // 2. Show selected files
        const files = fileInput.files;
        if (files) {
            Array.from(files).forEach(file => {
                const reader = new FileReader();
                reader.onload = function (e) {
                    const div = document.createElement('div');
                    div.className = 'position-relative';
                    div.innerHTML = `
                        <img src="${e.target.result}" style="width: 80px; height: 80px; object-fit: cover; border-radius: 4px; border: 1px dashed #aaa;">
                         <small class="d-block text-center" style="font-size: 10px; width: 80px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">New</small>
                    `;
                    previewContainer.appendChild(div);
                };
                reader.readAsDataURL(file);
            });
        }
    }

    function removeImage(urlToRemove) {
        const hiddenInput = document.getElementById('news-image');
        let urls = hiddenInput.value ? hiddenInput.value.split(/[\n,;]+/).filter(s => s.trim()) : [];
        urls = urls.filter(u => u !== urlToRemove);
        hiddenInput.value = urls.join('\n');
        previewNewsFlagNodes();
    }

    function saveNews() {
        // 1. Get Form Data
        const id = document.getElementById('news-id').value;
        const title = document.getElementById('news-title').value;
        const currentImages = document.getElementById('news-image').value;
        const content = document.getElementById('news-content').value;
        const status = document.getElementById('news-status').value;
        const fileInput = document.getElementById('news-image-db');

        if (!title) {
            Swal.fire('Lỗi', 'Tiêu đề là bắt buộc', 'warning');
            return;
        }

        const btn = document.querySelector('#newsModal .btn-primary');
        const originalText = btn.innerHTML;
        btn.innerHTML = 'Đang tải ảnh & lưu...';
        btn.disabled = true;

        // 2. Process Files
        const files = fileInput.files;
        const promises = [];

        if (files && files.length > 0) {
            Array.from(files).forEach(file => {
                promises.push(new Promise((resolve, reject) => {
                    const reader = new FileReader();
                    reader.onload = (e) => resolve({
                        name: file.name,
                        type: file.type,
                        data: e.target.result.split(',')[1] // Base64 content
                    });
                    reader.onerror = reject;
                    reader.readAsDataURL(file);
                }));
            });
        }

        Promise.all(promises).then(fileDataList => {
            const payload = {
                id: id,
                title: title,
                content: content,
                status: status,
                currentImages: currentImages, // Existing URLs
                newFiles: fileDataList        // New Files to upload
            };

            google.script.run.withSuccessHandler(function (res) {
                btn.innerHTML = originalText;
                btn.disabled = false;
                if (res.success) {
                    Swal.fire('Thành công', 'Đã lưu bài viết', 'success');
                    const modal = bootstrap.Modal.getInstance(document.getElementById('newsModal'));
                    if (modal) modal.hide();
                    loadNews();
                } else {
                    Swal.fire('Lỗi', res.message, 'error');
                }
            }).apiSaveNews(payload);

        }).catch(err => {
            console.error(err);
            btn.innerHTML = originalText;
            btn.disabled = false;
            Swal.fire('Lỗi', 'Không thể đọc file: ' + err, 'error');
        });
    }

    function deleteNews(id) {
        Swal.fire({
            title: 'Xóa bài viết?',
            text: "Không thể hoàn tác!",
            icon: 'warning',
            showCancelButton: true,
            confirmButtonColor: '#d33',
            confirmButtonText: 'Xóa'
        }).then((result) => {
            if (result.isConfirmed) {
                google.script.run.withSuccessHandler(function (res) {
                    if (res.success) {
                        Swal.fire('Đã xóa', '', 'success');
                        loadNews();
                    } else {
                        Swal.fire('Lỗi', res.message, 'error');
                    }
                }).apiDeleteNews(id);
            }
        });
    }

    // ============================================
    // INTERVIEW EVALUATION LOGIC
    // ============================================

    // 1. Recruiter: Open Request Modal
    function openRequestEvaluationModal() {
        const candidateId = document.getElementById('current-candidate-id').value;
        const candidateName = document.getElementById('detail-name').value;

        if (!candidateId) {
            Swal.fire('Lỗi', 'Không xác định được ứng viên', 'error');
            return;
        }

        document.getElementById('req-eval-candidate-id').value = candidateId;
        document.getElementById('req-eval-candidate-name').innerText = candidateName;

        // Populate Manager Select
        const select = document.getElementById('req-eval-manager');
        select.innerHTML = '<option value="">Chọn Manager</option>';

        // Filter users with role 'Manager' or 'Admin'
        // NOTE: usersData comes from apiGetTableData('Users'), so keys match Sheet Headers (Capitalized)
        const managers = usersData.filter(u => (u.Role === 'Manager' || u.Role === 'Admin') && u.Email);

        if (managers.length === 0) {
            console.warn('No managers found in usersData', usersData);
        }

        managers.forEach(m => {
            const option = document.createElement('option');
            option.value = m.Email; // Use email for notification
            // Use Full_Name if available, else Name, else Username
            const dName = m.Full_Name || m.Name || m.Username;
            option.text = `${dName} (${m.Role})`;
            select.appendChild(option);
        });

        // Hide Detail Modal? No, keep it open or stack. 
        // Bootstrap modals stack fine if configured, but let's hide Detail for clarity or keep it.
        // Let's keep it (z-index might handle it).
        new bootstrap.Modal(document.getElementById('requestEvaluationModal')).show();
    }

    // 2. Recruiter: Submit Request
    function submitEvaluationRequest() {
        const candidateId = document.getElementById('req-eval-candidate-id').value;
        const managerEmail = document.getElementById('req-eval-manager').value;

        if (!managerEmail) {
            Swal.fire('Lỗi', 'Vui lòng chọn người đánh giá', 'warning');
            return;
        }

        const btn = document.querySelector('#requestEvaluationModal .btn-primary');
        const originalText = btn.innerText;
        btn.innerText = 'Đang gửi...';
        btn.disabled = true;

        google.script.run.withSuccessHandler(function (res) {
            btn.innerText = originalText;
            btn.disabled = false;
            if (res.success) {
                Swal.fire('Thành công', 'Đã gửi yêu cầu đánh giá', 'success');
                bootstrap.Modal.getInstance(document.getElementById('requestEvaluationModal')).hide();
            } else {
                Swal.fire('Lỗi', res.message, 'error');
            }
        }).apiCreateEvaluationRequest(candidateId, managerEmail, currentUser.email);
    }

    // 3. Manager: Check Pending Evaluations
    function checkForPendingEvaluations() {
        if (!currentUser) return;

        console.log('Checking pending evaluations for:', currentUser.email);
        google.script.run.withSuccessHandler(function (list) {
            // Save for modal usage form Pending List
            if (!window.currentEvaluationList) window.currentEvaluationList = [];
            // Merge or set? Valid question. 
            // If we are on "Evaluations" page, currentEvaluationList might be the full list.
            // If we are on Dashboard, it might be empty.
            // Let's simplified: We append pending ones if they are not there?
            // Or just store pending ones in a separate var?
            // Simpler: Just ensure they are in currentList.
            list.forEach(item => {
                if (!window.currentEvaluationList.find(x => x.ID === item.ID)) {
                    window.currentEvaluationList.push(item);
                }
            });

            console.log('Pending evaluations:', list);
            const badge = document.getElementById('nav-eval-badge');
            const container = document.getElementById('dashboard-eval-container');

            if (list.length > 0) {
                if (badge) {
                    badge.innerText = list.length;
                    badge.style.display = 'inline-block';
                }

                if (container) {
                    let html = `
                   <div class="card shadow-sm mb-4 border-warning">
                       <div class="card-header bg-warning text-dark fw-bold">
                           <i class="fas fa-clipboard-list me-2"></i> Đánh giá cần thực hiện (${list.length})
                       </div>
                       <div class="list-group list-group-flush">
                   `;

                    list.forEach(evalItem => {
                        html += `
                        <a href="#" class="list-group-item list-group-item-action d-flex justify-content-between align-items-center" onclick="openEvaluationForm('${evalItem.ID}')">
                            <div>
                                <strong>${evalItem.Candidate_Name}</strong> - <span class="text-muted">${evalItem.Position}</span>
                                <div class="small text-muted">Phòng ban: ${evalItem.Department} | Gửi lúc: ${new Date(evalItem.Created_At).toLocaleString('vi-VN')}</div>
                            </div>
                            <button class="btn btn-sm btn-primary">Chấm điểm</button>
                        </a>
                       `;
                    });

                    html += `</div></div>`;
                    container.innerHTML = html;
                    container.style.display = 'block';
                }
            } else {
                if (badge) badge.style.display = 'none';
                if (container) container.style.display = 'none';
            }
        }).apiGetPendingEvaluations(currentUser.email || currentUser.username);
    }

    // 4. Manager: Open Evaluation Form
    function openEvaluationForm(id) {
        const item = window.currentEvaluationList.find(x => x.ID === id);
        if (!item) return;

        document.getElementById('do-eval-id').value = id;
        document.getElementById('do-eval-cname').innerText = item.Candidate_Name;
        document.getElementById('do-eval-position').innerText = item.Position;
        document.getElementById('eval-comment').value = '';
        document.getElementById('eval-proposed-salary').value = '';
        document.getElementById('eval-signature-confirm').checked = false;

        // Populate Manager Info from currentUser
        if (currentUser) {
            document.getElementById('do-eval-mgr-name').innerText = currentUser.name || currentUser.username;
            document.getElementById('do-eval-mgr-pos').innerText = currentUser.role === 'Admin' ? 'Administrator' : (currentUser.position || 'Manager');
            document.getElementById('do-eval-mgr-dept').innerText = currentUser.department || 'HR';
        }

        // Render Dynamic Inputs
        const container = document.getElementById('dynamic-scores-container');
        if (container) {
            container.innerHTML = '';

            let criteria = item.Criteria_Config;
            if (!criteria || criteria.length === 0) {
                // Try parsing if string? No, API returns object.
                // If legacy data, criteria might be empty.
                criteria = ['Chuyên môn', 'Kỹ năng mềm', 'Văn hóa'];
            }

            criteria.forEach((c, index) => {
                const div = document.createElement('div');
                div.className = 'col-md-4 mb-3';
                div.innerHTML = `
                    <label class="form-label fw-bold">${c} (1-10)</label>
                    <select class="form-select score-input" data-label="${c}" onchange="calculateEvaluationTotal()">
                        <option value="">-- Chọn điểm --</option>
                        <option value="10">10 - Rất tốt</option>
                        <option value="9">9 - Rất tốt</option>
                        <option value="8">8 - Tốt</option>
                        <option value="7">7 - Tốt</option>
                        <option value="6">6 - Hài lòng</option>
                        <option value="5">5 - Hài lòng</option>
                        <option value="4">4 - Tạm được</option>
                        <option value="3">3 - Tạm được</option>
                        <option value="2">2 - Không đạt</option>
                        <option value="1">1 - Không đạt</option>
                    </select>
                `;
                container.appendChild(div);
            });
            calculateEvaluationTotal();
        }

        // Reset Radio
        document.querySelectorAll('input[name="eval-result"]').forEach(r => r.checked = false);

        new bootstrap.Modal(document.getElementById('doEvaluationModal')).show();
    }

    // 5. Manager: Submit Evaluation
    function submitEvaluation() {
        const id = document.getElementById('do-eval-id').value;
        const comment = document.getElementById('eval-comment').value;
        const resultEls = document.querySelector('input[name="eval-result"]:checked');

        if (!comment || !resultEls) {
            Swal.fire('Cảnh báo', 'Vui lòng nhập nhận xét và chọn kết quả cuối cùng', 'warning');
            return;
        }

        // Collect Scores
        const scores = {};
        const inputs = document.querySelectorAll('.score-input');
        let missingScore = false;
        inputs.forEach(input => {
            if (!input.value) missingScore = true;
            scores[input.getAttribute('data-label')] = input.value;
        });

        // Map to legacy if standard names (optional but good for sheet columns 10-12)
        if (scores['Chuyên môn']) scores.professional = scores['Chuyên môn'];
        if (scores['Kỹ năng mềm']) scores.softSkills = scores['Kỹ năng mềm'];
        if (scores['Văn hóa']) scores.culture = scores['Văn hóa'];

        const result = resultEls.value;
        const totalScore = document.getElementById('eval-total-score').value;
        const proposedSalary = document.getElementById('eval-proposed-salary').value;
        const isSigned = document.getElementById('eval-signature-confirm').checked;

        if (!isSigned) {
            Swal.fire('Cảnh báo', 'Vui lòng xác nhận và ký điện tử trước khi hoàn thành.', 'warning');
            return;
        }

        const additionalData = {
            managerName: currentUser.name || currentUser.username,
            managerDept: currentUser.department || 'HR',
            managerPos: currentUser.role === 'Admin' ? 'Administrator' : (currentUser.position || 'Manager'),
            totalScore: totalScore,
            proposedSalary: proposedSalary,
            signatureStatus: 'Digitally Signed'
        };

        const btn = document.querySelector('#doEvaluationModal .btn-success');
        btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Đang xử lý...';
        btn.disabled = true;

        google.script.run.withSuccessHandler(function (res) {
            btn.innerHTML = '<i class="fas fa-check-circle"></i> Hoàn thành Đánh giá';
            btn.disabled = false;

            if (res.success) {
                bootstrap.Modal.getInstance(document.getElementById('doEvaluationModal')).hide();
                Swal.fire('Thành công', 'Đã lưu kết quả đánh giá và ký xác nhận!', 'success');
                // ... (rest of success logic)

                // Refresh Pending List (Dashboard)
                checkForPendingEvaluations();

                // Refresh Evaluations Page List (if open)
                if (typeof loadEvaluations === 'function') loadEvaluations();

                // Refresh Dashboard Stats
                if (typeof loadDashboardData === 'function') loadDashboardData();
            } else {
                Swal.fire('Lỗi', res.message, 'error');
            }
        }).apiSubmitEvaluation(id, scores, result, comment, additionalData);
    }

    function calculateEvaluationTotal() {
        const inputs = document.querySelectorAll('.score-input');
        let total = 0;
        let count = 0;
        inputs.forEach(input => {
            if (input.value) {
                total += parseInt(input.value);
                count++;
            }
        });

        const avg = count > 0 ? (total / count).toFixed(1) : 0;
        const totalInput = document.getElementById('eval-total-score');
        const labelEl = document.getElementById('eval-score-label');

        if (totalInput) totalInput.value = avg;

        if (labelEl) {
            if (count === 0) {
                labelEl.innerText = 'Chưa đánh giá';
                labelEl.className = 'text-muted';
            } else {
                if (avg >= 9) { labelEl.innerText = 'Rất tốt'; labelEl.className = 'text-success fw-bold'; }
                else if (avg >= 7) { labelEl.innerText = 'Tốt'; labelEl.className = 'text-primary fw-bold'; }
                else if (avg >= 5) { labelEl.innerText = 'Hài lòng'; labelEl.className = 'text-info fw-bold'; }
                else if (avg >= 3) { labelEl.innerText = 'Tạm được'; labelEl.className = 'text-warning fw-bold'; }
                else { labelEl.innerText = 'Không đạt'; labelEl.className = 'text-danger fw-bold'; }
            }
        }
    }


    // 7. NEW: CREATE EVALUATION LOGIC
    let managerOptionsCache = []; // To store manager options for reuse

    function openCreateEvaluationModal(candidateId = null) {
        // Populate Candidates
        const cSelect = document.getElementById('create-eval-candidate');
        cSelect.innerHTML = '<option value="">-- Chọn ứng viên --</option>';

        const interviewCandidates = candidatesData.filter(c => {
            const s = (c.Stage || c.Status || '').toLowerCase();
            return s.includes('phỏng vấn') || s.includes('interview');
        });

        interviewCandidates.forEach(c => {
            const opt = document.createElement('option');
            opt.value = c.ID;
            opt.text = `${c.Name} - ${c.Position}`;
            cSelect.appendChild(opt);
        });

        if (candidateId) cSelect.value = candidateId;

        // Cache Manager Options
        managerOptionsCache = usersData.filter(u => u.Role === 'Manager').map(m => ({
            value: m.Email || m.Username,
            text: `${m.Name} (${m.Email})`
        }));

        // Reset Manager Inputs
        const container = document.getElementById('manager-select-container');
        container.innerHTML = '';
        addManagerInput(); // Add first one

        // Reset Criteria Inputs
        const critContainer = document.getElementById('criteria-container');
        critContainer.innerHTML = '';
        // Add default criteria
        ['Chuyên môn', 'Kỹ năng mềm', 'Văn hóa'].forEach(c => addCriteriaInput(c));

        new bootstrap.Modal(document.getElementById('createEvaluationModal')).show();
    }

    function addManagerInput() {
        const container = document.getElementById('manager-select-container');
        if (container.children.length >= 3) {
            Swal.fire('Thông báo', 'Tối đa 3 người đánh giá', 'info');
            return;
        }

        const div = document.createElement('div');
        div.className = 'input-group mb-2 manager-row';

        let optionsHtml = '<option value="">-- Chọn Manager --</option>';
        managerOptionsCache.forEach(m => {
            optionsHtml += `<option value="${m.value}">${m.text}</option>`;
        });

        div.innerHTML = `
            <select class="form-select manager-select">
                ${optionsHtml}
            </select>
            <button class="btn btn-outline-danger remove-manager" type="button" onclick="this.parentElement.remove()">
                <i class="fas fa-times"></i>
            </button>
        `;
        container.appendChild(div);

        // Hide remove button if only one
        toggleRemoveButtons();
    }

    function toggleRemoveButtons() {
        const rows = document.querySelectorAll('.manager-row');
        rows.forEach(row => {
            const btn = row.querySelector('.remove-manager');
            if (rows.length > 1) btn.style.display = 'block';
            else btn.style.display = 'none';
        });
    }

    function addCriteriaInput(value = '') {
        const container = document.getElementById('criteria-container');
        const div = document.createElement('div');
        div.className = 'input-group mb-2 criteria-row';
        div.innerHTML = `
            <input type="text" class="form-control criteria-input" value="${value}" placeholder="Nhập tên tiêu chí">
            <button class="btn btn-outline-danger" type="button" onclick="this.parentElement.remove()">
                <i class="fas fa-times"></i>
            </button>
        `;
        container.appendChild(div);
    }

    function submitCreateEvaluation() {
        const cId = document.getElementById('create-eval-candidate').value;
        const doNow = document.getElementById('create-eval-now').checked;

        if (!cId) {
            Swal.fire('Lỗi', 'Vui lòng chọn ứng viên', 'warning');
            return;
        }

        // Get Managers
        const managerSelects = document.querySelectorAll('.manager-select');
        const managers = [];
        managerSelects.forEach(s => {
            if (s.value) managers.push(s.value);
        });

        if (managers.length === 0) {
            Swal.fire('Lỗi', 'Vui lòng chọn ít nhất 1 người đánh giá', 'warning');
            return;
        }

        // Get Criteria
        const criteriaInputs = document.querySelectorAll('.criteria-input');
        const criteria = [];
        criteriaInputs.forEach(i => {
            if (i.value.trim()) criteria.push(i.value.trim());
        });

        if (criteria.length === 0) {
            Swal.fire('Lỗi', 'Vui lòng nhập ít nhất 1 tiêu chí', 'warning');
            return;
        }

        const btn = document.querySelector('#createEvaluationModal .btn-primary');
        const originalText = btn.innerText;
        btn.innerText = 'Đang tạo...';
        btn.disabled = true;

        google.script.run.withSuccessHandler(function (res) {
            btn.innerText = originalText;
            btn.disabled = false;

            if (res.success) {
                bootstrap.Modal.getInstance(document.getElementById('createEvaluationModal')).hide();
                Swal.fire({
                    title: 'Thành công',
                    text: res.message,
                    icon: 'success'
                }).then(() => {
                    loadEvaluations();
                });
            } else {
                Swal.fire('Lỗi', res.message, 'error');
            }
        }).apiCreateEvaluationRequest(cId, managers, currentUser.email, criteria);
    }

    // 6. LOAD EVALUATIONS LIST (For Evaluations Page)
    function loadEvaluations() {
        console.log('🔄 Calling apiGetEvaluationsList for:', currentUser.email || currentUser.username, currentUser.role);
        const tbody = document.querySelector('#evaluations-table tbody');
        if (tbody) tbody.innerHTML = '<tr><td colspan="8" class="text-center py-4 text-muted"><div class="spinner-border spinner-border-sm text-primary"></div> Đang tải...</td></tr>';

        google.script.run
            .withSuccessHandler(renderEvaluations)
            .withFailureHandler(function (err) {
                console.error('❌ Error loading evaluations:', err);
                const tbody = document.querySelector('#evaluations-table tbody');
                if (tbody) tbody.innerHTML = '<tr><td colspan="8" class="text-center py-4 text-danger">Lỗi tải dữ liệu: ' + err.message + '</td></tr>';
            })
            .apiGetEvaluationsList(currentUser.email || currentUser.username, currentUser.role);
    }

    function confirmResetEvaluationSheet() {
        Swal.fire({
            title: 'RESET BẢNG ĐÁNH GIÁ?',
            text: 'Thao tác này sẽ xóa sạch dữ liệu hiện có trong sheet EVALUATIONS và tạo lại cấu trúc 24 cột chuẩn. Bạn có chắc chắn không?',
            icon: 'warning',
            showCancelButton: true,
            confirmButtonColor: '#d33',
            cancelButtonColor: '#3085d6',
            confirmButtonText: 'Đồng ý Reset',
            cancelButtonText: 'Hủy'
        }).then((result) => {
            if (result.isConfirmed) {
                Swal.fire({
                    title: 'Đang xử lý...',
                    allowOutsideClick: false,
                    didOpen: () => { Swal.showLoading(); }
                });

                google.script.run
                    .withSuccessHandler(function (res) {
                        Swal.fire('Thành công!', res.message, 'success');
                        loadEvaluations();
                    })
                    .withFailureHandler(function (err) {
                        Swal.fire('Lỗi', err.message, 'error');
                    })
                    .apiResetEvaluationSheet();
            }
        });
    }

    function renderEvaluations(list) {
        console.log('✅ Received evaluations list:', list);
        // Save list globally for detail view and modal
        window.currentEvaluationList = list;

        const tbody = document.querySelector('#evaluations-table tbody');
        if (!tbody) return;
        tbody.innerHTML = '';

        if (!list) {
            tbody.innerHTML = '<tr><td colspan="8" class="text-center py-4 text-danger">Lỗi: Phản hồi từ máy chủ không hợp lệ (null). Hãy thử tải lại trang.</td></tr>';
            return;
        }

        if (list.length === 0) {
            tbody.innerHTML = '<tr><td colspan="8" class="text-center py-4 text-muted">Hệ thống không tìm thấy đánh giá nào cho tài khoản của bạn.</td></tr>';
            return;
        }

        // Group by Batch_ID if present, otherwise treat as individual
        const batches = {};
        const singles = [];

        list.forEach(item => {
            if (item.Batch_ID) {
                if (!batches[item.Batch_ID]) {
                    batches[item.Batch_ID] = {
                        Items: [],
                        Candidate: item.Candidate_Name,
                        Position: item.Position,
                        Created: item.Created_At
                    };
                }
                batches[item.Batch_ID].Items.push(item);
            } else {
                singles.push(item);
            }
        });

        // 1. Render Batches
        Object.keys(batches).forEach(bId => {
            const batch = batches[bId];
            const items = batch.Items;

            // Calculate summary
            const total = items.length;
            const completed = items.filter(i => {
                const s = (i.Status || '').toString().trim().toLowerCase();
                return s === 'completed' || s === 'hoàn thành' || s === 'đã hoàn thành';
            }).length;
            const statusClass = completed === total ? 'success' : (completed > 0 ? 'info' : 'warning');
            const statusText = `${completed}/${total} Hoàn thành`;

            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><span class="badge bg-secondary">Group</span></td>
                <td><strong>${batch.Candidate}</strong></td>
                <td>${batch.Position}</td>
                <td>
                    <small>
                    ${items.map(i => {
                const s = (i.Status || '').toString().trim().toLowerCase();
                const isDone = s === 'completed' || s === 'hoàn thành' || s === 'đã hoàn thành';
                return `<div class="${isDone ? 'text-success' : 'text-muted'}">
                            <i class="fas fa-${isDone ? 'check-circle' : 'clock'}"></i> ${i.Manager_Email}
                        </div>`;
            }).join('')}
                    </small>
                </td>
                <td><span class="badge bg-${statusClass}">${statusText}</span></td>
                <td>-</td>
                <td>${batch.Created ? new Date(batch.Created).toLocaleDateString('vi-VN') : '-'}</td>
                <td>
                    ${renderBatchActions(items)}
                </td>
            `;
            tbody.appendChild(tr);
        });

        // 2. Render Singles
        singles.forEach(item => {
            const tr = document.createElement('tr');
            const s = (item.Status || '').toString().trim().toLowerCase();
            const isDone = s === 'completed' || s === 'hoàn thành' || s === 'đã hoàn thành';
            let badgeClass = isDone ? 'success' : 'warning';
            tr.innerHTML = `
                <td>${item.ID}</td>
                <td>${item.Candidate_Name}</td>
                <td>${item.Position}</td> 
                <td>${item.Manager_Email}</td>
                <td><span class="badge bg-${badgeClass}">${isDone ? 'Đã hoàn thành' : 'Chờ đánh giá'}</span></td>
                <td>${item.Final_Result || '-'}</td>
                <td>${item.Created_At ? new Date(item.Created_At).toLocaleDateString('vi-VN') : '-'}</td>
                <td>
                    ${renderSingleAction(item)}
                </td>
            `;
            tbody.appendChild(tr);
        });
    }

    function renderBatchActions(items) {
        // Find if I have a pending task in this batch
        const myEmail = String(currentUser.email || '').toLowerCase().trim();
        const myUser = String(currentUser.username || '').toLowerCase().trim();

        const myTask = items.find(i => {
            const mgrEmail = String(i.Manager_Email || '').toLowerCase().trim();
            const s = (i.Status || '').toString().trim().toLowerCase();
            return (mgrEmail === myEmail || mgrEmail === myUser) && s === 'pending';
        });

        let html = '';
        if (myTask) {
            html += `<button class="btn btn-sm btn-primary mb-1" onclick="openEvaluationForm('${myTask.ID}')">Chấm điểm</button> `;
        }

        // View detail button (pass first ID, and flag to show all in batch)
        html += `<button class="btn btn-sm btn-outline-info me-1" onclick="viewEvaluationDetail('${items[0].ID}', true)"><i class="fas fa-eye"></i> Xem</button>`;

        // Direct PDF Button
        html += `<button class="btn btn-sm btn-success" onclick="exportEvaluationPDF('${items[0].ID}')"><i class="fas fa-file-pdf"></i> PDF</button>`;
        return html;
    }

    function renderSingleAction(item) {
        const myEmail = String(currentUser.email || '').toLowerCase().trim();
        const myUser = String(currentUser.username || '').toLowerCase().trim();
        const mgrEmail = String(item.Manager_Email || '').toLowerCase().trim();
        const isMyTask = (currentUser.role === 'Manager' && (myEmail === mgrEmail || myUser === mgrEmail)) || currentUser.role === 'Admin';

        const s = (item.Status || '').toString().trim().toLowerCase();

        if (s === 'pending' && isMyTask) {
            return `<button class="btn btn-sm btn-primary me-1" onclick="openEvaluationForm('${item.ID}')"><i class="fas fa-pen"></i> Chấm điểm</button>`;
        } else {
            return `
                <button class="btn btn-sm btn-outline-info me-1" onclick="viewEvaluationDetail('${item.ID}')"><i class="fas fa-eye"></i> Xem</button>
                <button class="btn btn-sm btn-success" onclick="exportEvaluationPDF('${item.ID}')"><i class="fas fa-file-pdf"></i> PDF</button>
            `;
        }
    }

    function viewEvaluationDetail(id, isBatch = false) {
        if (!window.currentEvaluationList) return;
        const item = window.currentEvaluationList.find(x => x.ID == id);
        if (!item) {
            Swal.fire('Lỗi', 'Không tìm thấy dữ liệu', 'error');
            return;
        }

        let itemsToShow = [item];
        if (isBatch && item.Batch_ID) {
            itemsToShow = window.currentEvaluationList.filter(x => x.Batch_ID === item.Batch_ID);
        }

        let html = '<div class="text-start p-1" style="max-height: 80vh; overflow-y: auto; font-family: \'Segoe UI\', system-ui, -apple-system, sans-serif;">';

        // Professional Header
        html += '<div class="p-4 mb-4" style="background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%); border-radius: 12px; border-left: 6px solid #0056b3; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">';
        html += '<div class="row align-items-center">';
        html += '<div class="col-md-7">';
        html += '<div class="text-uppercase small fw-bold text-muted mb-1 ls-1">Phiếu kết quả phỏng vấn</div>';
        html += '<h4 class="fw-bold text-dark mb-1" style="letter-spacing: -0.5px;">' + (item.Candidate_Name || '---') + '</h4>';
        html += '<div class="d-flex align-items-center mt-2">';
        html += '<span class="badge bg-white text-primary border border-primary me-2 shadow-sm">' + (item.Candidate_ID || 'ID:---') + '</span>';
        html += '<span class="text-muted small"><i class="fas fa-calendar-alt me-1"></i> ' + (item.Created_At ? new Date(item.Created_At).toLocaleDateString('vi-VN') : '---') + '</span>';
        html += '</div></div>';
        html += '<div class="col-md-5 text-md-end mt-3 mt-md-0">';
        html += '<div class="fw-bold text-primary h5 mb-0">' + (item.Position || '---') + '</div>';
        html += '<div class="text-muted fw-medium"><i class="fas fa-layer-group me-1"></i>' + (item.Department || '---') + '</div>';
        html += '</div></div></div>';

        // Comparison Table
        html += '<div class="table-responsive rounded-3 border">';
        html += '<table class="table table-sm table-borderless align-middle mb-0">';
        html += '<thead style="background-color: #f1f3f5; border-bottom: 2px solid #dee2e6;"><tr>';
        html += '<th class="ps-3 py-3 text-muted small fw-bold" style="width: 30%;">TIÊU CHÍ ĐÁNH GIÁ</th>';

        itemsToShow.forEach(i => {
            const mName = i.Manager_Name || i.Manager_Email.split('@')[0];
            const mPos = i.Manager_Position || 'PV';
            const mDept = i.Manager_Department || '';
            html += '<th class="text-center py-3">';
            html += '<div class="fw-bold text-dark">' + mName + '</div>';
            html += '<div class="text-muted" style="font-size: 0.65rem; font-weight: 500; text-transform: uppercase;">' + mPos + '<br>' + mDept + '</div>';
            html += '</th>';
        });
        html += '</tr></thead>';
        html += '<tbody class="bg-white">';

        let allCriteria = new Set();
        itemsToShow.forEach(i => {
            if (i.Scores_JSON) {
                Object.keys(i.Scores_JSON).forEach(k => {
                    if (!['professional', 'softSkills', 'culture'].includes(k)) allCriteria.add(k);
                });
            }
        });
        if (allCriteria.size === 0) {
            allCriteria.add("Chuyên môn");
            allCriteria.add("Kỹ năng mềm");
            allCriteria.add("Văn hóa & Thái độ");
        }

        allCriteria.forEach(crit => {
            html += '<tr style="border-bottom: 1px solid #f1f3f5;"><td class="ps-3 fw-medium text-dark">' + crit + '</td>';
            itemsToShow.forEach(i => {
                let score = i.Scores_JSON ? i.Scores_JSON[crit] : null;
                if (score === null) {
                    if (crit === "Chuyên môn") score = i.Score_Professional;
                    if (crit === "Kỹ năng mềm") score = i.Score_Soft_Skills;
                    if (crit === "Văn hóa & Thái độ") score = i.Score_Culture;
                }
                const scoreNum = parseFloat(score);
                const scoreColor = isNaN(scoreNum) ? 'text-muted' : (scoreNum >= 8 ? 'text-success' : (scoreNum >= 5 ? 'text-primary' : 'text-danger'));
                html += '<td class="text-center fw-bold ' + scoreColor + '">' + (score || '-') + '</td>';
            });
            html += '</tr>';
        });

        // Average Row Calculation (Per Manager)
        html += '<tr class="fw-bold" style="background-color: #f8f9fa;"><td class="ps-3 py-3 text-dark">ĐIỂM TRUNG BÌNH</td>';
        itemsToShow.forEach(i => {
            let personSum = 0;
            let count = 0;
            allCriteria.forEach(crit => {
                let score = i.Scores_JSON ? i.Scores_JSON[crit] : null;
                if (score === null) {
                    if (crit === "Chuyên môn") score = i.Score_Professional;
                    if (crit === "Kỹ năng mềm") score = i.Score_Soft_Skills;
                    if (crit === "Văn hóa & Thái độ") score = i.Score_Culture;
                }
                const scoreNum = parseFloat(score);
                if (!isNaN(scoreNum)) {
                    personSum += scoreNum;
                    count++;
                }
            });
            const avgVal = count > 0 ? (personSum / count).toFixed(1) : '0.0';
            html += '<td class="text-center text-danger h5 mb-0 py-3">' + avgVal + '</td>';
        });
        html += '</tr>';

        // Integrated Comments Row
        html += '<tr><td class="ps-3 py-2 fw-bold bg-light">NHẬN XÉT CHI TIẾT</td>';
        itemsToShow.forEach(i => {
            const comment = i.Manager_Comment || 'Chưa có nhận xét.';
            html += '<td class="small text-muted p-2" style="font-style: italic; vertical-align: top; border: 1px solid #f1f3f5;">' + comment + '</td>';
        });
        html += '</tr>';

        // Salary Row
        html += '<tr class="fw-bold"><td class="ps-3 py-2 text-muted small">MỨC LƯƠNG ĐỀ XUẤT</td>';
        itemsToShow.forEach(i => {
            html += '<td class="text-center small text-primary py-2">' + (i.Proposed_Salary || '-') + '</td>';
        });
        html += '</tr>';
        html += '</tbody></table></div>';

        // Admin Override
        const isOfficial = (currentUser && (currentUser.role === 'Admin' || currentUser.role === 'Recruiter'));
        if (isOfficial) {
            html += '<div class="card mt-4 border-0 shadow-sm" style="background-color: #fff9db; border-radius: 12px;">';
            html += '<div class="card-header bg-warning border-0 text-dark py-2 rounded-top-4" style="font-size: 0.8rem; letter-spacing: 0.5px;"><i class="fas fa-gavel me-2"></i>DÀNH CHO QUẢN TRỊ (PHÊ DUYỆT CUỐI CÙNG)</div>';
            html += '<div class="card-body p-3"><div class="row g-3"><div class="col-md-6">';
            html += '<label class="form-label small fw-bold text-muted">KẾT LUẬN CHÍNH THỨC</label>';
            html += '<select id="pdf-override-result" class="form-select shadow-sm border-0">';
            html += '<option value="' + (item.Final_Result || 'Pending') + '">Giữ nguyên (' + (item.Final_Result || 'Chờ') + ')</option>';
            html += '<option value="Pass">PASS (Đạt yêu cầu)</option>';
            html += '<option value="Consider">CONSIDER (Xem xét thêm)</option>';
            html += '<option value="Reject">REJECT (Không đạt)</option></select></div>';
            html += '<div class="col-md-6"><label class="form-label small fw-bold text-muted">LƯƠNG CHÍNH THỨC (XUẤT FILE)</label>';
            html += '<input type="text" id="pdf-official-salary" class="form-control shadow-sm border-0" placeholder="VD: 15.000.000 VNĐ" value="' + (item.Proposed_Salary || '') + '"></div>';
            html += '</div></div></div>';
        }

        // Comments
        html += '<div class="mt-4"><h6 class="fw-bold text-dark mb-3"><i class="fas fa-comment-alt text-primary me-2"></i>NHẬN XÉT CHI TIẾT</h6>';
        itemsToShow.forEach(i => {
            const hasComment = i.Manager_Comment && i.Manager_Comment !== 'Không có nhận xét.';
            const resColor = i.Final_Result === 'Pass' ? '#198754' : '#dc3545';
            const badgeClass = i.Final_Result === 'Pass' ? 'bg-success' : (i.Final_Result === 'Reject' ? 'bg-danger' : 'bg-warning text-dark');
            html += '<div class="mb-3 p-3 rounded-3" style="background-color: #f8f9fa; border-left: 4px solid ' + resColor + ';">';
            html += '<div class="d-flex justify-content-between align-items-center mb-2">';
            html += '<span class="fw-bold text-dark">' + (i.Manager_Name || i.Manager_Email) + '</span>';
            html += '<span class="badge ' + badgeClass + '" style="font-size: 0.7rem; padding: 4px 8px;">' + (i.Final_Result || 'Pending') + '</span></div>';
            html += '<p class="mb-0 text-muted" style="font-size: 0.9rem; line-height: 1.5; font-style: ' + (hasComment ? 'normal' : 'italic') + ';">' + (i.Manager_Comment || 'Chưa có nhận xét cụ thể.') + '</p></div>';
        });
        html += '</div>';

        html += '<div class="text-center mt-4 text-muted small opacity-50" style="letter-spacing: 2px;">HỆ THỐNG ATS RECRUIT - PROFESSIONAL EVALUATION</div>';

        const finalResColor = item.Final_Result === 'Pass' ? 'text-success' : 'text-danger';
        html += '<div class="alert alert-warning mt-4 text-center py-2">';
        html += '<h5 class="mb-0 fw-bold">KẾT LUẬN CUỐI CÙNG: <span class="text-uppercase ms-2 ' + finalResColor + '">' + (item.Final_Result || 'ĐANG CHỜ') + '</span></h5></div>';
        html += '</div>';

        Swal.fire({
            title: 'Kết quả Đánh giá (' + itemsToShow.length + ' người)',
            html: html,
            width: '900px',
            showCancelButton: true,
            confirmButtonText: 'Đóng',
            cancelButtonText: '<i class="fas fa-file-pdf"></i> Xuất file PDF',
            cancelButtonColor: '#198754'
        }).then((res) => {
            if (res.dismiss === Swal.DismissReason.cancel) {
                exportEvaluationPDF(id);
            }
        });
    }

    function exportEvaluationPDF(id) {
        // Collect overrides if available
        const overrideRes = document.getElementById('pdf-override-result') ? document.getElementById('pdf-override-result').value : null;
        const officialSal = document.getElementById('pdf-official-salary') ? document.getElementById('pdf-official-salary').value : null;

        Swal.fire({
            title: 'Đang tạo PDF...',
            text: 'Vui lòng chờ trong giây lát',
            allowOutsideClick: false,
            didOpen: () => { Swal.showLoading(); }
        });

        google.script.run
            .withSuccessHandler(function (res) {
                if (res.success) {
                    Swal.fire({
                        icon: 'success',
                        title: 'Đã tạo PDF thành công!',
                        text: 'File của bạn đang được mở/tải về...',
                        timer: 2000,
                        showConfirmButton: false
                    });

                    // Convert base64 to Blob and Download
                    const fileName = res.filename;
                    const byteCharacters = atob(res.data);
                    const byteNumbers = new Array(byteCharacters.length);
                    for (let i = 0; i < byteCharacters.length; i++) {
                        byteNumbers[i] = byteCharacters.charCodeAt(i);
                    }
                    const byteArray = new Uint8Array(byteNumbers);
                    const blob = new Blob([byteArray], { type: 'application/pdf' });

                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.style.display = 'none';
                    a.href = url;
                    a.download = fileName;
                    document.body.appendChild(a);
                    a.click();

                    // Also open in new tab
                    window.open(url, '_blank');

                    window.URL.revokeObjectURL(url);
                } else {
                    Swal.fire('Lỗi', res.message, 'error');
                }
            })
            .withFailureHandler(function (err) {
                Swal.fire('Lỗi kết nối', err.message, 'error');
            })
            .apiExportEvaluationPDF(id, overrideRes, officialSal);
    }
    function debugUserInfo() {
        Swal.fire({
            title: 'Hệ thống đang kiểm tra...',
            didOpen: () => { Swal.showLoading(); }
        });

        google.script.run
            .withSuccessHandler(function (res) {
                Swal.fire({
                    title: 'Dữ liệu Hệ thống (DEBUG)',
                    html: '<pre style="text-align: left; max-height: 400px; overflow: auto; font-size: 11px;">' + res + '</pre>',
                    width: '800px'
                });
            })
            .withFailureHandler(function (err) {
                Swal.fire('Lỗi Debug', err.toString(), 'error');
            })
            .debugSheetData();
    }
    // ============================================
    // BULK UPLOAD LOGIC
    // ============================================

    function handleImportAction() {
        // Check active tab
        const activeTab = document.querySelector('#bulkUploadTabs .active').id;

        if (activeTab === 'tab-cvs-btn') {
            handleBulkUpload();
        } else {
            handleSheetImport();
        }
    }

    function toggleImportInputs() {
        const type = document.querySelector('input[name="importType"]:checked').value;
        if (type === 'URL') {
            document.getElementById('input-url-container').classList.remove('d-none');
            document.getElementById('input-file-container').classList.add('d-none');
        } else {
            document.getElementById('input-url-container').classList.add('d-none');
            document.getElementById('input-file-container').classList.remove('d-none');
        }
    }

    function handleSheetImport() {
        const type = document.querySelector('input[name="importType"]:checked').value;
        const source = document.getElementById('bulk-source').value;
        const stage = document.getElementById('bulk-stage').value;

        // Progress UI
        const progressDiv = document.getElementById('bulk-upload-progress');
        const progressBar = document.getElementById('bulk-progress-bar');
        const statusText = document.getElementById('bulk-status-text');

        progressDiv.classList.remove('d-none');
        progressBar.style.width = '10%';
        statusText.innerText = 'Đang chuẩn bị...';

        if (type === 'URL') {
            const url = document.getElementById('sheet-url').value;
            if (!url) { Swal.fire('Lỗi', 'Vui lòng nhập Link Google Sheet', 'error'); return; }

            statusText.innerText = 'Đang đọc dữ liệu từ Google Sheet...';
            progressBar.style.width = '30%';

            google.script.run.withSuccessHandler(res => {
                finishImport(res);
            }).withFailureHandler(err => { failImport(err); }).apiImportFromSheet('URL', url, source, stage);

        } else {
            const fileInput = document.getElementById('sheet-file');
            if (!fileInput.files || fileInput.files.length === 0) {
                Swal.fire('Lỗi', 'Vui lòng chọn file Excel', 'error'); return;
            }
            const file = fileInput.files[0];

            statusText.innerText = 'Đang đọc file Excel...';
            progressBar.style.width = '20%';

            const reader = new FileReader();
            reader.onload = function (e) {
                const base64 = e.target.result.split(',')[1];
                const data = { name: file.name, base64: base64 };

                statusText.innerText = 'Đang gửi lên server xử lý...';
                progressBar.style.width = '50%';

                google.script.run.withSuccessHandler(res => {
                    finishImport(res);
                }).withFailureHandler(err => { failImport(err); }).apiImportFromSheet('FILE', data, source, stage);
            };
            reader.readAsDataURL(file);
        }
    }

    function finishImport(res) {
        const progressBar = document.getElementById('bulk-progress-bar');
        const statusText = document.getElementById('bulk-status-text');

        progressBar.style.width = '100%';
        if (res.success) {
            Swal.fire('Thành công', res.message, 'success');
            // Close & Reload
            loadDashboardData();
            const modalEl = document.getElementById('bulkUploadModal');
            const modal = bootstrap.Modal.getInstance(modalEl);
            if (modal) modal.hide();
        } else {
            statusText.innerText = 'Có lỗi xảy ra.';
            progressBar.classList.add('bg-danger');
            Swal.fire('Thát bại', res.message, 'error');
        }
    }

    function failImport(err) {
        const progressBar = document.getElementById('bulk-progress-bar');
        const statusText = document.getElementById('bulk-status-text');
        progressBar.classList.add('bg-danger');
        statusText.innerText = 'Lỗi kết nối!';
        Swal.fire('Lỗi Server', err.message, 'error');
    }

    function handleBulkUpload() {
        const fileInput = document.getElementById('bulk-files');
        const sourceInput = document.getElementById('bulk-source');
        const stageInput = document.getElementById('bulk-stage');

        if (!fileInput.files || fileInput.files.length === 0) {
            Swal.fire('Lỗi', 'Vui lòng chọn ít nhất 1 file CV.', 'error');
            return;
        }

        // Limit to 5 for safety in Phase 1
        if (fileInput.files.length > 5) {
            Swal.fire('Lưu ý', 'Để đảm bảo hiệu năng trong bản thử nghiệm, vui lòng chọn tối đa 5 file một lúc.', 'warning');
            return;
        }

        const files = Array.from(fileInput.files);
        const source = sourceInput.value;
        const stage = stageInput.value;

        // Visual Progress
        const progressDiv = document.getElementById('bulk-upload-progress');
        const progressBar = document.getElementById('bulk-progress-bar');
        const statusText = document.getElementById('bulk-status-text');

        progressDiv.classList.remove('d-none');
        progressBar.style.width = '10%';
        statusText.innerText = `Đang đọc ${files.length} files...`;

        // 1. Read all files as Base64
        const promises = files.map(file => {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (e) => {
                    const data = e.target.result.split(',')[1]; // Get Base64 part
                    resolve({
                        name: file.name,
                        type: file.type,
                        data: data
                    });
                };
                reader.onerror = reject;
                reader.readAsDataURL(file);
            });
        });

        Promise.all(promises).then(filesData => {
            progressBar.style.width = '40%';
            statusText.innerText = 'Đang gửi lên server (có thể mất vài giây)...';

            // 2. Send to Backend
            google.script.run.withSuccessHandler(res => {
                progressBar.style.width = '100%';

                if (res.errors && res.errors.length > 0) {
                    let msg = `Thành công: ${res.success.length} file.\nLỗi: ${res.errors.length} file.\nChi tiết lỗi: ${res.errors.join('\n')} `;
                    Swal.fire('Hoàn tất một phần', msg, 'warning');
                } else {
                    Swal.fire('Thành công', `Đã tải lên ${res.success.length} hồ sơ!`, 'success');
                    // Reset
                    fileInput.value = '';
                    progressDiv.classList.add('d-none');
                    // Reload Data
                    loadDashboardData();
                    // Close Modal
                    const modalEl = document.getElementById('bulkUploadModal');
                    const modal = bootstrap.Modal.getInstance(modalEl);
                    if (modal) modal.hide();
                }
            }).withFailureHandler(err => {
                progressBar.classList.add('bg-danger');
                statusText.innerText = 'Lỗi kết nối!';
                Swal.fire('Lỗi Server', err.message, 'error');
            }).apiBulkUploadCandidates(filesData, source, stage);

        }).catch(err => {
            console.error(err);
            Swal.fire('Lỗi đọc file', 'Không thể đọc file từ máy tính của bạn.', 'error');
        });
    }

    // ============================================
    // CANDIDATE SOURCE MANAGEMENT
    // ============================================

    let candidateSourcesData = [];

    function loadCandidateSources() {
        console.log('Fetching candidate sources...');
        google.script.run.withSuccessHandler(function (sources) {
            console.log('Sources fetched:', sources);
            candidateSourcesData = sources || [];
            if (candidateSourcesData.length === 0) {
                // Fallback defaults if empty (should be handled by backend but good safety)
                candidateSourcesData = ['Website', 'LinkedIn', 'Facebook', 'Referral', 'Job Portal'];
            }
            renderCandidateSources();
            populateCandidateSources(); // Update dropdowns immediately
        }).withFailureHandler(function (err) {
            console.error('Failed to fetch sources:', err);
            // Fallback on error
            candidateSourcesData = ['Website', 'LinkedIn', 'Facebook', 'Referral', 'Job Portal'];
            populateCandidateSources();
        }).apiGetCandidateSources();
    }

    function renderCandidateSources() {
        const tbody = document.querySelector('#sources-table tbody');
        if (!tbody) return;
        tbody.innerHTML = '';

        if (candidateSourcesData.length === 0) {
            tbody.innerHTML = '<tr><td colspan="2" class="text-center">Chưa có nguồn nào</td></tr>';
            return;
        }

        candidateSourcesData.forEach(source => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
            < td > ${source}</td >
                <td class="text-center">
                    <button class="btn btn-sm btn-primary me-1" onclick="editCandidateSource('${source}')" title="Sửa">
                        <i class="fas fa-edit"></i>
                    </button>
                    <button class="btn btn-sm btn-danger" onclick="deleteCandidateSource('${source}')" title="Xóa">
                        <i class="fas fa-trash"></i>
                    </button>
                </td>
        `;
            tbody.appendChild(tr);
        });
    }

    function editCandidateSource(oldName) {
        Swal.fire({
            title: 'Sửa tên nguồn',
            input: 'text',
            inputValue: oldName,
            inputPlaceholder: 'Nhập tên mới',
            showCancelButton: true,
            confirmButtonText: 'Lưu',
            cancelButtonText: 'Hủy',
            inputValidator: (value) => {
                if (!value) {
                    return 'Vui lòng nhập tên nguồn!';
                }
            }
        }).then((result) => {
            if (result.isConfirmed && result.value) {
                const newName = result.value.trim();
                if (newName === oldName) return; // No change

                google.script.run.withSuccessHandler(function (res) {
                    if (res.success) {
                        Swal.fire('Thành công', 'Đã cập nhật tên nguồn', 'success');
                        loadCandidateSources();
                    } else {
                        Swal.fire('Lỗi', res.message, 'error');
                    }
                }).apiEditCandidateSource(oldName, newName);
            }
        });
    }

    function addCandidateSource() {
        const input = document.getElementById('new-source-name');
        const name = input.value.trim();
        if (!name) return Swal.fire('Lỗi', 'Vui lòng nhập tên nguồn', 'warning');

        google.script.run.withSuccessHandler(function (res) {
            if (res.success) {
                Swal.fire('Thành công', 'Đã thêm nguồn mới', 'success');
                input.value = '';
                loadCandidateSources();
            } else {
                Swal.fire('Lỗi', res.message, 'error');
            }
        }).apiAddCandidateSource(name);
    }

    function deleteCandidateSource(name) {
        Swal.fire({
            title: 'Xóa nguồn?',
            text: `Bạn có chắc muốn xóa nguồn "${name}" ? `,
            icon: 'warning',
            showCancelButton: true,
            confirmButtonColor: '#d33',
            confirmButtonText: 'Xóa'
        }).then((result) => {
            if (result.isConfirmed) {
                google.script.run.withSuccessHandler(function (res) {
                    if (res.success) {
                        Swal.fire('Đã xóa', '', 'success');
                        loadCandidateSources();
                    } else {
                        Swal.fire('Lỗi', res.message, 'error');
                    }
                }).apiDeleteCandidateSource(name);
            }
        });
    }

    function populateCandidateSources() {
        // Helper to fill a select element
        const fill = (id) => {
            const select = document.getElementById(id);
            if (!select) return;

            const currentVal = select.getAttribute('data-value') || select.value;
            const previouslySelected = select.value; // Keep current selection if valid

            select.innerHTML = '<option value="">Chọn nguồn</option>';

            candidateSourcesData.forEach(source => {
                const opt = document.createElement('option');
                opt.value = source;
                opt.text = source;
                if (source === currentVal || source === previouslySelected) opt.selected = true;
                select.appendChild(opt);
            });
        };

        fill('add-source');
        fill('detail-source');
    }

    // Hook into Tab Change
    document.addEventListener('shown.bs.tab', function (e) {
        if (e.target.getAttribute('href') === '#tab-sources') {
            loadCandidateSources();
        }
    });

    // INITIAL LOAD
    document.addEventListener('DOMContentLoaded', function () {
        // Initial fetch to ensure data is available for dropdowns
        loadCandidateSources();
    });
    // ============================================
    // NOTIFICATION DROPDOWN LOGIC
    // ============================================
    let currentNotifications = [];

    function loadNotifications() {
        if (!currentUser) return;

        // Show loading state if needed
        // document.getElementById('notification-list').innerHTML = '<li class="text-center p-3"><div class="spinner-border spinner-border-sm"></div></li>';

        google.script.run.withSuccessHandler(function (notifs) {
            currentNotifications = notifs;
            renderNotifications(notifs);
        }).apiGetNotifications(currentUser.username, currentUser.email);
    }

    function renderNotifications(notifs) {
        const badge = document.getElementById('notification-badge');
        const list = document.getElementById('notification-list');

        if (!list) return;

        // Update Badge
        if (notifs.length > 0) {
            badge.innerText = notifs.length;
            badge.style.display = 'inline-block';
        } else {
            badge.style.display = 'none';
        }

        // Update List
        list.innerHTML = '';
        if (notifs.length === 0) {
            list.innerHTML = '<li class="text-center p-3 text-muted">Không có thông báo mới</li>';
            return;
        }

        notifs.forEach(n => {
            const time = n.CreatedAt ? new Date(n.CreatedAt).toLocaleString('vi-VN') : '';
            let icon = 'fa-info-circle';
            // let bgClass = 'bg-light'; // Unused

            const isRead = n.IsRead;
            const bgClass = isRead ? 'bg-white' : 'bg-light'; // Read = White, Unread = Light Grey
            const textClass = isRead ? 'text-muted' : 'fw-bold'; // Unread = Bold

            if (n.Type === 'Evaluation') icon = 'fa-clipboard-check text-primary';
            if (n.Type === 'Mention') icon = 'fa-at text-warning';
            if (n.Type === 'Email') icon = 'fa-envelope text-info';

            const li = document.createElement('li');
            li.className = `p - 2 border - bottom notification - item ${bgClass} `;
            li.style.cursor = 'pointer';
            li.onclick = () => handleNotificationClick(n);

            li.innerHTML = `
            < div class="d-flex align-items-start" >
                    <div class="me-2 mt-1"><i class="fas ${icon}" style="width: 20px; text-align: center;"></i></div>
                    <div class="flex-grow-1">
                        <div class="small ${textClass}" style="line-height: 1.2;">${n.Message}</div>
                        <div class="text-muted" style="font-size: 0.7rem; margin-top: 2px;">${time}</div>
                    </div>
                    ${!isRead ? '<span class="badge bg-danger rounded-pill ms-1" style="font-size: 0.5rem;">New</span>' : ''}
                </div >
            `;
            list.appendChild(li);
        });

        // Update badge count (only count unread)
        const unreadCount = notifs.filter(n => !n.IsRead).length;
        badge.innerText = unreadCount;
        badge.style.display = unreadCount > 0 ? 'inline-block' : 'none';
    }

    function handleNotificationClick(n) {
        // 1. Mark as read
        google.script.run.withSuccessHandler(() => {
            // Reload to sync state
            loadNotifications();
        }).apiMarkNotificationRead(n.ID);

        // 2. Navigate
        if (n.Type === 'Evaluation') {
            // Updated: Click Evaluation notification -> Go to Evaluation Section
            if (document.getElementById('evaluations')) {
                showSection('evaluations', document.getElementById('nav-evaluations'));
            } else {
                // Fallback if section not found (e.g. permission issue)
                Swal.fire('Thông báo', 'Vui lòng truy cập mục Đánh giá PV để xem chi tiết.', 'info');
            }
        } else if (n.Type === 'Mention' || n.Type === 'Email') { // NEW
            // Open Candidate Detail
            if (n.RelatedId) {
                openCandidateDetail(n.RelatedId);
            }
        }
    }

    function markAllNotificationsRead() {
        if (!currentUser) return;
        if (!confirm('Đánh dấu tất cả là đã đọc?')) return;

        google.script.run.withSuccessHandler(() => {
            loadNotifications();
        }).apiMarkAllNotificationsRead(currentUser.username, currentUser.email);
    }

    // Auto-refresh notifications every 60s
    setInterval(() => {
        if (currentUser) loadNotifications();
    }, 60000);

    // ============================================
    // MENTION SUGGESTIONS LOGIC
    // ============================================
    // Ensure elements exist
    const mentionSuggestions = document.getElementById('mention-suggestions');
    const noteTeaxarea = document.getElementById('detail-new-note');

    // Make insertMention global so it can be called from onclick in HTML string
    window.insertMention = function (identifier) {
        if (!noteTeaxarea) return;

        const cursorPosition = noteTeaxarea.selectionStart;
        const textBefore = noteTeaxarea.value.substring(0, cursorPosition);
        const textAfter = noteTeaxarea.value.substring(cursorPosition);

        // Find where the @ started
        const lastAt = textBefore.lastIndexOf('@');

        if (lastAt !== -1) {
            // Identifier is either username or full name (with spaces?)
            // If full name has spaces, we should replace with underscores for robustness?
            // OR we just insert as is?
            // User wants "Chị Oanh".
            // Backend now supports finding by Slugified Name.
            // So let's insert "Chị_Oanh" or "Chi_Oanh"?
            // Let's insert "Chị_Oanh" (preserving accents but replacing spaces).

            const safeIdentifier = identifier.replace(/\s+/g, '_');

            // We replace from @ up to cursor with @Identifier + space
            const newTextBefore = textBefore.substring(0, lastAt) + '@' + safeIdentifier + ' ';
            noteTeaxarea.value = newTextBefore + textAfter;

            hideMentionSuggestions();
            noteTeaxarea.focus();

            const newPos = newTextBefore.length;
            noteTeaxarea.setSelectionRange(newPos, newPos);
        }
    };

    if (noteTeaxarea && mentionSuggestions) {
        noteTeaxarea.addEventListener('input', function (e) {
            const cursorPosition = this.selectionStart;
            const textBeforeCursor = this.value.substring(0, cursorPosition);

            // Regex to find @username at end of string or after a space
            // Case 1: Start of line: ^@...
            // Case 2: After space: \s@...
            // Or just check lastIndexOf('@') ?

            // Simple approach: Check if "word" at cursor starts with @
            // We want to capture the text after the last @
            const lastAt = textBeforeCursor.lastIndexOf('@');

            if (lastAt !== -1) {
                const textAfterAt = textBeforeCursor.substring(lastAt + 1);

                // Allow spaces and unicode (Vietnamese)
                // Stop if we hit a newline
                if (!textAfterAt.includes('\n')) {
                    const query = textAfterAt.toLowerCase();
                    // Optional: enforce max length to avoid searching whole paragraph
                    if (query.length < 50) {
                        showMentionSuggestions(query);
                        return;
                    }
                }
            }

            hideMentionSuggestions();
        });

        noteTeaxarea.addEventListener('keydown', function (e) {
            if (mentionSuggestions.style.display === 'block') {
                if (e.key === 'ArrowDown') {
                    e.preventDefault();
                    navigateSuggestions('down');
                } else if (e.key === 'ArrowUp') {
                    e.preventDefault();
                    navigateSuggestions('up');
                } else if (e.key === 'Enter') {
                    e.preventDefault();
                    selectSuggestion();
                } else if (e.key === 'Escape') {
                    hideMentionSuggestions();
                }
            }
        });

        // Close suggestions when clicking outside
        document.addEventListener('click', function (e) {
            if (e.target !== noteTeaxarea && !mentionSuggestions.contains(e.target)) {
                hideMentionSuggestions();
            }
        });
    }

    function showMentionSuggestions(query) {
        if (!usersData || usersData.length === 0) return;

        // Filter users
        const filteredUsers = usersData.filter(u => {
            const uName = String(u.Username).toLowerCase();
            const fName = u.Full_Name ? String(u.Full_Name).toLowerCase() : '';
            return uName.includes(query) || fName.includes(query);
        });

        if (filteredUsers.length === 0) {
            hideMentionSuggestions();
            return;
        }

        let html = '';
        // Limit to 5 suggestions
        filteredUsers.slice(0, 5).forEach((u, index) => {
            const activeClass = index === 0 ? 'active' : '';
            const fullName = u.Full_Name || u.Username;
            // Pass FullName if available, else Username
            const insertValue = u.Full_Name || u.Username;

            html += `< div class="mention-item ${activeClass}" onclick = "insertMention('${insertValue}')" >
                        <div class="fw-bold me-2">${fullName}</div>
                        <small class="text-muted">@${u.Username}</small>
                     </div > `;
        });

        mentionSuggestions.innerHTML = html;
        mentionSuggestions.style.display = 'block';

        // Position logic: Use Fixed to stick to visual position
        const rect = noteTeaxarea.getBoundingClientRect();

        mentionSuggestions.style.position = 'fixed';
        mentionSuggestions.style.top = (rect.bottom + 5) + 'px'; // 5px gap
        mentionSuggestions.style.left = rect.left + 'px';
        mentionSuggestions.style.width = Math.max(rect.width, 200) + 'px';
    }

    function hideMentionSuggestions() {
        if (mentionSuggestions) mentionSuggestions.style.display = 'none';
    }

    function navigateSuggestions(direction) {
        const items = mentionSuggestions.querySelectorAll('.mention-item');
        if (items.length === 0) return;

        let activeIndex = -1;
        items.forEach((item, index) => {
            if (item.classList.contains('active')) activeIndex = index;
            item.classList.remove('active');
        });

        if (direction === 'down') {
            activeIndex = (activeIndex + 1) % items.length;
        } else {
            activeIndex = (activeIndex - 1 + items.length) % items.length;
        }

        items[activeIndex].classList.add('active');
        items[activeIndex].scrollIntoView({ block: 'nearest' });
    }

    function selectSuggestion() {
        const activeItem = mentionSuggestions.querySelector('.mention-item.active');
        if (activeItem) {
            // Extract the value passed to insertMention in the onclick attribute requires parsing
            // Or easier: store it in data-attribute
            // But we didn't add data attribute in previous step.
            // Let's rely on simulated click.
            activeItem.click();
        }
    }


    // DOCUMENT HUB LOGIC
    function openDocumentHub() {
        const canID = document.getElementById('current-candidate-id').value;
        if (!canID) return Swal.fire('Lỗi', 'Không xác định được ứng viên', 'error');

        const bootstrapModal = new bootstrap.Modal(document.getElementById('documentHubModal'));
        bootstrapModal.show();

        // Default some data if available from detail form
        if (document.getElementById('doc-salary')) {
            document.getElementById('doc-salary').value = document.getElementById('detail-expected-salary') ? document.getElementById('detail-expected-salary').value : '';
        }
        if (document.getElementById('doc-manager')) {
            document.getElementById('doc-manager').value = document.getElementById('current-user-display').innerText;
        }

        // Populate Address and Signer dropdowns
        const locationSelect = document.getElementById('doc-location-select');
        const signerSelect = document.getElementById('doc-signer-select');

        if (locationSelect && signerSelect && initialData && initialData.companyInfo) {
            const info = initialData.companyInfo;

            locationSelect.innerHTML = '<option value="">-- Chọn địa điểm --</option>';
            if (info.addresses && Array.isArray(info.addresses)) {
                info.addresses.forEach(addr => {
                    const opt = document.createElement('option');
                    opt.value = addr;
                    opt.textContent = addr;
                    locationSelect.appendChild(opt);
                });
                if (info.addresses.length > 0) locationSelect.selectedIndex = 1;
            }

            signerSelect.innerHTML = '<option value="">-- Chọn người ký --</option>';
            if (info.signers && Array.isArray(info.signers)) {
                info.signers.forEach((s, idx) => {
                    const opt = document.createElement('option');
                    const name = typeof s === 'object' ? s.name : s;
                    const pos = typeof s === 'object' ? s.position : '';
                    opt.value = idx; // Use index to retrieve object later
                    opt.textContent = pos ? `${name} (${pos})` : name;
                    signerSelect.appendChild(opt);
                });
                if (info.signers.length > 0) signerSelect.selectedIndex = 1;
            }
        }

        // Populate Issuance Location if empty
        const issuanceLoc = document.getElementById('doc-issuance-location');
        if (issuanceLoc && !issuanceLoc.value && initialData && initialData.companyInfo && initialData.companyInfo.addresses) {
            issuanceLoc.value = initialData.companyInfo.addresses[0].split(',')[0].trim();
        }

        toggleDocExtraFields();
    }

    function toggleDocExtraFields() {
        const template = document.getElementById('doc-template-select').value;
        const fields = {
            'salary': document.getElementById('field-salary'),
            'start': document.getElementById('field-start-date'),
            'deadline': document.getElementById('field-deadline'),
            'manager': document.getElementById('field-manager'),
            'probationEnd': document.getElementById('field-probation-end'),
            'contractPeriod': document.getElementById('field-contract-period'),
            'issuanceLocation': document.getElementById('field-issuance-location')
        };

        // Reset display
        Object.values(fields).forEach(f => { if (f) f.style.display = 'block'; });

        if (template === 'TemplateCandidateProfile') {
            ['salary', 'start', 'deadline', 'manager', 'probationEnd', 'contractPeriod'].forEach(k => fields[k].style.display = 'none');
        } else if (template === 'TemplateOffer') {
            if (fields.probationEnd) fields.probationEnd.style.display = 'none';
            if (fields.contractPeriod) fields.contractPeriod.style.display = 'none';
        } else if (template === 'TemplateContractProbation') {
            if (fields.deadline) fields.deadline.style.display = 'none';
            if (fields.contractPeriod) fields.contractPeriod.style.display = 'none';
        } else if (template === 'TemplateContractOfficial') {
            if (fields.deadline) fields.deadline.style.display = 'none';
            if (fields.probationEnd) fields.probationEnd.style.display = 'none';
        }
    }

    function generateDocument() {
        const canID = document.getElementById('current-candidate-id').value;
        const template = document.getElementById('doc-template-select').value;
        const loader = document.getElementById('doc-loading');
        const btn = document.getElementById('btn-generate-doc');

        const startDateVal = document.getElementById('doc-start-date').value;
        const probationPeriod = 2; // Default

        // 1. Calculate Probation End Date (Auto or Manual)
        let probationEndDate = document.getElementById('doc-probation-end').value;
        if (!probationEndDate && startDateVal) {
            const sd = new Date(startDateVal);
            sd.setMonth(sd.getMonth() + probationPeriod);
            sd.setDate(sd.getDate() - 1);
            probationEndDate = sd.toISOString().split('T')[0];
        }

        const compInfo = (initialData && initialData.companyInfo) || {};
        const compName = compInfo.name || '';
        const compNameShort = compName.split(' ').map(w => w[0]).join('').toUpperCase();

        // 2. Get selected location and signer
        const locationSelect = document.getElementById('doc-location-select');
        const signerSelect = document.getElementById('doc-signer-select');
        const selectedLocation = locationSelect ? locationSelect.value : '';

        let selectedSignerName = '';
        let selectedSignerPos = '';
        if (signerSelect && signerSelect.value !== '' && initialData && initialData.companyInfo && initialData.companyInfo.signers) {
            const s = initialData.companyInfo.signers[parseInt(signerSelect.value)];
            if (s) {
                selectedSignerName = typeof s === 'object' ? s.name : s;
                selectedSignerPos = typeof s === 'object' ? s.position : '';
            }
        }

        // 3. Document Number Override
        const docNoOverride = document.getElementById('doc-number-override').value.trim();
        const contractNo = docNoOverride || Math.floor(1000 + Math.random() * 9000);

        const extraData = {
            'Salary': document.getElementById('doc-salary').value,
            'SalaryInWords': document.getElementById('doc-salary-words') ? document.getElementById('doc-salary-words').value : '',
            'StartDate': formatDateDisplay(startDateVal),
            'ProbationEndDate': formatDateDisplay(probationEndDate),
            'Deadline': formatDateDisplay(document.getElementById('doc-deadline').value),
            'ManagerName': document.getElementById('doc-manager').value,
            'ContractNo': contractNo,
            'ProbationPeriod': '02',
            'ContractPeriod': document.getElementById('doc-contract-period') ? document.getElementById('doc-contract-period').value : '',
            'CompanyNameShort': compNameShort,
            'CompanyAddress': selectedLocation,
            'CompanySignerName': selectedSignerName,
            'CompanySignerPosition': selectedSignerPos,
            'IssuanceLocation': document.getElementById('doc-issuance-location') ? document.getElementById('doc-issuance-location').value : ''
        };

        // 4. Calculate EndDate for Official Contract
        if (template === 'TemplateContractOfficial' && startDateVal) {
            let periodStr = extraData['ContractPeriod'] || '';
            let months = parseInt(periodStr);
            if (!isNaN(months)) {
                const ed = new Date(startDateVal);
                ed.setMonth(ed.getMonth() + months);
                ed.setDate(ed.getDate() - 1);
                extraData['EndDate'] = formatDateDisplay(ed.toISOString().split('T')[0]);
            }
        }

        // Calculate Probation Days for the new template
        if (startDateVal && probationEndDate) {
            const start = new Date(startDateVal);
            const end = new Date(probationEndDate);
            const diffTime = Math.abs(end - start);
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1; // Include start day
            extraData['ProbationPeriodDays'] = diffDays;
        } else {
            extraData['ProbationPeriodDays'] = '60'; // Default if dates missing
        }

        if (loader) loader.style.display = 'block';
        if (btn) btn.disabled = true;

        google.script.run.withSuccessHandler(function (res) {
            if (loader) loader.style.display = 'none';
            if (btn) btn.disabled = false;

            if (res.success) {
                Swal.fire({
                    title: 'Tạo văn bản thành công!',
                    text: 'File PDF đã được lưu trên Google Drive.',
                    icon: 'success',
                    showCancelButton: true,
                    confirmButtonText: 'Mở File ngay',
                    cancelButtonText: 'Đóng'
                }).then((result) => {
                    if (result.isConfirmed) {
                        window.open(res.url, '_blank');
                    }
                });
            } else {
                Swal.fire('Lỗi', res.message, 'error');
            }
        }).apiGenerateDocument(canID, template, extraData);
    }

    function generateCandidateProfilePDF() {
        const canID = document.getElementById('current-candidate-id').value;
        if (!canID) return Swal.fire('Lỗi', 'Không tìm thấy ID ứng viên', 'error');

        Swal.fire({
            title: 'Hệ thống đang xuất hồ sơ...',
            allowOutsideClick: false,
            didOpen: () => { Swal.showLoading(); }
        });

        google.script.run.withSuccessHandler(function (res) {
            Swal.close();
            if (res.success) {
                Swal.fire({
                    title: 'Đã xuất Hồ sơ PDF',
                    icon: 'success',
                    showCancelButton: true,
                    confirmButtonText: 'Xem File',
                    cancelButtonText: 'Để sau'
                }).then(r => {
                    if (r.isConfirmed) window.open(res.url, '_blank');
                });
            } else {
                Swal.fire('Lỗi', res.message, 'error');
            }
        }).apiGenerateDocument(canID, 'TemplateCandidateProfile', {});
    }

    function formatDateDisplay(dateStr) {
        if (!dateStr) return '';
        const parts = dateStr.split('-');
        if (parts.length === 3) return `${parts[2]} /${parts[1]}/${parts[0]}`;
        return dateStr;
    }

    // DATABASE CLEANUP UTILITY
    function runDatabaseCleanup() {
        Swal.fire({
            title: 'Xác nhận dọn dẹp?',
            text: "Hệ thống sẽ hợp nhất các cột trùng lặp và xóa các cột dư thừa trong Google Sheet. Bạn nên sao lưu file trước khi thực hiện.",
            icon: 'warning',
            showCancelButton: true,
            confirmButtonColor: '#d33',
            cancelButtonColor: '#3085d6',
            confirmButtonText: 'Đồng ý, dọn dẹp ngay',
            cancelButtonText: 'Hủy'
        }).then((result) => {
            if (result.isConfirmed) {
                const loader = document.getElementById('cleanup-loader');
                const btn = event.currentTarget;
                if (loader) loader.style.display = 'block';
                if (btn) btn.disabled = true;

                google.script.run.withSuccessHandler(function (res) {
                    if (loader) loader.style.display = 'none';
                    if (btn) btn.disabled = false;

                    if (res.success) {
                        Swal.fire({
                            icon: 'success',
                            title: 'Thành công!',
                            text: res.message + (res.details ? '\nDanh sách đã xóa: ' + res.details : ''),
                        }).then(() => {
                            // Reload everything to get fresh headers
                            location.reload();
                        });
                    } else {
                        Swal.fire('Kết quả', res.message, 'info');
                    }
                }).withFailureHandler(function (err) {
                    if (loader) loader.style.display = 'none';
                    if (btn) btn.disabled = false;
                    Swal.fire('Lỗi Server', err.toString(), 'error');
                }).apiCleanupCandidateSheet();
            }
        });
    }

    // ============================================
    // AI CV PARSER LOGIC (GEMINI)
    // ============================================
    async function handleAIParsing() {
        const fileInput = document.getElementById('detail-cv-file');
        if (fileInput.files.length === 0) {
            Swal.fire('Thông báo', 'Vui lòng chọn file CV (PDF) trước khi dùng AI. Bạn có thể nhấn nút "Upload" để chọn file PDF trước.', 'warning');
            return;
        }

        const file = fileInput.files[0];
        if (file.type !== 'application/pdf') {
            Swal.fire('Lỗi', 'Tính năng AI hiện chỉ hỗ trợ mạnh nhất cho file định dạng PDF thông qua OCR.', 'error');
            return;
        }

        // Hiển thị trạng thái loading
        document.getElementById('ai-loading').style.display = 'block';
        document.getElementById('btn-ai-parse').disabled = true;

        const reader = new FileReader();
        reader.onload = function (e) {
            const base64Data = e.target.result.split(',')[1];
            const fileData = {
                base64: base64Data,
                type: file.type,
                name: file.name
            };

            // Gọi lên Backend
            google.script.run
                .withSuccessHandler(function (response) {
                    document.getElementById('ai-loading').style.display = 'none';
                    document.getElementById('btn-ai-parse').disabled = false;

                    if (response.error) {
                        Swal.fire('Lỗi AI', response.error, 'error');
                    } else {
                        fillCandidateForm(response);
                        Swal.fire('Thành công', 'AI đã trích xuất dữ liệu xong! Vui lòng kiểm tra lại các thông tin trước khi lưu.', 'success');
                    }
                })
                .withFailureHandler(function (err) {
                    document.getElementById('ai-loading').style.display = 'none';
                    document.getElementById('btn-ai-parse').disabled = false;
                    Swal.fire('Lỗi hệ thống', 'Không thể kết nối với dịch vụ AI: ' + err.toString(), 'error');
                })
                .apiParseCV(fileData);
        };
        reader.readAsDataURL(file);
    }

    function fillCandidateForm(data) {
        if (data.Name) document.getElementById('detail-name').value = data.Name;
        if (data.Phone) document.getElementById('detail-phone').value = data.Phone;
        if (data.Email) document.getElementById('detail-email').value = data.Email;
        if (data.Birth_Year) document.getElementById('detail-dob').value = data.Birth_Year;

        if (data.Experience) document.getElementById('detail-experience').value = data.Experience;
        if (data.School) document.getElementById('detail-school').value = data.School;
        if (data.Major) document.getElementById('detail-major').value = data.Major;

        // Handle Position (Select dropdown)
        if (data.Position) {
            const posSelect = document.getElementById('detail-position');
            if (posSelect) {
                for (let i = 0; i < posSelect.options.length; i++) {
                    if (posSelect.options[i].text.toLowerCase().includes(data.Position.toLowerCase())) {
                        posSelect.selectedIndex = i;
                        break;
                    }
                }
            }
        }
    }

    // AI MATCHING
    function resetAiMatchUI() {
        const bar = document.getElementById('ai-score-bar');
        if (bar) {
            bar.style.width = '0%';
            bar.innerText = '0%';
            bar.classList.remove('bg-info', 'bg-success', 'bg-warning', 'bg-danger');
            document.getElementById('ai-score-text').innerText = '0%';
            document.getElementById('ai-pros').innerHTML = '';
            document.getElementById('ai-cons').innerHTML = '';
            document.getElementById('ai-summary').innerText = 'Chưa có dữ liệu phân tích.';
        }
    }

    function startAiMatching() {
        const candidateId = document.getElementById('current-candidate-id').value;
        const ticketId = document.getElementById('detail-ticket-id').value;

        if (!candidateId) return;
        if (!ticketId) {
            Swal.fire('Chú ý', 'Cần gắn ứng viên vào một Ticket (Yêu cầu tuyển dụng) để so sánh JD.', 'info');
            return;
        }

        const bar = document.getElementById('ai-score-bar');
        const text = document.getElementById('ai-score-text');

        // Show loading state in the bar
        bar.style.width = '100%';
        bar.innerText = 'Đang phân tích...';
        bar.classList.add('bg-info');

        google.script.run
            .withSuccessHandler(function (res) {
                if (res.success) {
                    const data = res.analysis;
                    bar.style.width = data.score + '%';
                    bar.innerText = data.score + '%';
                    bar.classList.remove('bg-info', 'bg-success', 'bg-warning', 'bg-danger');
                    bar.classList.add(data.score >= 75 ? 'bg-success' : (data.score >= 50 ? 'bg-warning' : 'bg-danger'));
                    text.innerText = data.score + '%';

                    document.getElementById('ai-pros').innerHTML = (data.pros || []).map(i => `<li>${i}</li>`).join('');
                    document.getElementById('ai-cons').innerHTML = (data.cons || []).map(i => `<li>${i}</li>`).join('');
                    document.getElementById('ai-summary').innerText = "Kết luận: " + (data.summary || "N/A");
                } else {
                    bar.style.width = '0%';
                    bar.innerText = 'Lỗi';
                    Swal.fire('Lỗi AI', res.message, 'error');
                }
            })
            .withFailureHandler(err => {
                bar.style.width = '0%';
                Swal.fire('Lỗi hệ thống', err.toString(), 'error');
            })
            .apiAnalyzeCandidateMatching(candidateId, ticketId);
    }

    // AI JD GENERATION
    function handleGenerateJD() {
        const title = document.getElementById('job-title').value;
        const currentDesc = document.getElementById('job-description').value;

        if (!title) {
            Swal.fire('Thông báo', 'Vui lòng nhập Tiêu đề công việc trước.', 'warning');
            return;
        }

        // Use a more robust SweetAlert config for input fields inside modals
        Swal.fire({
            title: 'AI Soạn JD',
            text: 'Dựa trên tiêu đề "' + title + '", AI sẽ soạn thảo bản mô tả chi tiết. Bạn có muốn bổ sung yêu cầu gì không?',
            input: 'text',
            inputPlaceholder: 'VD: Cần 2 năm kinh nghiệm, biết tiếng Anh...',
            showCancelButton: true,
            confirmButtonText: 'Soạn thảo ngay',
            showLoaderOnConfirm: true,
            backdrop: false, // Prevent backdrop issues
            target: document.getElementById('addJobModal'), // Render inside the modal
            preConfirm: (req) => {
                return new Promise((resolve, reject) => {
                    google.script.run
                        .withSuccessHandler(res => {
                            if (res && res.success && res.jd) {
                                // Strip Markdown (bold **, headers ##)
                                res.jd = res.jd.replace(/\*\*/g, '').replace(/###/g, '').replace(/##/g, '').replace(/#/g, '');
                                resolve(res);
                            } else {
                                resolve(res);
                            }
                        })
                        .withFailureHandler(err => reject(err))
                        .apiGenerateJD(title, req || currentDesc);
                });
            },
            allowOutsideClick: () => !Swal.isLoading()
        }).then((result) => {
            if (result.isConfirmed) {
                if (result.value && result.value.success) {
                    document.getElementById('job-description').value = result.value.jd;
                    Swal.fire({
                        icon: 'success',
                        title: 'Hoàn tất',
                        text: 'Bản JD đã được AI soạn thảo!',
                        target: document.getElementById('addJobModal')
                    });
                } else {
                    Swal.fire({
                        icon: 'error',
                        title: 'Lỗi',
                        text: result.value ? result.value.message : 'Lỗi không xác định',
                        target: document.getElementById('addJobModal')
                    });
                }
            }
        });
    }

    // ============================================
    // CALENDAR INTERVIEW LOGIC
    // ============================================
    function handleCreateSchedule() {
        const scheduleData = {
            candidateName: document.getElementById('detail-name').value,
            candidateEmail: document.getElementById('detail-email').value,
            managerEmail: document.getElementById('interview-manager-email').value,
            startTime: document.getElementById('interview-start').value,
            endTime: document.getElementById('interview-end').value,
            location: "Vòng 1 - Văn phòng bTaskee",
            note: "Phỏng vấn chuyên môn vòng 1"
        };

        if (!scheduleData.startTime || !scheduleData.endTime || !scheduleData.managerEmail) {
            Swal.fire('Chú ý', 'Vui lòng chọn đầy đủ thời gian và nhập Email Manager', 'warning');
            return;
        }

        Swal.fire({
            title: 'Đang đặt lịch...',
            text: 'Vui lòng chờ trong giây lát',
            allowOutsideClick: false,
            didOpen: () => { Swal.showLoading(); }
        });

        google.script.run
            .withSuccessHandler(function (res) {
                if (res.success) {
                    Swal.fire('Thành công', res.message || 'Lịch hẹn đã được tạo và gửi email mời họp!', 'success');
                } else {
                    Swal.fire('Lỗi', res.message, 'error');
                }
            })
            .withFailureHandler(function (err) {
                Swal.fire('Lỗi hệ thống', err.toString(), 'error');
            })
            .apiCreateInterviewSchedule(scheduleData);
    }
    // Helper to render notes as bubbles
    function renderNoteHistory(notesText, container) {
        if (!container) return;
        if (!notesText) {
            container.innerHTML = '<div class="text-center text-muted py-3">Chưa có lịch sử ghi chú.</div>';
            return;
        }

        const lines = notesText.split('\n');
        let html = '';

        lines.forEach(line => {
            if (!line.trim()) return;

            // Format: [HH:mm dd/MM/yyyy] User: Content
            const match = line.match(/^\[(.*?)\] (.*?): (.*)$/);
            if (match) {
                const [_, time, user, content] = match;
                const isCurrentUser = (currentUser && (user === currentUser.username || user === currentUser.name));

                html += `
                    <div class="note-bubble ${isCurrentUser ? 'note-bubble-admin' : ''}">
                        <div class="note-header">
                            <span class="note-user">${user}</span>
                            <span class="note-time">${time}</span>
                        </div>
                        <div class="note-content">${content}</div>
                    </div>
                `;
            } else {
                // Fallback for legacy or unstructured notes
                html += `
                    <div class="note-bubble">
                        <div class="note-content">${line}</div>
                    </div>
                `;
            }
        });

        container.innerHTML = html;
        // Scroll to bottom
        container.scrollTop = container.scrollHeight;
    }
    // Aliases for compatibility with different button clicks
    function saveCandidate() { saveCandidateDetail(); }

    function openSendEmail() {
        const id = document.getElementById('current-candidate-id').value;
        if (id) openSendEmailModal(id);
    }

    // EXPORT CANDIDATE PROFILE PDF
    function exportCandidateToPDF() {
        const id = document.getElementById('current-candidate-id').value;
        if (!id) return;

        Swal.fire({
            title: 'Đang chuẩn bị...',
            text: 'Hệ thống đang trích xuất dữ liệu và tạo PDF',
            allowOutsideClick: false,
            didOpen: () => { Swal.showLoading(); }
        });

        google.script.run
            .withSuccessHandler(function (res) {
                if (res.success) {
                    Swal.fire({
                        icon: 'success',
                        title: 'Đã tạo xong!',
                        text: 'File của bạn đang được tải xuống...',
                        timer: 2000,
                        showConfirmButton: false
                    });

                    if (res.url) {
                        window.open(res.url, '_blank');
                    }
                } else {
                    Swal.fire('Lỗi', res.message, 'error');
                }
            })
            .withFailureHandler(function (err) {
                Swal.fire('Lỗi hệ thống', err.message, 'error');
            })
            .apiGenerateDocument(id, 'TemplateCandidateProfile', {});
    }

    // DOCUMENT HUB LOGIC
    function openDocumentHub(candidateId = null) {
        if (!candidateId) {
            candidateId = document.getElementById('current-candidate-id').value;
        }
        if (!candidateId) return;

        const modal = document.getElementById('documentHubModal');
        // Pre-fill hidden or data attribute if needed
        modal.setAttribute('data-candidate-id', candidateId);

        // Reset sub-fields in Hub Modal
        const salaryField = document.getElementById('doc-salary');
        if (salaryField) {
            const c = candidatesData.find(x => x.ID == candidateId);
            if (c) salaryField.value = c.Salary_Expectation || '';
        }

        new bootstrap.Modal(modal).show();
    }

    function generateDocument() {
        const modal = document.getElementById('documentHubModal');
        const candidateId = modal.getAttribute('data-candidate-id');
        const template = document.getElementById('doc-template-select').value;

        if (!candidateId || !template) {
            Swal.fire('Lỗi', 'Thông tin không hợp lệ', 'warning');
            return;
        }

        const extraData = {
            docNumber: document.getElementById('doc-number-override')?.value,
            salary: document.getElementById('doc-salary')?.value,
            salaryWords: document.getElementById('doc-salary-words')?.value,
            startDate: document.getElementById('doc-start-date')?.value,
            probationEnd: document.getElementById('doc-probation-end')?.value,
            deadline: document.getElementById('doc-deadline')?.value,
            contractPeriod: document.getElementById('doc-contract-period')?.value,
            location: document.getElementById('doc-location-select')?.value,
            signer: document.getElementById('doc-signer-select')?.value,
            manager: document.getElementById('doc-manager')?.value
        };

        const btn = document.getElementById('btn-generate-doc');
        const originalText = btn.innerHTML;
        btn.innerHTML = '<span class="spinner-border spinner-border-sm me-2"></span> Đang tạo...';
        btn.disabled = true;

        google.script.run
            .withSuccessHandler(function (res) {
                btn.innerHTML = originalText;
                btn.disabled = false;
                if (res.success) {
                    Swal.fire({
                        icon: 'success',
                        title: 'Tạo văn bản thành công',
                        text: 'Bấm OK để xem file PDF',
                    }).then(() => {
                        window.open(res.url, '_blank');
                    });
                } else {
                    Swal.fire('Lỗi', res.message, 'error');
                }
            })
            .withFailureHandler(function (err) {
                btn.innerHTML = originalText;
                btn.disabled = false;
                Swal.fire('Lỗi hệ thống', err.message, 'error');
            })
            .apiGenerateDocument(candidateId, template, extraData);
    }

</script>
