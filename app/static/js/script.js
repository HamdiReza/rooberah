

const VALID_ROLES = ['admin', 'supervisor', 'expert', 'monitoring', 'agency'];
const ROLE_PAGE_MAP = { admin: 'dashboard', supervisor: 'dashboard', expert: 'dashboard', monitoring: 'monitoring', agency: 'agency' };
const ROLE_DEFAULT_PAGE = { admin: 'dashboard', supervisor: 'dashboard', expert: 'dashboard', monitoring: 'monitoring', agency: 'agency' };
const ROLE_DEFAULT_VIEW = { admin: 'views/view-dashboard', supervisor: 'views/view-dashboard', expert: 'views/view-dashboard', monitoring: 'views/monitoring', agency: 'views/agency' };

function restoreLastViewByRole(role) {
    const normalizedRole = (role || '').toLowerCase();
    const lastView = localStorage.getItem('lastView');

    if (lastView) {
        const normalizedView = (lastView === 'view-assignment' || lastView === 'view-comparison')
            ? 'view-users'
            : lastView;
        app.switchView(normalizedView);
        return;
    }

    const fallbackView = ROLE_DEFAULT_VIEW[normalizedRole] || 'view-dashboard';
    if (fallbackView === 'monitoring' || fallbackView === 'agency') {
        app.navigateTo(fallbackView, null, false);
    } else {
        app.switchView(fallbackView);
    }
}


// ===================================================================================
// === بخش تعاریف و متغیرهای اصلی (بدون تغییر) ===
// ===================================================================================
let isAdmin = false;
let globalData = [];
let globalHeaders = [];
let autoRefreshInterval = null;
let allUsers = [];
let pricingData = [];
let agencyData = []; // Store agency data separately

const COLUMN_INDEXES = {
    ROW_NUM: 0,
    EXPERT_NAME: 1,
    PHONE: 2,
    FULL_NAME: 3,
    NATIONAL_ID: 4,   // ستون جدید
    CUSTOMER_TYPE: 5,
    PRODUCT: 6,
    DATE: 7,
    CALL: 8,
    CALL_RESULT: 9,
    SMART_CARD: 10,
    AGE: 11,
    CREDIT_RATING: 12,
    SCRAP_TRUCK: 13,
    GUARANTEES: 14,
    SAYYAD_CHECK: 15,
    SCRAPPING: 16,
    PREPAYMENT: 17,
    MORTGAGE: 18,
    CONTRACT: 19,
    DELIVERY: 20,
    CANCELLATION: 21,

    FILE_NATIONAL_ID: 22,
    FILE_SAHA: 23,
    FILE_RECEIPT: 24,

    CALC_BRAND: 25,
    CALC_AXLE: 26,
    CALC_USAGE: 27,
    CALC_DESC: 28,
    CALC_PAYMENT: 29,
    CALC_PRICE: 30,

    CUSTOMER_PAID_PREPAYMENT: 31,
    TARGET_PREPAYMENT: 32,
    ASSIGNED_BY: 33,

    SHEET_ROW_INDEX: 34
};



const SELECTABLE_FIELDS = [6, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21];
const HIDDEN_COLUMNS_IN_TABLE = [7, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34];


const CUSTOM_OPTIONS = {
    [COLUMN_INDEXES.CALL]: ["۱", "۲", "بیشتر", "عدم پاسخگویی"],
    [COLUMN_INDEXES.CALL_RESULT]: [
        "تماس گرفته نشده",
        "پاسخگو نبوده",
        "در حال بررسی توسط مشتری",
        "در پروسه انجام",
        "انصراف",
        "عدم پاسخگویی",
        "سایر"
    ],
    [COLUMN_INDEXES.SMART_CARD]: ["دارد", "ندارد", "سایر اعضای خانواده"],
    [COLUMN_INDEXES.AGE]: ["بالای ۶۵", "پایین ۶۵"],
    [COLUMN_INDEXES.CREDIT_RATING]: ["A", "B", "C", "D", "E", "نمی داند"],
    [COLUMN_INDEXES.SCRAP_TRUCK]: ["دارد", "ندارد", "تریلی قیمت بالا", "سایر"],
    [COLUMN_INDEXES.GUARANTEES]: ["کشنده", "ملک", "سایر"],
    [COLUMN_INDEXES.SAYYAD_CHECK]: ["دارد", "ندارد", "سایر اعضا"],
    [COLUMN_INDEXES.SCRAPPING]: ["درحال انجام", "درحال پرداخت", "برگه صادر شده", "بخشی از پیش پرداخت"],
    [COLUMN_INDEXES.PREPAYMENT]: ["درحال پرداخت", "پرداخت شد"],
    [COLUMN_INDEXES.MORTGAGE]: ["در حال انجام", "انجام شد", "استرداد وجه"],
    [COLUMN_INDEXES.CONTRACT]: ["در حال انجام", "انجام شد", "استرداد وجه"],
    [COLUMN_INDEXES.DELIVERY]: ["نشد", "شد"],
    [COLUMN_INDEXES.CANCELLATION]: ["قیمت بالای کالا", "سود بالای اقساط", "دلایل شخصی", "سایر", "عدم داشتن پیش پرداخت", "عدم داشتن تضامین", "عدم توان پرداخت اقساط"]
};

const app = {
    // فرمت ساده تاریخ برای نمایش در کلاینت
    formatClientDateTime: function (ts) {
        if (!ts) return '';
        try {
            const d = ts instanceof Date ? ts : new Date(ts);
            const pad = n => (n < 10 ? '0' + n : '' + n);
            const y = d.getFullYear();
            const m = pad(d.getMonth() + 1);
            const day = pad(d.getDate());
            const hh = pad(d.getHours());
            const mm = pad(d.getMinutes());
            return `${y}/${m}/${day} ${hh}:${mm}`;
        } catch (e) {
            return '';
        }
    },
    // نرمال‌سازی متن جستجو (حذف فاصله اضافی، کوچک‌نوشتن، تبدیل اعداد فارسی/عربی به انگلیسی)
    normalizeSearchText: function (txt) {
        if (!txt) return '';
        let s = String(txt).trim().toLowerCase();
        const persianDigits = [/۰/g, /۱/g, /۲/g, /۳/g, /۴/g, /۵/g, /۶/g, /۷/g, /۸/g, /۹/g];
        const arabicDigits = [/٠/g, /١/g, /٢/g, /٣/g, /٤/g, /٥/g, /٦/g, /٧/g, /٨/g, /٩/g];
        for (let i = 0; i < 10; i++) {
            s = s.replace(persianDigits[i], String(i));
            s = s.replace(arabicDigits[i], String(i));
        }
        return s;
    },
    data: [],
    headers: [],
    isAdmin: false,
    chartBar: null,
    chartPie: null,
    activeExpertSection: 'manage',
    expertComparisonChart: null,
    changeExpertData: [],
    filteredChangeExpertData: [],
    processTransfers: {},
    currentFilteredData: [],
    segmentFilters: {
        "new-calls": ["تماس گرفته نشده", "پاسخگو نبوده", ""],
        "under-review": ["در حال بررسی توسط مشتری", "درحال بررسی توسط مشتری", "درحال بررسی توسط مستری"],
        "following": ["در پروسه انجام", "درپروسه انجام"],
        "cancelled": ["انصراف"]
    },

    getOptions: function (columnIndex) {
        let options = [{ value: "", label: "- انتخاب کنید -" }];
        if (CUSTOM_OPTIONS[columnIndex]) {
            CUSTOM_OPTIONS[columnIndex].forEach(opt => options.push({ value: opt, label: opt }));
        }
        return options;
    },

    doLogin: function () {
        const user = document.getElementById('loginUser').value;
        const pass = document.getElementById('loginPass').value;
        const btn = document.getElementById('loginBtn');
        const spinner = document.getElementById('loginSpinner');
        const err = document.getElementById('loginError');

        if (!user || !pass) { err.innerText = "لطفا نام کاربری و رمز عبور را وارد کنید"; return; }

        btn.disabled = true; btn.classList.add('opacity-80'); spinner.classList.remove('hidden'); err.innerText = "";

        google.script.run.withSuccessHandler(res => {
            btn.disabled = false; btn.classList.remove('opacity-80'); spinner.classList.add('hidden');
            if (res.success) {
                sessionStorage.setItem('appUser', res.username);
                sessionStorage.setItem('appName', res.fullName);
                const normalizedRole = VALID_ROLES.includes((res.role || '').toLowerCase()) ? (res.role || '').toLowerCase() : 'expert';
                sessionStorage.setItem('role', normalizedRole);
                if (res.supervisor) {
                    sessionStorage.setItem('supervisor', String(res.supervisor || '').toLowerCase().trim());
                } else {
                    sessionStorage.removeItem('supervisor');
                }

                const loginPage = document.getElementById('login');
                if (loginPage) loginPage.classList.remove('active');

                const defaultPage = ROLE_DEFAULT_PAGE[normalizedRole] || 'dashboard';
                app.applyRoleVisibility(normalizedRole);
                app.navigateTo(defaultPage);
            } else {
                err.innerText = res.message;
            }
        }).withFailureHandler(() => {
            btn.disabled = false; spinner.classList.add('hidden'); err.innerText = "خطا در اتصال";
        }).loginUser(user, pass);
    },

    logout: function () {
        sessionStorage.clear();
        document.getElementById('loginUser').value = "";
        document.getElementById('loginPass').value = "";
        app.navigateTo('login');
    },

    // === تب پیام سرپرست برای کارشناس ===
    expertMessagesActiveStatus: 'unread',

    showExpertMessagesTab: function (status) {
        const allowed = ['unread', 'read', 'all'];
        if (!allowed.includes(status)) status = 'unread';
        app.expertMessagesActiveStatus = status;

        ['unread', 'read', 'all'].forEach(s => {
            const btn = document.getElementById(`expertMessagesFilter-${s}`);
            if (btn) {
                if (s === status) {
                    btn.classList.add('bg-indigo-600', 'text-white');
                    btn.classList.remove('bg-slate-100', 'dark:bg-slate-700', 'text-slate-600', 'dark:text-slate-300');
                } else {
                    btn.classList.remove('bg-indigo-600', 'text-white');
                    btn.classList.add('bg-slate-100', 'dark:bg-slate-700', 'text-slate-600', 'dark:text-slate-300');
                }
            }
        });

        app.loadSupervisorRepliesForExpert(status);
    },

    loadSupervisorRepliesForExpert: function (status) {
        const expert = (sessionStorage.getItem('appUser') || '').toLowerCase();
        if (!expert) return;

        const normalizedStatus = status || app.expertMessagesActiveStatus || 'unread';
        const list = document.getElementById('expertSupervisorMessagesList');
        const empty = document.getElementById('expertSupervisorMessagesEmpty');
        const loader = document.getElementById('expertSupervisorMessagesLoading');

        if (list) {
            list.innerHTML = '';
            list.classList.add('hidden');
        }
        if (empty) empty.classList.add('hidden');
        if (loader) loader.classList.remove('hidden');

        google.script.run.withSuccessHandler(res => {
            if (loader) loader.classList.add('hidden');
            if (!res || !res.success || !Array.isArray(res.messages) || res.messages.length === 0) {
                if (empty) empty.classList.remove('hidden');
                return;
            }
            if (empty) empty.classList.add('hidden');

            let html = '';
            res.messages.forEach(msg => {
                const dateStr = app.formatClientDateTime(msg.timestamp);
                const preview = (msg.text || '').length > 80 ? (msg.text || '').substring(0, 80) + '…' : (msg.text || '');
                const isUnread = !msg.readByExpert;
                html += `
                        <button
                          type="button"
                          class="w-full text-right px-3 py-2 rounded-xl border border-slate-200 dark:border-slate-700 ${isUnread ? 'bg-indigo-50 dark:bg-indigo-900/20' : 'bg-white/70 dark:bg-slate-800/70'} hover:bg-indigo-100 dark:hover:bg-indigo-900/40 transition flex flex-col gap-1"
                          onclick="app.selectExpertSupervisorMessage('${msg.messageId}')">
                          <div class="flex items-center justify-between gap-2">
                            <span class="text-xs font-bold text-slate-700 dark:text-slate-200">
                              پیام سرپرست
                            </span>
                            <span class="text-[10px] text-slate-400 dark:text-slate-500">
                              ${dateStr}
                            </span>
                          </div>
                          <p class="text-xs text-slate-600 dark:text-slate-300 line-clamp-2">${preview}</p>
                        </button>
                      `;
            });
            if (list) {
                list.innerHTML = html;
                list.classList.remove('hidden');
            }
        }).withFailureHandler(() => {
            if (loader) loader.classList.add('hidden');
            if (empty) {
                empty.classList.remove('hidden');
                empty.innerText = 'خطا در بارگذاری پیام‌ها';
            }
        }).getSupervisorRepliesForExpert(expert, normalizedStatus);
    },

    expertSelectedSupervisorMessage: null,

    selectExpertSupervisorMessage: function (messageId) {
        const expert = (sessionStorage.getItem('appUser') || '').toLowerCase();
        if (!expert || !messageId) return;

        const status = app.expertMessagesActiveStatus || 'unread';
        google.script.run.withSuccessHandler(res => {
            if (!res || !res.success || !Array.isArray(res.messages)) return;
            const msg = res.messages.find(m => m.messageId === messageId);
            if (!msg) return;
            app.expertSelectedSupervisorMessage = msg;

            const titleEl = document.getElementById('expertSupervisorSelectedTitle');
            const metaEl = document.getElementById('expertSupervisorSelectedMeta');
            const bodyEl = document.getElementById('expertSupervisorSelectedBody');

            if (titleEl) titleEl.innerText = 'جزئیات پیام سرپرست';
            if (metaEl) {
                const dateStr = app.formatClientDateTime(msg.timestamp);
                metaEl.innerText = `مرتبط با ردیف CRM: ${msg.relatedRowIndex || '-'} | ${dateStr}`;
            }
            if (bodyEl) bodyEl.innerText = msg.text || '';
        }).withFailureHandler(() => {
            // نادیده گرفتن
        }).getSupervisorRepliesForExpert(expert, status);
    },

    markExpertSupervisorMessageAsRead: function () {
        const msg = app.expertSelectedSupervisorMessage;
        if (!msg || !msg.messageId) return;
        const statusLabel = document.getElementById('expertSupervisorMarkStatus');
        if (statusLabel) {
            statusLabel.innerText = 'در حال بروزرسانی...';
            statusLabel.classList.remove('text-rose-500');
        }
        google.script.run.withSuccessHandler(() => {
            if (statusLabel) {
                statusLabel.innerText = 'پیام به‌عنوان خوانده‌شده علامت خورد';
            }
            app.loadSupervisorRepliesForExpert(app.expertMessagesActiveStatus || 'unread');
        }).withFailureHandler(() => {
            if (statusLabel) {
                statusLabel.innerText = 'خطا در بروزرسانی وضعیت پیام';
                statusLabel.classList.add('text-rose-500');
            }
        }).markMessageRead(msg.messageId, 'expert');
    },

    navigateTo: function (pageId, param = null, pushHistory = true) {
        const u = sessionStorage.getItem('appUser');
        if (!u && pageId !== 'login') { pageId = 'login'; pushHistory = false; }

        const rolePages = ['admin', 'supervisor', 'expert', 'monitoring', 'agency'];
        const isRoleDashboard = rolePages.includes(pageId);
        const isDashboardLike = pageId === 'dashboard' || pageId === 'supervisor' || pageId === 'expert';

        if (isDashboardLike || isRoleDashboard) {
            if (pageId === 'agency') {
                app.startAgencyAutoRefresh();
            } else {
                app.startAutoRefresh();
            }
        } else {
            app.stopAutoRefresh();
        }

        const allPages = document.querySelectorAll('.page-section');
        const targetPage = document.getElementById(pageId);

        if (!targetPage) {
            if (u && isRoleDashboard) {
                const role = (sessionStorage.getItem('role') || '').toLowerCase();
                const fallback = ROLE_PAGE_MAP[role] || 'dashboard';
                const fallbackEl = document.getElementById(fallback);
                if (fallbackEl) {
                    allPages.forEach(el => el.classList.remove('active'));
                    fallbackEl.classList.add('active');
                    window.scrollTo(0, 0);
                    if (pushHistory) try { window.history.pushState({ page: fallback, param: param }, "", "#" + fallback); } catch (e) { }
                    return;
                }
            }
            console.warn('Page not found:', pageId);
            return;
        }

        allPages.forEach(el => el.classList.remove('active'));
        targetPage.classList.add('active');
        window.scrollTo(0, 0);

        try { if (pushHistory) window.history.pushState({ page: pageId, param: param }, "", "#" + pageId); } catch (e) { }
        if (param) sessionStorage.setItem('currentDetailParam', param);

        if (isDashboardLike) {
            const n = sessionStorage.getItem('appName');
            const role = (sessionStorage.getItem('role') || '').toLowerCase();
            app.isAdmin = role === 'admin';
            const greetingEl = document.getElementById('greetingMsg');
            const usernameEl = document.getElementById('usernameDisplay');
            if (greetingEl) greetingEl.innerText = `سلام، ${n}`;
            if (usernameEl) usernameEl.innerText = `@${u}`;

            const els = ['sidebar-experts-label', 'expertListContainer', 'adminMenuContainer'];
            const allowExpertMenu = role === 'admin' || role === 'supervisor';
            els.forEach(id => {
                const el = document.getElementById(id);
                if (el) allowExpertMenu ? el.classList.remove('hidden') : el.classList.add('hidden');
            });
            if (allowExpertMenu) app.loadUsers();

            app.switchView('view-dashboard');
            app.refreshTable(false);

            google.script.run.withSuccessHandler(d => {
                pricingData = d;
                app.initCalculator();
            }).getPricingData();

        } else if (pageId === 'agency') {
            const n = sessionStorage.getItem('appName');
            const role = (sessionStorage.getItem('role') || '').toLowerCase();

            // Check if user has access (supervisor or admin only)

            const greetingEl = document.getElementById('agencyGreetingMsg');
            const usernameEl = document.getElementById('agencyUsernameDisplay');
            if (greetingEl) greetingEl.innerText = `پنل نمایندگی - ${n}`;
            if (usernameEl) usernameEl.innerText = `@${u}`;

            app.refreshAgencyTable(false);
            app.startAgencyAutoRefresh();

        } else if (pageId === 'details') {
            const rowId = param || sessionStorage.getItem('currentDetailParam') || sessionStorage.getItem('customerRow');
            if (rowId) loadDetails(rowId);
        }
    },

    applyRoleVisibility: function (role) {
        const roleLower = (role || '').toLowerCase();

        // Elements that should only be visible to supervisors
        const supervisorOnlyElements = [
            'adminMenuContainer',           // Expert management tab
            'sidebar-experts-label',        // Expert list label
            'expertListContainer'           // Expert list container
        ];

        app.isAdmin = roleLower === 'admin';
        const allowExpertMenu = roleLower === 'admin' || roleLower === 'supervisor';

        supervisorOnlyElements.forEach(id => {
            const el = document.getElementById(id);
            if (el) {
                allowExpertMenu ? el.classList.remove('hidden') : el.classList.add('hidden');
            }
        });

        if (allowExpertMenu) {
            app.loadUsers();
        }

        // نمایش تب پیام‌های سرپرست فقط برای کارشناس
        const expertMessagesMenu = document.getElementById('menu-view-expert-messages');
        if (expertMessagesMenu) {
            if (roleLower === 'expert') {
                expertMessagesMenu.classList.remove('hidden');
            } else {
                expertMessagesMenu.classList.add('hidden');
            }
        }
    },

    switchView: function (viewId) {
        document.querySelectorAll('.view-section').forEach(el => el.classList.add('hidden'));
        document.getElementById(viewId).classList.remove('hidden');
        document.querySelectorAll('.sidebar-link').forEach(el => el.classList.remove('active-link'));
        const menuId = 'menu-' + viewId;
        const menuEl = document.getElementById(menuId);
        if (menuEl) menuEl.classList.add('active-link');

        if (viewId === 'view-users') {
            const role = (sessionStorage.getItem('role') || '').toLowerCase();
            const isAdminUser = role === 'admin';
            if (role !== 'supervisor' && !isAdminUser) {
                alert("شما مجوز دسترسی به مدیریت کارشناس‌ها را ندارید.");
                return;
            }
            this.showExpertSection(this.activeExpertSection || 'manage');
        } else if (viewId === 'view-expert-messages') {
            // تب پیام سرپرست برای کارشناس
            app.showExpertMessagesTab(app.expertMessagesActiveStatus || 'unread');
        }
    },

    showExpertSection: function (section) {
        const sections = ['manage', 'compare', 'assign'];
        if (!sections.includes(section)) section = 'manage';
        this.activeExpertSection = section;

        sections.forEach(sec => {
            const sectionEl = document.getElementById(`expert-section-${sec}`);
            const btnEl = document.getElementById(`expert-tab-${sec}`);

            if (sectionEl) {
                if (sec === section) {
                    sectionEl.classList.remove('hidden');
                } else {
                    sectionEl.classList.add('hidden');
                }
            }

            if (btnEl) {
                const activeClasses = ['bg-blue-50', 'text-blue-700', 'dark:bg-blue-900/30', 'dark:text-blue-200', 'border', 'border-blue-200'];
                const inactiveClasses = ['border', 'border-transparent', 'text-slate-700', 'dark:text-slate-200'];

                if (sec === section) {
                    btnEl.classList.remove('border-transparent');
                    activeClasses.forEach(c => btnEl.classList.add(c));
                } else {
                    activeClasses.forEach(c => btnEl.classList.remove(c));
                    inactiveClasses.forEach(c => btnEl.classList.add(c));
                }
            }
        });

        if (section === 'assign') {
            this.loadAssignmentView();
            this.loadChangeExpertData();
        } else if (section === 'compare') {
            this.renderExpertComparison();
        }
    },

    // ===========================
    // تغییر کارشناس مشتری
    // ===========================
    loadChangeExpertData: function () {
        const role = (sessionStorage.getItem('role') || '').toLowerCase();
        const isAdminUser = role === 'admin';
        if (role !== 'supervisor' && !isAdminUser) return;

        const refreshBtn = document.getElementById('changeExpertRefreshBtn');
        if (refreshBtn) {
            refreshBtn.disabled = true;
            refreshBtn.classList.add('opacity-60');
        }

        google.script.run.withSuccessHandler((rows) => {
            this.changeExpertData = Array.isArray(rows) ? rows : [];
            this.prepareChangeExpertFilters();
            this.filterChangeExpertTable();
        }).withFailureHandler((err) => {
            this.showChangeExpertToast(err && err.message ? err.message : 'خطا در بارگذاری داده‌ها', true);
        }).getLeadsForSupervisor({
            username: sessionStorage.getItem('appUser') || '',
            role,
            displayName: sessionStorage.getItem('appName') || ''
        });

        if (refreshBtn) {
            setTimeout(() => {
                refreshBtn.disabled = false;
                refreshBtn.classList.remove('opacity-60');
            }, 400);
        }
    },

    prepareChangeExpertFilters: function () {
        const filterEl = document.getElementById('changeExpertFilter');
        if (!filterEl) return;

        const experts = [...new Set(this.changeExpertData.map(r => r.expertName).filter(Boolean))];
        filterEl.innerHTML = `<option value="">فیلتر بر اساس کارشناس</option>` + experts.map(e => `<option value="${e}">${e}</option>`).join('');
    },

    filterChangeExpertTable: function () {
        const search = (document.getElementById('changeExpertSearch')?.value || '').trim().toLowerCase();
        const expertFilter = (document.getElementById('changeExpertFilter')?.value || '').trim();

        const filtered = this.changeExpertData.filter(row => {


            const supervisorName = sessionStorage.getItem('appUser') || '';
            const assignedBy = row.assignedBy || '';
            // نمایش فقط زیرمجموعه همین سرپرست (assigned_by یا نقش فعلی admin)
            const role = (sessionStorage.getItem('role') || '').toLowerCase();
            const isAdminUser = role === 'admin';
            if (!isAdminUser && assignedBy && assignedBy.toLowerCase() !== supervisorName.toLowerCase()) return false;

            if (expertFilter && row.expertName !== expertFilter) return false;

            if (search.length > 0) {
                const phone = String(row.phone || '').toLowerCase();
                const fullname = String(row.fullName || '').toLowerCase();
                const product = String(row.product || '').toLowerCase();
                if (!fullname.includes(search) && !phone.includes(search) && !product.includes(search)) return false;
            }

            return true;
        });

        // مرتب‌سازی ساده بر اساس نام مشتری
        filtered.sort((a, b) => String(a.fullName || '').localeCompare(String(b.fullName || '')));

        this.filteredChangeExpertData = filtered;
        this.renderChangeExpertTable();
    },

    renderChangeExpertTable: function () {
        const tbody = document.getElementById('changeExpertTableBody');
        const empty = document.getElementById('changeExpertEmpty');
        if (!tbody) return;

        tbody.innerHTML = '';

        if (!this.filteredChangeExpertData || this.filteredChangeExpertData.length === 0) {
            if (empty) empty.classList.remove('hidden');
            return;
        }
        if (empty) empty.classList.add('hidden');

        // آماده‌سازی گزینه‌های کارشناس فعال
        let expertOptions = '<option value="">انتخاب کارشناس</option>';
        allUsers.forEach(user => {
            if (user.username.toLowerCase() !== 'admin') {
                expertOptions += `<option value="${user.name}">${user.name}</option>`;
            }
        });

        this.filteredChangeExpertData.forEach(row => {
            const rowIndex = row.rowIndex;
            const currentExpert = row.expertName || '';
            const phone = row.phone || '';
            const fullname = row.fullName || '';
            const product = row.product || '';
            const status = row.status || '';

            const tr = document.createElement('tr');
            tr.className = 'hover:bg-slate-50 dark:hover:bg-slate-700/40';
            tr.innerHTML = `
                      <td class="p-3 truncate">${fullname}</td>
                      <td class="p-3 font-mono text-xs">${phone}</td>
                      <td class="p-3 truncate">${product}</td>
                      <td class="p-3 truncate">${status}</td>
                      <td class="p-3 font-semibold text-slate-700 dark:text-slate-200">${currentExpert}</td>
                      <td class="p-3">
                          <select class="input-modern rounded-lg w-full p-2 text-sm change-expert-select" data-row="${rowIndex}">
                              ${expertOptions.replace(`value="${currentExpert}"`, `value="${currentExpert}" selected`)}
                          </select>
                      </td>
                      <td class="p-3 text-center">
                          <button
                              class="change-expert-save bg-indigo-600 hover:bg-indigo-700 text-white px-3 py-2 rounded-lg text-xs font-bold transition disabled:opacity-50 disabled:cursor-not-allowed"
                              data-row="${rowIndex}"
                              disabled
                          >
                              ذخیره
                          </button>
                      </td>
                  `;
            tbody.appendChild(tr);
        });

        // Event listeners برای select و دکمه‌ها
        tbody.querySelectorAll('.change-expert-select').forEach(sel => {
            sel.addEventListener('change', (e) => {
                const rowId = e.target.getAttribute('data-row');
                const saveBtn = tbody.querySelector(`.change-expert-save[data-row="${rowId}"]`);
                const row = this.filteredChangeExpertData.find(r => String(r.rowIndex) === String(rowId));
                const currentExpert = row ? row.expertName : '';
                const newVal = e.target.value;
                if (saveBtn) {
                    saveBtn.disabled = (newVal === '' || newVal === currentExpert);
                }
            });
        });

        tbody.querySelectorAll('.change-expert-save').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const rowId = e.target.getAttribute('data-row');
                const selectEl = tbody.querySelector(`.change-expert-select[data-row="${rowId}"]`);
                if (!selectEl) return;
                const newExpert = selectEl.value;
                const row = this.filteredChangeExpertData.find(r => String(r.rowIndex) === String(rowId));
                if (!row) return;
                const currentExpert = row.expertName;
                if (!newExpert || newExpert === currentExpert) return;

                this.saveChangeExpert(rowId, newExpert, btn, selectEl);
            });
        });
    },

    saveChangeExpert: function (rowIndex, newExpert, btnEl, selectEl) {
        if (!btnEl) return;
        const originalText = btnEl.innerText;
        btnEl.disabled = true;
        btnEl.innerText = 'در حال ذخیره...';

        const assignedBy = sessionStorage.getItem('appUser') || 'supervisor';
        google.script.run.withSuccessHandler((res) => {
            if (res && res.success) {
                // بروزرسانی داده در حافظه
                const target = this.changeExpertData.find(r => String(r.rowIndex) === String(rowIndex));
                if (target) {
                    target.expertName = newExpert;
                    target.assignedBy = assignedBy;
                }
                this.filterChangeExpertTable();
                this.showChangeExpertToast('کارشناس مشتری با موفقیت تغییر یافت', false);
            } else {
                this.showChangeExpertToast(res && res.message ? res.message : 'خطا در ذخیره تغییر', true);
            }
        }).withFailureHandler((err) => {
            this.showChangeExpertToast(err && err.message ? err.message : 'خطا در ذخیره تغییر', true);
        }).updateLeadExpert(Number(rowIndex), newExpert, assignedBy);

        setTimeout(() => {
            btnEl.disabled = false;
            btnEl.innerText = originalText;
            if (selectEl) selectEl.disabled = false;
        }, 400);
    },

    showChangeExpertToast: function (message, isError = false) {
        const toast = document.getElementById('changeExpertToast');
        if (!toast) return;
        toast.innerText = message;
        toast.className = `text-sm px-4 py-2 rounded-lg ${isError ? 'bg-rose-50 text-rose-700 border border-rose-200 dark:bg-rose-900/20 dark:text-rose-200 dark:border-rose-800' : 'bg-emerald-50 text-emerald-700 border border-emerald-200 dark:bg-emerald-900/20 dark:text-emerald-200 dark:border-emerald-800'}`;
        toast.classList.remove('hidden');
        setTimeout(() => {
            toast.classList.add('hidden');
        }, 3000);
    },

    renderExpertComparison: function () {
        const dataSource = (Array.isArray(globalData) && globalData.length > 0)
            ? globalData
            : (typeof monitoringDashboard !== 'undefined' ? monitoringDashboard.allData : []);

        if (!dataSource || dataSource.length === 0) return;

        const calculator = (typeof monitoringDashboard !== 'undefined' && typeof monitoringDashboard.calculateExpertStats === 'function')
            ? monitoringDashboard.calculateExpertStats
            : this.calculateExpertStatsForComparison;

        const stats = calculator.call(monitoringDashboard || this, dataSource);

        const tbody = document.getElementById('expertComparisonTableBody');
        if (tbody) {
            tbody.innerHTML = '';
            Object.entries(stats).forEach(([expert, s]) => {
                const row = document.createElement('tr');
                row.className = 'border-b border-slate-200 dark:border-slate-700';
                row.innerHTML = `
                          <td class="py-3 px-4 font-semibold text-slate-800 dark:text-slate-200">${expert || 'نامشخص'}</td>
                          <td class="py-3 px-4 text-right">${s.totalLeads}</td>
                          <td class="py-3 px-4 text-right">${s.activeLeads}</td>
                          <td class="py-3 px-4 text-right">${s.contracts}</td>
                          <td class="py-3 px-4 text-right">${s.deliveries}</td>
                          <td class="py-3 px-4 text-right">${s.cancellations}</td>
                          <td class="py-3 px-4 text-right">${s.conversionRate.toFixed(1)}%</td>
                      `;
                tbody.appendChild(row);
            });
        }

        const ctx = document.getElementById('expertComparisonChart');
        if (ctx) {
            if (this.expertComparisonChart) {
                this.expertComparisonChart.destroy();
            }

            const experts = Object.keys(stats);
            const contracts = experts.map(e => stats[e].contracts);
            const deliveries = experts.map(e => stats[e].deliveries);

            this.expertComparisonChart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: experts,
                    datasets: [
                        {
                            label: 'قرارداد',
                            data: contracts,
                            backgroundColor: 'rgba(59, 130, 246, 0.8)',
                            borderColor: 'rgb(59, 130, 246)',
                            borderWidth: 1.5
                        },
                        {
                            label: 'تحویل',
                            data: deliveries,
                            backgroundColor: 'rgba(16, 185, 129, 0.7)',
                            borderColor: 'rgb(16, 185, 129)',
                            borderWidth: 1.5
                        }
                    ]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { legend: { position: 'top' } },
                    scales: {
                        y: { beginAtZero: true }
                    }
                }
            });
        }
    },

    calculateExpertStatsForComparison: function (data) {
        const stats = {};

        data.forEach(row => {
            const expert = row[COLUMN_INDEXES.EXPERT_NAME] || 'نامشخص';
            if (!stats[expert]) {
                stats[expert] = {
                    totalLeads: 0,
                    activeLeads: 0,
                    contracts: 0,
                    deliveries: 0,
                    cancellations: 0,
                    conversionRate: 0
                };
            }

            stats[expert].totalLeads++;

            const callResult = row[COLUMN_INDEXES.CALL_RESULT] || '';
            if (callResult.includes('در حال بررسی') || callResult.includes('در پروسه')) {
                stats[expert].activeLeads++;
            }

            const contract = row[COLUMN_INDEXES.CONTRACT] || '';
            if (contract === 'انجام شد') {
                stats[expert].contracts++;
            }

            const delivery = row[COLUMN_INDEXES.DELIVERY] || '';
            if (delivery === 'شد') {
                stats[expert].deliveries++;
            }

            const cancellation = row[COLUMN_INDEXES.CANCELLATION] || '';
            if (cancellation && cancellation !== '') {
                stats[expert].cancellations++;
            }
        });

        Object.keys(stats).forEach(expert => {
            const s = stats[expert];
            s.conversionRate = s.totalLeads > 0 ? (s.contracts / s.totalLeads) * 100 : 0;
        });

        return stats;
    },

    initDetailCalculator: function (savedBrand) {
        const select = document.getElementById('calcBrand');
        if (!select) return;
        const brands = [...new Set(pricingData.map(d => d.brand))];
        select.innerHTML = '<option value="">- انتخاب -</option>' + brands.map(b => `<option value="${b}">${b}</option>`).join('');
        if (savedBrand) select.value = savedBrand;
    },

    initCalculator: function () { },

    updateCalcStep1: function () {
        const brand = document.getElementById('calcBrand').value;
        const axleSelect = document.getElementById('calcAxle');
        const usageSelect = document.getElementById('calcUsage');
        const descSelect = document.getElementById('calcDesc');

        axleSelect.innerHTML = '<option value="">- ابتدا برند را انتخاب کنید -</option>'; axleSelect.disabled = true;
        usageSelect.innerHTML = '<option value="">- ابتدا محور را انتخاب کنید -</option>'; usageSelect.disabled = true;
        descSelect.innerHTML = '<option value="">- ابتدا کاربری را انتخاب کنید -</option>'; descSelect.disabled = true;
        document.getElementById('displayTotalPrice').innerText = 0;

        if (!brand) return;

        const filteredData = pricingData.filter(d => d.brand === brand);
        const axles = [...new Set(filteredData.map(d => d.axle))];
        axleSelect.innerHTML = '<option value="">- انتخاب -</option>' + axles.map(a => `<option value="${a}">${a}</option>`).join('');
        axleSelect.disabled = false;
    },

    updateCalcStep2: function () {
        const brand = document.getElementById('calcBrand').value;
        const axle = document.getElementById('calcAxle').value;
        const usageSelect = document.getElementById('calcUsage');
        const descSelect = document.getElementById('calcDesc');

        usageSelect.innerHTML = '<option value="">- ابتدا محور را انتخاب کنید -</option>'; usageSelect.disabled = true;
        descSelect.innerHTML = '<option value="">- ابتدا کاربری را انتخاب کنید -</option>'; descSelect.disabled = true;
        document.getElementById('displayTotalPrice').innerText = 0;

        if (!axle) return;

        const filteredData = pricingData.filter(d => d.brand === brand && d.axle === axle);
        const usages = [...new Set(filteredData.map(d => d.usage))];
        usageSelect.innerHTML = '<option value="">- انتخاب -</option>' + usages.map(u => `<option value="${u}">${u}</option>`).join('');
        usageSelect.disabled = false;
    },

    updateCalcStep3: function () {
        const brand = document.getElementById('calcBrand').value;
        const axle = document.getElementById('calcAxle').value;
        const usage = document.getElementById('calcUsage').value;
        const descSelect = document.getElementById('calcDesc');

        descSelect.innerHTML = '<option value="">- ابتدا کاربری را انتخاب کنید -</option>'; descSelect.disabled = true;
        document.getElementById('displayTotalPrice').innerText = 0;

        if (!usage) return;

        const filteredData = pricingData.filter(d => d.brand === brand && d.axle === axle && d.usage === usage);
        const descs = [...new Set(filteredData.map(d => d.desc))];
        descSelect.innerHTML = '<option value="">- انتخاب -</option>' + descs.map(d => `<option value="${d}">${d}</option>`).join('');
        descSelect.disabled = false;
    },

    calculateFinalPrice: function () {
        const brand = document.getElementById('calcBrand').value;
        const axle = document.getElementById('calcAxle').value;
        const usage = document.getElementById('calcUsage').value;
        const desc = document.getElementById('calcDesc').value;
        const paymentMethod = document.getElementById('calcPaymentMethod').value;

        if (!brand || !axle || !usage || !desc) {
            document.getElementById('displayTotalPrice').innerText = 0;
            // Clear target prepayment if calculator is incomplete
            const targetPrepaymentEl = document.getElementById('prepayment_target');
            if (targetPrepaymentEl) {
                targetPrepaymentEl.value = "";
                updatePrepaymentDifference();
            }
            return;
        }

        const item = pricingData.find(d => d.brand === brand && d.axle === axle && d.usage === usage && d.desc === desc);

        if (item) {
            let total = item.price;
            const installmentBox = document.getElementById('installmentDetails');
            let calculatedPrepayment = 0;

            if (paymentMethod === 'cash') {
                installmentBox.classList.add('hidden');
                // For cash payment, prepayment is typically 0 or can be set differently
                calculatedPrepayment = 0;
            } else {
                installmentBox.classList.remove('hidden');
                const months = Number(paymentMethod);
                const interestRate = months === 12 ? 0.02 : 0.025;
                const totalWithInterest = total * (1 + (interestRate * months));
                calculatedPrepayment = total * 0.5; // 50% prepayment
                const remaining = totalWithInterest - calculatedPrepayment;
                const monthly = remaining / months;

                document.getElementById('displayPrepayment').innerText = calculatedPrepayment.toLocaleString();
                document.getElementById('displayMonthly').innerText = Math.round(monthly).toLocaleString();
                total = totalWithInterest;
            }

            document.getElementById('displayTotalPrice').innerText = Math.round(total).toLocaleString();

            // Auto-fill Target Prepayment from calculator result
            const targetPrepaymentEl = document.getElementById('prepayment_target');
            if (targetPrepaymentEl) {
                targetPrepaymentEl.value = Math.round(calculatedPrepayment);
                updatePrepaymentDifference();
            }
        }
    },

    populateCalculator: function (row) {
        const brand = row[COLUMN_INDEXES.CALC_BRAND];
        const axle = row[COLUMN_INDEXES.CALC_AXLE];
        const usage = row[COLUMN_INDEXES.CALC_USAGE];
        const desc = row[COLUMN_INDEXES.CALC_DESC];
        const pay = row[COLUMN_INDEXES.CALC_PAYMENT];
        const price = row[COLUMN_INDEXES.CALC_PRICE];

        app.initDetailCalculator(brand);
        if (brand) {
            app.updateCalcStep1();
            document.getElementById('calcAxle').value = axle || "";
            if (axle) {
                app.updateCalcStep2();
                document.getElementById('calcUsage').value = usage || "";
                if (usage) {
                    app.updateCalcStep3();
                    document.getElementById('calcDesc').value = desc || "";
                }
            }
        }
        if (pay) document.getElementById('calcPaymentMethod').value = pay;
        if (price) document.getElementById('displayTotalPrice').innerText = price;
    },

    loadUsers: function () {
        google.script.run.withSuccessHandler(users => {
            allUsers = users;
            app.renderMenu(users);
            if (app.isAdmin) app.renderUsersList(users);
            app.populateSupervisorOptions(users);
            app.applyRoleFormVisibility();
            // پس از بارگذاری لیست کاربران، اگر سرپرست هستیم شمارش پیام‌های نخوانده را بگیر
            const currentRole = (sessionStorage.getItem('role') || '').toLowerCase();
            if (currentRole === 'supervisor' || currentRole === 'admin') {
                app.refreshUnreadCountsForSupervisor();
            }
        }).getUsersList();
    },

    // نگهداری شمارش پیام‌های نخوانده برای هر کارشناس (کلید: username نرمال‌شده)
    unreadCountsByExpert: {},

    refreshUnreadCountsForSupervisor: function () {
        const supervisor = (sessionStorage.getItem('appUser') || '').toLowerCase();
        if (!supervisor) return;
        google.script.run.withSuccessHandler(res => {
            const map = {};
            if (res && res.success && Array.isArray(res.counts)) {
                res.counts.forEach(item => {
                    const u = String(item.expertUsername || '').toLowerCase();
                    const c = Number(item.unreadCount) || 0;
                    if (u) map[u] = c;
                });
            }
            app.unreadCountsByExpert = map;
            // رندر مجدد منو تا badgeها بروزرسانی شوند
            if (Array.isArray(allUsers)) {
                app.renderMenu(allUsers);
            }
        }).withFailureHandler(() => {
            // در صورت خطا، چیزی نمایش نمی‌دهیم تا UI خراب نشود
        }).getUnreadCountsForSupervisor(supervisor);
    },

    renderMenu: function (users) {
        const container = document.getElementById('expertListContainer');
        if (!container) return;
        let html = '';
        const currentRole = (sessionStorage.getItem('role') || '').toLowerCase();
        const currentUser = (sessionStorage.getItem('appUser') || '').toLowerCase();

        const renderExpertRow = (u, indexPrefix) => {
            const expertUsername = String(u.username || '').toLowerCase();
            const unread = app.unreadCountsByExpert[expertUsername] || 0;
            const hasUnread = unread > 0;
            return `
                    <div onclick="app.showExpertPanel('${u.username}', '${u.name}')" id="menu-expert-${indexPrefix}-${expertUsername}" class="sidebar-link">
                      <svg class="sidebar-icon" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z"></path>
                      </svg>
                      <span class="flex-1 flex items-center justify-between gap-2">
                        <span>${u.name}</span>
                        ${hasUnread ? `
                        <span class="inline-flex items-center justify-center min-w-[22px] h-5 rounded-full bg-rose-100 text-rose-600 dark:bg-rose-900/40 dark:text-rose-300 text-[11px] font-bold">
                          ${unread > 99 ? '99+' : unread}
                        </span>` : ''}
                      </span>
                    </div>
                  `;
        };

        if (currentRole === 'supervisor') {
            // فقط کارشناسان زیرمجموعه همین سرپرست
            const experts = users.filter(u => {
                const role = (u.role || '').toLowerCase();
                const isExpertRole = role === 'expert' || role === '';
                if (!isExpertRole) return false;
                return String(u.supervisor || '').toLowerCase() === currentUser;
            });
            experts.forEach((u, index) => {
                html += renderExpertRow(u, index);
            });
        } else if (currentRole === 'admin') {
            // ادمین: نمایش سرپرست‌ها + زیرمنوی کارشناسان آن‌ها
            const supervisors = users.filter(u => (u.role || '').toLowerCase() === 'supervisor');
            supervisors.forEach((sup, sIndex) => {
                const supUsername = String(sup.username || '').toLowerCase();
                const expertsOfSup = users.filter(u => {
                    const role = (u.role || '').toLowerCase();
                    const isExpertRole = role === 'expert' || role === '';
                    return isExpertRole && String(u.supervisor || '').toLowerCase() === supUsername;
                });
                html += `
                        <div class="sidebar-link cursor-default">
                          <svg class="sidebar-icon" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5.121 17.804A13.937 13.937 0 0112 15c2.5 0 4.847.655 6.879 1.804M15 10a3 3 0 11-6 0 3 3 0 016 0z"></path>
                          </svg>
                          <span class="flex-1 flex items-center justify-between gap-2">
                            <span>${sup.name} (@${supUsername})</span>
                          </span>
                        </div>
                      `;
                if (expertsOfSup.length > 0) {
                    html += `<div class="ml-6 space-y-1">`;
                    expertsOfSup.forEach((exp, eIndex) => {
                        html += renderExpertRow(exp, `${sIndex}-${eIndex}`);
                    });
                    html += `</div>`;
                }
            });
        } else {
            // سایر نقش‌ها: فعلاً لیستی نمایش نده
            html = '';
        }

        container.innerHTML = html;
    },

    // اطلاعات کارشناس انتخاب‌شده برای پنل سرپرست
    currentSupervisorExpertUsername: null,
    currentSupervisorExpertName: null,

    showExpertPanel: function (username, name) {
        app.switchView('view-expert-dynamic');
        app.currentSupervisorExpertUsername = (username || '').toLowerCase();
        app.currentSupervisorExpertName = name || '';
        const titleEl = document.getElementById('dynamicExpertTitle');
        if (titleEl) {
            titleEl.innerText = `پنل کارشناس: ${name}`;
        }
        const nameEl = document.getElementById('dynamicExpertName');
        if (nameEl) {
            nameEl.innerText = name || '';
        }
        // به‌صورت پیش‌فرض تب پیام‌های خوانده‌نشده را نمایش بده
        app.showSupervisorExpertTab('unread');
    },

    // سوییچ بین تب‌های آیکون‌محور در پنل سرپرست برای هر کارشناس
    supervisorExpertActiveTab: 'unread',

    showSupervisorExpertTab: function (tab) {
        const allowed = ['unread', 'read', 'analytics'];
        if (!allowed.includes(tab)) tab = 'unread';
        app.supervisorExpertActiveTab = tab;

        const tabs = ['unread', 'read', 'analytics'];
        tabs.forEach(t => {
            const card = document.getElementById(`supervisor-expert-tab-${t}`);
            if (card) {
                if (t === tab) {
                    card.classList.add('ring-2', 'ring-amber-500');
                    card.classList.remove('opacity-70');
                } else {
                    card.classList.remove('ring-2', 'ring-amber-500');
                    card.classList.add('opacity-70');
                }
            }
        });

        const listContainer = document.getElementById('supervisorExpertMessagesList');
        const emptyAnalytics = document.getElementById('supervisorExpertAnalyticsPlaceholder');
        if (tab === 'analytics') {
            if (listContainer) listContainer.classList.add('hidden');
            if (emptyAnalytics) emptyAnalytics.classList.remove('hidden');
            return;
        }

        if (emptyAnalytics) emptyAnalytics.classList.add('hidden');
        if (listContainer) listContainer.classList.remove('hidden');

        app.loadMessagesForSupervisor(tab);
    },

    // بارگذاری پیام‌های کارشناس انتخاب‌شده برای سرپرست
    loadMessagesForSupervisor: function (status) {
        const supervisor = (sessionStorage.getItem('appUser') || '').toLowerCase();
        const expert = app.currentSupervisorExpertUsername || '';
        if (!supervisor || !expert) return;

        const normalizedStatus = status || app.supervisorExpertActiveTab || 'unread';
        const listContainer = document.getElementById('supervisorExpertMessagesList');
        const empty = document.getElementById('supervisorExpertEmptyState');
        const loader = document.getElementById('supervisorExpertLoading');

        if (listContainer) {
            listContainer.innerHTML = '';
            listContainer.classList.add('hidden');
        }
        if (empty) empty.classList.add('hidden');
        if (loader) loader.classList.remove('hidden');

        google.script.run.withSuccessHandler(res => {
            if (loader) loader.classList.add('hidden');
            if (!res || !res.success || !Array.isArray(res.messages) || res.messages.length === 0) {
                if (empty) empty.classList.remove('hidden');
                return;
            }
            if (empty) empty.classList.add('hidden');

            const msgs = res.messages;
            let html = '';
            msgs.forEach(msg => {
                const dateStr = app.formatClientDateTime(msg.timestamp);
                const preview = (msg.text || '').length > 80 ? (msg.text || '').substring(0, 80) + '…' : (msg.text || '');
                html += `
                        <button
                          type="button"
                          class="w-full text-right px-3 py-2 rounded-xl border border-slate-200 dark:border-slate-700 bg-white/70 dark:bg-slate-800/70 hover:bg-amber-50 dark:hover:bg-amber-900/10 transition flex flex-col gap-1"
                          onclick="app.selectSupervisorMessage('${msg.messageId}')">
                          <div class="flex items-center justify-between gap-2">
                            <span class="text-xs font-bold text-slate-700 dark:text-slate-200">
                              گزارش ردیف ${msg.relatedRowIndex || '-'}
                            </span>
                            <span class="text-[10px] text-slate-400 dark:text-slate-500">
                              ${dateStr}
                            </span>
                          </div>
                          <p class="text-xs text-slate-600 dark:text-slate-300 line-clamp-2">${preview}</p>
                        </button>
                      `;
            });
            if (listContainer) {
                listContainer.innerHTML = html;
                listContainer.classList.remove('hidden');
            }
        }).withFailureHandler(() => {
            if (loader) loader.classList.add('hidden');
            if (empty) {
                empty.classList.remove('hidden');
                empty.innerText = 'خطا در بارگذاری پیام‌ها';
            }
        }).getMessagesForSupervisor(supervisor, expert, normalizedStatus);
    },

    // ذخیره آخرین پیام انتخاب‌شده برای سرپرست
    selectedSupervisorMessage: null,

    selectSupervisorMessage: function (messageId) {
        const supervisor = (sessionStorage.getItem('appUser') || '').toLowerCase();
        const expert = app.currentSupervisorExpertUsername || '';
        if (!supervisor || !expert || !messageId) return;

        // برای سادگی، دوباره لیست فعلی را می‌خوانیم و پیام را پیدا می‌کنیم
        const status = app.supervisorExpertActiveTab || 'unread';
        google.script.run.withSuccessHandler(res => {
            if (!res || !res.success || !Array.isArray(res.messages)) return;
            const msg = res.messages.find(m => m.messageId === messageId);
            if (!msg) return;
            app.selectedSupervisorMessage = msg;

            const titleEl = document.getElementById('supervisorExpertSelectedTitle');
            const metaEl = document.getElementById('supervisorExpertSelectedMeta');
            const bodyEl = document.getElementById('supervisorExpertSelectedBody');

            if (titleEl) {
                titleEl.innerText = `پیام کارشناس (${app.currentSupervisorExpertName || expert})`;
            }
            if (metaEl) {
                const dateStr = app.formatClientDateTime(msg.timestamp);
                metaEl.innerText = `ردیف CRM: ${msg.relatedRowIndex || '-'} | ${dateStr}`;
            }
            if (bodyEl) {
                bodyEl.innerText = msg.text || '';
            }
        }).withFailureHandler(() => {
            // نادیده گرفتن خطا
        }).getMessagesForSupervisor(supervisor, expert, status);
    },

    // علامت‌گذاری پیام به‌عنوان خوانده‌شده توسط سرپرست
    markSelectedSupervisorMessageAsRead: function () {
        if (!app.selectedSupervisorMessage || !app.selectedSupervisorMessage.messageId) return;
        const msgId = app.selectedSupervisorMessage.messageId;
        google.script.run.withSuccessHandler(() => {
            // پس از بروزرسانی، لیست پیام‌ها و badgeها را بروزرسانی کن
            app.loadMessagesForSupervisor(app.supervisorExpertActiveTab || 'unread');
            app.refreshUnreadCountsForSupervisor();
        }).withFailureHandler(() => {
            // نادیده گرفتن خطا
        }).markMessageRead(msgId, 'supervisor');
    },

    // ارسال پاسخ سرپرست به کارشناس (threaded)
    sendReplyToExpert: function () {
        const msg = app.selectedSupervisorMessage;
        if (!msg || !msg.messageId) return;
        const textarea = document.getElementById('supervisorExpertReplyText');
        const statusEl = document.getElementById('supervisorExpertReplyStatus');
        if (!textarea) return;
        const text = textarea.value.trim();
        if (!text) {
            if (statusEl) {
                statusEl.innerText = 'متن پاسخ را وارد کنید';
                statusEl.classList.add('text-rose-500');
            }
            return;
        }

        const supervisor = (sessionStorage.getItem('appUser') || '').toLowerCase();
        const expert = app.currentSupervisorExpertUsername || '';
        if (!supervisor || !expert) return;

        if (statusEl) {
            statusEl.innerText = 'در حال ارسال پاسخ...';
            statusEl.classList.remove('text-rose-500');
        }

        const payload = {
            type: 'supervisor_to_expert',
            fromUser: supervisor,
            toUser: expert,
            expertUsername: expert,
            supervisorUsername: supervisor,
            relatedRowIndex: msg.relatedRowIndex || 0,
            text: text,
            threadId: msg.threadId || msg.messageId
        };

        google.script.run.withSuccessHandler(res => {
            if (statusEl) {
                if (res && res.success) {
                    statusEl.innerText = 'پاسخ با موفقیت ارسال شد';
                } else {
                    statusEl.innerText = (res && res.message) || 'خطا در ارسال پاسخ';
                    statusEl.classList.add('text-rose-500');
                }
            }
            textarea.value = '';
        }).withFailureHandler(() => {
            if (statusEl) {
                statusEl.innerText = 'خطا در ارسال پاسخ';
                statusEl.classList.add('text-rose-500');
            }
        }).sendMessage(payload);
    },

    renderUsersList: function (users) {
        const tbody = document.getElementById('usersListBody');
        let html = '';
        const currentRole = (sessionStorage.getItem('role') || '').toLowerCase();
        const currentUser = (sessionStorage.getItem('appUser') || '').toLowerCase();

        const roleLabelMap = {
            admin: 'ادمین',
            supervisor: 'سرپرست',
            expert: 'کارشناس',
            monitoring: 'مانیتورینگ',
            agency: 'نمایندگی'
        };

        const addRow = (user, level) => {
            const role = (user.role || 'expert').toLowerCase();
            const badgeClass = role === 'admin'
                ? 'bg-rose-100 text-rose-700 dark:bg-rose-900/30 dark:text-rose-200'
                : role === 'supervisor'
                    ? 'bg-amber-100 text-amber-700 dark:bg-amber-900/30 dark:text-amber-200'
                    : 'bg-slate-100 text-slate-700 dark:bg-slate-700/50 dark:text-slate-200';
            const indentClass = level > 0 ? 'pl-6' : '';
            const canDelete = currentRole === 'admin' || (currentRole === 'supervisor' && role === 'expert' && String(user.supervisor || '').toLowerCase() === currentUser);

            html += `
                  <tr class="hover:bg-slate-50 dark:hover:bg-slate-600/50 transition">
                      <td class="p-3 text-slate-700 dark:text-slate-200 ${indentClass}">${user.username}</td>
                      <td class="p-3 text-slate-700 dark:text-slate-200">
                        <div class="flex items-center gap-2">
                          <span>${user.name}</span>
                          <span class="px-2 py-1 text-[11px] font-bold rounded ${badgeClass}">${roleLabelMap[role] || 'کارشناس'}</span>
                        </div>
                      </td>
                      <td class="p-3 text-center">
                          ${canDelete ? `<button onclick="app.deleteUser('${user.username}', '${user.name}')" class="text-rose-500 hover:bg-rose-50 dark:hover:bg-rose-900/50 p-2 rounded-lg transition" title="حذف">
                              <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path></svg>
                          </button>` : ''}
                      </td>
                  </tr>`;
        };

        const filteredUsers = users.filter(u => {
            if (u.username.toLowerCase() === 'admin') return false;
            if (currentRole === 'supervisor') {
                const role = (u.role || 'expert').toLowerCase();
                if (role !== 'expert') return false;
                return String(u.supervisor || '').toLowerCase() === currentUser;
            }
            return true;
        });

        if (currentRole === 'admin') {
            const supervisors = filteredUsers.filter(u => (u.role || '').toLowerCase() === 'supervisor');
            const experts = filteredUsers.filter(u => (u.role || '').toLowerCase() === 'expert');
            const others = filteredUsers.filter(u => ['monitoring', 'agency', 'admin', 'supervisor', 'expert'].indexOf((u.role || '').toLowerCase()) === -1 || ['monitoring', 'agency'].includes((u.role || '').toLowerCase()));

            supervisors.forEach(sup => {
                addRow(sup, 0);
                experts.filter(ex => String(ex.supervisor || '').toLowerCase() === sup.username.toLowerCase())
                    .forEach(ex => addRow(ex, 1));
            });

            // experts بدون supervisor
            experts.filter(ex => !ex.supervisor).forEach(ex => addRow(ex, 0));
            // سایر نقش‌ها (monitoring/agency)
            others.forEach(o => addRow(o, 0));
        } else {
            filteredUsers.forEach(u => addRow(u, 0));
        }
        tbody.innerHTML = html;
    },

    addUser: function () {
        const u = document.getElementById('newUsername').value;
        const p = document.getElementById('newPassword').value;
        const n = document.getElementById('newDisplayName').value;
        const role = document.getElementById('newRole') ? document.getElementById('newRole').value : 'expert';
        const supervisor = document.getElementById('newSupervisor') ? document.getElementById('newSupervisor').value : '';
        if (!u || !p || !n || !role) { alert("لطفا تمام فیلدها را پر کنید"); return; }
        const currentUser = sessionStorage.getItem('appUser') || '';
        const currentRole = sessionStorage.getItem('role') || '';
        // سرپرست فقط کارشناس می‌تواند اضافه کند
        if (currentRole.toLowerCase() === 'supervisor' && role !== 'expert') {
            alert("سرپرست فقط مجاز به افزودن کارشناس است.");
            return;
        }
        google.script.run.withSuccessHandler(res => {
            if (res.success) {
                const roleLower = (role || '').toLowerCase();
                const currentRoleLower = (currentRole || '').toLowerCase();
                if (roleLower === 'supervisor') {
                    alert(`سلام، ${n}`);
                } else if (currentRoleLower === 'supervisor' && roleLower === 'expert') {
                    alert(`سلام، ${n}`);
                }
                document.getElementById('newUsername').value = "";
                document.getElementById('newPassword').value = "";
                document.getElementById('newDisplayName').value = "";
                if (document.getElementById('newRole')) document.getElementById('newRole').value = "supervisor";
                if (document.getElementById('newSupervisor')) document.getElementById('newSupervisor').value = "";
                app.loadUsers();
            } else {
                alert(res.message);
            }
        }).addUser({
            username: u,
            password: p,
            name: n,
            role: role,
            supervisor: supervisor,
            currentUser: { username: currentUser, role: currentRole }
        });
    },

    populateSupervisorOptions: function (users) {
        const select = document.getElementById('newSupervisor');
        if (!select) return;
        const supervisors = users.filter(u => (u.role || '').toLowerCase() === 'supervisor');
        const options = ['<option value="">(سرپرست)</option>']
            .concat(supervisors.map(s => `<option value="${s.username}">${s.name || s.username}</option>`));
        select.innerHTML = options.join('');
    },

    applyRoleFormVisibility: function () {
        const role = (sessionStorage.getItem('role') || '').toLowerCase();
        const roleSelect = document.getElementById('newRole');
        const supervisorSelect = document.getElementById('newSupervisor');
        if (!roleSelect) return;
        if (role === 'supervisor') {
            roleSelect.innerHTML = `<option value="expert">کارشناس</option>`;
            if (supervisorSelect) {
                supervisorSelect.value = sessionStorage.getItem('appUser') || '';
                supervisorSelect.disabled = true;
            }
        } else {
            // admin: همه گزینه‌ها
            roleSelect.innerHTML = `
                    <option value="expert">کارشناس</option>
                    <option value="supervisor">سرپرست</option>
                    <option value="monitoring">مانیتورینگ</option>
                    <option value="agency">نمایندگی</option>
                    <option value="admin">ادمین</option>
                  `;
            if (supervisorSelect) supervisorSelect.disabled = false;
        }
    },

    deleteUser: function (username, name) {
        const modal = document.getElementById('confirmModal');
        const msg = document.getElementById('confirmMessage');
        const btn = document.getElementById('confirmYesBtn');
        msg.innerText = `آیا مطمئن هستید که می‌خواهید کاربر "${name}" را حذف کنید؟`;
        modal.classList.remove('hidden');
        btn.onclick = function () {
            modal.classList.add('hidden');
            google.script.run.withSuccessHandler(res => {
                if (res.success) app.loadUsers();
                else alert(res.message);
            }).deleteUser({
                username,
                currentUser: { username: sessionStorage.getItem('appUser') || '', role: sessionStorage.getItem('role') || '' }
            });
        };
    },
    applySearchIfNeeded: function () {
        const input = document.getElementById("leadSearchInput");
        if (!input) return;

        const text = input.value.trim();
        if (text.length > 0) {
            app.searchLeads();   // 🔥 دوباره فیلتر اعمال می‌شود
        }
    },

    refreshTable: function (isAuto = false) {
        const user = sessionStorage.getItem('appUser');
        const role = (sessionStorage.getItem('role') || '').toLowerCase();
        const name = sessionStorage.getItem('appName');
        if (!user) return;
        if (!isAuto && globalData.length === 0) {
            document.getElementById('emptyState').classList.add('hidden');
            document.getElementById('loadingState').classList.remove('hidden');
        }
        const btn = document.getElementById('refreshBtn');
        if (btn) btn.classList.add('animate-spin');

        const userContext = { username: user || '', role, displayName: name || '' };

        google.script.run.withSuccessHandler((res) => {
            if (btn) btn.classList.remove('animate-spin');
            globalHeaders = res.headers;
            globalData = res.data;
            const rowIndices = res.data.map(r => r[COLUMN_INDEXES.SHEET_ROW_INDEX]);
            app.loadProcessTransfers(rowIndices, () => {
                app.renderData(res.data);
                app.renderSegments();
                app.applySearchIfNeeded();
            });

        }).getData(userContext);
    },
    searchLeads: function () {
        const inputEl = document.getElementById("leadSearchInput");
        if (!inputEl) return;
        const raw = inputEl.value || "";
        const text = app.normalizeSearchText(raw);

        if (!text) {
            // بازگشت به نمایش کامل
            app.renderData(globalData);
            return;
        }

        const filtered = globalData.filter(r => {
            const rowNumber = r[COLUMN_INDEXES.SHEET_ROW_INDEX];
            const transfer = app.processTransfers[rowNumber];

            const parts = [];
            // فیلدهای اصلی
            parts.push(transfer ? transfer.fullName : (r[COLUMN_INDEXES.FULL_NAME] || ""));
            parts.push(transfer ? transfer.phone : (r[COLUMN_INDEXES.PHONE] || ""));
            parts.push(transfer ? transfer.nationalId : (r[COLUMN_INDEXES.NATIONAL_ID] || ""));
            parts.push(r[COLUMN_INDEXES.PRODUCT] || "");
            parts.push(r[COLUMN_INDEXES.EXPERT_NAME] || "");
            parts.push(r[COLUMN_INDEXES.CALL_RESULT] || "");
            // کل ردیف به صورت متن
            parts.push((r || []).join(" "));

            const haystack = app.normalizeSearchText(parts.join(" "));
            return haystack.includes(text);
        });

        app.renderData(filtered);
    },


    renderData: function (data) {
        document.getElementById('loadingState').classList.add('hidden');
        const tbody = document.getElementById('tableBody');
        const thead = document.getElementById('tableHeader');
        if (data.length === 0) {
            tbody.innerHTML = ""; thead.innerHTML = "";
            document.getElementById('emptyState').classList.remove('hidden');
            document.getElementById('totalCountDisplay').innerText = "0";
            app.updateCharts([]);
            return;
        }
        document.getElementById('emptyState').classList.add('hidden');
        let headerHtml = '<tr>';
        globalHeaders.forEach((header, index) => {
            if (index === COLUMN_INDEXES.SHEET_ROW_INDEX) return;  // مخفی کردن ستون شماره ردیف شیت
            if (index === COLUMN_INDEXES.ROW_NUM) return;           // 🔥 مخفی کردن ستون "ردیف"
            if (index > COLUMN_INDEXES.DATE) return;
            if (HIDDEN_COLUMNS_IN_TABLE.includes(index)) return;

            headerHtml += `<th class="p-4 whitespace-nowrap bg-slate-50 dark:bg-slate-700 text-center">${header}</th>`;
        });

        headerHtml += '</tr>';
        thead.innerHTML = headerHtml;

        let rowsHtml = "";
        data.slice().reverse().forEach(row => {
            const rowNumber = row[COLUMN_INDEXES.SHEET_ROW_INDEX];
            const transfer = app.processTransfers[rowNumber];
            const displayName = transfer ? (transfer.fullName || row[COLUMN_INDEXES.FULL_NAME]) : row[COLUMN_INDEXES.FULL_NAME];
            const displayPhone = transfer ? (transfer.phone || row[COLUMN_INDEXES.PHONE]) : row[COLUMN_INDEXES.PHONE];
            const displayNationalId = transfer ? (transfer.nationalId || row[COLUMN_INDEXES.NATIONAL_ID]) : row[COLUMN_INDEXES.NATIONAL_ID];
            const isManual = !String(row[COLUMN_INDEXES.FULL_NAME]).includes('\u200B');
            const isAssignedBySupervisor = String(row[COLUMN_INDEXES.ASSIGNED_BY] || "").trim() !== "";

            // Determine row background class
            let rowBgClass;
            if (isAssignedBySupervisor) {
                // Highlight assigned rows with a distinct color
                rowBgClass = "bg-blue-50 dark:bg-blue-900/20 hover:bg-blue-100 dark:hover:bg-blue-900/40 border-blue-200 dark:border-blue-800";
            } else if (isManual) {
                rowBgClass = "bg-emerald-50 dark:bg-emerald-900/20 hover:bg-emerald-100 dark:hover:bg-emerald-900/40 border-emerald-100 dark:border-emerald-800";
            } else {
                rowBgClass = "hover:bg-slate-50 dark:hover:bg-slate-700/50 border-slate-100 dark:border-slate-700";
            }

            let cellsHtml = '';
            for (let i = 0; i <= COLUMN_INDEXES.DATE; i++) {
                if (i === COLUMN_INDEXES.SHEET_ROW_INDEX) continue;
                if (i === COLUMN_INDEXES.ROW_NUM) continue;
                if (HIDDEN_COLUMNS_IN_TABLE.includes(i)) continue;

                let value = row[i] || "";
                if (i === COLUMN_INDEXES.FULL_NAME) value = displayName || value;
                if (i === COLUMN_INDEXES.PHONE) value = displayPhone || value;
                if (i === COLUMN_INDEXES.NATIONAL_ID) value = displayNationalId || value;
                let cellContent;
                if (i === COLUMN_INDEXES.ROW_NUM) {
                    const deleteButton = isAdmin ? `<button onclick="app.deleteRow(${rowNumber}, event)" class="text-rose-400 hover:text-rose-600 transition" title="حذف"><svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path></svg></button>` : '';
                    // Add badge for assigned rows
                    const assignedBadge = isAssignedBySupervisor ? `<span class="text-[10px] bg-blue-500 text-white px-1.5 py-0.5 rounded ml-1" title="تخصیص داده شده توسط supervisor">تخصیص</span>` : '';
                    cellContent = `<div class="flex items-center justify-center gap-2 text-xs text-slate-400 dark:text-slate-600">${rowNumber} ${assignedBadge} ${deleteButton}</div>`;
                } else {
                    let badge = '';
                    if (i === COLUMN_INDEXES.FULL_NAME && transfer) {
                        badge = `<span class="ml-2 text-[10px] font-bold bg-amber-100 text-amber-700 dark:bg-amber-900/30 dark:text-amber-200 px-1.5 py-0.5 rounded" title="انتقال فرایند از ${transfer.refName || ''}">انتقال فرایند</span>`;
                    }
                    cellContent = `<span class="text-xs font-medium text-slate-700 dark:text-slate-200">${value}</span>${badge}`;
                }
                cellsHtml += `<td class="p-2 whitespace-nowrap text-center">${cellContent}</td>`;
            }
            rowsHtml += `<tr class="${rowBgClass} transition group border-b last:border-0 cursor-pointer" onclick="app.openCustomer(${rowNumber})">${cellsHtml}</tr>`;
        });
        tbody.innerHTML = rowsHtml;
        document.getElementById('totalCountDisplay').innerText = data.length;
        app.updateCharts(data);
    },

    filterByCallResult: function (filters) {
        if (!Array.isArray(filters)) return [];
        const normalized = filters.map(f => String(f || "").trim().toLowerCase());
        return globalData.filter(row => {
            const value = String(row[COLUMN_INDEXES.CALL_RESULT] || "").trim().toLowerCase();
            if (value === "" && normalized.includes("")) return true;
            return normalized.includes(value);
        });
    },

    renderSegments: function () {
        const segments = [
            { id: "new-calls", filters: app.segmentFilters["new-calls"] },
            { id: "under-review", filters: app.segmentFilters["under-review"] },
            { id: "following", filters: app.segmentFilters["following"] },
            { id: "cancelled", filters: app.segmentFilters["cancelled"] },
        ];
        segments.forEach(seg => app.renderSegment(seg));
    },

    renderSegment: function ({ id, filters }) {
        const headerEl = document.getElementById(`tableHeader-${id}`);
        const bodyEl = document.getElementById(`tableBody-${id}`);
        const emptyEl = document.getElementById(`empty-${id}`);
        if (!headerEl || !bodyEl) return;

        const rows = app.filterByCallResult(filters);
        const columns = [
            COLUMN_INDEXES.FULL_NAME,
            COLUMN_INDEXES.PHONE,
            COLUMN_INDEXES.NATIONAL_ID,
            COLUMN_INDEXES.PRODUCT,
            COLUMN_INDEXES.CALL_RESULT,
            COLUMN_INDEXES.DATE
        ];

        let headerHtml = "<tr>";
        columns.forEach(idx => {
            const label = globalHeaders[idx] || "";
            headerHtml += `<th class="p-3 text-center whitespace-nowrap">${label}</th>`;
        });
        headerHtml += "</tr>";
        headerEl.innerHTML = headerHtml;

        if (rows.length === 0) {
            bodyEl.innerHTML = "";
            if (emptyEl) emptyEl.classList.remove("hidden");
            return;
        }
        if (emptyEl) emptyEl.classList.add("hidden");

        let rowsHtml = "";
        rows.slice().reverse().forEach(row => {
            const sheetRow = row[COLUMN_INDEXES.SHEET_ROW_INDEX];
            const transfer = app.processTransfers[sheetRow];
            const displayName = transfer ? (transfer.fullName || row[COLUMN_INDEXES.FULL_NAME]) : row[COLUMN_INDEXES.FULL_NAME];
            const displayPhone = transfer ? (transfer.phone || row[COLUMN_INDEXES.PHONE]) : row[COLUMN_INDEXES.PHONE];
            const displayNationalId = transfer ? (transfer.nationalId || row[COLUMN_INDEXES.NATIONAL_ID]) : row[COLUMN_INDEXES.NATIONAL_ID];
            rowsHtml += `<tr class="hover:bg-slate-50 dark:hover:bg-slate-700/40 cursor-pointer" onclick="app.openCustomer(${sheetRow})">`;
            columns.forEach(idx => {
                let val = row[idx] || "";
                if (idx === COLUMN_INDEXES.FULL_NAME) val = displayName || val;
                if (idx === COLUMN_INDEXES.PHONE) val = displayPhone || val;
                if (idx === COLUMN_INDEXES.NATIONAL_ID) val = displayNationalId || val;
                let badge = '';
                if (idx === COLUMN_INDEXES.FULL_NAME && transfer) {
                    badge = `<span class="ml-2 text-[10px] font-bold bg-amber-100 text-amber-700 dark:bg-amber-900/30 dark:text-amber-200 px-1.5 py-0.5 rounded" title="انتقال فرایند از ${transfer.refName || ''}">انتقال فرایند</span>`;
                }
                rowsHtml += `<td class="p-2 text-center whitespace-nowrap text-sm text-slate-700 dark:text-slate-200">${val}${badge}</td>`;
            });
            rowsHtml += `</tr>`;
        });
        bodyEl.innerHTML = rowsHtml;
    },

    updateCharts: function (data) {
        const productTotals = {};
        data.forEach(r => productTotals[r[COLUMN_INDEXES.PRODUCT]] = (productTotals[r[COLUMN_INDEXES.PRODUCT]] || 0) + 1);
        const ctxBar = document.getElementById('barChart');
        if (this.chartBar) this.chartBar.destroy();
        this.chartBar = new Chart(ctxBar, {
            type: 'bar',
            data: { labels: Object.keys(productTotals), datasets: [{ data: Object.values(productTotals), borderRadius: 6 }] },
            options: {
                responsive: true, maintainAspectRatio: false,
                onClick: (e, elements, chart) => {
                    if (elements[0]) {
                        const i = elements[0].index;
                        app.openChartModal(chart.data.labels[i], COLUMN_INDEXES.PRODUCT);
                    }
                },
                layout: { padding: 10 },
                scales: {
                    x: { reverse: true, grid: { display: false }, ticks: { color: Chart.defaults.color } },
                    y: { position: 'right', grid: { color: document.documentElement.classList.contains('dark') ? '#334155' : '#f1f5f9' }, ticks: { color: Chart.defaults.color } }
                },
                plugins: { legend: { display: false } }
            }
        });
        const smartCardBuckets = {};
        data.forEach(r => { const s = String(r[COLUMN_INDEXES.SMART_CARD]) || 'نامشخص'; smartCardBuckets[s] = (smartCardBuckets[s] || 0) + 1; });
        const ctxPie = document.getElementById('pieChart');
        if (this.chartPie) this.chartPie.destroy();
        this.chartPie = new Chart(ctxPie, {
            type: 'doughnut',
            data: { labels: Object.keys(smartCardBuckets), datasets: [{ data: Object.values(smartCardBuckets), borderWidth: 0 }] },
            options: {
                responsive: true, maintainAspectRatio: false,
                onClick: (e, elements, chart) => {
                    if (elements[0]) {
                        const i = elements[0].index;
                        app.openChartModal(chart.data.labels[i], COLUMN_INDEXES.SMART_CARD);
                    }
                },
                layout: { padding: 10 },
                plugins: { legend: { position: 'bottom', rtl: true, labels: { boxWidth: 10, font: { family: 'Vazirmatn' }, padding: 15, color: Chart.defaults.color } } }
            }
        });
    },

    openChartModal: function (filterValue, columnIndex) {
        app.currentFilteredData = globalData.filter(row => {
            const rawValue = String(row[columnIndex] || "").trim();
            const effectiveLabel = rawValue === "" ? "نامشخص" : rawValue;
            return effectiveLabel === String(filterValue);
        });

        app.navigateTo('view-chart-details');

        const title = document.getElementById('chartDetailTitle');
        const thead = document.getElementById('chartDetailHeader');
        const tbody = document.getElementById('chartDetailBody');

        title.innerText = `لیست: ${filterValue} (${app.currentFilteredData.length} مورد)`;

        const excludedIndices = [20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31];

        let headerHtml = '<tr>';
        for (let i = 0; i < globalHeaders.length; i++) {
            if (!excludedIndices.includes(i)) {
                headerHtml += `<th class="p-3 whitespace-nowrap text-center">${globalHeaders[i]}</th>`;
            }
        }
        headerHtml += '</tr>';
        thead.innerHTML = headerHtml;

        let rowsHtml = '';
        app.currentFilteredData.forEach(row => {
            rowsHtml += `<tr class="border-b border-slate-100 dark:border-slate-700 hover:bg-slate-50 dark:hover:bg-slate-700/50">`;
            for (let i = 0; i < globalHeaders.length; i++) {
                if (!excludedIndices.includes(i)) {
                    rowsHtml += `<td class="p-3 whitespace-nowrap text-center text-slate-600 dark:text-slate-300">${row[i] || '-'}</td>`;
                }
            }
            rowsHtml += `</tr>`;
        });
        tbody.innerHTML = rowsHtml || '<tr><td colspan="5" class="p-4 text-center text-slate-400">داده‌ای یافت نشد</td></tr>';
    },

    downloadChartCSV: function () {
        if (!app.currentFilteredData || app.currentFilteredData.length === 0) return;
        const title = document.getElementById('chartDetailTitle').innerText;

        const excludedIndices = [20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31];

        let csv = "\uFEFF";
        csv += globalHeaders.filter((_, i) => !excludedIndices.includes(i)).join(",") + "\n";

        app.currentFilteredData.forEach(row => {
            const cleanRow = [];
            for (let i = 0; i < globalHeaders.length; i++) {
                if (!excludedIndices.includes(i)) {
                    let str = String(row[i] || "");
                    if (str.includes(',')) str = `"${str.replace(/"/g, '""')}"`;
                    cleanRow.push(str);
                }
            }
            csv += cleanRow.join(",") + "\n";
        });

        const a = document.createElement('a');
        a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
        a.download = `Chart_Data.csv`;
        a.click();
    },

    downloadCSV: function () {

        // اولویت ۱: مقدار ذخیره شده در sessionStorage (همیشه صحیح است)
        let rowId = Number(sessionStorage.getItem("customerRow") || 0);

        // اولویت ۲: مقدار موجود در صفحه (اگر ست شده باشد)
        const detailInput = document.getElementById("detailRowIndex");
        if (detailInput && Number(detailInput.value) > 0) {
            rowId = Number(detailInput.value);
        }

        if (!rowId) {
            alert("شناسه مشتری یافت نشد.");
            return;
        }

        // پیدا کردن ردیف از دیتای اصلی
        const row = globalData.find(r => Number(r[COLUMN_INDEXES.SHEET_ROW_INDEX]) === rowId);

        if (!row) {
            alert("داده‌های مشتری یافت نشد.");
            return;
        }

        // ساختن CSV
        const headers = globalHeaders.slice(0);
        let csv = "\uFEFF" + headers.join(",") + "\n";

        const cleanRow = row.map(cell => {
            if (!cell) return "";

            let v = String(cell);

            // تمیزکردن فیلدهای فایل: فقط اسم فایل‌ها را می‌گیرد
            if (v.includes("|")) {
                v = v
                    .split(",")
                    .map(p => p.split("|")[1] || "")
                    .join("; ");
            }

            // فرمت CSV: جدا کردن با دابل کوتیشن اگر ویرگول داشت
            if (v.includes(",")) v = `"${v.replace(/"/g, '""')}"`;

            return v;
        });

        csv += cleanRow.join(",") + "\n";

        const a = document.createElement("a");
        a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" }));
        a.download = `Customer_${rowId}.csv`;
        a.click();
    },

    openCustomer: function (rowIndex) {
        sessionStorage.setItem('customerRow', rowIndex);
        app.navigateTo('details', rowIndex);
    },

    saveCustomerDetails: function () {
        const rowIndex = Number(document.getElementById('detailRowIndex').value || 0);
        if (!rowIndex) return;
        const statusDiv = document.getElementById('detailStatus');
        const rowData = globalData.find(r => r[COLUMN_INDEXES.SHEET_ROW_INDEX] === rowIndex);
        if (!rowData) { statusDiv.innerText = "ردیف یافت نشد"; return; }

        const updatedRow = [...rowData];

        for (let i = COLUMN_INDEXES.CALL; i <= COLUMN_INDEXES.CANCELLATION; i++) {
            const el = document.getElementById('field_' + i);
            if (!el) continue;
            updatedRow[i] = el.value;
        }

        updatedRow[COLUMN_INDEXES.CALC_BRAND] = document.getElementById('calcBrand').value;
        updatedRow[COLUMN_INDEXES.CALC_AXLE] = document.getElementById('calcAxle').value;
        updatedRow[COLUMN_INDEXES.CALC_USAGE] = document.getElementById('calcUsage').value;
        updatedRow[COLUMN_INDEXES.CALC_DESC] = document.getElementById('calcDesc').value;
        updatedRow[COLUMN_INDEXES.CALC_PAYMENT] = document.getElementById('calcPaymentMethod').value;
        updatedRow[COLUMN_INDEXES.CALC_PRICE] = document.getElementById('displayTotalPrice').innerText;

        // Save prepayment values
        const customerPaidEl = document.getElementById('prepayment_customer_paid');
        const targetPrepaymentEl = document.getElementById('prepayment_target');
        if (customerPaidEl) updatedRow[COLUMN_INDEXES.CUSTOMER_PAID_PREPAYMENT] = customerPaidEl.value || "";
        if (targetPrepaymentEl) updatedRow[COLUMN_INDEXES.TARGET_PREPAYMENT] = targetPrepaymentEl.value || "";

        const dataToSend = updatedRow.slice(0, globalHeaders.length);

        statusDiv.innerText = "در حال ذخیره...";
        google.script.run.withSuccessHandler((msg) => {
            statusDiv.innerText = msg + " ✔";
            const idx = globalData.findIndex(r => r[COLUMN_INDEXES.SHEET_ROW_INDEX] === rowIndex);
            if (idx !== -1) globalData[idx] = updatedRow;
            setTimeout(() => statusDiv.innerText = "", 3000);
            app.updateCharts(globalData);
        }).withFailureHandler(() => {
            statusDiv.innerText = "خطا در ذخیره";
        }).updateLeadData(rowIndex, dataToSend);
    },

    startAutoRefresh: function () {
        if (autoRefreshInterval) clearInterval(autoRefreshInterval);
        autoRefreshInterval = setInterval(() => { app.refreshTable(true); }, 15000);
    },
    stopAutoRefresh: function () {
        if (autoRefreshInterval) { clearInterval(autoRefreshInterval); autoRefreshInterval = null; }
    },

    transferToAgency: function (sheetRowIndex) {
        const city = document.getElementById('agencyCitySelect')?.value || '';
        const agency = document.getElementById('agencyNameSelect')?.value || '';
        const btn = document.getElementById('transferToAgencyBtn');

        if (!city || !agency) {
            this.showTransferAgencyToast('لطفاً شهر و نمایندگی را انتخاب کنید', true);
            return;
        }

        if (btn) {
            btn.disabled = true;
            btn.innerText = 'در حال انتقال...';
        }

        google.script.run.withSuccessHandler((res) => {
            if (res && res.success) {
                this.showTransferAgencyToast('انتقال به نمایندگی ثبت شد', false);
                app.refreshTable(false);
            } else {
                this.showTransferAgencyToast(res && res.message ? res.message : 'خطا در انتقال', true);
            }
        }).withFailureHandler((error) => {
            this.showTransferAgencyToast(error && error.message ? error.message : 'خطا در انتقال', true);
        }).transferToAgency(sheetRowIndex, city, agency);

        setTimeout(() => {
            if (btn) {
                btn.disabled = false;
                btn.innerText = 'انتقال به نمایندگی';
            }
        }, 600);
    },

    loadAgencyCities: function () {
        google.script.run.withSuccessHandler((cities) => {
            const citySelect = document.getElementById('agencyCitySelect');
            if (!citySelect) return;
            const uniqueCities = Array.isArray(cities) ? cities : [];
            citySelect.innerHTML = '<option value="">انتخاب شهر</option>' + uniqueCities.map(c => `<option value="${c}">${c}</option>`).join('');
            this.onAgencyCityChange(); // ریست لیست نمایندگی
        }).withFailureHandler(() => {
            this.showTransferAgencyToast('خطا در بارگذاری شهرها', true);
        }).getAgencyCities();
    },

    onAgencyCityChange: function () {
        const city = document.getElementById('agencyCitySelect')?.value || '';
        const agencySelect = document.getElementById('agencyNameSelect');
        if (!agencySelect) return;

        if (!city) {
            agencySelect.innerHTML = '<option value="">ابتدا شهر را انتخاب کنید</option>';
            this.updateTransferAgencyBtnState();
            return;
        }

        google.script.run.withSuccessHandler((list) => {
            const agencies = Array.isArray(list) ? list : [];
            agencySelect.innerHTML = '<option value="">انتخاب نمایندگی</option>' + agencies.map(a => `<option value="${a}">${a}</option>`).join('');
            this.updateTransferAgencyBtnState();
        }).withFailureHandler(() => {
            agencySelect.innerHTML = '<option value="">خطا در بارگذاری</option>';
            this.updateTransferAgencyBtnState();
        }).getAgenciesByCity(city);
    },

    updateTransferAgencyBtnState: function () {
        const city = document.getElementById('agencyCitySelect')?.value || '';
        const agency = document.getElementById('agencyNameSelect')?.value || '';
        const btn = document.getElementById('transferToAgencyBtn');
        if (btn) {
            btn.disabled = !(city && agency);
        }
    },

    showTransferAgencyToast: function (message, isError = false) {
        const toast = document.getElementById('transferAgencyToast');
        if (!toast) return;
        toast.innerText = message;
        toast.className = `text-sm px-3 py-2 rounded-lg ${isError ? 'bg-rose-50 text-rose-700 border border-rose-200 dark:bg-rose-900/30 dark:text-rose-200 dark:border-rose-800' : 'bg-emerald-50 text-emerald-700 border border-emerald-200 dark:bg-emerald-900/30 dark:text-emerald-200 dark:border-emerald-800'}`;
        toast.classList.remove('hidden');
        setTimeout(() => toast.classList.add('hidden'), 3000);
    },

    startAgencyAutoRefresh: function () {
        app.stopAutoRefresh(); // Stop any existing refresh
        autoRefreshInterval = setInterval(() => { app.refreshAgencyTable(true); }, 15000);
    },

    refreshAgencyTable: function (isAuto = false) {
        const user = sessionStorage.getItem('appUser');
        if (!user) return;

        const btn = document.getElementById('agencyRefreshBtn');
        const loadingState = document.getElementById('agencyLoadingState');
        const emptyState = document.getElementById('agencyEmptyState');

        if (btn) btn.classList.add('animate-spin');
        if (loadingState && !isAuto) loadingState.classList.remove('hidden');
        if (emptyState) emptyState.classList.add('hidden');

        google.script.run.withSuccessHandler((res) => {
            if (btn) btn.classList.remove('animate-spin');
            if (loadingState) loadingState.classList.add('hidden');

            agencyData = res.data;
            app.renderAgencyTable(res.data, res.headers);
            app.applyAgencySearchIfNeeded();
        }).withFailureHandler((error) => {
            if (btn) btn.classList.remove('animate-spin');
            if (loadingState) loadingState.classList.add('hidden');
            alert("خطا در بارگذاری داده‌های نمایندگی: " + error.message);
        }).getAgencyData();
    },

    renderAgencyTable: function (data, headers) {
        const tbody = document.getElementById('agencyTableBody');
        const thead = document.getElementById('agencyTableHeader');
        const emptyState = document.getElementById('agencyEmptyState');

        if (!tbody || !thead) return;

        // Use headers from parameter or fallback to globalHeaders or CRM_HEADERS
        const headerArray = Array.isArray(headers) ? headers : (globalHeaders.length > 0 ? globalHeaders : [
            "ردیف", "نام کارشناس", "تلفن همراه", "نام و نام خانوادگی",
            "کد ملی", "نوع مشتری", "محصول مورد نیاز", "تاریخ ثبت نام", "تماس", "نتیجه تماس",
            "وضعیت کارت هوشمند", "سن", "وضعیت رتبه اعتباری", "وضعیت تریلی فرسوده", "وضعیت تضامین",
            "چک صیادی", "وضعیت اسقاط", "واریز پیش پرداخت", "وضعیت ترهین تضامین",
            "وضعیت قرار داد", "وضعیت تحویل", "دلایل انصراف",
            "فایل کارت ملی", "فایل استعلام سخا", "فایل فیش واریزی",
            "برند انتخابی", "محور انتخابی", "کاربری انتخابی", "توضیحات محصول",
            "روش پرداخت", "قیمت نهایی", "Customer_Paid_Prepayment", "Target_Prepayment"
        ]);

        if (data.length === 0) {
            tbody.innerHTML = '';
            thead.innerHTML = "";
            if (emptyState) emptyState.classList.remove('hidden');
            return;
        }

        if (emptyState) emptyState.classList.add('hidden');

        // Render header - show key columns
        let headerHtml = '<tr>';
        const displayColumns = [
            COLUMN_INDEXES.FULL_NAME,
            COLUMN_INDEXES.PHONE,
            COLUMN_INDEXES.NATIONAL_ID,
            COLUMN_INDEXES.PRODUCT,
            COLUMN_INDEXES.DATE,
            COLUMN_INDEXES.EXPERT_NAME,
            COLUMN_INDEXES.CALL_RESULT,
            COLUMN_INDEXES.CUSTOMER_TYPE
        ];

        displayColumns.forEach(idx => {
            const label = headerArray[idx] || "";
            headerHtml += `<th class="p-4 whitespace-nowrap bg-slate-50 dark:bg-slate-700 text-center text-slate-500 dark:text-slate-400 font-bold">${label}</th>`;
        });
        headerHtml += '</tr>';
        thead.innerHTML = headerHtml;

        // Render rows
        let rowsHtml = "";
        data.slice().reverse().forEach(row => {
            const rowNumber = row[COLUMN_INDEXES.SHEET_ROW_INDEX];
            rowsHtml += `<tr class="hover:bg-slate-50 dark:hover:bg-slate-700/50 border-b border-slate-100 dark:border-slate-700 cursor-pointer" onclick="app.openAgencyCustomer(${rowNumber})">`;
            displayColumns.forEach(idx => {
                const val = row[idx] || "";
                rowsHtml += `<td class="p-2 whitespace-nowrap text-center text-sm text-slate-700 dark:text-slate-200">${val}</td>`;
            });
            rowsHtml += `</tr>`;
        });
        tbody.innerHTML = rowsHtml;
    },

    openAgencyCustomer: function (rowIndex) {
        sessionStorage.setItem('customerRow', rowIndex);
        sessionStorage.setItem('isAgencyCustomer', 'true');
        app.navigateTo('details', rowIndex);
    },

    // ===========================
    // گزارش تماس
    // ===========================
    showCallReportForm: function () {
        const form = document.getElementById('callReportForm');
        if (form) form.classList.remove('hidden');

        const role = (sessionStorage.getItem('role') || '').toLowerCase();
        const labelSup = document.getElementById('labelSendToSupervisor');
        const labelAdmin = document.getElementById('labelSendToAdmin');

        if (role === 'expert') {
            if (labelSup) labelSup.classList.remove('hidden');
            if (labelAdmin) labelAdmin.classList.add('hidden');
        } else if (role === 'supervisor') {
            if (labelSup) labelSup.classList.add('hidden');
            if (labelAdmin) labelAdmin.classList.remove('hidden');
        } else {
            // سایر نقش‌ها: هر دو را مخفی کن
            if (labelSup) labelSup.classList.add('hidden');
            if (labelAdmin) labelAdmin.classList.add('hidden');
        }
    },
    hideCallReportForm: function () {
        const form = document.getElementById('callReportForm');
        const textarea = document.getElementById('callReportText');
        const status = document.getElementById('callReportStatus');
        if (form) form.classList.add('hidden');
        if (textarea) textarea.value = '';
        if (status) status.innerText = '';
    },
    loadProcessTransfers: function (rowIndices, cb) {
        google.script.run.withSuccessHandler((res) => {
            if (res && res.success && res.transfers) {
                this.processTransfers = res.transfers;
            } else {
                this.processTransfers = {};
            }
            if (typeof cb === 'function') cb();
        }).withFailureHandler(() => {
            this.processTransfers = {};
            if (typeof cb === 'function') cb();
        }).getProcessTransfers(rowIndices || []);
    },
    loadCallReports: function (rowIndex) {
        const container = document.getElementById('callReportsContainer');
        const empty = document.getElementById('callReportsEmpty');
        const status = document.getElementById('callReportStatus');
        if (!container) return;

        const numericRowIndex = Number(rowIndex);
        if (Number.isNaN(numericRowIndex)) {
            console.warn('loadCallReports: invalid rowIndex', rowIndex);
            return;
        }

        if (status) { status.innerText = 'در حال بارگذاری...'; status.classList.remove('text-rose-500'); }
        console.log('loadCallReports -> rowIndex:', numericRowIndex);

        google.script.run.withSuccessHandler((result) => {
            console.log('getCallReports result:', result);
            console.log('result.success:', result ? result.success : 'undefined');
            console.log('Array.isArray(result?.reports):', result ? Array.isArray(result.reports) : 'undefined');
            console.log('reports length:', result && Array.isArray(result.reports) ? result.reports.length : 'undefined');

            container.innerHTML = '';
            const validList = result && result.success === true && Array.isArray(result.reports) ? result.reports : [];

            if (validList.length === 0) {
                if (empty) empty.classList.remove('hidden');
                if (status) status.innerText = result && result.success === false ? (result.message || 'خطا در بارگذاری گزارش‌ها') : '';
                return;
            }

            if (empty) empty.classList.add('hidden');

            const total = validList.length;
            validList.forEach((item, idx) => {
                const num = total - idx; // جدیدترین = شماره بزرگتر
                const card = document.createElement('div');
                card.className = 'p-3 rounded-xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 shadow-sm';
                card.innerHTML = `
                        <div class="flex items-center justify-between mb-1">
                          <div class="text-sm font-black text-slate-900 dark:text-white">گزارش تماس ${num}</div>
                          <div class="text-sm font-bold text-slate-800 dark:text-white">${item.expertName || 'نامشخص'}</div>
                          <div class="text-xs text-slate-400 dark:text-slate-500">${item.date || ''} ${item.time || ''}</div>
                        </div>
                        <div class="text-sm text-slate-700 dark:text-slate-200 whitespace-pre-line">${item.reportText || ''}</div>
                      `;
                container.appendChild(card);
            });
            if (status) status.innerText = '';
        }).withFailureHandler((err) => {
            console.error('getCallReports error:', err);
            if (empty) {
                empty.classList.remove('hidden');
                empty.innerText = 'خطا در بارگذاری گزارش‌ها';
            }
            if (status) { status.innerText = 'خطا در بارگذاری گزارش‌ها'; status.classList.add('text-rose-500'); }
        }).getCallReports(numericRowIndex);
    },
    saveCallReport: function (rowIndex) {
        const textarea = document.getElementById('callReportText');
        const status = document.getElementById('callReportStatus');
        const btn = document.getElementById('saveCallReportBtn');
        const sendToSupervisorCheckbox = document.getElementById('sendToSupervisorCheckbox');
        const sendToAdminCheckbox = document.getElementById('sendToAdminCheckbox');
        if (!textarea) return;
        const text = textarea.value.trim();
        if (!text) {
            if (status) status.innerText = 'متن گزارش را وارد کنید';
            return;
        }
        if (btn) { btn.disabled = true; btn.innerText = 'در حال ثبت...'; }
        if (status) { status.innerText = 'در حال ذخیره...'; status.classList.remove('text-rose-500'); }

        const expert = sessionStorage.getItem('appUser') || '';
        const role = (sessionStorage.getItem('role') || '').toLowerCase();
        google.script.run.withSuccessHandler((res) => {
            if (res && res.success) {
                const shouldSendToSupervisor = role === 'expert' && sendToSupervisorCheckbox && sendToSupervisorCheckbox.checked;
                const shouldSendToAdmin = role === 'supervisor' && sendToAdminCheckbox && sendToAdminCheckbox.checked;
                const numericRow = Number(rowIndex) || 0;

                const afterReportSaved = () => {
                    app.hideCallReportForm();
                    app.loadCallReports(numericRow);
                };

                if (shouldSendToSupervisor) {
                    const supervisor = (sessionStorage.getItem('supervisor') || '').toLowerCase();
                    if (supervisor) {
                        const payload = {
                            type: 'expert_to_supervisor',
                            fromUser: expert,
                            toUser: supervisor,
                            expertUsername: expert,
                            supervisorUsername: supervisor,
                            relatedRowIndex: numericRow,
                            text: text
                        };
                        google.script.run.withSuccessHandler(() => {
                            afterReportSaved();
                        }).withFailureHandler(() => {
                            // اگر ارسال پیام خطا داد، فقط لاگ کن و بقیه فرایند ادامه یابد
                            console.error('sendMessage error');
                            afterReportSaved();
                        }).sendMessage(payload);
                    } else {
                        // اگر supervisor در سشن موجود نبود، بدون ارسال پیام ادامه بده
                        afterReportSaved();
                    }
                } else if (shouldSendToAdmin) {
                    const payloadAdmin = {
                        type: 'supervisor_to_admin',
                        fromUser: expert.toLowerCase(),
                        toUser: 'admin',
                        expertUsername: expert.toLowerCase(),
                        supervisorUsername: expert.toLowerCase(),
                        relatedRowIndex: numericRow,
                        text: text
                    };
                    google.script.run.withSuccessHandler(() => {
                        afterReportSaved();
                    }).withFailureHandler(() => {
                        console.error('sendMessage to admin error');
                        afterReportSaved();
                    }).sendMessage(payloadAdmin);
                } else {
                    afterReportSaved();
                }
            } else {
                if (status) { status.innerText = res && res.message ? res.message : 'خطا در ذخیره'; status.classList.add('text-rose-500'); }
            }
        }).withFailureHandler(() => {
            if (status) { status.innerText = 'خطا در ذخیره'; status.classList.add('text-rose-500'); }
        }).saveCallReport({
            rowIndex: rowIndex,
            expertName: expert,
            reportText: text
        });

        setTimeout(() => {
            if (btn) { btn.disabled = false; btn.innerText = 'ثبت گزارش'; }
        }, 800);
    },

    searchAgencyLeads: function () {
        const text = document.getElementById("agencySearchInput").value.trim().toLowerCase();

        if (!text) {
            const headerArray = globalHeaders.length > 0 ? globalHeaders : [
                "ردیف", "نام کارشناس", "تلفن همراه", "نام و نام خانوادگی",
                "کد ملی", "نوع مشتری", "محصول مورد نیاز", "تاریخ ثبت نام", "تماس", "نتیجه تماس",
                "وضعیت کارت هوشمند", "سن", "وضعیت رتبه اعتباری", "وضعیت تریلی فرسوده", "وضعیت تضامین",
                "چک صیادی", "وضعیت اسقاط", "واریز پیش پرداخت", "وضعیت ترهین تضامین",
                "وضعیت قرار داد", "وضعیت تحویل", "دلایل انصراف",
                "فایل کارت ملی", "فایل استعلام سخا", "فایل فیش واریزی",
                "برند انتخابی", "محور انتخابی", "کاربری انتخابی", "توضیحات محصول",
                "روش پرداخت", "قیمت نهایی", "Customer_Paid_Prepayment", "Target_Prepayment"
            ];
            app.renderAgencyTable(agencyData, headerArray);
            return;
        }

        const filtered = agencyData.filter(r => {
            return (
                String(r[COLUMN_INDEXES.FULL_NAME]).toLowerCase().includes(text) ||
                String(r[COLUMN_INDEXES.PHONE]).toLowerCase().includes(text) ||
                String(r[COLUMN_INDEXES.NATIONAL_ID]).toLowerCase().includes(text) ||
                String(r[COLUMN_INDEXES.PRODUCT]).toLowerCase().includes(text)
            );
        });

        const headerArray = globalHeaders.length > 0 ? globalHeaders : [
            "ردیف", "نام کارشناس", "تلفن همراه", "نام و نام خانوادگی",
            "کد ملی", "نوع مشتری", "محصول مورد نیاز", "تاریخ ثبت نام", "تماس", "نتیجه تماس",
            "وضعیت کارت هوشمند", "سن", "وضعیت رتبه اعتباری", "وضعیت تریلی فرسوده", "وضعیت تضامین",
            "چک صیادی", "وضعیت اسقاط", "واریز پیش پرداخت", "وضعیت ترهین تضامین",
            "وضعیت قرار داد", "وضعیت تحویل", "دلایل انصراف",
            "فایل کارت ملی", "فایل استعلام سخا", "فایل فیش واریزی",
            "برند انتخابی", "محور انتخابی", "کاربری انتخابی", "توضیحات محصول",
            "روش پرداخت", "قیمت نهایی", "Customer_Paid_Prepayment", "Target_Prepayment"
        ];
        app.renderAgencyTable(filtered, headerArray);
    },

    applyAgencySearchIfNeeded: function () {
        const input = document.getElementById("agencySearchInput");
        if (!input) return;

        const text = input.value.trim();
        if (text.length > 0) {
            app.searchAgencyLeads();
        }
    },

    // ===========================
    // تخصیص مشتری (ادمین → سرپرست، سرپرست → کارشناس)
    // ===========================
    adminAssignableCount: 0,
    supervisorAssignableCount: 0,

    loadAssignmentView: function () {
        const user = sessionStorage.getItem('appUser');
        const role = (sessionStorage.getItem('role') || '').toLowerCase();

        if (!user) {
            alert("ابتدا وارد سیستم شوید.");
            return;
        }

        const adminCard = document.getElementById('admin-assignment-card');
        const supervisorCard = document.getElementById('supervisor-assignment-card');

        if (role === 'admin') {
            if (adminCard) adminCard.classList.remove('hidden');
            if (supervisorCard) supervisorCard.classList.add('hidden');
            google.script.run.withSuccessHandler((res) => {
                const safe = res || { success: false, total: 0 };
                app.adminAssignableCount = Number(safe.total) || 0;
                const span = document.getElementById('adminAssignSourceCount');
                if (span) {
                    span.innerText = `تعداد ردیف‌های قابل تخصیص (نام کارشناس = admin): ${app.adminAssignableCount}`;
                }
                app.clearAdminAssignmentRows();
            }).withFailureHandler((err) => {
                alert((err && err.message) || "خطا در دریافت اطلاعات تخصیص ادمین.");
            }).getAdminAssignmentInfo({
                username: user,
                role: role
            });
        } else if (role === 'supervisor') {
            if (supervisorCard) supervisorCard.classList.remove('hidden');
            if (adminCard) adminCard.classList.add('hidden');
            google.script.run.withSuccessHandler((res) => {
                const safe = res || { success: false, total: 0 };
                app.supervisorAssignableCount = Number(safe.total) || 0;
                const span = document.getElementById('supervisorAssignSourceCount');
                if (span) {
                    span.innerText = `تعداد ردیف‌های قابل تخصیص (نام کارشناس = ${user}): ${app.supervisorAssignableCount}`;
                }
                app.clearSupervisorAssignmentRows();
            }).withFailureHandler((err) => {
                alert((err && err.message) || "خطا در دریافت اطلاعات تخصیص سرپرست.");
            }).getSupervisorAssignmentInfo({
                username: user,
                role: role,
                displayName: sessionStorage.getItem('appName') || ''
            });
        } else {
            alert("فقط ادمین یا سرپرست مجاز به استفاده از این بخش هستند.");
        }
    },

    // --- ادمین → سرپرست‌ها ---
    clearAdminAssignmentRows: function () {
        const container = document.getElementById('adminAssignmentRows');
        if (container) container.innerHTML = '';
        const totalEl = document.getElementById('adminTotalAssignmentCount');
        if (totalEl) totalEl.innerText = '0';
        const errBox = document.getElementById('adminAssignmentErrorBox');
        if (errBox) errBox.classList.add('hidden');
    },

    addAdminAssignmentRow: function () {
        const container = document.getElementById('adminAssignmentRows');
        if (!container) return;

        const rowIndex = container.children.length;
        let options = '<option value="">- انتخاب سرپرست -</option>';
        (allUsers || []).forEach(u => {
            const role = (u.role || '').toLowerCase();
            if (role === 'supervisor') {
                const val = (u.username || '').toLowerCase();
                options += `<option value="${val}">${u.name} (@${val})</option>`;
            }
        });

        const html = `
                <div class="flex items-center gap-3 p-3 bg-slate-50 dark:bg-slate-700/50 rounded-lg border border-slate-200 dark:border-slate-700" data-admin-row="${rowIndex}">
                  <div class="flex-1">
                    <select
                      class="input-modern rounded-lg w-full p-2 text-sm"
                      onchange="app.updateAdminAssignmentSummary()">
                      ${options}
                    </select>
                  </div>
                  <div class="w-32">
                    <input
                      type="number"
                      min="1"
                      value=""
                      oninput="app.updateAdminAssignmentSummary()"
                      placeholder="تعداد"
                      class="input-modern rounded-lg w-full p-2 text-sm text-center">
                  </div>
                  <button
                    type="button"
                    onclick="(function(){ const row = this.closest('[data-admin-row]'); if(row){ row.remove(); app.updateAdminAssignmentSummary(); } }).call(this)"
                    class="text-rose-500 hover:text-rose-700 hover:bg-rose-50 dark:hover:bg-rose-900/30 p-2 rounded-lg transition">
                    <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path>
                    </svg>
                  </button>
                </div>
              `;

        container.insertAdjacentHTML('beforeend', html);
        app.updateAdminAssignmentSummary();
    },

    updateAdminAssignmentSummary: function () {
        const container = document.getElementById('adminAssignmentRows');
        const totalEl = document.getElementById('adminTotalAssignmentCount');
        const errBox = document.getElementById('adminAssignmentErrorBox');
        const errText = document.getElementById('adminAssignmentErrorText');

        if (!container) return;

        let total = 0;
        const rows = container.querySelectorAll('[data-admin-row]');
        rows.forEach(row => {
            const input = row.querySelector('input[type="number"]');
            if (input) {
                const v = parseInt(input.value) || 0;
                total += v;
            }
        });

        if (totalEl) totalEl.innerText = String(total);

        if (total > app.adminAssignableCount) {
            if (errBox) errBox.classList.remove('hidden');
            if (errText) {
                errText.innerText = `مجموع درخواستی (${total}) بیشتر از تعداد ردیف‌های قابل تخصیص (${app.adminAssignableCount}) است.`;
            }
        } else {
            if (errBox) errBox.classList.add('hidden');
        }
    },

    runAdminAssignment: function () {
        const role = (sessionStorage.getItem('role') || '').toLowerCase();
        const user = sessionStorage.getItem('appUser') || '';
        if (role !== 'admin') {
            alert("فقط ادمین می‌تواند این عملیات را انجام دهد.");
            return;
        }

        const container = document.getElementById('adminAssignmentRows');
        if (!container) return;

        const rows = container.querySelectorAll('[data-admin-row]');
        const assignments = [];
        rows.forEach(row => {
            const select = row.querySelector('select');
            const input = row.querySelector('input[type="number"]');
            if (!select || !input) return;
            const supervisorUsername = (select.value || '').toLowerCase().trim();
            const count = parseInt(input.value) || 0;
            if (supervisorUsername && count > 0) {
                assignments.push({ supervisorUsername, count });
            }
        });

        if (assignments.length === 0) {
            alert("لطفاً حداقل یک ردیف تخصیص معتبر وارد کنید.");
            return;
        }

        const totalRequested = assignments.reduce((s, a) => s + (a.count || 0), 0);
        if (totalRequested > app.adminAssignableCount) {
            alert(`مجموع درخواستی (${totalRequested}) بیشتر از تعداد ردیف‌های قابل تخصیص (${app.adminAssignableCount}) است.`);
            return;
        }

        if (!confirm(`آیا مطمئن هستید که می‌خواهید ${totalRequested} ردیف را بین سرپرست‌ها تقسیم کنید؟`)) {
            return;
        }

        google.script.run.withSuccessHandler((res) => {
            if (res && res.success) {
                const assigned = res.assigned || {};
                const parts = [];
                Object.keys(assigned).forEach(k => {
                    if (assigned[k] > 0) parts.push(`${assigned[k]} ردیف به ${k}`);
                });
                alert("✓ " + (parts.join('، ') || 'تخصیص انجام شد.'));
                app.loadAssignmentView();
                app.refreshTable(false);
            } else {
                alert("✗ " + ((res && res.message) || 'خطا در تخصیص.'));
            }
        }).withFailureHandler(err => {
            alert((err && err.message) || 'خطا در تخصیص.');
        }).adminAssignToSupervisors(assignments, { username: user, role });
    },

    // --- سرپرست → کارشناسان خودش ---
    clearSupervisorAssignmentRows: function () {
        const container = document.getElementById('supervisorAssignmentRows');
        if (container) container.innerHTML = '';
        const totalEl = document.getElementById('supervisorTotalAssignmentCount');
        if (totalEl) totalEl.innerText = '0';
        const errBox = document.getElementById('supervisorAssignmentErrorBox');
        if (errBox) errBox.classList.add('hidden');
    },

    addSupervisorAssignmentRow: function () {
        const container = document.getElementById('supervisorAssignmentRows');
        if (!container) return;

        const rowIndex = container.children.length;
        const currentUser = (sessionStorage.getItem('appUser') || '').toLowerCase();

        let options = '<option value="">- انتخاب کارشناس -</option>';
        (allUsers || []).forEach(u => {
            const role = (u.role || '').toLowerCase();
            const supervisor = (u.supervisor || '').toLowerCase();
            if (role === 'expert' && supervisor === currentUser) {
                const val = (u.username || '').toLowerCase();
                options += `<option value="${val}">${u.name} (@${val})</option>`;
            }
        });

        const html = `
                <div class="flex items-center gap-3 p-3 bg-slate-50 dark:bg-slate-700/50 rounded-lg border border-slate-200 dark:border-slate-700" data-supervisor-row="${rowIndex}">
                  <div class="flex-1">
                    <select
                      class="input-modern rounded-lg w-full p-2 text-sm"
                      onchange="app.updateSupervisorAssignmentSummary()">
                      ${options}
                    </select>
                  </div>
                  <div class="w-32">
                    <input
                      type="number"
                      min="1"
                      value=""
                      oninput="app.updateSupervisorAssignmentSummary()"
                      placeholder="تعداد"
                      class="input-modern rounded-lg w-full p-2 text-sm text-center">
                  </div>
                  <button
                    type="button"
                    onclick="(function(){ const row = this.closest('[data-supervisor-row]'); if(row){ row.remove(); app.updateSupervisorAssignmentSummary(); } }).call(this)"
                    class="text-rose-500 hover:text-rose-700 hover:bg-rose-50 dark:hover:bg-rose-900/30 p-2 rounded-lg transition">
                    <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path>
                    </svg>
                  </button>
                </div>
              `;

        container.insertAdjacentHTML('beforeend', html);
        app.updateSupervisorAssignmentSummary();
    },

    updateSupervisorAssignmentSummary: function () {
        const container = document.getElementById('supervisorAssignmentRows');
        const totalEl = document.getElementById('supervisorTotalAssignmentCount');
        const errBox = document.getElementById('supervisorAssignmentErrorBox');
        const errText = document.getElementById('supervisorAssignmentErrorText');

        if (!container) return;

        let total = 0;
        const rows = container.querySelectorAll('[data-supervisor-row]');
        rows.forEach(row => {
            const input = row.querySelector('input[type="number"]');
            if (input) {
                const v = parseInt(input.value) || 0;
                total += v;
            }
        });

        if (totalEl) totalEl.innerText = String(total);

        if (total > app.supervisorAssignableCount) {
            if (errBox) errBox.classList.remove('hidden');
            if (errText) {
                errText.innerText = `مجموع درخواستی (${total}) بیشتر از تعداد ردیف‌های قابل تخصیص (${app.supervisorAssignableCount}) است.`;
            }
        } else {
            if (errBox) errBox.classList.add('hidden');
        }
    },

    runSupervisorAssignment: function () {
        const role = (sessionStorage.getItem('role') || '').toLowerCase();
        const user = sessionStorage.getItem('appUser') || '';
        if (role !== 'supervisor') {
            alert("فقط سرپرست می‌تواند این عملیات را انجام دهد.");
            return;
        }

        const container = document.getElementById('supervisorAssignmentRows');
        if (!container) return;

        const rows = container.querySelectorAll('[data-supervisor-row]');
        const assignments = [];
        rows.forEach(row => {
            const select = row.querySelector('select');
            const input = row.querySelector('input[type="number"]');
            if (!select || !input) return;
            const expertUsername = (select.value || '').toLowerCase().trim();
            const count = parseInt(input.value) || 0;
            if (expertUsername && count > 0) {
                assignments.push({ expertUsername, count });
            }
        });

        if (assignments.length === 0) {
            alert("لطفاً حداقل یک ردیف تخصیص معتبر وارد کنید.");
            return;
        }

        const totalRequested = assignments.reduce((s, a) => s + (a.count || 0), 0);
        if (totalRequested > app.supervisorAssignableCount) {
            alert(`مجموع درخواستی (${totalRequested}) بیشتر از تعداد ردیف‌های قابل تخصیص (${app.supervisorAssignableCount}) است.`);
            return;
        }

        if (!confirm(`آیا مطمئن هستید که می‌خواهید ${totalRequested} ردیف را بین کارشناسان خود تقسیم کنید؟`)) {
            return;
        }

        google.script.run.withSuccessHandler((res) => {
            if (res && res.success) {
                const assigned = res.assigned || {};
                const parts = [];
                Object.keys(assigned).forEach(k => {
                    if (assigned[k] > 0) parts.push(`${assigned[k]} ردیف به ${k}`);
                });
                alert("✓ " + (parts.join('، ') || 'تخصیص انجام شد.'));
                app.loadAssignmentView();
                app.refreshTable(false);
            } else {
                alert("✗ " + ((res && res.message) || 'خطا در تخصیص.'));
            }
        }).withFailureHandler(err => {
            alert((err && err.message) || 'خطا در تخصیص.');
        }).supervisorAssignToExperts(assignments, { username: user, role });
    }
};

function getDriveViewLink(fileId) {
    if (!fileId || typeof fileId !== 'string' || fileId.length < 25) return null;
    return `${fileId}`;
}

// === تابع کلیدی: ساخت لینک‌های واقعی و دکمه حذف ===
function renderFileLinks(cellValue, rowIndex, colIndex) {
    if (!cellValue || cellValue === "-") return `<span class="text-xs text-slate-400 dark:text-slate-600">فایلی نیست</span>`;

    const parts = String(cellValue).split(',');
    let html = '';

    parts.forEach(part => {
        let id = part.trim();
        let name = "فایل";

        if (part.includes('|')) {
            const split = part.split('|');
            id = split[0];
            name = split[1] || "فایل";
        }

        const link = getDriveViewLink(id);
        if (link) {
            // ساختار لینک + دکمه حذف کوچک
            html += `
                 <div class="flex items-center justify-between bg-slate-100 dark:bg-slate-600/30 px-2 py-1 rounded mb-1">
                     <a href="${link}" target="_blank" class="text-xs text-blue-600 dark:text-blue-400 hover:underline truncate max-w-[120px]" title="${name}">
                        📎 ${name}
                     </a>
                     <button onclick="app.deleteFile(${rowIndex}, ${colIndex}, '${id}')" class="text-rose-500 hover:text-rose-700 ml-2" title="حذف فایل">
                        <svg class="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path></svg>
                     </button>
                 </div>`;
        }
    });
    return html;
}

// === تابع جدید: حذف فایل ===
app.deleteFile = function (rowIndex, colIndex, fileId) {
    if (!confirm("آیا از حذف این فایل مطمئن هستید؟")) return;

    const listContainer = document.getElementById(`file_list_${colIndex}`);
    if (listContainer) listContainer.classList.add('opacity-50');

    google.script.run.withSuccessHandler((res) => {
        if (listContainer) listContainer.classList.remove('opacity-50');
        if (res.success) {
            const row = globalData.find(r => r[COLUMN_INDEXES.SHEET_ROW_INDEX] === rowIndex);
            if (row) {
                row[colIndex] = res.fullCellValue;
                if (listContainer) listContainer.innerHTML = renderFileLinks(res.fullCellValue, rowIndex, colIndex);
            }
        } else {
            alert("خطا در حذف: " + res.message);
        }
    }).withFailureHandler((err) => {
        if (listContainer) listContainer.classList.remove('opacity-50');
        alert("خطا در ارتباط با سرور");
    }).deleteFileFromSheet(rowIndex, colIndex, fileId);
};

function submitClientData(e) {
    e.preventDefault();
    const name = document.getElementById('clientName').value;
    const lastName = document.getElementById('clientLastName').value;
    const phone = document.getElementById('clientPhone').value;
    const requiredProduct = document.getElementById('clientRequiredProduct').value;
    const u = sessionStorage.getItem('appUser');
    const n = sessionStorage.getItem('appName');
    const expertName = u || '';
    if (!name || !lastName || !phone || !requiredProduct) return;
    const status = document.getElementById('clientStatus');
    status.innerText = "در حال ذخیره...";
    const nationalId = document.getElementById('clientNationalId').value;
    const clientType = document.getElementById("clientType").value;

    const formData = {
        name,
        lastName,
        phone,
        nationalId,   // 🔥 اضافه شد
        requiredProduct,
        clientType,
        expertName

    };


    google.script.run.withSuccessHandler((msg) => {
        status.innerText = msg + " ✔";
        document.getElementById('clientName').value = "";
        document.getElementById('clientLastName').value = "";
        document.getElementById('clientPhone').value = "";
        document.getElementById('clientRequiredProduct').value = "";
        document.getElementById('clientNationalId').value = "";
        document.getElementById('clientType').value = "";

        app.refreshTable(false);
        setTimeout(() => status.innerText = "", 3000);
    }).withFailureHandler(() => { status.innerText = "خطا"; }).registerNewClient(
        formData,
        {
            username: u || '',
            displayName: n || '',
            role: sessionStorage.getItem('role') || ''
        }
    );
}

function goBack() { if (window.history.length > 1) window.history.back(); else app.navigateTo('dashboard'); }

function toggleTheme() {
    if (document.documentElement.classList.contains('dark')) {
        document.documentElement.classList.remove('dark'); localStorage.setItem('theme', 'light');
        document.getElementById('moonIcon').classList.remove('hidden'); document.getElementById('sunIcon').classList.add('hidden');
    } else {
        document.documentElement.classList.add('dark'); localStorage.setItem('theme', 'dark');
        document.getElementById('moonIcon').classList.add('hidden'); document.getElementById('sunIcon').classList.remove('hidden');
    }
    if (globalData.length > 0) app.updateCharts(globalData);
}

app.uploadFile = function (element, sheetRowIndex, columnIndex) {
    const file = element.files[0];
    if (!file) return;

    const statusDiv = document.getElementById(`upload_status_${columnIndex}`);
    statusDiv.innerText = "در حال آپلود...";
    statusDiv.classList.remove('text-rose-500'); statusDiv.classList.add('text-blue-500');

    const reader = new FileReader();
    reader.onload = function (e) {
        const fileData = { fileName: file.name, mimeType: file.type, bytes: Array.from(new Int8Array(e.target.result)) };

        google.script.run.withSuccessHandler((res) => {
            statusDiv.classList.remove('text-blue-500');
            if (res.success) {
                statusDiv.innerText = "آپلود شد ✔";

                const originalRow = globalData.find(r => r[COLUMN_INDEXES.SHEET_ROW_INDEX] === sheetRowIndex);
                if (originalRow) {
                    originalRow[columnIndex] = res.fullCellValue;
                }

                const fileListContainer = document.getElementById(`file_list_${columnIndex}`);
                if (fileListContainer) {
                    fileListContainer.innerHTML = renderFileLinks(res.fullCellValue, sheetRowIndex, columnIndex);
                }

            } else {
                statusDiv.innerText = "خطا"; statusDiv.classList.add('text-rose-500');
                alert(res.message);
            }
            setTimeout(() => { if (res.success) { statusDiv.innerText = ""; } }, 3000);
        }).withFailureHandler((error) => {
            statusDiv.innerText = "خطا"; statusDiv.classList.add('text-rose-500');
        }).handleFileUpload(fileData, sheetRowIndex, columnIndex);
    };
    reader.readAsArrayBuffer(file);
};

app.saveCellValue = function (element, sheetRowIndex, columnIndex) {
};

app.deleteRow = function (rowIndex) {
    if (!app.isAdmin) { alert("مجوز ندارید."); return; }
    const modal = document.getElementById('confirmModal');
    const msg = document.getElementById('confirmMessage');
    const btn = document.getElementById('confirmYesBtn');
    msg.innerText = `آیا مطمئن هستید که می‌خواهید ردیف ${rowIndex} را حذف کنید؟`;
    modal.classList.remove('hidden');
    btn.onclick = function () {
        modal.classList.add('hidden');
        const loading = document.getElementById('loadingState');
        loading.classList.remove('hidden');
        google.script.run.withSuccessHandler((res) => {
            loading.classList.add('hidden');
            if (res.success) app.refreshTable(false); else alert(res.message);
        }).withFailureHandler(() => {
            loading.classList.add('hidden'); alert("خطا");
        }).deleteData(rowIndex);
    };
};

document.getElementById('themeBtn').onclick = toggleTheme;

window.onload = function () {
    const theme = localStorage.getItem('theme');
    if (theme === 'dark') {
        document.documentElement.classList.add('dark');
        document.getElementById('moonIcon').classList.add('hidden');
        document.getElementById('sunIcon').classList.remove('hidden');
    } else {
        document.documentElement.classList.remove('dark');
        localStorage.setItem('theme', 'light');
        document.getElementById('moonIcon').classList.remove('hidden');
        document.getElementById('sunIcon').classList.add('hidden');
    }

    const u = sessionStorage.getItem('appUser');
    if (!u) {
        app.navigateTo('login', null, false);
        return;
    }

    const role = (sessionStorage.getItem('role') || '').toLowerCase();

    const defaultPage = ROLE_DEFAULT_PAGE[role] || 'dashboard';
    app.navigateTo(defaultPage, null, false);

    // 🔥 2. بعد آخرین view ذخیره‌شده را restore کن
    restoreLastViewByRole(role);

    // 🔥 3. visibility بر اساس نقش
    app.applyRoleVisibility(role);
};

window.onpopstate = function (event) {
    if (event.state) app.navigateTo(event.state.page, event.state.param, false);
    else app.navigateTo('login', null, false);
};

// ===================================================================================
// === تابع اصلی loadDetails (با درج بخش انتقال فرایند) ===
// ===================================================================================
function loadDetails(rowId) {
    const numericRowId = Number(rowId);
    let rowData = globalData.find(r => r[COLUMN_INDEXES.SHEET_ROW_INDEX] === numericRowId);
    // If not found in globalData, check agencyData
    if (!rowData && agencyData.length > 0) {
        rowData = agencyData.find(r => r[COLUMN_INDEXES.SHEET_ROW_INDEX] === numericRowId);
    }
    if (!rowData) {
        document.getElementById('detailName').innerText = "یافت نشد";
        document.getElementById('detailTable').innerHTML = '<p class="text-center text-slate-500 dark:text-slate-400">رکوردی یافت نشد</p>';
        return;
    }
    const transfer = app.processTransfers[numericRowId];
    const displayName = transfer ? (transfer.fullName || rowData[COLUMN_INDEXES.FULL_NAME]) : rowData[COLUMN_INDEXES.FULL_NAME];
    const displayPhone = transfer ? (transfer.phone || rowData[COLUMN_INDEXES.PHONE]) : rowData[COLUMN_INDEXES.PHONE];
    const displayNationalId = transfer ? (transfer.nationalId || rowData[COLUMN_INDEXES.NATIONAL_ID]) : rowData[COLUMN_INDEXES.NATIONAL_ID];
    document.getElementById('detailName').innerText = displayName;

    let html = '';

    html += `
            ${transfer ? `<div class="mb-4 p-3 rounded-xl bg-amber-50 dark:bg-amber-900/30 border border-amber-200 dark:border-amber-800 text-sm font-bold text-amber-700 dark:text-amber-200">این پرونده انتقال فرایند خورده و از طرف ${transfer.refName || 'مشتری قبلی'} می‌باشد.</div>` : ''}

            <div class="mb-6 grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-6">
              <div class="space-y-1">
                <div class="text-xs font-bold text-slate-400 dark:text-slate-500">نام</div>
                <div class="text-sm font-bold text-slate-800 dark:text-white">${displayName || '-'}</div>
              </div>
              <div class="space-y-1">
                <div class="text-xs font-bold text-slate-400 dark:text-slate-500">کارشناس</div>
                <div class="text-sm font-bold text-slate-700 dark:text-slate-200">${rowData[COLUMN_INDEXES.EXPERT_NAME] || '-'}</div>
              </div>
              <div class="space-y-1">
                <div class="text-xs font-bold text-slate-400 dark:text-slate-500">تلفن</div>
                <div class="text-sm font-mono text-slate-700 dark:text-slate-200">${displayPhone || '-'}</div>
              </div>
              <div class="space-y-1">
                <div class="text-xs font-bold text-slate-400 dark:text-slate-500">کد ملی</div>
                <div class="text-sm font-bold text-slate-800 dark:text-white">${displayNationalId || '-'}</div>
              </div>
    
              <div class="space-y-1">
                <div class="text-xs font-bold text-slate-400 dark:text-slate-500">محصول درخواستی اولیه</div>
                <div class="text-sm text-slate-700 dark:text-slate-200">${rowData[COLUMN_INDEXES.PRODUCT] || '-'}</div>
              </div>
              <div class="space-y-1">
                <div class="text-xs font-bold text-slate-400 dark:text-slate-500">تاریخ ثبت</div>
                <div class="text-sm text-slate-700 dark:text-slate-200">${rowData[COLUMN_INDEXES.DATE] || '-'}</div>
              </div>
              <div class="space-y-1">
                <div class="text-xs font-bold text-slate-400 dark:text-slate-500">نوع مشتری</div>
                <div class="text-sm font-bold text-slate-800 dark:text-white">${rowData[COLUMN_INDEXES.CUSTOMER_TYPE] || '-'}</div>
              </div>
    
            </div>
    
            <div class="border-t border-slate-100 dark:border-slate-700 pt-4 mt-2">
              <div class="flex justify-between items-center mb-4">
                <h3 class="text-sm font-bold text-slate-700 dark:text-slate-200">وضعیت و پیگیری‌ها</h3>
                <div id="detailStatus" class="text-xs font-bold text-emerald-500 h-4 text-left"></div>
              </div>
    
              <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 mb-8">
          `;

    for (let i = COLUMN_INDEXES.CALL; i <= COLUMN_INDEXES.CANCELLATION; i++) {
        const header = globalHeaders[i];
        const value = rowData[i] || "";
        html += `<div class="space-y-1">
                        <label class="block text-xs font-bold text-slate-500 dark:text-slate-300 mb-1">${header}</label>`;
        if (SELECTABLE_FIELDS.includes(i)) {
            const options = app.getOptions(i);
            html += `<select id="field_${i}" class="input-modern rounded-xl w-full p-2 text-sm">`;
            options.forEach(opt => {
                html += `<option value="${opt.value}" ${String(opt.value) === String(value) ? 'selected' : ''}>${opt.label}</option>`;
            });
            html += `</select>`;
        } else {
            html += `<input id="field_${i}" class="input-modern rounded-xl w-full p-2 text-sm" value="${value}">`;
        }
        html += `</div>`;
    }
    html += `</div>`; // Closing the grid for follow-up fields

    // ====================================================================
    // === جدول مقایسه پیش پرداخت ===
    // ====================================================================
    const customerPaid = rowData[COLUMN_INDEXES.CUSTOMER_PAID_PREPAYMENT] || "";
    const targetPrepayment = rowData[COLUMN_INDEXES.TARGET_PREPAYMENT] || "";
    const difference = (parseFloat(customerPaid) || 0) - (parseFloat(targetPrepayment) || 0);

    html += `
            <div class="mt-6 p-4 border border-slate-200 dark:border-slate-700 rounded-2xl bg-slate-50 dark:bg-slate-800/50">
              <h3 class="text-sm font-bold text-slate-700 dark:text-slate-200 mb-4">مقایسه پیش پرداخت</h3>
              <div class="grid grid-cols-3 gap-4">
                <div class="space-y-1">
                  <label class="block text-xs font-bold text-slate-500 dark:text-slate-300 mb-1">پیش پرداخت پرداخت شده توسط مشتری</label>
                  <input 
                    id="prepayment_customer_paid" 
                    type="number" 
                    step="0.01"
                    value="${customerPaid}"
                    oninput="updatePrepaymentDifference()"
                    class="input-modern rounded-lg w-full p-2 text-sm"
                    placeholder="0"
                  />
                </div>
                <div class="space-y-1">
                  <label class="block text-xs font-bold text-slate-500 dark:text-slate-300 mb-1">پیش پرداخت هدف/مورد انتظار</label>
                  <input 
                    id="prepayment_target" 
                    type="number" 
                    step="0.01"
                    value="${targetPrepayment}"
                    oninput="updatePrepaymentDifference()"
                    class="input-modern rounded-lg w-full p-2 text-sm"
                    placeholder="0"
                  />
                </div>
                <div class="space-y-1">
                  <label class="block text-xs font-bold text-slate-500 dark:text-slate-300 mb-1">تفاوت</label>
                  <div 
                    id="prepayment_difference" 
                    class="input-modern rounded-lg w-full p-2 text-sm font-bold text-center ${difference >= 0 ? 'text-emerald-600 dark:text-emerald-400' : 'text-rose-600 dark:text-rose-400'}"
                    style="background: rgba(255, 255, 255, 0.9); border: 1px solid #e2e8f0;"
                  >
                    ${difference.toLocaleString('fa-IR')}
                  </div>
                </div>
              </div>
            </div>
          `;

    // ====================================================================
    // === بخش انتقال فرایند (با اعمال اصلاحات Dark Mode، اندازه و فعال‌سازی pt_fullname) ===
    // ====================================================================
    html += `
            <div class="mt-6 p-4 border border-slate-200 dark:border-slate-700 rounded-2xl">
    
              <label class="flex items-center gap-2 text-sm font-bold text-slate-700 dark:text-slate-200">
                <input type="checkbox" id="transferCheckbox" onchange="toggleTransferBox()" />
                انتقال فرایند
              </label>
    
              <div id="transferBox" class="hidden mt-4 grid grid-cols-1 md:grid-cols-4 gap-4 bg-slate-50 dark:bg-slate-800 p-4 rounded-xl">
    
                  <div>
                    <label class="text-xs font-bold text-slate-500 dark:text-slate-300">نام و نام خانوادگی </label>
                    <input 
                      id="pt_fullname" 
                      class="input-modern rounded-lg w-full p-2 text-xs bg-white dark:bg-slate-700 dark:text-white" 
                      placeholder="نام و نام خانوادگی"
                      
                      />
                  </div>
    
                  <div>
                    <label class="text-xs font-bold text-slate-500 dark:text-slate-300">کد ملی</label>
                    <input 
                      id="pt_nationalId" 
                      class="input-modern rounded-lg w-full p-2 text-xs bg-white dark:bg-slate-700 dark:text-white" 
                      placeholder="کد ملی" 
                    />
                  </div>
    
                  <div>
                    <label class="text-xs font-bold text-slate-500 dark:text-slate-300">شماره تماس</label>
                    <input 
                      id="pt_phone"
                      class="input-modern rounded-lg w-full p-2 text-xs bg-white dark:bg-slate-700 dark:text-white" 
                      placeholder="شماره تماس" 
                    />
                  </div>
    
                  <div>
                    <label class="text-xs font-bold text-slate-500 dark:text-slate-300">نام معرف (مشتری فعلی)</label>
                    <input 
                      id="pt_refName" 
                      class="input-modern rounded-lg w-full p-2 text-xs bg-slate-200 dark:bg-slate-700 dark:text-white" 
                      value="${rowData[COLUMN_INDEXES.FULL_NAME] || ''}"
                      disabled 
                    />
                    <input type="hidden" id="transferRowIndex" value="${rowData[COLUMN_INDEXES.SHEET_ROW_INDEX]}">
                  </div>
    
                  <button 
                    onclick="submitProcessTransfer()" 
                    class="w-full bg-emerald-600 hover:bg-emerald-700 text-white py-2 rounded-lg text-sm font-bold shadow-lg transition"
                  >
                     ثبت انتقال فرایند
                  </button>
    
              </div>
    
            </div>
          `;
    // ====================================================================
    // === پایان بخش انتقال فرایند ===
    // ====================================================================

    // ====================================================================
    // === بخش انتقال به نمایندگی ===
    // ====================================================================
    html += `
            <div class="mt-6 p-4 border border-slate-200 dark:border-slate-700 rounded-2xl">
              <label class="flex items-center gap-2 text-sm font-bold text-slate-700 dark:text-slate-200">
                <input type="checkbox" id="transferToAgencyCheckbox" onchange="toggleTransferToAgencyBox()" />
                انتقال به نمایندگی
              </label>
              
              <div id="transferToAgencyBox" class="hidden mt-4 p-4 bg-rose-50 dark:bg-rose-900/20 rounded-xl border border-rose-200 dark:border-rose-800 space-y-3">
                <p class="text-xs text-slate-600 dark:text-slate-300">
                  شهر و نمایندگی مقصد را انتخاب کنید تا برای این مشتری ثبت شود.
                </p>
                <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
                  <div>
                    <label class="text-xs font-bold text-slate-600 dark:text-slate-300 mb-1 block">شهر</label>
                    <select id="agencyCitySelect" class="input-modern rounded-lg w-full p-2 text-sm" onchange="app.onAgencyCityChange()">
                      <option value="">انتخاب شهر</option>
                    </select>
                  </div>
                  <div>
                    <label class="text-xs font-bold text-slate-600 dark:text-slate-300 mb-1 block">نمایندگی</label>
                    <select id="agencyNameSelect" class="input-modern rounded-lg w-full p-2 text-sm" onchange="app.updateTransferAgencyBtnState()">
                      <option value="">ابتدا شهر را انتخاب کنید</option>
                    </select>
                  </div>
                </div>
                <div id="transferAgencyToast" class="hidden text-sm px-3 py-2 rounded-lg"></div>
                <button 
                  id="transferToAgencyBtn"
                  onclick="app.transferToAgency(${rowData[COLUMN_INDEXES.SHEET_ROW_INDEX]})" 
                  class="w-full bg-rose-600 hover:bg-rose-700 text-white py-2 rounded-lg text-sm font-bold shadow-lg transition disabled:opacity-50 disabled:cursor-not-allowed"
                  disabled
                >
                  انتقال به نمایندگی
                </button>
              </div>
            </div>
          `;

    // ====================================================================
    // === بخش گزارش تماس ===
    // ====================================================================
    html += `
            <div class="mt-6 p-4 border border-slate-200 dark:border-slate-700 rounded-2xl space-y-4">
              <div class="flex items-center justify-between">
                <div class="flex items-center gap-2">
                  <span class="w-2 h-2 rounded-full bg-amber-500"></span>
                  <h4 class="font-bold text-slate-800 dark:text-white">گزارش تماس</h4>
                </div>
                <button
                  id="addCallReportBtn"
                  onclick="app.showCallReportForm()"
                  class="flex items-center gap-2 bg-amber-500 hover:bg-amber-600 text-white text-sm font-bold px-4 py-2 rounded-lg transition">
                  <span class="text-lg leading-none">+</span>
                  <span>افزودن گزارش تماس</span>
                </button>
              </div>

              <div id="callReportForm" class="hidden space-y-3">
                <textarea
                  id="callReportText"
                  class="input-modern rounded-xl w-full p-3 text-sm min-h-[120px]"
                  placeholder="متن گزارش تماس را بنویسید..."></textarea>
                <div class="flex flex-col md:flex-row md:items-center md:justify-between gap-3 mt-2">
                  <div class="flex flex-col gap-1 text-xs md:text-sm text-slate-600 dark:text-slate-300">
                    <label id="labelSendToSupervisor" class="inline-flex items-center gap-2">
                    <input
                      type="checkbox"
                      id="sendToSupervisorCheckbox"
                      class="rounded border-slate-300 text-amber-600 focus:ring-amber-500">
                    <span>ارسال این گزارش به سرپرست</span>
                    </label>
                    <label id="labelSendToAdmin" class="inline-flex items-center gap-2 hidden">
                      <input
                        type="checkbox"
                        id="sendToAdminCheckbox"
                        class="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500">
                      <span>ارسال این گزارش به ادمین</span>
                    </label>
                  </div>
                  <div class="flex items-center gap-2">
                  <button
                    id="saveCallReportBtn"
                    onclick="app.saveCallReport(${rowData[COLUMN_INDEXES.SHEET_ROW_INDEX]})"
                    class="bg-amber-600 hover:bg-amber-700 text-white text-sm font-bold px-4 py-2 rounded-lg transition">
                    ثبت گزارش
                  </button>
                  <button
                    onclick="app.hideCallReportForm()"
                    class="text-sm font-bold px-4 py-2 rounded-lg bg-slate-100 dark:bg-slate-700 text-slate-700 dark:text-slate-200 hover:bg-slate-200 dark:hover:bg-slate-600 transition">
                    انصراف
                  </button>
                  <div id="callReportStatus" class="text-sm text-slate-500 dark:text-slate-400"></div>
                  </div>
                </div>
              </div>

              <div id="callReportsContainer" class="space-y-3 max-h-[300px] overflow-y-auto pr-1">
                <div id="callReportsEmpty" class="text-sm text-slate-400">گزارشی ثبت نشده است</div>
              </div>
            </div>
          `;
    // ====================================================================
    // === پایان بخش گزارش تماس ===
    // ====================================================================
    // ====================================================================
    // === پایان بخش انتقال به نمایندگی ===
    // ====================================================================

    html += `<div class="grid grid-cols-1 lg:grid-cols-2 gap-8 border-t border-slate-100 dark:border-slate-700 pt-6 mt-6">`;

    html += `
             <div class="glass-panel rounded-2xl p-5 border border-indigo-100 dark:border-indigo-900/30">
                 <div class="flex items-center gap-2 mb-4 border-b border-slate-100 dark:border-slate-700 pb-2">
                     <span class="w-2 h-2 rounded-full bg-indigo-500"></span>
                     <h4 class="font-bold text-slate-800 dark:text-white">ماشین حساب اختصاصی</h4>
                 </div>
                 
                 <div class="space-y-3">
                    <select id="calcBrand" onchange="app.updateCalcStep1()" class="input-modern rounded-lg w-full p-2 text-sm"><option value="">۱. انتخاب برند</option></select>
                    <div class="grid grid-cols-2 gap-2">
                        <select id="calcAxle" onchange="app.updateCalcStep2()" class="input-modern rounded-lg w-full p-2 text-sm" disabled><option value="">۲. محور</option></select>
                        <select id="calcUsage" onchange="app.updateCalcStep3()" class="input-modern rounded-lg w-full p-2 text-sm" disabled><option value="">۳. کاربری</option></select>
                    </div>
                    <select id="calcDesc" class="input-modern rounded-lg w-full p-2 text-sm" disabled><option value="">۴. توضیحات</option></select>
                    <select id="calcPaymentMethod" class="input-modern rounded-lg w-full p-2 text-sm text-indigo-700 font-bold bg-indigo-50">
                        <option value="cash">پرداخت نقدی</option>
                        <option value="12">اقساط ۱۲ ماهه</option>
                        <option value="24">اقساط ۲۴ ماهه</option>
                    </select>
                 </div>
    
                 <div class="mt-4 bg-slate-50 dark:bg-slate-900 rounded-xl p-4 text-center">
                    <div class="text-xs text-slate-500 mb-1">قیمت نهایی (تومان)</div>
                    <div id="displayTotalPrice" class="text-xl font-black text-slate-800 dark:text-white">0</div>
                    <div id="installmentDetails" class="hidden mt-2 pt-2 border-t border-slate-200 dark:border-slate-700 text-[10px] text-slate-500">
                        <div class="flex justify-between"><span>پیش:</span> <span id="displayPrepayment" class="font-bold">0</span></div>
                        <div class="flex justify-between"><span>ماهانه:</span> <span id="displayMonthly" class="font-bold">0</span></div>
                    </div>
                 </div>
    
                 <div class="flex gap-2 mt-4">
                    <button onclick="app.calculateFinalPrice()" class="flex-1 bg-blue-50 text-blue-600 hover:bg-blue-100 dark:bg-blue-900/30 dark:text-blue-400 py-2 rounded-lg text-sm font-bold transition">محاسبه / نمایش</button>
                 </div>
             </div>
          `;

    const fileTypes = [
        { label: "کارت ملی مشتری", col: COLUMN_INDEXES.FILE_NATIONAL_ID },
        { label: "استعلام سخا", col: COLUMN_INDEXES.FILE_SAHA },
        { label: "فیش واریزی", col: COLUMN_INDEXES.FILE_RECEIPT }
    ];

    html += `
             <div class="glass-panel rounded-2xl p-5 border border-emerald-100 dark:border-emerald-900/30 flex flex-col h-full">
                 <div class="flex items-center gap-2 mb-4 border-b border-slate-100 dark:border-slate-700 pb-2">
                     <span class="w-2 h-2 rounded-full bg-emerald-500"></span>
                     <h4 class="font-bold text-slate-800 dark:text-white">مدارک و ضمائم</h4>
                 </div>
                 
                 <div class="flex-1 space-y-3 overflow-y-auto max-h-[300px] pr-1 scrollbar-thin">
          `;

    fileTypes.forEach(file => {
        html += `
                  <div class="flex items-center justify-between p-3 bg-slate-50 dark:bg-slate-700/50 rounded-xl border border-slate-100 dark:border-slate-700">
                      <div class="flex flex-col">
                          <span class="text-xs font-bold text-slate-600 dark:text-slate-300">${file.label}</span>
                          <span id="upload_status_${file.col}" class="text-[10px] h-3"></span>
                      </div>
                      <div class="flex flex-col items-end gap-1 w-1/2">
                          <div id="file_list_${file.col}" class="w-full flex flex-col gap-1 items-end">
                              ${renderFileLinks(rowData[file.col], rowData[COLUMN_INDEXES.SHEET_ROW_INDEX], file.col)}
                          </div>
                          
                          <label class="cursor-pointer text-emerald-600 hover:bg-emerald-50 dark:hover:bg-emerald-900/30 p-2 rounded-lg transition mt-1" title="افزودن فایل جدید">
                              <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 4v16m8-8H4"></path></svg>
                              <input type="file" class="hidden" onchange="app.uploadFile(this, ${rowData[COLUMN_INDEXES.SHEET_ROW_INDEX]}, ${file.col})">
                          </label>
                      </div>
                  </div>
              `;
    });

    html += `
                 </div>
                 <div class="mt-4 pt-4 border-t border-slate-100 dark:border-slate-700">
                    <button onclick="app.saveCustomerDetails()" class="w-full bg-indigo-600 hover:bg-indigo-700 text-white py-3 rounded-xl text-sm font-bold shadow-lg shadow-indigo-200 dark:shadow-indigo-900/40 transition flex items-center justify-center gap-2">
                       <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7"></path></svg>
                       ذخیره تمام اطلاعات (فرم و ماشین حساب)
                    </button>
                 </div>
             </div>
          `;

    html += `</div>`;

    html += `<input type="hidden" id="detailRowIndex" value="${rowData[COLUMN_INDEXES.SHEET_ROW_INDEX]}">`;

    document.getElementById('detailTable').innerHTML = html;

    setTimeout(() => {
        app.populateCalculator(rowData);
        app.loadCallReports(Number(rowData[COLUMN_INDEXES.SHEET_ROW_INDEX]));
    }, 50);
}

// ===================================================================================
// === تابع محاسبه تفاوت پیش پرداخت ===
// ===================================================================================
function updatePrepaymentDifference() {
    const customerPaid = parseFloat(document.getElementById('prepayment_customer_paid').value) || 0;
    const targetPrepayment = parseFloat(document.getElementById('prepayment_target').value) || 0;
    const difference = customerPaid - targetPrepayment;
    const diffEl = document.getElementById('prepayment_difference');

    if (diffEl) {
        diffEl.innerText = difference.toLocaleString('fa-IR');
        diffEl.className = `input-modern rounded-lg w-full p-2 text-sm font-bold text-center ${difference >= 0 ? 'text-emerald-600 dark:text-emerald-400' : 'text-rose-600 dark:text-rose-400'}`;
        diffEl.style.background = 'rgba(255, 255, 255, 0.9)';
        diffEl.style.border = '1px solid #e2e8f0';
    }
}

// ===================================================================================
// === توابع مربوط به انتقال به نمایندگی ===
// ===================================================================================
function toggleTransferToAgencyBox() {
    const box = document.getElementById("transferToAgencyBox");
    const chk = document.getElementById("transferToAgencyCheckbox");

    if (chk.checked) {
        box.classList.remove("hidden");
        app.loadAgencyCities();
    } else {
        box.classList.add("hidden");
    }
}

// ===================================================================================
// === توابع مربوط به انتقال فرایند (با اصلاحات) ===
// ===================================================================================

// تابع کنترل UI برای نمایش/پنهان کردن باکس انتقال فرایند
function toggleTransferBox() {
    const box = document.getElementById("transferBox");
    const chk = document.getElementById("transferCheckbox");

    if (chk.checked) {
        box.classList.remove("hidden");

        // مقداردهی فیلد نام معرف (pt_refName) با نام مشتری فعلی
        const fullName = document.getElementById("detailName").innerText;
        document.getElementById("pt_refName").value = fullName;

        // فیلد pt_fullname خالی می‌ماند تا نام مشتری جدید وارد شود (بر اساس درخواست)
        document.getElementById("pt_fullname").value = "";

    } else {
        box.classList.add("hidden");
    }
}

// تابع ارسال اطلاعات به Apps Script
function submitProcessTransfer() {
    const fullName = document.getElementById("pt_fullname").value.trim();
    const nationalId = document.getElementById("pt_nationalId").value.trim();
    const phone = document.getElementById("pt_phone").value.trim();
    const refName = document.getElementById("pt_refName").value.trim();
    const rowIndex = Number(document.getElementById("transferRowIndex")?.value || 0);
    const statusDiv = document.getElementById('detailStatus');

    if (!fullName || !nationalId || !phone || !rowIndex) {
        if (statusDiv) { statusDiv.innerText = "لطفاً همه فیلدها را وارد کنید."; statusDiv.classList.remove('text-emerald-500'); statusDiv.classList.add('text-rose-500'); }
        return;
    }

    if (statusDiv) { statusDiv.innerText = "در حال ثبت انتقال فرایند..."; statusDiv.classList.remove('text-rose-500'); statusDiv.classList.add('text-amber-500'); }

    const data = {
        rowIndex,
        fullName,
        nationalId,
        phone,
        refName
    };

    google.script.run
        .withSuccessHandler(() => {
            if (statusDiv) { statusDiv.innerText = "انتقال فرایند ثبت شد."; statusDiv.classList.remove('text-rose-500'); statusDiv.classList.add('text-emerald-500'); }
            document.getElementById("pt_fullname").value = "";
            document.getElementById("pt_nationalId").value = "";
            document.getElementById("pt_phone").value = "";
            document.getElementById("transferCheckbox").checked = false;
            toggleTransferBox();
            app.loadProcessTransfers([rowIndex], () => {
                app.renderData(globalData);
                app.renderSegments();
                loadDetails(rowIndex);
            });
        })
        .withFailureHandler((err) => {
            if (statusDiv) { statusDiv.innerText = "خطا در ثبت انتقال فرایند"; statusDiv.classList.remove('text-emerald-500'); statusDiv.classList.add('text-rose-500'); }
            console.error('saveProcessTransfer error', err);
        })
        .saveProcessTransfer(data);
}

// ===================================================================================
// === MONITORING DASHBOARD MODULE ===
// ===================================================================================

const monitoringDashboard = {
    charts: {},
    agg: null,
    isLoading: false,

    init() {
        this.bindEvents();
        this.setDateLimits();
        this.showSkeleton(true);
        this.fetchData();
    },

    bindEvents() {
        const refreshBtn = document.getElementById('monitoringRefresh');
        if (refreshBtn) refreshBtn.onclick = () => this.fetchData(true);

        const applyBtn = document.getElementById('monitoringApplyRange');
        if (applyBtn) applyBtn.onclick = () => this.fetchData(false);

        const exitBtn = document.getElementById('monitoringExit');
        if (exitBtn) exitBtn.onclick = () => app.navigateTo('login', null, false);
    },

    setDateLimits() {
        const max = new Date().toISOString().slice(0, 10);
        ['monitoringDateFrom', 'monitoringDateTo'].forEach(id => {
            const el = document.getElementById(id);
            if (el) {
                el.setAttribute('max', max);
            }
        });
    },

    showSkeleton(show) {
        const sk = document.getElementById('monitorSkeleton');
        const ct = document.getElementById('monitorContent');
        if (sk) sk.classList.toggle('hidden', !show);
        if (ct) ct.classList.toggle('hidden', show);
    },

    setLoading(state) {
        this.isLoading = state;
        const btn = document.getElementById('monitoringRefresh');
        if (!btn) return;
        const icon = btn.querySelector('[data-refresh-icon]');
        const txt = btn.querySelector('[data-refresh-text]');
        btn.disabled = state;
        btn.classList.toggle('opacity-70', state);
        btn.classList.toggle('cursor-not-allowed', state);
        if (icon) icon.classList.toggle('animate-spin', state);
        if (txt) txt.innerText = state ? 'در حال بروزرسانی...' : 'به‌روزرسانی';
    },

    fetchData(isRefresh = false) {
        if (isRefresh) this.setLoading(true);
        this.showSkeleton(true);
        const startDate = (document.getElementById('monitoringDateFrom') || {}).value || '';
        const endDate = (document.getElementById('monitoringDateTo') || {}).value || '';
        google.script.run.withSuccessHandler((res) => {
            this.setLoading(false);
            if (!res || res.success !== true) {
                this.showSkeleton(false);
                return;
            }
            this.agg = res;
            this.showSkeleton(false);
            this.renderAll();
        }).withFailureHandler(() => {
            this.setLoading(false);
            this.showSkeleton(false);
        }).getMonitoringDashboardData(startDate, endDate);
    },

    renderAll() {
        if (!this.agg) return;
        const summary = this.agg.summary || {};
        this.renderKpis(summary, this.agg.financial || {});
        this.renderFunnel(summary);
        this.renderTrend(this.agg.trend || []);
        this.renderBrandCharts(this.agg.brands || []);
        this.renderUsageCharts(this.agg.usages || []);
        this.renderBrandUsageMatrix(this.agg.brandUsage || []);
        this.renderCancelByUsage(this.agg.cancelByUsage || []);
        this.renderExpertPerformance(this.agg.experts || []);
        this.renderFinancial(this.agg.financial || {});
        this.renderBottlenecks(this.agg.cancelByUsage || []);
        this.renderCancellationReasons(this.agg.cancellationReasons || []);
        this.renderInsights(this.agg.insights || []);
    },

    renderKpis(summary, financial) {
        const setText = (id, val) => { const el = document.getElementById(id); if (el) el.innerText = val; };
        setText('kpiTotalLeads', this.fmt(summary.total));
        setText('kpiActiveLeads', this.fmt(Math.max(0, (summary.total || 0) - (summary.cancelled || 0))));
        setText('kpiCancelledLeads', this.fmt(summary.cancelled));
        setText('kpiWaitingLeads', this.fmt(summary.waiting));
        setText('kpiContractLeads', this.fmt(summary.contract));
        setText('kpiDeliveredLeads', this.fmt(summary.delivered));
        const paid = Number(financial.paidSum || 0);
        const target = Number(financial.targetSum || 0);
        const achievement = target > 0 ? (paid / target) * 100 : 0;
        setText('kpiTargetPrepaymentSum', this.fmt(target));
        setText('kpiPaidPrepaymentSum', this.fmt(paid));
        setText('kpiPaymentAchievement', `${achievement.toFixed(1)}%`);
        const bar = document.getElementById('kpiPaymentProgress');
        if (bar) bar.style.width = `${Math.min(100, Math.max(0, achievement))}%`;
    },

    renderFunnel(summary) {
        const labels = ['ثبت‌شده', 'قرارداد', 'تحویل', 'لغو', 'منتظر'];
        const data = [summary.total || 0, summary.contract || 0, summary.delivered || 0, summary.cancelled || 0, summary.waiting || 0];
        this.renderChart('monitorFunnelChart', {
            type: 'bar',
            data: { labels, datasets: [{ data, backgroundColor: ['#0ea5e9', '#22c55e', '#8b5cf6', '#ef4444', '#f59e0b'], borderRadius: 10 }] },
            options: { indexAxis: 'y', plugins: { legend: { display: false } }, scales: { x: { beginAtZero: true } } }
        });
    },

    renderBrandCharts(brands) {
        const labels = brands.map(b => b.name);
        const leads = brands.map(b => b.leads);
        const sold = brands.map(b => b.sold);
        const conversions = brands.map(b => Number((b.conversion || 0).toFixed(1)));

        this.renderChart('monitorBrandChart', {
            type: 'bar',
            data: {
                labels,
                datasets: [
                    { label: 'درخواست', data: leads, backgroundColor: '#0ea5e9', borderRadius: 6 },
                    { label: 'فروش', data: sold, backgroundColor: '#22c55e', borderRadius: 6 },
                    { label: 'نرخ تبدیل %', data: conversions, type: 'line', yAxisID: 'y1', borderColor: '#f97316', backgroundColor: 'rgba(249,115,22,0.15)', tension: 0.3 }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { position: 'top' } },
                scales: { y: { beginAtZero: true }, y1: { beginAtZero: true, position: 'right', grid: { drawOnChartArea: false } } }
            }
        });

        const summaryEl = document.getElementById('monitorBrandSummary');
        if (summaryEl) {
            const parts = brands.slice(0, 4).map(b => `${b.name}: ${this.fmt(b.sold)} از ${this.fmt(b.leads)} (${(b.conversion || 0).toFixed(1)}%)`);
            summaryEl.innerText = parts.join(' | ');
        }
    },

    renderUsageCharts(usages) {
        const labels = usages.map(u => u.name);
        const leads = usages.map(u => u.leads);
        const sold = usages.map(u => u.sold);
        const conversions = usages.map(u => Number((u.conversion || 0).toFixed(1)));

        this.renderChart('monitorUsageChart', {
            type: 'bar',
            data: {
                labels,
                datasets: [
                    { label: 'درخواست', data: leads, backgroundColor: '#38bdf8', borderRadius: 6 },
                    { label: 'فروش', data: sold, backgroundColor: '#22c55e', borderRadius: 6 },
                    { label: 'نرخ تبدیل %', data: conversions, type: 'line', yAxisID: 'y1', borderColor: '#f97316', backgroundColor: 'rgba(249,115,22,0.15)', tension: 0.3 }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { position: 'top' } },
                scales: { y: { beginAtZero: true }, y1: { beginAtZero: true, position: 'right', grid: { drawOnChartArea: false } } }
            }
        });
    },

    renderBrandUsageMatrix(list) {
        const { stats } = this.computeBrandUsageStats(list);
        const brands = Object.keys(stats);
        const usageSet = new Set();
        brands.forEach(b => Object.keys(stats[b].usages).forEach(u => usageSet.add(u)));
        const usages = Array.from(usageSet);
        const palette = ['#0ea5e9', '#22c55e', '#a855f7', '#f59e0b', '#14b8a6', '#6366f1', '#eab308'];

        if (!brands.length || !usages.length) {
            const canvas = document.getElementById('monitorBrandUsageChart');
            if (canvas && canvas.parentElement) {
                canvas.parentElement.innerHTML = '<div class="text-sm text-slate-500 dark:text-slate-400">داده‌ای برای برند × کاربری یافت نشد.</div>';
            }
            return;
        }
        // برچسب = برند، دیتاست = کاربری (روی برند پشته می‌شود)
        const datasets = usages.map((u, idx) => {
            const data = brands.map(b => stats[b].usages[u]?.leads || 0);
            const conversions = brands.map(b => {
                const rec = stats[b].usages[u];
                return rec && rec.leads ? (rec.sold / rec.leads) * 100 : 0;
            });
            return {
                label: u,
                data,
                conversions,
                backgroundColor: palette[idx % palette.length],
                borderRadius: 6,
                stack: 'brandUsage'
            };
        });

        this.renderChart('monitorBrandUsageChart', {
            type: 'bar',
            data: { labels: brands, datasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { position: 'top' },
                    tooltip: {
                        callbacks: {
                            label: (ctx) => {
                                const conv = ctx.dataset.conversions?.[ctx.dataIndex] || 0;
                                return `${ctx.dataset.label}: ${this.fmt(ctx.parsed.y)} (تبدیل ${conv.toFixed(1)}%)`;
                            }
                        }
                    }
                },
                scales: { y: { stacked: true, beginAtZero: true }, x: { stacked: true } }
            }
        });
    },

    computeBrandUsageStats(list) {
        const stats = {};
        (list || []).forEach(item => {
            const brand = item.brand || 'نامشخص';
            const usage = item.usage || 'نامشخص';
            if (!stats[brand]) stats[brand] = { usages: {} };
            if (!stats[brand].usages[usage]) stats[brand].usages[usage] = { leads: 0, sold: 0 };
            stats[brand].usages[usage].leads += item.leads || 0;
            stats[brand].usages[usage].sold += item.sold || 0;
        });
        return { stats };
    },

    renderCancelByUsage(list) {
        const map = {};
        (list || []).forEach(item => {
            const usage = item.usage || 'نامشخص';
            if (!map[usage]) map[usage] = { total: 0, reasons: {} };
            map[usage].total += item.count || 0;
            map[usage].reasons[item.reason || 'نامشخص'] = item.count || 0;
        });

        const usages = Object.keys(map);
        const reasonTotals = {};
        usages.forEach(u => {
            Object.entries(map[u].reasons).forEach(([reason, cnt]) => {
                reasonTotals[reason] = (reasonTotals[reason] || 0) + cnt;
            });
        });
        const topReasons = Object.entries(reasonTotals).sort((a, b) => b[1] - a[1]).slice(0, 6).map(r => r[0]);
        const palette = ['#0ea5e9', '#22c55e', '#a855f7', '#f59e0b', '#14b8a6', '#6366f1', '#eab308'];
        const datasets = topReasons.map((reason, idx) => {
            const data = usages.map(u => map[u].reasons[reason] || 0);
            return { label: reason, data, backgroundColor: palette[idx % palette.length], borderRadius: 6, stack: 'cancel' };
        });

        this.renderChart('monitorCancelByUsageChart', {
            type: 'bar',
            data: { labels: usages, datasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    tooltip: {
                        callbacks: {
                            label: (ctx) => {
                                const usage = usages[ctx.dataIndex];
                                const total = map[usage].total || 1;
                                const pct = (ctx.parsed.y / total) * 100;
                                return `${ctx.dataset.label}: ${this.fmt(ctx.parsed.y)} (${pct.toFixed(1)}% انصراف‌های ${usage})`;
                            }
                        }
                    }
                },
                scales: { y: { stacked: true, beginAtZero: true }, x: { stacked: true } }
            }
        });

        const topText = [];
        usages.forEach(u => {
            const total = map[u].total || 0;
            const entries = Object.entries(map[u].reasons || {}).sort((a, b) => b[1] - a[1]);
            if (entries.length && total > 0) {
                const pct = (entries[0][1] / total) * 100;
                topText.push(`کاربری ${u}: ${entries[0][0]} (${pct.toFixed(1)}%)`);
            }
        });
        const infoEl = document.getElementById('monitorCancelUsageTop');
        if (infoEl) infoEl.innerText = topText.join(' | ') || 'دلیلی ثبت نشده است.';
    },

    renderTrend(trendList) {
        if (!trendList || !trendList.length) {
            const canvas = document.getElementById('monitorTrendChart');
            if (canvas && canvas.parentElement) {
                canvas.parentElement.innerHTML = '<div class="text-sm text-slate-500 dark:text-slate-400">داده‌ای برای روند ثبت نیست.</div>';
            }
            return;
        }
        const labels = trendList.map(t => t.date);
        const counts = trendList.map(t => t.count || 0);
        this.renderChart('monitorTrendChart', {
            type: 'line',
            data: { labels, datasets: [{ label: 'سرنخ جدید', data: counts, borderColor: '#0ea5e9', backgroundColor: 'rgba(14,165,233,0.2)', tension: 0.35, fill: true }] },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: { y: { beginAtZero: true } }
            }
        });
    },

    renderExpertPerformance(experts) {
        const sorted = (experts || []).slice().sort((a, b) => b.leads - a.leads);
        const labels = sorted.map(e => e.name || 'نامشخص');
        const leads = sorted.map(e => e.leads || 0);
        const contracts = sorted.map(e => e.contracts || 0);
        const conversion = sorted.map(e => Number((e.conversion || 0).toFixed(1)));

        this.renderChart('monitorExpertChart', {
            type: 'bar',
            data: {
                labels,
                datasets: [
                    { label: 'سرنخ', data: leads, backgroundColor: '#0ea5e9', borderRadius: 6 },
                    { label: 'قرارداد', data: contracts, backgroundColor: '#22c55e', borderRadius: 6 },
                    { label: 'نرخ تبدیل %', data: conversion, type: 'line', yAxisID: 'y1', borderColor: '#f97316', backgroundColor: 'rgba(249,115,22,0.15)', tension: 0.3 }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { position: 'top' } },
                scales: { y: { beginAtZero: true }, y1: { beginAtZero: true, position: 'right', grid: { drawOnChartArea: false } } }
            }
        });

        const tbody = document.getElementById('monitorExpertTableBody');
        if (tbody) {
            if (!sorted.length) {
                tbody.innerHTML = `<tr><td colspan="6" class="py-3 text-center text-slate-500 dark:text-slate-400">داده‌ای موجود نیست</td></tr>`;
                return;
            }
            tbody.innerHTML = sorted.map(e => `
                        <tr class="border-b border-slate-100 dark:border-slate-800">
                          <td class="py-2 px-2 text-left">${e.name || 'نامشخص'}</td>
                          <td class="py-2 px-2 text-right">${this.fmt(e.leads)}</td>
                          <td class="py-2 px-2 text-right">${this.fmt(e.contracts)}</td>
                          <td class="py-2 px-2 text-right">${this.fmt(e.deliveries)}</td>
                          <td class="py-2 px-2 text-right">${this.fmt(e.cancellations)}</td>
                          <td class="py-2 px-2 text-right">${(e.conversion || 0).toFixed(1)}%</td>
                        </tr>
                      `).join('');
        }
    },

    renderFinancial(financial) {
        const paid = Number(financial.paidSum || 0);
        const target = Number(financial.targetSum || 0);
        const remaining = Math.max(target - paid, 0);
        const hasData = paid > 0 || remaining > 0;
        if (!hasData) {
            const canvas = document.getElementById('monitorFinancialChart');
            if (canvas && canvas.parentElement) {
                canvas.parentElement.innerHTML = '<div class="text-sm text-slate-500 dark:text-slate-400">داده مالی موجود نیست.</div>';
            }
            return;
        }
        this.renderChart('monitorFinancialChart', {
            type: 'doughnut',
            data: {
                labels: ['وصول شده', 'باقی تا هدف'],
                datasets: [{
                    data: [paid, remaining],
                    backgroundColor: ['#22c55e', '#f59e0b'],
                    borderWidth: 0
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { position: 'bottom' } }
            }
        });
    },

    renderBottlenecks(list) {
        const reasonTotals = {};
        (list || []).forEach(item => {
            const r = item.reason || 'نامشخص';
            reasonTotals[r] = (reasonTotals[r] || 0) + (item.count || 0);
        });
        const entries = Object.entries(reasonTotals).sort((a, b) => b[1] - a[1]).slice(0, 6);
        if (!entries.length) {
            const canvas = document.getElementById('monitorBottleneckChart');
            if (canvas && canvas.parentElement) {
                canvas.parentElement.innerHTML = '<div class="text-sm text-slate-500 dark:text-slate-400">گلوگاه شناسایی نشد.</div>';
            }
            return;
        }
        this.renderChart('monitorBottleneckChart', {
            type: 'bar',
            data: {
                labels: entries.map(e => e[0]),
                datasets: [{ label: 'تعداد', data: entries.map(e => e[1]), backgroundColor: '#a855f7', borderRadius: 6 }]
            },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } }
        });
    },

    renderCancellationReasons(list) {
        const entries = (list || []).slice().sort((a, b) => b.count - a.count);
        if (!entries.length) {
            const canvas = document.getElementById('monitorCancellationChart');
            if (canvas && canvas.parentElement) {
                canvas.parentElement.innerHTML = '<div class="text-sm text-slate-500 dark:text-slate-400">داده‌ای برای لغو وجود ندارد.</div>';
            }
            return;
        }
        this.renderChart('monitorCancellationChart', {
            type: 'bar',
            data: {
                labels: entries.map(e => e.reason || 'نامشخص'),
                datasets: [{ label: 'تعداد', data: entries.map(e => e.count || 0), backgroundColor: '#f59e0b', borderRadius: 6 }]
            },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } }
        });
    },

    renderInsights(list) {
        const container = document.getElementById('monitorInsightsList');
        if (!container) return;
        const items = (list || []).filter(Boolean);
        if (!items.length) {
            container.innerHTML = '<div class="text-sm text-slate-500 dark:text-slate-400">داده‌ای برای Insight نیست.</div>';
            return;
        }
        container.innerHTML = items.map(text => `
                    <div class="p-2 rounded-lg bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 text-slate-700 dark:text-slate-200">
                      ${text}
                    </div>
                  `).join('');
    },

    renderChart(id, config) {
        const el = document.getElementById(id);
        if (!el) return;
        if (this.charts[id]) {
            this.charts[id].destroy();
        }
        this.charts[id] = new Chart(el, config);
    },

    fmt(val) {
        return Number(val || 0).toLocaleString('fa-IR');
    }
};


// Initialize monitoring dashboard when page loads
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => monitoringDashboard.init());
} else {
    monitoringDashboard.init();
}

// Re-fetch when navigating to monitoring
const originalNavigateTo = app.navigateTo;
app.navigateTo = function (page, data, refresh = true) {
    originalNavigateTo.call(this, page, data, refresh);
    if (page === 'monitoring') {
        setTimeout(() => {
            monitoringDashboard.fetchData();
        }, 100);
    }
};

