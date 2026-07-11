// TalentStream AI - Application Logic and Agent Simulator

// Initial Database Structure
const DEFAULT_JOBS = [
    { id: "backend", title: "Senior Backend Engineer", department: "Engineering", reqSkills: ["Python", "Django", "PostgreSQL", "AWS"], count: 2 },
    { id: "frontend", title: "Frontend Developer", department: "Engineering", reqSkills: ["React", "TypeScript", "CSS3", "HTML5"], count: 1 },
    { id: "product", title: "Technical Product Manager", department: "Product", reqSkills: ["Agile", "Roadmapping", "SQL", "Jira"], count: 1 },
    { id: "devops", title: "DevOps Solutions Architect", department: "Operations", reqSkills: ["AWS", "Docker", "Kubernetes", "CI/CD"], count: 1 }
];

const DEFAULT_CANDIDATES = [
    {
        id: "c1",
        name: "Elena Rostova",
        email: "elena.rostova@techcorp.io",
        phone: "+1 (555) 723-9981",
        jobId: "backend",
        matchScore: 92,
        skills: ["Python", "Django", "PostgreSQL", "AWS", "Redis", "REST APIs"],
        experience: 6,
        summary: "Highly skilled software engineer with strong expertise in building scalable REST APIs and data processing systems in Python. Demonstrable cloud experience on AWS.",
        strengths: ["Strong backend coding skills", "Database performance optimization", "AWS architecture design"],
        weaknesses: ["Minimal experience with modern frontend frameworks"],
        outreachEmail: "Subject: Senior Backend Engineer opportunities at TechCorp\n\nDear Elena,\n\nI recently analyzed your background using our recruitment agent, and your 6 years of experience scaling python systems immediately stood out. We are currently looking for a Senior Backend Engineer who is proficient in Python and AWS to lead our microservices transition.\n\nI would love to set up a short 15-minute introductory call next Tuesday to tell you more about the role.\n\nBest regards,\nSarah Jenkins\nPrincipal Technical Recruiter",
        status: "Screened"
    },
    {
        id: "c2",
        name: "Marcus Vance",
        email: "marcus.vance@cloudstack.net",
        phone: "+1 (555) 129-8734",
        jobId: "devops",
        matchScore: 94,
        skills: ["AWS", "Docker", "Kubernetes", "CI/CD", "Terraform", "Linux"],
        experience: 8,
        summary: "Experienced DevOps engineer specializing in infrastructure-as-code, high-availability cluster management (Kubernetes), and continuous delivery pipelines.",
        strengths: ["Deep Kubernetes expertise", "Robust Terraform scripting", "Production-grade troubleshooting"],
        weaknesses: ["Higher salary expectation than average team bracket"],
        outreachEmail: "Subject: DevOps Architect Openings at TechCorp\n\nDear Marcus,\n\nYour profile was highlighted by our recruitment copilot due to your extensive experience with Kubernetes and Terraform. Our infrastructure team is currently automating our multi-region AWS deployments and we need an architect of your caliber.\n\nAre you available for a brief sync sometime this week?\n\nBest,\nSarah Jenkins\nPrincipal Technical Recruiter",
        status: "Interview Scheduled"
    },
    {
        id: "c3",
        name: "Sarah Chen",
        email: "schen.dev@gmail.com",
        phone: "+1 (555) 890-4321",
        jobId: "frontend",
        matchScore: 88,
        skills: ["React", "TypeScript", "CSS3", "HTML5", "Redux", "Webpack"],
        experience: 4,
        summary: "Product-minded frontend developer with a keen eye for UX. Proficient in building accessible, responsive, and performance-optimized React components.",
        strengths: ["Excellent CSS/UI styling capabilities", "TypeScript safety advocate", "Collaborative team player"],
        weaknesses: ["Limited unit test coverage on historical projects"],
        outreachEmail: "Subject: Frontend Engineering Role - TechCorp\n\nDear Sarah,\n\nOur screening agent highlighted your strong frontend layout achievements. We love your focus on design system accessibility. Our product team is building a new dashboard and your React/TypeScript skillset is a perfect match.\n\nLet me know if you have time for a virtual coffee next week.\n\nCheers,\nSarah Jenkins",
        status: "Screened"
    },
    {
        id: "c4",
        name: "Jonathan Smith",
        email: "jsmith.pm@outlook.com",
        phone: "+1 (555) 345-6789",
        jobId: "product",
        matchScore: 78,
        skills: ["Agile", "SQL", "Jira", "Product Strategy", "User Research"],
        experience: 5,
        summary: "Data-driven product manager who enjoys translating user research insights into functional technical product specs. Strong team collaboration skills.",
        strengths: ["Great user empathy", "Clear product backlog management"],
        weaknesses: ["Needs support on deeply technical backend architectures", "SQL queries are basic level only"],
        outreachEmail: "Subject: Product Manager Role - TechCorp\n\nDear Jonathan,\n\nI reviewed your portfolio using our TalentStream assistant. Your agile leadership and product strategy achievements are very interesting. We are looking for a product leader to drive client integration workflows.\n\nLet's coordinate a chat!\n\nBest,\nSarah Jenkins",
        status: "Screened"
    }
];

const SAMPLE_RESUMES = {
    backend: {
        name: "Liam Vance",
        contact: "liam.vance@codebase.org, (555) 678-4321",
        resumeText: "SUMMARY:\nInnovative software engineer with 5 years of professional experience building web services using Python, FastAPI, and Django. Passionate about PostgreSQL query tuning, containerized deployment, and serverless architectures.\n\nTECHNICAL SKILLS:\nLanguages: Python, JavaScript, SQL\nFrameworks: FastAPI, Django, Express.js\nDatabases: PostgreSQL, Redis, MongoDB\nCloud/Tools: AWS (S3, EC2, Lambda), Docker, Git\n\nEXPERIENCE:\nSoftware Engineer at LogiTech (2023 - Present)\n- Developed microservices in Python utilizing FastAPI, boosting processing throughput by 35%.\n- Refactored legacy SQL schemas and optimized PostgreSQL query execution speeds by 50%.\n- Integrated AWS cloud infrastructure for file upload servers."
    },
    product: {
        name: "Sophia Martinez",
        contact: "sophia.martinez@pmhub.com, (555) 987-1234",
        resumeText: "SUMMARY:\nProduct Manager with 7+ years of experience leading cross-functional teams to build SaaS applications. Champion of agile methodologies, database-driven decision making, and customer discovery interviews.\n\nTECHNICAL SKILLS:\nAgile Scrum, Product Strategy, Market Analysis, SQL, Jira, Confluence, Figma, Tableau\n\nEXPERIENCE:\nSenior Product Manager at CloudFlow (2022 - Present)\n- Guided the roadmap for core analytics platform, leading to 20% YoY growth in active users.\n- Formulated product specs based on extensive user analytics and SQL data extraction."
    }
};

// Application State
let candidates = [];
let jobs = [];
let auditLogs = [];
let currentRole = "recruiter"; // Default role
let currentTab = "dashboard";
let selectedCandidateId = null;

// Initialize App
document.addEventListener("DOMContentLoaded", () => {
    // Load database from localStorage or defaults
    candidates = JSON.parse(localStorage.getItem("ts_candidates")) || DEFAULT_CANDIDATES;
    jobs = JSON.parse(localStorage.getItem("ts_jobs")) || DEFAULT_JOBS;
    auditLogs = JSON.parse(localStorage.getItem("ts_audit_logs")) || [];
    currentRole = localStorage.getItem("ts_role") || "recruiter";

    // Set role selector visual state
    document.getElementById("roleSelect").value = currentRole;

    if (auditLogs.length === 0) {
        logEvent("SYSTEM", "system", "Agent Engine", "Database initialization complete. Initialized 4 default profiles.", "SUCCESS");
    }

    // Populate Job select inside modal
    const modalJobSelect = document.getElementById("modalJobSelect");
    modalJobSelect.innerHTML = jobs.map(j => `<option value="${j.id}">${j.title} (${j.department})</option>`).join('');

    // Trigger visual rendering
    updateDashboard();
    renderCandidateList();
    renderAuditLogs();
    renderJobTargets();
    
    // Switch to active role visual setup
    switchRole(currentRole, true);
});

// Switch Tab
window.switchTab = function(tabId) {
    currentTab = tabId;
    document.querySelectorAll(".tab-view").forEach(tab => tab.classList.remove("active"));
    document.querySelectorAll(".nav-item").forEach(btn => btn.classList.remove("active"));

    document.getElementById(`tab-${tabId}`).classList.add("active");
    document.getElementById(`tabBtn-${tabId}`).classList.add("active");

    // Add navigation audit entry
    logEvent(currentRole.toUpperCase(), "query", "UI Navigation", `Navigated to ${tabId} tab`, "SUCCESS");
};

// RBAC Switch Role
window.switchRole = function(role, isInit = false) {
    currentRole = role;
    localStorage.setItem("ts_role", role);
    
    // Log change
    if (!isInit) {
        logEvent(role.toUpperCase(), "rbac", "Security Control", `Role switched to ${role.toUpperCase()}`, "SUCCESS");
    }

    // Update security status bar
    const statusText = document.getElementById("terminalActiveRoleText");
    if (statusText) {
        statusText.innerText = role === "recruiter" ? "HR Recruiter" : role === "manager" ? "Hiring Manager" : "Compliance Auditor";
    }

    // Handle structural UI changes based on Role
    const outreachPanel = document.getElementById("candidateOutreachPanel");
    if (role === "recruiter") {
        document.body.classList.remove("role-restricted-manager", "role-restricted-auditor");
        document.body.classList.add("role-recruiter");
    } else if (role === "manager") {
        document.body.classList.remove("role-recruiter", "role-restricted-auditor");
        document.body.classList.add("role-restricted-manager");
    } else if (role === "auditor") {
        document.body.classList.remove("role-recruiter", "role-restricted-manager");
        document.body.classList.add("role-restricted-auditor");
    }

    // Refresh candidate list and detail panel to apply masking logic
    renderCandidateList();
    if (selectedCandidateId) {
        viewCandidateDetails(selectedCandidateId);
    }
    
    // Update live log on dashboard
    writeLiveLog("system", `Role context switched to [${role.toUpperCase()}]. Data visibility levels modified.`);
};

// Helper: Mask PII
function maskEmail(email) {
    if (!email) return "";
    const parts = email.split("@");
    if (parts.length !== 2) return email;
    const name = parts[0];
    const domain = parts[1];
    if (name.length <= 2) return name[0] + "***@" + domain;
    return name[0] + "***" + name[name.length - 1] + "@" + domain;
}

function maskPhone(phone) {
    if (!phone) return "";
    // Mask all but first and last few digits
    return phone.replace(/(\d{3})\D*(\d{3})\D*(\d{4})/, "+1 ($1) ***-$3");
}

// Log Event to Audit Trail
function logEvent(operator, actionType, targetComponent, details, status) {
    const event = {
        timestamp: new Date().toISOString().replace('T', ' ').substring(0, 19),
        operator: operator,
        actionType: actionType,
        targetComponent: targetComponent,
        details: details,
        status: status
    };
    auditLogs.unshift(event);
    localStorage.setItem("ts_audit_logs", JSON.stringify(auditLogs));
    renderAuditLogs();
}

// Write to active log block on dashboard
function writeLiveLog(type, text) {
    const logsBox = document.getElementById("dashboardAgentLogs");
    if (!logsBox) return;

    const line = document.createElement("div");
    line.className = `console-line ${type}`;
    
    let prefix = "[INFO]";
    if (type === "system") prefix = "[SYSTEM]";
    if (type === "agent-thinking") prefix = "[🤖 AI AGENT]";
    if (type === "agent-action") prefix = "[🔧 TOOL CALL]";
    if (type === "success") prefix = "[SUCCESS]";
    
    line.innerText = `${prefix} ${text}`;
    logsBox.appendChild(line);
    logsBox.scrollTop = logsBox.scrollHeight;
}

// Render Job Targets List on Sidebar
function renderJobTargets() {
    const list = document.getElementById("jobTargetsList");
    if (!list) return;

    // Recalculate candidates counts
    jobs.forEach(job => {
        job.count = candidates.filter(c => c.jobId === job.id).length;
    });

    list.innerHTML = jobs.map((job, index) => {
        const colors = ["blue", "green", "purple", "orange"];
        const color = colors[index % colors.length];
        return `
            <div class="job-target-item">
                <div class="job-title">
                    <span>${job.title}</span>
                    <span class="job-badge ${color}">${job.count} candidates</span>
                </div>
                <div class="job-meta">${job.department} • Req: ${job.reqSkills.slice(0, 3).join(', ')}</div>
            </div>
        `;
    }).join('');
}

// Render Dashboard Metrics & Charts
function updateDashboard() {
    const processed = candidates.length;
    const topFits = candidates.filter(c => c.matchScore >= 85).length;
    const interviews = candidates.filter(c => c.status === "Interview Scheduled").length;

    document.getElementById("metricProcessed").innerText = processed;
    document.getElementById("metricTopFits").innerText = topFits;
    document.getElementById("metricInterviews").innerText = interviews;
    document.getElementById("candidateCountBadge").innerText = processed;

    // Load Charts
    renderSVGChart();
    
    // Load Alert Center
    renderDashboardAlerts();
}

// Render Dashboard SVGs Charts
function renderSVGChart() {
    const container = document.getElementById("chartBarsGroup");
    if (!container) return;

    // We have 4 jobs. Let's calculate average score per job
    const stats = jobs.map(job => {
        const jobCands = candidates.filter(c => c.jobId === job.id);
        const avgScore = jobCands.length > 0 
            ? Math.round(jobCands.reduce((sum, c) => sum + c.matchScore, 0) / jobCands.length) 
            : 0;
        return {
            title: job.title.split(" ")[0] + " " + (job.title.split(" ")[1] || ""),
            avgScore: avgScore,
            count: jobCands.length
        };
    });

    // Draw SVG bars
    let html = '';
    const startX = 60;
    const width = 60;
    const gap = 40;
    const maxHeight = 150; // Max height inside SVG

    stats.forEach((stat, i) => {
        const x = startX + i * (width + gap);
        // Translate score 0-100 to height 0-130
        const barHeight = (stat.avgScore / 100) * 130;
        const y = 170 - barHeight;

        // Colors based on loop
        const colors = ["url(#gradBlue)", "url(#gradGreen)", "url(#gradPurple)", "url(#gradOrange)"];
        const color = colors[i % colors.length];

        html += `
            <g class="svg-bar">
                <title>${stat.title}: Avg Score ${stat.avgScore}%, Candidates: ${stat.count}</title>
                <!-- Background Bar -->
                <rect x="${x}" y="40" width="${width}" height="130" fill="rgba(255,255,255,0.02)" rx="4" />
                <!-- Active Bar -->
                <rect x="${x}" y="${y}" width="${width}" height="${barHeight}" fill="${color}" rx="4">
                    <animate attributeName="height" from="0" to="${barHeight}" dur="0.8s" fill="freeze" />
                    <animate attributeName="y" from="170" to="${y}" dur="0.8s" fill="freeze" />
                </rect>
                <!-- Text Value -->
                <text x="${x + width/2}" y="${y - 8}" fill="var(--text-primary)" font-size="10" font-weight="600" text-anchor="middle">${stat.avgScore}%</text>
                <!-- Label -->
                <text x="${x + width/2}" y="190" fill="var(--text-muted)" font-size="9" text-anchor="middle">${stat.title}</text>
            </g>
        `;
    });

    // Add Gradients to container if not already there
    let defs = `
        <defs>
            <linearGradient id="gradBlue" x1="0%" y1="0%" x2="100%" y2="100%">
                <stop offset="0%" stop-color="#0ea5e9" />
                <stop offset="100%" stop-color="#2563eb" />
            </linearGradient>
            <linearGradient id="gradGreen" x1="0%" y1="0%" x2="100%" y2="100%">
                <stop offset="0%" stop-color="#10b981" />
                <stop offset="100%" stop-color="#059669" />
            </linearGradient>
            <linearGradient id="gradPurple" x1="0%" y1="0%" x2="100%" y2="100%">
                <stop offset="0%" stop-color="#a855f7" />
                <stop offset="100%" stop-color="#6366f1" />
            </linearGradient>
            <linearGradient id="gradOrange" x1="0%" y1="0%" x2="100%" y2="100%">
                <stop offset="0%" stop-color="#f97316" />
                <stop offset="100%" stop-color="#ea580c" />
            </linearGradient>
        </defs>
    `;

    container.innerHTML = defs + html;

    // Render legend
    const legendContainer = document.getElementById("chartLegend");
    if (legendContainer) {
        const legendColors = ["#0ea5e9", "#10b981", "#a855f7", "#f97316"];
        legendContainer.innerHTML = stats.map((stat, i) => `
            <div class="legend-item">
                <span class="legend-color" style="background-color: ${legendColors[i]}"></span>
                <span>${stat.title} (Avg: ${stat.avgScore}%)</span>
            </div>
        `).join('');
    }
}

// Render Alerts List
function renderDashboardAlerts() {
    const container = document.getElementById("dashboardAlerts");
    if (!container) return;

    // Filter top candidates that have no interviews scheduled
    const topPending = candidates.filter(c => c.matchScore >= 88 && c.status === "Screened");
    
    let alerts = [];

    topPending.forEach(c => {
        alerts.push({
            type: "critical",
            title: "Outstanding Profile Pending Review",
            desc: `AI match rate for ${c.name} is ${c.matchScore}%. Review profile and initiate scheduler outreach immediately.`,
            time: "10 mins ago"
        });
    });

    alerts.push({
        type: "warning",
        title: "Compliance Checklist Warning",
        desc: "2 Hiring Managers are active. Candidate full PII details will automatically mask on their interface modules.",
        time: "1 hour ago"
    });

    alerts.push({
        type: "info",
        title: "Model Intelligence Synced",
        desc: "Autonomous screening parameters updated successfully based on job requirements mapping logic.",
        time: "3 hours ago"
    });

    document.getElementById("alertCountBadge").innerText = alerts.length;

    container.innerHTML = alerts.map(a => `
        <div class="alert-item ${a.type}">
            <div class="alert-icon">
                ${a.type === 'critical' ? `
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <polygon points="7.86 2 16.14 2 22 7.86 22 16.14 16.14 22 7.86 22 2 16.14 2 7.86 7.86 2"></polygon>
                        <line x1="12" y1="8" x2="12" y2="12"></line>
                        <line x1="12" y1="16" x2="12.01" y2="16"></line>
                    </svg>
                ` : a.type === 'warning' ? `
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path>
                        <line x1="12" y1="9" x2="12" y2="13"></line>
                        <line x1="12" y1="17" x2="12.01" y2="17"></line>
                    </svg>
                ` : `
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <circle cx="12" cy="12" r="10"></circle>
                        <line x1="12" y1="16" x2="12" y2="12"></line>
                        <line x1="12" y1="8" x2="12.01" y2="8"></line>
                    </svg>
                `}
            </div>
            <div class="alert-details">
                <h4>${a.title}</h4>
                <p>${a.desc}</p>
                <span class="alert-time">${a.time}</span>
            </div>
        </div>
    `).join('');
}

// Render Candidate Cards List (Left panel of Candidates View)
window.renderCandidateList = function() {
    const container = document.getElementById("candidateList");
    if (!container) return;

    const query = document.getElementById("candidateSearchInput").value.toLowerCase();
    const filterJob = document.getElementById("roleFilter").value;

    // Update filter options inside job select
    const roleFilter = document.getElementById("roleFilter");
    if (roleFilter.options.length <= 1) {
        roleFilter.innerHTML = '<option value="all">All Targets</option>' + 
            jobs.map(j => `<option value="${j.id}">${j.title}</option>`).join('');
    }

    // Filter list
    const filtered = candidates.filter(c => {
        const job = jobs.find(j => j.id === c.jobId);
        const jobTitle = job ? job.title.toLowerCase() : "";
        const skillsString = c.skills.join(" ").toLowerCase();
        const matchesQuery = c.name.toLowerCase().includes(query) || 
                             skillsString.includes(query) ||
                             jobTitle.includes(query);
        const matchesJob = filterJob === "all" || c.jobId === filterJob;

        return matchesQuery && matchesJob;
    });

    // Sort by Match Score descending
    filtered.sort((a, b) => b.matchScore - a.matchScore);

    if (filtered.length === 0) {
        container.innerHTML = `<div class="empty-detail-state">No candidates found for "${query}"</div>`;
        return;
    }

    container.innerHTML = filtered.map(c => {
        const isSelected = c.id === selectedCandidateId ? "selected" : "";
        const scoreClass = c.matchScore >= 90 ? "high" : c.matchScore >= 80 ? "medium" : "";
        const job = jobs.find(j => j.id === c.jobId);

        // RBAC check: mask names if role is Auditor, mask contact only for Manager
        const displayName = currentRole === "auditor" ? maskName(c.name) : c.name;

        return `
            <div class="candidate-card-item ${isSelected}" onclick="window.selectCandidate('${c.id}')">
                <div class="candidate-card-header">
                    <div>
                        <h4>${displayName}</h4>
                        <div class="candidate-card-role">${job ? job.title : "Unassigned"}</div>
                    </div>
                    <div class="score-badge-circle ${scoreClass}">${c.matchScore}%</div>
                </div>
                <div class="candidate-card-skills">
                    ${c.skills.slice(0, 3).map(s => `<span class="skill-tag">${s}</span>`).join('')}
                    ${c.skills.length > 3 ? `<span class="skill-tag">+${c.skills.length - 3}</span>` : ''}
                </div>
                <div class="candidate-card-footer">
                    <span>Exp: ${c.experience} Years</span>
                    <span>Status: ${c.status}</span>
                </div>
            </div>
        `;
    }).join('');
};

function maskName(name) {
    const parts = name.split(" ");
    return parts.map(p => p[0] + "***").join(" ");
}

// Select Candidate to View Details
window.selectCandidate = function(candId) {
    selectedCandidateId = candId;
    
    // Highlight list selection
    document.querySelectorAll(".candidate-card-item").forEach(item => item.classList.remove("selected"));
    
    renderCandidateList();
    viewCandidateDetails(candId);

    const cand = candidates.find(c => c.id === candId);
    logEvent(currentRole.toUpperCase(), "query", "Candidate database", `Viewed analysis scorecard for ${cand.name}`, "SUCCESS");
};

// View Candidate Scorecard and Detailed Analysis
function viewCandidateDetails(candId) {
    const panel = document.getElementById("candidateDetailsPanel");
    if (!panel) return;

    const cand = candidates.find(c => c.id === candId);
    if (!cand) {
        panel.innerHTML = `<div class="empty-detail-state">Candidate details not found.</div>`;
        return;
    }

    const job = jobs.find(j => j.id === cand.jobId);

    // Apply PII Masking depending on Active Role
    const showPII = currentRole === "recruiter";
    const displayedEmail = showPII ? cand.email : maskEmail(cand.email);
    const displayedPhone = showPII ? cand.phone : maskPhone(cand.phone);
    const displayedName = currentRole === "auditor" ? maskName(cand.name) : cand.name;

    const isManagerOrAuditor = currentRole === "manager" || currentRole === "auditor";
    const disabledClass = isManagerOrAuditor ? "disabled" : "";

    panel.innerHTML = `
        <!-- HEADER -->
        <div class="detail-header-card">
            <div class="detail-candidate-meta">
                <h3>${displayedName}</h3>
                <div class="subinfo">
                    <span>
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"></path>
                            <polyline points="22,6 12,13 2,6"></polyline>
                        </svg>
                        ${displayedEmail} ${!showPII ? `<span class="masked-pii">[MASKED]</span>` : ''}
                    </span>
                    <span>
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z"></path>
                        </svg>
                        ${displayedPhone} ${!showPII ? `<span class="masked-pii">[MASKED]</span>` : ''}
                    </span>
                </div>
            </div>
            <div class="detail-scoring-wheel">
                <div class="match-percentage">
                    <div class="percent">${cand.matchScore}%</div>
                    <div class="label">Match Rating</div>
                </div>
            </div>
        </div>

        <!-- DETAILS WRAPPER -->
        <div class="detail-body-container">
            <!-- TARGET SPECS -->
            <div>
                <h4 class="detail-section-title">Job Alignment Overview</h4>
                <div class="scorecard-metrics">
                    <div class="scorecard-metric-box">
                        <div class="val">${job ? job.title.split(' ')[0] : 'Unassigned'}</div>
                        <div class="lbl">Match Target</div>
                    </div>
                    <div class="scorecard-metric-box">
                        <div class="val">${cand.experience} Years</div>
                        <div class="lbl">Experience Duration</div>
                    </div>
                    <div class="scorecard-metric-box">
                        <div class="val">${cand.status}</div>
                        <div class="lbl">Workflow Stage</div>
                    </div>
                </div>
            </div>

            <!-- SKILLS ANALYZER -->
            <div>
                <h4 class="detail-section-title">Skill Analysis & Verification</h4>
                <div class="candidate-card-skills" style="gap: 8px;">
                    ${cand.skills.map(s => {
                        const isRequired = job && job.reqSkills.some(rs => rs.toLowerCase() === s.toLowerCase());
                        const borderStyle = isRequired ? "border: 1px solid var(--success-color); color: #34d399;" : "";
                        return `<span class="skill-tag" style="padding: 6px 12px; font-size: 0.75rem; ${borderStyle}">${s} ${isRequired ? '✓' : ''}</span>`;
                    }).join('')}
                </div>
            </div>

            <!-- COPILOT SYNTHESIS -->
            <div>
                <h4 class="detail-section-title">Copilot Screening Synthesis</h4>
                <p style="font-size: 0.85rem; line-height: 1.45; color: var(--text-secondary); margin-bottom: 12px;">
                    ${cand.summary}
                </p>
                
                <div style="display: flex; gap: 16px; flex-wrap: wrap;">
                    <div class="flex-1" style="background: rgba(16, 185, 129, 0.03); border: 1px solid rgba(16, 185, 129, 0.1); border-radius: 6px; padding: 12px;">
                        <h5 style="color: var(--success-color); font-size: 0.8rem; font-weight: 600; margin-bottom: 6px;">Identified Strengths</h5>
                        <ul class="analysis-points">
                            ${cand.strengths.map(st => `<li class="analysis-point" style="font-size: 0.75rem;">${st}</li>`).join('')}
                        </ul>
                    </div>
                    <div class="flex-1" style="background: rgba(239, 68, 68, 0.03); border: 1px solid rgba(239, 68, 68, 0.1); border-radius: 6px; padding: 12px;">
                        <h5 style="color: var(--danger-color); font-size: 0.8rem; font-weight: 600; margin-bottom: 6px;">Identified Development Gaps</h5>
                        <ul class="analysis-points">
                            ${(cand.weaknesses || ["None noted by Agent"]).map(wk => `<li class="analysis-point" style="font-size: 0.75rem;">${wk}</li>`).join('')}
                        </ul>
                    </div>
                </div>
            </div>

            <!-- AUTOMATED OUTREACH EMAIL -->
            <div id="candidateOutreachPanel" class="outreach-panel ${disabledClass}">
                <div class="outreach-header">
                    <h4>
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"></path>
                        </svg>
                        Autonomous Outreach Draft
                    </h4>
                    ${isManagerOrAuditor ? `
                        <span class="badge warning-badge" style="font-size: 0.65rem;">Recruiter Action Only</span>
                    ` : `
                        <span class="badge success-badge" style="font-size: 0.65rem;">Gemini Draft Generated</span>
                    `}
                </div>
                <div class="email-draft-box" id="emailDraftEditor" contenteditable="${!isManagerOrAuditor}">
                    ${cand.outreachEmail}
                </div>
                <div class="outreach-footer">
                    <button class="btn btn-secondary btn-iconic" onclick="window.regenerateEmail('${cand.id}')" ${isManagerOrAuditor ? 'disabled' : ''}>
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67"></path>
                        </svg>
                        <span>Rewrite Draft</span>
                    </button>
                    <button class="btn btn-primary" onclick="window.sendOutreach('${cand.id}')" ${isManagerOrAuditor ? 'disabled' : ''}>
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <line x1="22" y1="2" x2="11" y2="13"></line>
                            <polygon points="22 2 15 22 11 13 2 9 22 2"></polygon>
                        </svg>
                        <span>Send Candidate Email</span>
                    </button>
                </div>
            </div>
        </div>
    `;
}

// Regenerate outreach using mock AI parser
window.regenerateEmail = function(candId) {
    const cand = candidates.find(c => c.id === candId);
    if (!cand) return;

    writeLiveLog("agent-thinking", `Triggering outreach regeneration for Candidate: ${cand.name}.`);
    writeLiveLog("agent-action", `Model call: [gemini-3.5-flash] -> temperature: 0.85.`);

    // Mock delay
    const statusVal = document.getElementById("agentStatusText");
    const statusPulse = document.querySelector(".status-pulse-ring");
    statusVal.innerText = "Rewriting outreach...";
    statusPulse.className = "status-pulse-ring active thinking";

    setTimeout(() => {
        const variants = [
            `Subject: Engineering Career Path - TechCorp\n\nDear ${cand.name},\n\nI was reviewing your credentials and our scoring agent was impressed by your expertise in ${cand.skills.slice(0, 3).join(', ')}.\n\nWe have an immediate opening on our core product team. Your background aligns closely with what we're looking to achieve this quarter.\n\nCould we jump on a quick call this Thursday?\n\nWarmly,\nSarah Jenkins`,
            `Subject: Quick question from the TechCorp Engineering team\n\nHello ${cand.name},\n\nOur candidate screening copilot flagged your profile as an excellent technical match for our open pipeline. Your years of scaling applications fits our team composition perfectly.\n\nLet me know if you might be interested in exploring options.\n\nBest,\nSarah Jenkins`
        ];
        
        cand.outreachEmail = variants[Math.floor(Math.random() * variants.length)];
        localStorage.setItem("ts_candidates", JSON.stringify(candidates));
        
        viewCandidateDetails(candId);
        
        statusVal.innerText = "Ready";
        statusPulse.className = "status-pulse-ring active";
        writeLiveLog("success", `Personalized outreach template updated for ${cand.name}.`);
        logEvent(currentRole.toUpperCase(), "outreach", "Email generator", `Regenerated email template for ${cand.name}`, "SUCCESS");
    }, 1000);
};

// Send email (Mock Action)
window.sendOutreach = function(candId) {
    const cand = candidates.find(c => c.id === candId);
    if (!cand) return;

    const editorText = document.getElementById("emailDraftEditor").innerText;
    cand.outreachEmail = editorText;
    
    // Update state
    cand.status = "Interview Scheduled";
    localStorage.setItem("ts_candidates", JSON.stringify(candidates));

    writeLiveLog("system", `Email successfully sent to candidates address: ${cand.email}`);
    writeLiveLog("success", `Auto-shifted candidate ${cand.name} status to [Interview Scheduled].`);
    
    logEvent(currentRole.toUpperCase(), "outreach", "Outreach dispatcher", `Sent outreach email to ${cand.name}. Modified status to Interview Scheduled.`, "SUCCESS");

    // Re-render
    updateDashboard();
    renderCandidateList();
    viewCandidateDetails(candId);
};

// Filter candidates pool list
window.filterCandidates = function() {
    renderCandidateList();
};

// Render Audit Logs Table
function renderAuditLogs() {
    const tableBody = document.getElementById("auditLogsTableBody");
    const countText = document.getElementById("auditCountText");
    if (!tableBody) return;

    countText.innerText = auditLogs.length;

    const query = document.getElementById("auditSearchInput") ? document.getElementById("auditSearchInput").value.toLowerCase() : "";

    const filtered = auditLogs.filter(log => {
        return log.operator.toLowerCase().includes(query) ||
               log.actionType.toLowerCase().includes(query) ||
               log.targetComponent.toLowerCase().includes(query) ||
               log.details.toLowerCase().includes(query);
    });

    if (filtered.length === 0) {
        tableBody.innerHTML = `<tr><td colspan="6" style="text-align: center; color: var(--text-muted);">No compliance events match search query.</td></tr>`;
        return;
    }

    tableBody.innerHTML = filtered.map(log => `
        <tr>
            <td>${log.timestamp}</td>
            <td><span class="role-badge ${log.operator.toLowerCase() === 'ai agent' ? 'ai-agent' : log.operator.toLowerCase()}">${log.operator}</span></td>
            <td><span class="action-badge ${log.actionType.toLowerCase()}">${log.actionType.toUpperCase()}</span></td>
            <td>${log.targetComponent}</td>
            <td>${log.details}</td>
            <td><span class="badge success-badge">VERIFIED</span></td>
        </tr>
    `).join('');
}

// Filter Audit Trail
window.filterAuditLogs = function() {
    renderAuditLogs();
};

// Clear Audit Logs
window.clearAuditLogs = function() {
    auditLogs = [];
    localStorage.removeItem("ts_audit_logs");
    logEvent("SYSTEM", "system", "Agent Engine", "Compliance log cleared by user instruction.", "SUCCESS");
    renderAuditLogs();
    writeLiveLog("system", "Audit trails cleared.");
};

// Export Audit Logs (Mock file download)
window.exportAuditLogs = function() {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(auditLogs, null, 2));
    const downloadAnchor = document.createElement('a');
    downloadAnchor.setAttribute("href", dataStr);
    downloadAnchor.setAttribute("download", `talentstream_audit_log_${new Date().toISOString().slice(0,10)}.json`);
    document.body.appendChild(downloadAnchor);
    downloadAnchor.click();
    downloadAnchor.remove();

    logEvent(currentRole.toUpperCase(), "system", "Audit export", "Compliance audit logs exported as JSON file download", "SUCCESS");
    writeLiveLog("system", "Compliance audit logs exported.");
};

// Open Resume Upload Modal
window.openScreeningModal = function() {
    document.getElementById("screeningModal").classList.add("active");
    // Reset console & progress
    document.getElementById("modalAgentThinkingConsole").innerHTML = '<div class="console-line system">[IDLE] Waiting for screening execution to start...</div>';
    document.getElementById("modalProgressBar").style.width = '0%';
    
    // Clear inputs
    document.getElementById("modalCandidateName").value = '';
    document.getElementById("modalCandidateContact").value = '';
    document.getElementById("modalResumeText").value = '';
};

// Close Screening Modal
window.closeScreeningModal = function() {
    document.getElementById("screeningModal").classList.remove("active");
};

// Populate Template CV Data in modal
window.fillSampleResume = function(type) {
    const sample = SAMPLE_RESUMES[type];
    if (!sample) return;

    document.getElementById("modalCandidateName").value = sample.name;
    document.getElementById("modalCandidateContact").value = sample.contact;
    document.getElementById("modalResumeText").value = sample.resumeText;

    writeLiveLog("system", `Pre-loaded sample CV metadata for [${type.toUpperCase()}] candidate.`);
};

// Execute Autonomous Screening Simulator Pipeline
window.runScreeningPipeline = function() {
    const name = document.getElementById("modalCandidateName").value.trim();
    const contact = document.getElementById("modalCandidateContact").value.trim();
    const jobId = document.getElementById("modalJobSelect").value;
    const resumeText = document.getElementById("modalResumeText").value.trim();

    if (!name || !contact || !resumeText) {
        alert("Please fill out all required fields to initiate screening.");
        return;
    }

    // Disable button
    const btn = document.getElementById("btnRunScreening");
    btn.disabled = true;

    // Visual updates
    const modalConsole = document.getElementById("modalAgentThinkingConsole");
    const progress = document.getElementById("modalProgressBar");
    const statusVal = document.getElementById("agentStatusText");
    const statusPulse = document.querySelector(".status-pulse-ring");

    statusVal.innerText = "Screening Resume...";
    statusPulse.className = "status-pulse-ring active thinking";

    modalConsole.innerHTML = '';

    const steps = [
        { text: "[AGENT] Ingesting raw resume content structure...", pct: 10, delay: 600, type: "system" },
        { text: `[AGENT] Found contact info: extracting target node...`, pct: 25, delay: 1000, type: "info" },
        { text: `[AGENT] Target Profile identified as [${jobId.toUpperCase()}]. Scanning essential prerequisites...`, pct: 40, delay: 1200, type: "agent-thinking" },
        { text: `[AGENT] Extracting candidate technical skillset tags...`, pct: 55, delay: 900, type: "agent-action" },
        { text: "[AGENT] Executing scoring algorithm against job specifications...", pct: 75, delay: 1100, type: "agent-thinking" },
        { text: "[AGENT] Generating personalized outreach draft using Gemini LLM engine...", pct: 90, delay: 1400, type: "agent-action" },
        { text: "[AGENT] Writing record to secure candidates pool database and signing compliance trail...", pct: 98, delay: 600, type: "system" },
        { text: "[SUCCESS] Screening process complete. Match score computed.", pct: 100, delay: 500, type: "success" }
    ];

    let currentStep = 0;

    function runStep() {
        if (currentStep >= steps.length) {
            // Finish pipeline
            btn.disabled = false;
            statusVal.innerText = "Ready";
            statusPulse.className = "status-pulse-ring active";

            // Add new candidate to array
            const computedScore = calculateMockScore(resumeText, jobId);
            
            // Extract email/phone from contact string
            let email = "info@candidate.org";
            let phone = "+1 (555) 000-0000";
            const emailMatch = contact.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/);
            const phoneMatch = contact.match(/\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}/);
            if (emailMatch) email = emailMatch[0];
            if (phoneMatch) phone = phoneMatch[0];

            // Parse skills
            const defaultSkills = ["Python", "Docker", "Agile", "SQL", "TypeScript", "React"];
            const extractedSkills = defaultSkills.filter(s => resumeText.toLowerCase().includes(s.toLowerCase()));
            if (extractedSkills.length === 0) extractedSkills.push("Software Engineering", "Product Delivery");

            const newCandidate = {
                id: "c_" + Date.now(),
                name: name,
                email: email,
                phone: phone,
                jobId: jobId,
                matchScore: computedScore,
                skills: extractedSkills,
                experience: resumeText.match(/\d+\+?\s*(?:year|yr)/i) ? parseInt(resumeText.match(/(\d+)\+?\s*(?:year|yr)/i)[1]) : 5,
                summary: `AI screening agent synthesis for ${name}: Candidate demonstrates key competencies matching our vacancy requirements, specifically around ${extractedSkills.slice(0, 3).join(', ')}.`,
                strengths: [`Good match with vacancy requirements`, `Proficient core skillset: ${extractedSkills.slice(0, 2).join(', ')}`],
                weaknesses: [`Specific cloud certification validation pending`],
                outreachEmail: `Subject: Career opportunity at TechCorp\n\nDear ${name},\n\nOur candidate screening copilot flagged your profile as a high-potential fit. I reviewed your qualifications and would love to chat about the open position in our department.\n\nLet me know if next Tuesday works for a quick phone sync.\n\nBest,\nSarah Jenkins`,
                status: "Screened"
            };

            candidates.push(newCandidate);
            localStorage.setItem("ts_candidates", JSON.stringify(candidates));

            // Log event
            logEvent("AI AGENT", "screening", "Candidate database", `Screened resume for ${name}. Calculated Score: ${computedScore}%. Status: Screened`, "SUCCESS");
            writeLiveLog("success", `Autonomous screening complete. Ingested candidate profile for ${name}.`);

            // Update visuals
            updateDashboard();
            renderCandidateList();
            
            // Close and focus on candidates tab
            setTimeout(() => {
                closeScreeningModal();
                switchTab("candidates");
                selectCandidate(newCandidate.id);
            }, 600);

            return;
        }

        const step = steps[currentStep];
        progress.style.width = step.pct + '%';

        const line = document.createElement("div");
        line.className = `console-line ${step.type}`;
        line.innerText = step.text;
        modalConsole.appendChild(line);
        modalConsole.scrollTop = modalConsole.scrollHeight;

        writeLiveLog(step.type, step.text);

        currentStep++;
        setTimeout(runStep, step.delay);
    }

    runStep();
};

// Basic mock score calculation based on text matching
function calculateMockScore(text, jobId) {
    const textLower = text.toLowerCase();
    const job = jobs.find(j => j.id === jobId);
    if (!job) return 75;

    let hits = 0;
    job.reqSkills.forEach(skill => {
        if (textLower.includes(skill.toLowerCase())) {
            hits++;
        }
    });

    // Score ranges from 65 to 98
    const base = 65;
    const added = Math.min((hits / job.reqSkills.length) * 33, 33);
    return Math.round(base + added);
}


// ================= TAB 3: AGENT CONSOLE (NL COMMANDS) =================

// Trigger Preset commands
window.runPresetCommand = function(cmd) {
    const input = document.getElementById("consoleCommandLineInput");
    if (!input) return;
    input.value = cmd;
    submitConsoleCommand();
};

// Handle Keydown
window.handleConsoleInput = function(e) {
    if (e.key === "Enter") {
        submitConsoleCommand();
    }
};

// Submit Command
window.submitConsoleCommand = function() {
    const input = document.getElementById("consoleCommandLineInput");
    if (!input) return;

    const cmd = input.value.trim();
    if (!cmd) return;

    // Clear input
    input.value = '';

    // Append to Terminal body
    const termBody = document.getElementById("consoleTerminalBody");
    const userLine = document.createElement("div");
    userLine.className = "term-line";
    userLine.innerHTML = `<span class="prompt-arrow">&gt;</span> <span class="command-echo">${cmd}</span>`;
    termBody.appendChild(userLine);
    termBody.scrollTop = termBody.scrollHeight;

    // Log query audit trail
    logEvent(currentRole.toUpperCase(), "query", "Agent console CLI", `User inputted natural language instruction: "${cmd}"`, "SUCCESS");

    // Agent response pipeline
    const statusVal = document.getElementById("agentStatusText");
    const statusPulse = document.querySelector(".status-pulse-ring");
    statusVal.innerText = "Processing CLI command...";
    statusPulse.className = "status-pulse-ring active thinking";

    // Simulate thinking lines in terminal
    setTimeout(() => {
        appendTerminalLine("agent-thinking", "🤖 Parsing natural language syntax...");
        
        setTimeout(() => {
            // Execute the action based on parsing
            processInstruction(cmd);
            statusVal.innerText = "Ready";
            statusPulse.className = "status-pulse-ring active";
        }, 800);
    }, 400);
};

function appendTerminalLine(type, text) {
    const termBody = document.getElementById("consoleTerminalBody");
    const line = document.createElement("div");
    line.className = `term-line text-${type}`;
    line.innerText = text;
    termBody.appendChild(line);
    termBody.scrollTop = termBody.scrollHeight;
}

// Rules-based NLP processing
function processInstruction(cmd) {
    const lower = cmd.toLowerCase();

    // 1. Search / filter queries
    if (lower.includes("find") || lower.includes("filter") || lower.includes("search")) {
        // Check for score filter
        if (lower.includes("above") || lower.includes(">") || lower.includes("greater")) {
            const scoreMatch = lower.match(/(\d+)/);
            if (scoreMatch) {
                const targetScore = parseInt(scoreMatch[1]);
                appendTerminalLine("info", `🔧 Tool trigger: [db_query] -> filter candidates where score >= ${targetScore}.`);
                
                setTimeout(() => {
                    appendTerminalLine("success", `🤖 Query returned candidates matching criteria.`);
                    // Switch to Candidates view, filter in input
                    switchTab("candidates");
                    document.getElementById("candidateSearchInput").value = `score > ${targetScore}`;
                    // We custom override rendering logic for search
                    customFilterScore(targetScore);
                }, 400);
                return;
            }
        }

        // Check for skill matching
        let matchedSkill = null;
        const allSkills = ["python", "django", "react", "typescript", "kubernetes", "aws", "docker", "agile", "sql", "jira"];
        allSkills.forEach(sk => {
            if (lower.includes(sk)) {
                matchedSkill = sk;
            }
        });

        if (matchedSkill) {
            appendTerminalLine("info", `🔧 Tool trigger: [db_query] -> search tags field matching "${matchedSkill}".`);
            setTimeout(() => {
                appendTerminalLine("success", `🤖 Match index optimized. Filtered candidates pool.`);
                switchTab("candidates");
                document.getElementById("candidateSearchInput").value = matchedSkill;
                renderCandidateList();
            }, 400);
            return;
        }

        // General search
        appendTerminalLine("info", `🔧 Tool trigger: [db_query] -> string search for candidate fields.`);
        setTimeout(() => {
            appendTerminalLine("success", "🤖 Filtering completed.");
            switchTab("candidates");
            // Extract last words
            const words = cmd.split(" ");
            const lastWord = words[words.length - 1];
            document.getElementById("candidateSearchInput").value = lastWord;
            renderCandidateList();
        }, 400);
        return;
    }

    // 2. Data Scrubbing / Change Role
    if (lower.includes("mask") || lower.includes("role") || lower.includes("scrub") || lower.includes("hiring manager")) {
        appendTerminalLine("info", "🔧 Tool trigger: [rbac_controller] -> transition security level.");
        
        setTimeout(() => {
            let target = "manager";
            if (lower.includes("recruiter")) target = "recruiter";
            if (lower.includes("auditor")) target = "auditor";

            document.getElementById("roleSelect").value = target;
            switchRole(target);
            appendTerminalLine("success", `🤖 System state updated: Role changed to ${target.toUpperCase()} and access control rules compiled.`);
        }, 500);
        return;
    }

    // 3. Draft Email
    if (lower.includes("draft") || lower.includes("email") || lower.includes("outreach")) {
        // Search candidate name in command
        let foundCand = null;
        candidates.forEach(c => {
            if (lower.includes(c.name.toLowerCase().split(" ")[0])) {
                foundCand = c;
            }
        });

        if (foundCand) {
            appendTerminalLine("info", `🔧 Tool trigger: [outreach_draft_generator] -> payload: candidateId=${foundCand.id}, model=gemini-3.5-flash.`);
            setTimeout(() => {
                appendTerminalLine("success", `🤖 Personalization parameters synthesized. Outreach email drafted successfully for ${foundCand.name}.`);
                switchTab("candidates");
                selectCandidate(foundCand.id);
            }, 600);
            return;
        } else {
            appendTerminalLine("warning", "🤖 Draft instruction recognized, but candidate identity is ambiguous. Specify candidate's first name.");
            return;
        }
    }

    // 4. Generate scorecard/report
    if (lower.includes("generate") || lower.includes("report") || lower.includes("scorecard")) {
        appendTerminalLine("info", "🔧 Tool trigger: [compliance_reporter] -> fetching scores distribution data.");
        setTimeout(() => {
            const report = candidates.map(c => `${c.name}: Match: ${c.matchScore}%, Job: ${c.jobId}, Stage: ${c.status}`).join("\n");
            
            appendTerminalLine("success", "🤖 Report Generated successfully below:\n-------------------------------------\n" + report + "\n-------------------------------------");
            
            logEvent("AI AGENT", "query", "Compliance Reporter", "Generated candidate pipeline status scorecard overview report via NLP console.", "SUCCESS");
        }, 700);
        return;
    }

    // Default response if not matching rules
    appendTerminalLine("warning", "🤖 Command processed. Intent ambiguous. Try: 'find python developers', 'draft email for Elena', or 'change role to Manager'.");
}

// Custom scoring filter hook for console CLI
function customFilterScore(minScore) {
    const container = document.getElementById("candidateList");
    if (!container) return;

    // Filter list
    const filtered = candidates.filter(c => c.matchScore >= minScore);
    
    // Sort by Match Score descending
    filtered.sort((a, b) => b.matchScore - a.matchScore);

    if (filtered.length === 0) {
        container.innerHTML = `<div class="empty-detail-state">No candidates found scoring above ${minScore}%</div>`;
        return;
    }

    container.innerHTML = filtered.map(c => {
        const isSelected = c.id === selectedCandidateId ? "selected" : "";
        const scoreClass = c.matchScore >= 90 ? "high" : c.matchScore >= 80 ? "medium" : "";
        const job = jobs.find(j => j.id === c.jobId);
        const displayName = currentRole === "auditor" ? maskName(c.name) : c.name;

        return `
            <div class="candidate-card-item ${isSelected}" onclick="window.selectCandidate('${c.id}')">
                <div class="candidate-card-header">
                    <div>
                        <h4>${displayName}</h4>
                        <div class="candidate-card-role">${job ? job.title : "Unassigned"}</div>
                    </div>
                    <div class="score-badge-circle ${scoreClass}">${c.matchScore}%</div>
                </div>
                <div class="candidate-card-skills">
                    ${c.skills.slice(0, 3).map(s => `<span class="skill-tag">${s}</span>`).join('')}
                    ${c.skills.length > 3 ? `<span class="skill-tag">+${c.skills.length - 3}</span>` : ''}
                </div>
                <div class="candidate-card-footer">
                    <span>Exp: ${c.experience} Years</span>
                    <span>Status: ${c.status}</span>
                </div>
            </div>
        `;
    }).join('');
}
