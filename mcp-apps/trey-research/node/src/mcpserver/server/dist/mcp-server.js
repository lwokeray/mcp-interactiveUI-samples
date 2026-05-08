/**
 * HR MCP Server factory.
 *
 * All data comes directly from Microsoft Graph API — no local storage.
 * Uses the MCP Apps standard (@modelcontextprotocol/ext-apps) for
 * widget resources and tool registration.
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { registerAppTool, registerAppResource, RESOURCE_MIME_TYPE, } from "@modelcontextprotocol/ext-apps/server";
import { z } from "zod";
import * as graph from "./graph-api.js";
import { getPublicServerUrl } from "./index.js";
// ─── Widget HTML loader ────────────────────────────────────────────
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ASSETS_DIR = path.resolve(__dirname, "..", "..", "assets");
function readWidgetHtml(componentName) {
    if (!fs.existsSync(ASSETS_DIR)) {
        throw new Error(`Widget assets not found at ${ASSETS_DIR}. Run "npm run build:widgets" first.`);
    }
    let html;
    const directPath = path.join(ASSETS_DIR, `${componentName}.html`);
    if (fs.existsSync(directPath)) {
        html = fs.readFileSync(directPath, "utf8");
    }
    else {
        const candidates = fs
            .readdirSync(ASSETS_DIR)
            .filter((f) => f.startsWith(`${componentName}-`) && f.endsWith(".html"))
            .sort();
        const fallback = candidates[candidates.length - 1];
        if (fallback) {
            html = fs.readFileSync(path.join(ASSETS_DIR, fallback), "utf8");
        }
    }
    if (!html) {
        throw new Error(`Widget HTML for "${componentName}" not found in ${ASSETS_DIR}.`);
    }
    const serverUrl = getPublicServerUrl();
    const injection = `<script>window.__SERVER_BASE_URL__=${JSON.stringify(serverUrl)};</script>`;
    html = html.replace("<head>", `<head>${injection}`);
    return html;
}
// ─── Widget URI definitions ────────────────────────────────────────
const DASHBOARD_URI = "ui://trey-hr/hr-dashboard.html";
const PROFILE_URI = "ui://trey-hr/consultant-profile.html";
const BULK_EDITOR_URI = "ui://trey-hr/bulk-editor.html";
// ─── Server factory ────────────────────────────────────────────────
export function createHRServer() {
    const server = new McpServer({ name: "trey-hr-consultant", version: "1.0.0" });
    // ─── Widget Resources ──────────────────────────────────────────
    registerAppResource(server, "HR Dashboard", DASHBOARD_URI, {
        mimeType: RESOURCE_MIME_TYPE,
        description: "HR Dashboard widget markup",
    }, async () => {
        const html = readWidgetHtml("hr-dashboard");
        return { contents: [{ uri: DASHBOARD_URI, mimeType: RESOURCE_MIME_TYPE, text: html }] };
    });
    registerAppResource(server, "Consultant Profile", PROFILE_URI, {
        mimeType: RESOURCE_MIME_TYPE,
        description: "Consultant Profile widget markup",
    }, async () => {
        const html = readWidgetHtml("consultant-profile");
        return { contents: [{ uri: PROFILE_URI, mimeType: RESOURCE_MIME_TYPE, text: html }] };
    });
    registerAppResource(server, "Bulk Editor", BULK_EDITOR_URI, {
        mimeType: RESOURCE_MIME_TYPE,
        description: "Bulk Editor widget markup",
    }, async () => {
        const html = readWidgetHtml("bulk-editor");
        return { contents: [{ uri: BULK_EDITOR_URI, mimeType: RESOURCE_MIME_TYPE, text: html }] };
    });
    // ─── Widget Tools (render UI) ──────────────────────────────────
    // show-hr-dashboard
    registerAppTool(server, "show-hr-dashboard", {
        title: "Show HR Dashboard",
        description: "Display the HR organization dashboard with KPIs including total employees, departments, and organization info. Accepts optional filters: department, name, jobTitle.",
        inputSchema: {
            department: z.string().optional().describe("Optional department name to pre-filter the dashboard (partial match, case-insensitive)."),
            name: z.string().optional().describe("Optional employee name to pre-filter the dashboard (partial match, case-insensitive)."),
            jobTitle: z.string().optional().describe("Optional job title to pre-filter the dashboard (partial match, case-insensitive)."),
        },
        annotations: { readOnlyHint: true },
        _meta: { ui: { resourceUri: DASHBOARD_URI } },
    }, async ({ department, name, jobTitle }) => {
        const [users, org] = await Promise.all([
            graph.listUsers(),
            graph.getOrganization(),
        ]);
        let filtered = users;
        const activeFilters = {};
        const filterDescParts = [];
        if (department) {
            const q = department.toLowerCase();
            filtered = filtered.filter((u) => u.department.toLowerCase().includes(q));
            activeFilters.department = department;
            filterDescParts.push(`department: "${department}"`);
        }
        if (name) {
            const q = name.toLowerCase();
            filtered = filtered.filter((u) => u.displayName.toLowerCase().includes(q));
            activeFilters.name = name;
            filterDescParts.push(`name: "${name}"`);
        }
        if (jobTitle) {
            const q = jobTitle.toLowerCase();
            filtered = filtered.filter((u) => u.jobTitle.toLowerCase().includes(q));
            activeFilters.jobTitle = jobTitle;
            filterDescParts.push(`jobTitle: "${jobTitle}"`);
        }
        // Compute department breakdown
        const deptMap = new Map();
        for (const u of filtered) {
            const dept = u.department || "Unassigned";
            deptMap.set(dept, (deptMap.get(dept) ?? 0) + 1);
        }
        const departments = Array.from(deptMap.entries())
            .map(([name, count]) => ({ name, count }))
            .sort((a, b) => b.count - a.count);
        // Compute employee type breakdown
        const typeMap = new Map();
        for (const u of filtered) {
            const type = u.employeeType || "Unknown";
            typeMap.set(type, (typeMap.get(type) ?? 0) + 1);
        }
        const employeeTypes = Array.from(typeMap.entries())
            .map(([type, count]) => ({ type, count }))
            .sort((a, b) => b.count - a.count);
        const dashboardData = {
            users: filtered.map((u) => ({
                id: u.id,
                displayName: u.displayName,
                email: u.mail || u.userPrincipalName,
                jobTitle: u.jobTitle,
                department: u.department,
                officeLocation: u.officeLocation,
                employeeType: u.employeeType,
                employeeId: u.employeeId,
            })),
            summary: {
                totalEmployees: filtered.length,
                totalInOrganization: users.length,
                organizationName: org?.displayName ?? "Unknown",
                departments,
                employeeTypes,
            },
            ...(Object.keys(activeFilters).length > 0 ? { filters: activeFilters } : {}),
        };
        const filterDesc = filterDescParts.length > 0
            ? ` (filtered by ${filterDescParts.join(", ")})`
            : "";
        // ── Build rich markdown text ──
        const mdLines = [];
        mdLines.push(`# HR Dashboard — ${org?.displayName ?? "Organization"}`);
        mdLines.push("");
        mdLines.push(`**${filtered.length}** employees across **${departments.length}** departments.${filterDesc}`);
        mdLines.push("");
        // Summary KPIs
        mdLines.push("## 📊 Summary");
        mdLines.push("");
        mdLines.push(`| Metric | Value |`);
        mdLines.push(`|--------|-------|`);
        mdLines.push(`| Total Employees (filtered) | ${filtered.length} |`);
        mdLines.push(`| Total in Organization | ${users.length} |`);
        mdLines.push(`| Departments | ${departments.length} |`);
        mdLines.push(`| Employee Types | ${employeeTypes.length} |`);
        mdLines.push("");
        // Department breakdown
        if (departments.length > 0) {
            mdLines.push("## 🏢 Department Breakdown");
            mdLines.push("");
            mdLines.push(`| Department | Employees |`);
            mdLines.push(`|------------|-----------|`);
            for (const d of departments) {
                const bar = "█".repeat(Math.max(1, Math.round((d.count / departments[0].count) * 20)));
                mdLines.push(`| ${d.name} | ${d.count} ${bar} |`);
            }
            mdLines.push("");
        }
        // Employee type breakdown
        if (employeeTypes.length > 0) {
            mdLines.push("## 👥 Employee Type Breakdown");
            mdLines.push("");
            mdLines.push(`| Type | Count |`);
            mdLines.push(`|------|-------|`);
            for (const et of employeeTypes) {
                mdLines.push(`| ${et.type} | ${et.count} |`);
            }
            mdLines.push("");
        }
        // Recent employees (top 10)
        if (filtered.length > 0) {
            mdLines.push("## 👤 Employees");
            mdLines.push("");
            mdLines.push(`| Name | Job Title | Department | Email |`);
            mdLines.push(`|------|-----------|------------|-------|`);
            const displayUsers = filtered.slice(0, 20);
            for (const u of displayUsers) {
                mdLines.push(`| ${u.displayName} | ${u.jobTitle || "—"} | ${u.department || "—"} | ${u.mail || u.userPrincipalName} |`);
            }
            if (filtered.length > 20) {
                mdLines.push(`| *… and ${filtered.length - 20} more* | | | |`);
            }
            mdLines.push("");
        }
        return {
            content: [
                {
                    type: "text",
                    text: mdLines.join("\n"),
                },
            ],
            structuredContent: dashboardData,
        };
    });
    // show-consultant-profile
    registerAppTool(server, "show-consultant-profile", {
        title: "Show Employee Profile",
        description: "Display a detailed profile for an employee (by ID, email, or name), including contact info, job details, skills, certifications, education, languages, work history, manager, direct reports, and groups.",
        inputSchema: {
            employeeId: z.string().describe("The employee ID, email, or display name (partial match, case-insensitive) to view."),
        },
        annotations: { readOnlyHint: true },
        _meta: { ui: { resourceUri: PROFILE_URI } },
    }, async ({ employeeId }) => {
        // Try to resolve the identifier to a user
        let user = await graph.getUser(employeeId);
        if (!user) {
            // Try searching by name/email
            const results = await graph.searchUsers(employeeId);
            if (results.length === 0) {
                return {
                    content: [{ type: "text", text: `Employee "${employeeId}" not found.` }],
                    isError: true,
                };
            }
            user = results[0];
        }
        // Fetch rich profile data in parallel
        const [profile, manager, directReports, groups, activityStats] = await Promise.all([
            graph.getUserProfile(user.id),
            graph.getUserManager(user.id),
            graph.getUserDirectReports(user.id),
            graph.getUserGroups(user.id),
            graph.getUserActivityStats(user.id),
        ]);
        const profileData = {
            user,
            profile,
            manager,
            directReports,
            groups,
            activityStats,
        };
        const skillCount = profile.skills.length;
        const certCount = profile.certifications.length;
        const reportCount = directReports.length;
        // ── Build rich markdown text ──
        const mdLines = [];
        mdLines.push(`# 👤 ${user.displayName}`);
        mdLines.push("");
        // Contact & Job Info
        mdLines.push("## 📋 Profile Overview");
        mdLines.push("");
        mdLines.push(`| Field | Value |`);
        mdLines.push(`|-------|-------|`);
        mdLines.push(`| **Job Title** | ${user.jobTitle || "—"} |`);
        mdLines.push(`| **Department** | ${user.department || "—"} |`);
        mdLines.push(`| **Email** | ${user.mail || user.userPrincipalName} |`);
        mdLines.push(`| **Phone** | ${user.mobilePhone || user.businessPhones?.[0] || "—"} |`);
        mdLines.push(`| **Office Location** | ${user.officeLocation || "—"} |`);
        mdLines.push(`| **Employee Type** | ${user.employeeType || "—"} |`);
        mdLines.push(`| **Employee ID** | ${user.employeeId || "—"} |`);
        mdLines.push("");
        // Manager
        if (manager) {
            mdLines.push(`**Manager:** ${manager.displayName} (${manager.jobTitle || "—"})`);
        }
        else {
            mdLines.push("**Manager:** None (top-level)");
        }
        mdLines.push("");
        // Skills
        if (profile.skills.length > 0) {
            mdLines.push("## 🛠️ Skills");
            mdLines.push("");
            mdLines.push(profile.skills.map((s) => `- ${s}`).join("\n"));
            mdLines.push("");
        }
        // Certifications
        if (profile.certifications.length > 0) {
            mdLines.push("## 📜 Certifications");
            mdLines.push("");
            for (const cert of profile.certifications) {
                mdLines.push(`- ${cert}`);
            }
            mdLines.push("");
        }
        // Education
        if (profile.educationalActivities?.length > 0) {
            mdLines.push("## 🎓 Education");
            mdLines.push("");
            for (const edu of profile.educationalActivities) {
                mdLines.push(`- **${edu.degree || "Degree"}** in ${edu.fieldOfStudy || "—"} — ${edu.institution || "—"}`);
            }
            mdLines.push("");
        }
        // Languages
        if (profile.languages?.length > 0) {
            mdLines.push(`## 🌐 Languages (${profile.languages.length})`);
            mdLines.push("");
            for (const lang of profile.languages) {
                mdLines.push(`- ${lang.displayName}${lang.proficiency ? ` (${lang.proficiency})` : ""}`);
            }
            mdLines.push("");
        }
        // Work History (positions)
        if (profile.positions?.length > 0) {
            mdLines.push("## 💼 Work History");
            mdLines.push("");
            for (const pos of profile.positions) {
                mdLines.push(`- **${pos.title || "—"}** at ${pos.company || "—"}${pos.startDate ? ` (${pos.startDate}${pos.endDate ? ` — ${pos.endDate}` : " — Present"})` : ""}`);
            }
            mdLines.push("");
        }
        // Direct Reports
        if (directReports.length > 0) {
            mdLines.push(`## 👥 Direct Reports (${directReports.length})`);
            mdLines.push("");
            mdLines.push(`| Name | Job Title |`);
            mdLines.push(`|------|-----------|`);
            for (const dr of directReports) {
                mdLines.push(`| ${dr.displayName} | ${dr.jobTitle || "—"} |`);
            }
            mdLines.push("");
        }
        // Groups
        if (groups.length > 0) {
            mdLines.push(`## 👪 Groups (${groups.length})`);
            mdLines.push("");
            for (const g of groups) {
                mdLines.push(`- ${g.displayName || g.id}`);
            }
            mdLines.push("");
        }
        // Activity Stats
        if (activityStats) {
            mdLines.push("## 📊 Activity Stats (7 days)");
            mdLines.push("");
            mdLines.push(`| Metric | Value |`);
            mdLines.push(`|--------|-------|`);
            mdLines.push(`| Emails Sent | ${activityStats.email?.send ?? "—"} |`);
            mdLines.push(`| Emails Received | ${activityStats.email?.receive ?? "—"} |`);
            mdLines.push(`| Chats Sent | ${activityStats.chats?.send ?? "—"} |`);
            mdLines.push(`| Chats Received | ${activityStats.chats?.receive ?? "—"} |`);
            mdLines.push(`| Meetings Organized | ${activityStats.meetings?.organize ?? "—"} |`);
            mdLines.push(`| Meetings Attended | ${activityStats.meetings?.attend ?? "—"} |`);
            mdLines.push(`| Focus Hours | ${activityStats.focus?.hours ?? "—"} |`);
            mdLines.push("");
        }
        return {
            content: [
                {
                    type: "text",
                    text: mdLines.join("\n"),
                },
            ],
            structuredContent: profileData,
        };
    });
    // search-employees
    registerAppTool(server, "search-employees", {
        title: "Search Employees",
        description: "Search employees by name, email, job title, or department. Returns matching employees from the organization directory.",
        inputSchema: {
            query: z.string().describe("Search query to find employees (matches displayName, mail, userPrincipalName, jobTitle, department)."),
        },
        annotations: { readOnlyHint: true },
        _meta: { ui: { resourceUri: BULK_EDITOR_URI } },
    }, async ({ query }) => {
        const users = await graph.searchUsers(query);
        // ── Build rich markdown text ──
        const mdLines = [];
        mdLines.push(`# 🔍 Employee Search Results`);
        mdLines.push("");
        mdLines.push(`Found **${users.length}** employee(s) matching "${query}".`);
        mdLines.push("");
        if (users.length > 0) {
            mdLines.push(`| Name | Job Title | Department | Email |`);
            mdLines.push(`|------|-----------|------------|-------|`);
            for (const u of users) {
                mdLines.push(`| ${u.displayName} | ${u.jobTitle || "—"} | ${u.department || "—"} | ${u.mail || u.userPrincipalName} |`);
            }
            mdLines.push("");
        }
        return {
            content: [
                {
                    type: "text",
                    text: mdLines.join("\n"),
                },
            ],
            structuredContent: {
                users: users.map((u) => ({
                    id: u.id,
                    displayName: u.displayName,
                    email: u.mail || u.userPrincipalName,
                    jobTitle: u.jobTitle,
                    department: u.department,
                    officeLocation: u.officeLocation,
                    employeeType: u.employeeType,
                    employeeId: u.employeeId,
                })),
            },
        };
    });
    // list-employees
    registerAppTool(server, "list-employees", {
        title: "List All Employees",
        description: "List all employees in the organization directory. Returns up to 200 employees with their profile information.",
        inputSchema: {
            _: z.string().optional().describe("No parameters needed."),
        },
        annotations: { readOnlyHint: true },
        _meta: { ui: { resourceUri: BULK_EDITOR_URI } },
    }, async () => {
        const users = await graph.listUsers();
        // ── Build rich markdown text ──
        const mdLines = [];
        mdLines.push(`# 📋 All Employees`);
        mdLines.push("");
        mdLines.push(`Found **${users.length}** employee(s) in the organization.`);
        mdLines.push("");
        if (users.length > 0) {
            mdLines.push(`| Name | Job Title | Department | Email |`);
            mdLines.push(`|------|-----------|------------|-------|`);
            const displayUsers = users.slice(0, 50);
            for (const u of displayUsers) {
                mdLines.push(`| ${u.displayName} | ${u.jobTitle || "—"} | ${u.department || "—"} | ${u.mail || u.userPrincipalName} |`);
            }
            if (users.length > 50) {
                mdLines.push(`| *… and ${users.length - 50} more* | | | |`);
            }
            mdLines.push("");
        }
        return {
            content: [
                {
                    type: "text",
                    text: mdLines.join("\n"),
                },
            ],
            structuredContent: {
                users: users.map((u) => ({
                    id: u.id,
                    displayName: u.displayName,
                    email: u.mail || u.userPrincipalName,
                    jobTitle: u.jobTitle,
                    department: u.department,
                    officeLocation: u.officeLocation,
                    employeeType: u.employeeType,
                    employeeId: u.employeeId,
                })),
            },
        };
    });
    // show-org-chart
    registerAppTool(server, "show-org-chart", {
        title: "Show Organization Chart",
        description: "Display the organization hierarchy for an employee, showing their manager, direct reports, and team structure.",
        inputSchema: {
            employeeId: z.string().describe("The employee ID, email, or display name (partial match, case-insensitive) to view the org chart for."),
        },
        annotations: { readOnlyHint: true },
        _meta: { ui: { resourceUri: DASHBOARD_URI } },
    }, async ({ employeeId }) => {
        let user = await graph.getUser(employeeId);
        if (!user) {
            const results = await graph.searchUsers(employeeId);
            if (results.length === 0) {
                return {
                    content: [{ type: "text", text: `Employee "${employeeId}" not found.` }],
                    isError: true,
                };
            }
            user = results[0];
        }
        const [manager, directReports] = await Promise.all([
            graph.getUserManager(user.id),
            graph.getUserDirectReports(user.id),
        ]);
        // ── Build rich markdown text ──
        const mdLines = [];
        mdLines.push(`# 🏢 Organization Chart`);
        mdLines.push("");
        mdLines.push(`## 👤 ${user.displayName}`);
        mdLines.push(`**${user.jobTitle || "—"}** · ${user.department || "—"}`);
        mdLines.push("");
        // Manager
        mdLines.push("### ⬆️ Manager");
        mdLines.push("");
        if (manager) {
            mdLines.push(`- **${manager.displayName}** — ${manager.jobTitle || "—"}`);
        }
        else {
            mdLines.push(`- *(No manager — top-level)*`);
        }
        mdLines.push("");
        // Direct Reports
        if (directReports.length > 0) {
            mdLines.push(`### ⬇️ Direct Reports (${directReports.length})`);
            mdLines.push("");
            mdLines.push(`| Name | Job Title |`);
            mdLines.push(`|------|-----------|`);
            for (const dr of directReports) {
                mdLines.push(`| ${dr.displayName} | ${dr.jobTitle || "—"} |`);
            }
            mdLines.push("");
        }
        else {
            mdLines.push("### ⬇️ Direct Reports");
            mdLines.push("");
            mdLines.push("*(No direct reports)*");
            mdLines.push("");
        }
        // Hierarchy visualization
        mdLines.push("### 📊 Hierarchy");
        mdLines.push("");
        mdLines.push("```");
        if (manager) {
            mdLines.push(`${manager.displayName} (Manager)`);
            mdLines.push(`  └── ${user.displayName} (You)`);
        }
        else {
            mdLines.push(`${user.displayName} (You)`);
        }
        for (const dr of directReports) {
            mdLines.push(`       ├── ${dr.displayName}`);
        }
        if (directReports.length > 0) {
            mdLines.push(`       └── (${directReports.length} total)`);
        }
        mdLines.push("```");
        mdLines.push("");
        return {
            content: [
                {
                    type: "text",
                    text: mdLines.join("\n"),
                },
            ],
            structuredContent: {
                employee: user,
                manager,
                directReports,
            },
        };
    });
    return server;
}
//# sourceMappingURL=mcp-server.js.map