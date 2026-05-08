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
import {
  registerAppTool,
  registerAppResource,
  RESOURCE_MIME_TYPE,
} from "@modelcontextprotocol/ext-apps/server";
import { z } from "zod";
import * as graph from "./graph-api.js";
import { getPublicServerUrl } from "./index.js";

// ─── Widget HTML loader ────────────────────────────────────────────
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ASSETS_DIR = path.resolve(__dirname, "..", "..", "assets");

function readWidgetHtml(componentName: string): string {
  if (!fs.existsSync(ASSETS_DIR)) {
    throw new Error(
      `Widget assets not found at ${ASSETS_DIR}. Run "npm run build:widgets" first.`
    );
  }
  let html: string | undefined;
  const directPath = path.join(ASSETS_DIR, `${componentName}.html`);
  if (fs.existsSync(directPath)) {
    html = fs.readFileSync(directPath, "utf8");
  } else {
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

export function createHRServer(): McpServer {
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

  // ─── Tools ─────────────────────────────────────────────────────

  // show-hr-dashboard
  registerAppTool(server, "show-hr-dashboard", {
    title: "Show HR Dashboard",
    description:
      "Display the HR organization dashboard with KPIs including total employees, departments, and organization info. Accepts optional filters: department, name, jobTitle.",
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
    const activeFilters: Record<string, unknown> = {};
    const filterDescParts: string[] = [];

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
    const deptMap = new Map<string, number>();
    for (const u of filtered) {
      const dept = u.department || "Unassigned";
      deptMap.set(dept, (deptMap.get(dept) ?? 0) + 1);
    }
    const departments = Array.from(deptMap.entries())
      .map(([name, count]) => ({ name, count }))
      .sort((a, b) => b.count - a.count);

    // Compute employee type breakdown
    const typeMap = new Map<string, number>();
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

    return {
      content: [
        {
          type: "text" as const,
          text: `HR Dashboard: ${filtered.length} employees across ${departments.length} departments.${filterDesc}`,
        },
      ],
      structuredContent: dashboardData,
    };
  });

  // show-consultant-profile
  registerAppTool(server, "show-consultant-profile", {
    title: "Show Employee Profile",
    description:
      "Display a detailed profile for an employee (by ID, email, or name), including contact info, job details, skills, certifications, education, languages, work history, manager, direct reports, and groups.",
    inputSchema: {
      employeeId: z.string().describe("The employee ID, email, or display name (partial match, case-insensitive) to view."),
    },
    annotations: { readOnlyHint: true },
    _meta: { ui: { resourceUri: PROFILE_URI } },
  }, async ({ employeeId }) => {
    let user = await graph.getUser(employeeId);
    if (!user) {
      const results = await graph.searchUsers(employeeId);
      if (results.length === 0) {
        return {
          content: [{ type: "text" as const, text: `Employee "${employeeId}" not found.` }],
          isError: true,
        };
      }
      user = results[0];
    }

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

    return {
      content: [
        {
          type: "text" as const,
          text: `Profile for ${user.displayName}: ${user.jobTitle} in ${user.department}. ${skillCount} skills, ${certCount} certifications, ${reportCount} direct report(s).`,
        },
      ],
      structuredContent: profileData,
    };
  });

  // search-employees
  registerAppTool(server, "search-employees", {
    title: "Search Employees",
    description:
      "Search employees by name, email, job title, or department. Returns matching employees from the organization directory.",
    inputSchema: {
      query: z.string().describe("Search query to find employees (matches displayName, mail, userPrincipalName, jobTitle, department)."),
    },
    annotations: { readOnlyHint: true },
    _meta: { ui: { resourceUri: BULK_EDITOR_URI } },
  }, async ({ query }) => {
    const users = await graph.searchUsers(query);

    return {
      content: [
        {
          type: "text" as const,
          text: `Found ${users.length} employee(s) matching "${query}".`,
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
    description:
      "List all employees in the organization directory. Returns up to 200 employees with their profile information.",
    inputSchema: {
      _: z.string().optional().describe("No parameters needed."),
    },
    annotations: { readOnlyHint: true },
    _meta: { ui: { resourceUri: BULK_EDITOR_URI } },
  }, async () => {
    const users = await graph.listUsers();

    return {
      content: [
        {
          type: "text" as const,
          text: `Found ${users.length} employee(s) in the organization.`,
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
    description:
      "Display the organization hierarchy for an employee, showing their manager, direct reports, and team structure.",
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
          content: [{ type: "text" as const, text: `Employee "${employeeId}" not found.` }],
          isError: true,
        };
      }
      user = results[0];
    }

    const [manager, directReports] = await Promise.all([
      graph.getUserManager(user.id),
      graph.getUserDirectReports(user.id),
    ]);

    return {
      content: [
        {
          type: "text" as const,
          text: `Org chart for ${user.displayName}: ${manager ? `Manager: ${manager.displayName}` : "No manager"} | ${directReports.length} direct report(s).`,
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
