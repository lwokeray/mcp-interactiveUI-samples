/**
 * Microsoft Graph API unified query module.
 *
 * Single source of truth for all HR data. No local storage.
 * Uses client credentials flow (application permission) to read
 * the entire directory without a signed-in user.
 *
 * Credentials from environment:
 *   GRAPH_APP_ID       – the app (client) ID
 *   GRAPH_APP_SECRET   – a client secret
 *   TEAMS_APP_TENANT_ID – the tenant ID
 */
import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/lib/es/src/authentication/azureTokenCredentials/index.js";
// ─── Types ──────────────────────────────────────────────────────────
/** Fields we query from /users (v1.0, stable). */
const USER_SELECT_FIELDS = [
    "id",
    "displayName",
    "mail",
    "userPrincipalName",
    "jobTitle",
    "department",
    "officeLocation",
    "mobilePhone",
    "businessPhones",
    "employeeId",
    "employeeType",
    "employeeHireDate",
    "employeeLeaveDateTime",
    "companyName",
    "city",
    "country",
    "postalCode",
    "preferredLanguage",
];
// ─── Graph client singleton ─────────────────────────────────────────
function getConfig() {
    const clientId = process.env.GRAPH_APP_ID;
    const clientSecret = process.env.GRAPH_APP_SECRET;
    const tenantId = process.env.TEAMS_APP_TENANT_ID;
    if (!clientId)
        throw new Error("GRAPH_APP_ID is not set");
    if (!clientSecret)
        throw new Error("GRAPH_APP_SECRET is not set – create a client secret in Entra admin center for the Graph app");
    if (!tenantId)
        throw new Error("TEAMS_APP_TENANT_ID is not set");
    return { clientId, clientSecret, tenantId };
}
let _graphClient = null;
function getGraphClient() {
    if (_graphClient)
        return _graphClient;
    const { clientId, clientSecret, tenantId } = getConfig();
    const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
        scopes: ["https://graph.microsoft.com/.default"],
    });
    _graphClient = Client.initWithMiddleware({
        authProvider,
        defaultVersion: "v1.0",
    });
    return _graphClient;
}
/**
 * Reset the cached client (useful for testing or credential refresh).
 */
export function resetGraphClient() {
    _graphClient = null;
}
// ─── User query helpers ─────────────────────────────────────────────
function toHRUserSummary(u) {
    return {
        id: u.id,
        displayName: u.displayName ?? "",
        mail: u.mail ?? "",
        userPrincipalName: u.userPrincipalName ?? "",
        jobTitle: u.jobTitle ?? "",
        department: u.department ?? "",
        officeLocation: u.officeLocation ?? "",
        mobilePhone: u.mobilePhone ?? null,
        businessPhones: u.businessPhones ?? [],
        employeeId: u.employeeId ?? null,
        employeeType: u.employeeType ?? null,
        employeeHireDate: u.employeeHireDate ?? null,
        employeeLeaveDateTime: u.employeeLeaveDateTime ?? null,
        companyName: u.companyName ?? null,
        city: u.city ?? null,
        country: u.country ?? null,
        postalCode: u.postalCode ?? null,
        preferredLanguage: u.preferredLanguage ?? null,
    };
}
/**
 * Search M365 users by a free-text query.
 * Matches against displayName, mail, userPrincipalName, jobTitle, department.
 */
export async function searchUsers(query) {
    const client = getGraphClient();
    const response = await client
        .api("/users")
        .header("ConsistencyLevel", "eventual")
        .search(`"displayName:${query}" OR "mail:${query}" OR "userPrincipalName:${query}"`)
        .select(USER_SELECT_FIELDS)
        .top(25)
        .get();
    const users = response.value ?? [];
    return users.map(toHRUserSummary);
}
/**
 * List all M365 users (up to 200) with basic info.
 */
export async function listUsers() {
    const client = getGraphClient();
    const response = await client
        .api("/users")
        .header("ConsistencyLevel", "eventual")
        .select(USER_SELECT_FIELDS)
        .top(200)
        .orderby("displayName")
        .get();
    const users = response.value ?? [];
    return users.map(toHRUserSummary);
}
/**
 * Get a single user by ID.
 */
export async function getUser(userId) {
    try {
        const client = getGraphClient();
        const user = await client
            .api(`/users/${userId}`)
            .select(USER_SELECT_FIELDS)
            .get();
        return toHRUserSummary(user);
    }
    catch {
        return null;
    }
}
// ─── Profile helpers (beta API) ─────────────────────────────────────
/**
 * Get rich employee profile data from the beta Profile API.
 * Returns skills, certifications, education, languages, positions, interests, awards.
 */
export async function getUserProfile(userId) {
    const client = getGraphClient();
    const profile = {
        skills: [],
        certifications: [],
        educationalActivities: [],
        languages: [],
        positions: [],
        interests: [],
        awards: [],
    };
    try {
        // Skills
        const skillsResp = await client.api(`/users/${userId}/profile/skills`).get();
        profile.skills = (skillsResp.value ?? []).map((s) => s.displayName);
        // Certifications
        const certsResp = await client.api(`/users/${userId}/profile/certifications`).get();
        profile.certifications = (certsResp.value ?? []).map((c) => c.displayName);
        // Educational activities
        const eduResp = await client.api(`/users/${userId}/profile/educationalActivities`).get();
        profile.educationalActivities = (eduResp.value ?? []).map((e) => ({
            institution: e.institution?.displayName ?? "",
            degree: e.degree ?? "",
            fieldOfStudy: e.fieldOfStudy ?? "",
            startDate: e.startDate ?? "",
            endDate: e.endDate ?? "",
        }));
        // Languages
        const langResp = await client.api(`/users/${userId}/profile/languages`).get();
        profile.languages = (langResp.value ?? []).map((l) => ({
            displayName: l.displayName ?? "",
            proficiency: l.proficiency ?? "",
        }));
        // Positions
        const posResp = await client.api(`/users/${userId}/profile/positions`).get();
        profile.positions = (posResp.value ?? []).map((p) => ({
            title: p.detail?.jobTitle ?? "",
            company: p.detail?.company?.displayName ?? "",
            startDate: p.startDate ?? "",
            endDate: p.endDate ?? null,
        }));
        // Interests
        const intResp = await client.api(`/users/${userId}/profile/interests`).get();
        profile.interests = (intResp.value ?? []).map((i) => i.displayName);
        // Awards
        const awResp = await client.api(`/users/${userId}/profile/awards`).get();
        profile.awards = (awResp.value ?? []).map((a) => a.displayName);
    }
    catch {
        // Beta API may not be available in all tenants; return empty profile
    }
    return profile;
}
// ─── Org hierarchy helpers ──────────────────────────────────────────
/**
 * Get a user's manager.
 */
export async function getUserManager(userId) {
    try {
        const client = getGraphClient();
        const manager = await client
            .api(`/users/${userId}/manager`)
            .select(["id", "displayName", "mail", "jobTitle"])
            .get();
        return {
            id: manager.id,
            displayName: manager.displayName ?? "",
            mail: manager.mail ?? "",
            jobTitle: manager.jobTitle ?? "",
        };
    }
    catch {
        return null;
    }
}
/**
 * Get a user's direct reports.
 */
export async function getUserDirectReports(userId) {
    try {
        const client = getGraphClient();
        const resp = await client
            .api(`/users/${userId}/directReports`)
            .select(["id", "displayName", "mail", "jobTitle"])
            .get();
        return (resp.value ?? []).map((u) => ({
            id: u.id,
            displayName: u.displayName ?? "",
            mail: u.mail ?? "",
            jobTitle: u.jobTitle ?? "",
        }));
    }
    catch {
        return [];
    }
}
/**
 * Get groups a user is a member of.
 */
export async function getUserGroups(userId) {
    try {
        const client = getGraphClient();
        const resp = await client
            .api(`/users/${userId}/memberOf`)
            .select(["id", "displayName", "description"])
            .get();
        return (resp.value ?? []).map((g) => ({
            id: g.id ?? "",
            displayName: g.displayName ?? "",
            description: g.description ?? "",
        }));
    }
    catch {
        return [];
    }
}
// ─── Organization info ──────────────────────────────────────────────
/**
 * Get organization information.
 */
export async function getOrganization() {
    try {
        const client = getGraphClient();
        const resp = await client.api("/organization").get();
        const org = resp.value?.[0];
        if (!org)
            return null;
        return {
            id: org.id ?? "",
            displayName: org.displayName ?? "",
            technicalNotificationMails: org.technicalNotificationMails ?? [],
        };
    }
    catch {
        return null;
    }
}
// ─── Activity stats (beta) ──────────────────────────────────────────
/**
 * Get a user's activity statistics (email, chats, meetings, focus).
 */
export async function getUserActivityStats(userId) {
    try {
        const client = getGraphClient();
        const resp = await client
            .api(`/users/${userId}/analytics/activitystatistics`)
            .get();
        const stats = resp.value ?? [];
        const result = {
            email: { send: 0, receive: 0 },
            chats: { send: 0, receive: 0 },
            meetings: { organize: 0, attend: 0 },
            focus: { hours: 0 },
        };
        for (const s of stats) {
            if (s.activity === "Email") {
                result.email.send = s.summary?.send ?? 0;
                result.email.receive = s.summary?.receive ?? 0;
            }
            else if (s.activity === "Chat") {
                result.chats.send = s.summary?.send ?? 0;
                result.chats.receive = s.summary?.receive ?? 0;
            }
            else if (s.activity === "Meetings") {
                result.meetings.organize = s.summary?.organize ?? 0;
                result.meetings.attend = s.summary?.attend ?? 0;
            }
            else if (s.activity === "Focus") {
                result.focus.hours = s.summary?.hours ?? 0;
            }
        }
        return result;
    }
    catch {
        return null;
    }
}
//# sourceMappingURL=graph-api.js.map