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
import type { User } from "@microsoft/microsoft-graph-types";

// ─── Types ──────────────────────────────────────────────────────────

/** Fields we query from /users (v1.0, stable). */
const USER_SELECT_FIELDS: string[] = [
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

export interface HRUserSummary {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
  jobTitle: string;
  department: string;
  officeLocation: string;
  mobilePhone: string | null;
  businessPhones: string[];
  employeeId: string | null;
  employeeType: string | null;
  employeeHireDate: string | null;
  employeeLeaveDateTime: string | null;
  companyName: string | null;
  city: string | null;
  country: string | null;
  postalCode: string | null;
  preferredLanguage: string | null;
}

export interface HRProfile {
  skills: string[];
  certifications: string[];
  educationalActivities: Array<{
    institution: string;
    degree: string;
    fieldOfStudy: string;
    startDate: string;
    endDate: string;
  }>;
  languages: Array<{
    displayName: string;
    proficiency: string;
  }>;
  positions: Array<{
    title: string;
    company: string;
    startDate: string;
    endDate: string | null;
  }>;
  interests: string[];
  awards: string[];
}

export interface HRManager {
  id: string;
  displayName: string;
  mail: string;
  jobTitle: string;
}

export interface HRDirectReport {
  id: string;
  displayName: string;
  mail: string;
  jobTitle: string;
}

export interface HRGroup {
  id: string;
  displayName: string;
  description: string;
}

export interface HRActivityStats {
  email: { send: number; receive: number };
  chats: { send: number; receive: number };
  meetings: { organize: number; attend: number };
  focus: { hours: number };
}

export interface HROrganization {
  id: string;
  displayName: string;
  technicalNotificationMails: string[];
}

// ─── Token & fetch helpers ──────────────────────────────────────────

let _credential: ClientSecretCredential | null = null;

function getCredential(): ClientSecretCredential {
  if (_credential) return _credential;

  const clientId = process.env.GRAPH_APP_ID;
  const clientSecret = process.env.GRAPH_APP_SECRET;
  const tenantId = process.env.TEAMS_APP_TENANT_ID;

  if (!clientId) throw new Error("GRAPH_APP_ID is not set");
  if (!clientSecret) throw new Error("GRAPH_APP_SECRET is not set");
  if (!tenantId) throw new Error("TEAMS_APP_TENANT_ID is not set");

  _credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  return _credential;
}

let _token: string | null = null;
let _tokenExpiry = 0;

async function getToken(): Promise<string> {
  if (_token && Date.now() < _tokenExpiry) return _token;
  const credential = getCredential();
  const result = await credential.getToken("https://graph.microsoft.com/.default");
  _token = result.token;
  _tokenExpiry = Date.now() + (result.expiresOnTimestamp - Date.now()) * 0.8;
  return _token!;
}

/**
 * Reset cached token (useful for testing or credential refresh).
 */
export function resetGraphClient(): void {
  _token = null;
  _tokenExpiry = 0;
  _credential = null;
}

async function graphGet<T = unknown>(url: string): Promise<T> {
  const token = await getToken();
  const resp = await fetch(`https://graph.microsoft.com/v1.0${url}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!resp.ok) {
    const body = await resp.text();
    throw new Error(`Graph API error ${resp.status}: ${body}`);
  }
  return resp.json() as Promise<T>;
}

async function graphGetBeta<T = unknown>(url: string): Promise<T> {
  const token = await getToken();
  const resp = await fetch(`https://graph.microsoft.com/beta${url}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!resp.ok) {
    const body = await resp.text();
    throw new Error(`Graph API beta error ${resp.status}: ${body}`);
  }
  return resp.json() as Promise<T>;
}

// ─── User query helpers ─────────────────────────────────────────────

function toHRUserSummary(u: User): HRUserSummary {
  return {
    id: u.id!,
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
export async function searchUsers(query: string): Promise<HRUserSummary[]> {
  const fields = USER_SELECT_FIELDS.join(",");
  const response = await graphGet<{ value: User[] }>(
    `/users?$select=${fields}&$top=25&$search="${encodeURIComponent(query)}"`,
  );
  return (response.value ?? []).map(toHRUserSummary);
}

/**
 * List all M365 users (up to 200) with basic info.
 */
export async function listUsers(): Promise<HRUserSummary[]> {
  const fields = USER_SELECT_FIELDS.join(",");
  const response = await graphGet<{ value: User[] }>(
    `/users?$select=${fields}&$top=200&$orderby=displayName`,
  );
  return (response.value ?? []).map(toHRUserSummary);
}

/**
 * Get a single user by ID.
 */
export async function getUser(userId: string): Promise<HRUserSummary | null> {
  try {
    const fields = USER_SELECT_FIELDS.join(",");
    const user = await graphGet<User>(`/users/${encodeURIComponent(userId)}?$select=${fields}`);
    return toHRUserSummary(user);
  } catch {
    return null;
  }
}

// ─── Profile helpers (beta API) ─────────────────────────────────────

/**
 * Get rich employee profile data from the beta Profile API.
 * Returns skills, certifications, education, languages, positions, interests, awards.
 */
export async function getUserProfile(userId: string): Promise<HRProfile> {
  const profile: HRProfile = {
    skills: [],
    certifications: [],
    educationalActivities: [],
    languages: [],
    positions: [],
    interests: [],
    awards: [],
  };

  try {
    const uid = encodeURIComponent(userId);

    // Skills
    const skillsResp = await graphGetBeta<{ value: any[] }>(`/users/${uid}/profile/skills`);
    profile.skills = (skillsResp.value ?? []).map((s: any) => s.displayName);

    // Certifications
    const certsResp = await graphGetBeta<{ value: any[] }>(`/users/${uid}/profile/certifications`);
    profile.certifications = (certsResp.value ?? []).map((c: any) => c.displayName);

    // Educational activities
    const eduResp = await graphGetBeta<{ value: any[] }>(`/users/${uid}/profile/educationalActivities`);
    profile.educationalActivities = (eduResp.value ?? []).map((e: any) => ({
      institution: e.institution?.displayName ?? "",
      degree: e.degree ?? "",
      fieldOfStudy: e.fieldOfStudy ?? "",
      startDate: e.startDate ?? "",
      endDate: e.endDate ?? "",
    }));

    // Languages
    const langResp = await graphGetBeta<{ value: any[] }>(`/users/${uid}/profile/languages`);
    profile.languages = (langResp.value ?? []).map((l: any) => ({
      displayName: l.displayName ?? "",
      proficiency: l.proficiency ?? "",
    }));

    // Positions
    const posResp = await graphGetBeta<{ value: any[] }>(`/users/${uid}/profile/positions`);
    profile.positions = (posResp.value ?? []).map((p: any) => ({
      title: p.detail?.jobTitle ?? "",
      company: p.detail?.company?.displayName ?? "",
      startDate: p.startDate ?? "",
      endDate: p.endDate ?? null,
    }));

    // Interests
    const intResp = await graphGetBeta<{ value: any[] }>(`/users/${uid}/profile/interests`);
    profile.interests = (intResp.value ?? []).map((i: any) => i.displayName);

    // Awards
    const awResp = await graphGetBeta<{ value: any[] }>(`/users/${uid}/profile/awards`);
    profile.awards = (awResp.value ?? []).map((a: any) => a.displayName);
  } catch {
    // Beta API may not be available in all tenants; return empty profile
  }

  return profile;
}

// ─── Org hierarchy helpers ──────────────────────────────────────────

/**
 * Get a user's manager.
 */
export async function getUserManager(userId: string): Promise<HRManager | null> {
  try {
    const uid = encodeURIComponent(userId);
    const manager = await graphGet<User>(`/users/${uid}/manager?$select=id,displayName,mail,jobTitle`);
    return {
      id: manager.id!,
      displayName: manager.displayName ?? "",
      mail: manager.mail ?? "",
      jobTitle: manager.jobTitle ?? "",
    };
  } catch {
    return null;
  }
}

/**
 * Get a user's direct reports.
 */
export async function getUserDirectReports(userId: string): Promise<HRDirectReport[]> {
  try {
    const uid = encodeURIComponent(userId);
    const resp = await graphGet<{ value: User[] }>(`/users/${uid}/directReports?$select=id,displayName,mail,jobTitle`);
    return (resp.value ?? []).map((u: User) => ({
      id: u.id!,
      displayName: u.displayName ?? "",
      mail: u.mail ?? "",
      jobTitle: u.jobTitle ?? "",
    }));
  } catch {
    return [];
  }
}

/**
 * Get groups a user is a member of.
 */
export async function getUserGroups(userId: string): Promise<HRGroup[]> {
  try {
    const uid = encodeURIComponent(userId);
    const resp = await graphGet<{ value: any[] }>(`/users/${uid}/memberOf?$select=id,displayName,description`);
    return (resp.value ?? []).map((g: any) => ({
      id: g.id ?? "",
      displayName: g.displayName ?? "",
      description: g.description ?? "",
    }));
  } catch {
    return [];
  }
}

// ─── Organization info ──────────────────────────────────────────────

/**
 * Get organization information.
 */
export async function getOrganization(): Promise<HROrganization | null> {
  try {
    const resp = await graphGet<{ value: any[] }>("/organization");
    const org = resp.value?.[0];
    if (!org) return null;
    return {
      id: org.id ?? "",
      displayName: org.displayName ?? "",
      technicalNotificationMails: org.technicalNotificationMails ?? [],
    };
  } catch {
    return null;
  }
}

// ─── Activity stats (beta) ──────────────────────────────────────────

/**
 * Get a user's activity statistics (email, chats, meetings, focus).
 */
export async function getUserActivityStats(userId: string): Promise<HRActivityStats | null> {
  try {
    const uid = encodeURIComponent(userId);
    const resp = await graphGetBeta<{ value: any[] }>(`/users/${uid}/analytics/activitystatistics`);
    const stats = resp.value ?? [];
    const result: HRActivityStats = {
      email: { send: 0, receive: 0 },
      chats: { send: 0, receive: 0 },
      meetings: { organize: 0, attend: 0 },
      focus: { hours: 0 },
    };
    for (const s of stats) {
      if (s.activity === "Email") {
        result.email.send = s.summary?.send ?? 0;
        result.email.receive = s.summary?.receive ?? 0;
      } else if (s.activity === "Chat") {
        result.chats.send = s.summary?.send ?? 0;
        result.chats.receive = s.summary?.receive ?? 0;
      } else if (s.activity === "Meetings") {
        result.meetings.organize = s.summary?.organize ?? 0;
        result.meetings.attend = s.summary?.attend ?? 0;
      } else if (s.activity === "Focus") {
        result.focus.hours = s.summary?.hours ?? 0;
      }
    }
    return result;
  } catch {
    return null;
  }
}