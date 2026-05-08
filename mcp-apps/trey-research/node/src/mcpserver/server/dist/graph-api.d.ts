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
    email: {
        send: number;
        receive: number;
    };
    chats: {
        send: number;
        receive: number;
    };
    meetings: {
        organize: number;
        attend: number;
    };
    focus: {
        hours: number;
    };
}
export interface HROrganization {
    id: string;
    displayName: string;
    technicalNotificationMails: string[];
}
/**
 * Reset the cached client (useful for testing or credential refresh).
 */
export declare function resetGraphClient(): void;
/**
 * Search M365 users by a free-text query.
 * Matches against displayName, mail, userPrincipalName, jobTitle, department.
 */
export declare function searchUsers(query: string): Promise<HRUserSummary[]>;
/**
 * List all M365 users (up to 200) with basic info.
 */
export declare function listUsers(): Promise<HRUserSummary[]>;
/**
 * Get a single user by ID.
 */
export declare function getUser(userId: string): Promise<HRUserSummary | null>;
/**
 * Get rich employee profile data from the beta Profile API.
 * Returns skills, certifications, education, languages, positions, interests, awards.
 */
export declare function getUserProfile(userId: string): Promise<HRProfile>;
/**
 * Get a user's manager.
 */
export declare function getUserManager(userId: string): Promise<HRManager | null>;
/**
 * Get a user's direct reports.
 */
export declare function getUserDirectReports(userId: string): Promise<HRDirectReport[]>;
/**
 * Get groups a user is a member of.
 */
export declare function getUserGroups(userId: string): Promise<HRGroup[]>;
/**
 * Get organization information.
 */
export declare function getOrganization(): Promise<HROrganization | null>;
/**
 * Get a user's activity statistics (email, chats, meetings, focus).
 */
export declare function getUserActivityStats(userId: string): Promise<HRActivityStats | null>;
//# sourceMappingURL=graph-api.d.ts.map