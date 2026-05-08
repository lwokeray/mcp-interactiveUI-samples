/**
 * Quick Graph API connectivity test.
 * Uses raw fetch + @azure/identity to avoid @microsoft/microsoft-graph-client ESM issues.
 * Run: npx tsx src/test-graph.ts
 */
import { ClientSecretCredential } from "@azure/identity";

async function getToken(): Promise<string> {
  const credential = new ClientSecretCredential(
    process.env.TEAMS_APP_TENANT_ID!,
    process.env.GRAPH_APP_ID!,
    process.env.GRAPH_APP_SECRET!
  );
  const token = await credential.getToken("https://graph.microsoft.com/.default");
  return token.token;
}

async function graphGet(url: string, token: string) {
  const resp = await fetch(`https://graph.microsoft.com/v1.0${url}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!resp.ok) {
    const body = await resp.text();
    throw new Error(`HTTP ${resp.status}: ${body}`);
  }
  return resp.json();
}

async function main() {
  if (!process.env.GRAPH_APP_ID || !process.env.GRAPH_APP_SECRET || !process.env.TEAMS_APP_TENANT_ID) {
    console.error("❌ Missing env vars. Ensure GRAPH_APP_ID, GRAPH_APP_SECRET, TEAMS_APP_TENANT_ID are set.");
    process.exit(1);
  }

  console.log(`App ID: ${process.env.GRAPH_APP_ID.substring(0, 8)}...`);
  console.log(`Tenant: ${process.env.TEAMS_APP_TENANT_ID}`);

  console.log("\n🔑 Getting token...");
  const token = await getToken();
  console.log(`✅ Token acquired (${token.substring(0, 20)}...)`);

  // Test 1: List users
  console.log("\n📋 Test 1: Listing up to 5 users...");
  const usersResp = await graphGet("/users?$select=id,displayName,mail,jobTitle,department&$top=5", token);
  console.log(`✅ Found ${usersResp.value?.length ?? 0} users`);
  for (const u of usersResp.value ?? []) {
    console.log(`   - ${u.displayName} | ${u.jobTitle ?? "(no title)"} | ${u.department ?? "(no dept)"}`);
  }

  // Test 2: Get organization
  console.log("\n🏢 Test 2: Getting organization info...");
  const orgResp = await graphGet("/organization", token);
  const org = orgResp.value?.[0];
  if (org) {
    console.log(`✅ Organization: ${org.displayName}`);
  } else {
    console.log("⚠️  No organization found");
  }

  // Test 3: Manager relationship
  if (usersResp.value?.length > 0) {
    const firstUser = usersResp.value[0];
    console.log(`\n🔗 Test 3: Getting manager for ${firstUser.displayName}...`);
    try {
      const mgr = await graphGet(`/users/${firstUser.id}/manager?$select=id,displayName,jobTitle`, token);
      console.log(`✅ Manager: ${mgr.displayName} (${mgr.jobTitle ?? "N/A"})`);
    } catch {
      console.log("⚠️  No manager or insufficient permissions");
    }
  }

  // Test 4: Profile API (beta) - skills
  if (usersResp.value?.length > 0) {
    const firstUser = usersResp.value[0];
    console.log(`\n🎓 Test 4: Getting profile skills for ${firstUser.displayName}...`);
    try {
      const skillsResp = await fetch(`https://graph.microsoft.com/beta/users/${firstUser.id}/profile/skills`, {
        headers: { Authorization: `Bearer ${token}` },
      });
      if (skillsResp.ok) {
        const data = await skillsResp.json();
        console.log(`✅ Skills: ${(data.value ?? []).map((s: any) => s.displayName).join(", ") || "(none)"}`);
      } else {
        console.log("⚠️  Profile API not available (beta may not be enabled)");
      }
    } catch {
      console.log("⚠️  Profile API error");
    }
  }

  console.log("\n🎉 All tests complete!");
}

main().catch((err) => {
  console.error("\n❌ Test failed:", err.message);
  process.exit(1);
});