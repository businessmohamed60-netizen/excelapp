import { createClient } from "npm:@supabase/supabase-js@2";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Client-Info, Apikey",
};

const MASTER_KEY = "IATFHConsult@nt2026Premium!";

function hmacSha256Hex(key: string, message: string): string {
  // Simple hash for license key generation (not crypto-secure, but matches VBA pattern)
  let h1 = 0xdeadbeef;
  let h2 = 0x41c6ce57;
  const data = key + message;
  for (let i = 0; i < data.length; i++) {
    const ch = data.charCodeAt(i);
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507);
  h1 ^= Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507);
  h2 ^= Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  const hash = (h2 >>> 0).toString(16).padStart(8, "0") + (h1 >>> 0).toString(16).padStart(8, "0");
  // Extend to 20 chars by repeating with offset
  let extended = hash;
  let offset = 0;
  while (extended.length < 20) {
    offset++;
    let h = 0;
    for (let i = 0; i < data.length; i++) {
      h = ((h << 5) - h + data.charCodeAt(i) + offset) | 0;
    }
    extended += Math.abs(h).toString(16).padStart(8, "0");
  }
  return extended.substring(0, 20).toUpperCase();
}

function generateLicenseKey(machineId: string, expiresAt: string, edition: string): string {
  const payload = `${machineId}|${expiresAt.replace(/-/g, "")}|${edition}`;
  const core = hmacSha256Hex(MASTER_KEY, payload);
  let result = "IATF";
  for (let i = 0; i < 20; i += 5) {
    result += "-" + core.substring(i, i + 5);
  }
  return result;
}

const supabase = createClient(
  Deno.env.get("SUPABASE_URL")!,
  Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
  { auth: { persistSession: false } }
);

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { status: 200, headers: corsHeaders });
  }

  const url = new URL(req.url);
  const path = url.pathname.replace("/functions/v1/license-api", "");
  const segments = path.split("/").filter(Boolean);

  try {
    // GET /verify?key=XXXX&machineId=YYYY — Verify a license
    if (req.method === "GET" && (segments[0] === "verify" || path === "" || path === "/")) {
      const licenseKey = (url.searchParams.get("key") || "").toUpperCase().trim();
      const machineId = url.searchParams.get("machineId") || "";

      if (!licenseKey) {
        return json({ valid: false, error: "Missing license key" }, 400);
      }

      const { data: license, error } = await supabase
        .from("licenses")
        .select("*")
        .eq("license_key", licenseKey)
        .maybeSingle();

      if (error || !license) {
        return json({ valid: false, error: "License not found" }, 404);
      }

      if (license.status === "REVOKED") {
        return json({ valid: false, error: "License revoked" }, 403);
      }

      const now = new Date();
      const expires = new Date(license.expires_at);
      const isExpired = expires < now;

      if (isExpired) {
        await supabase.from("licenses").update({ status: "EXPIRED" }).eq("id", license.id);
        return json({ valid: false, error: "License expired", expires_at: license.expires_at }, 403);
      }

      if (license.activation_count >= license.max_activations && license.machine_id !== machineId) {
        return json({ valid: false, error: "Max activations reached" }, 403);
      }

      return json({
        valid: true,
        edition: license.edition,
        expires_at: license.expires_at,
        customer_name: license.customer_name,
        activation_count: license.activation_count,
        max_activations: license.max_activations,
      });
    }

    // POST /activate — Activate a license on a machine
    if (req.method === "POST" && segments[0] === "activate") {
      const body = await req.json();
      const { license_key, machine_id, machine_info } = body;

      if (!license_key || !machine_id) {
        return json({ error: "Missing license_key or machine_id" }, 400);
      }

      const key = license_key.toUpperCase().trim();
      const { data: license, error } = await supabase
        .from("licenses")
        .select("*")
        .eq("license_key", key)
        .maybeSingle();

      if (error || !license) {
        return json({ error: "License not found" }, 404);
      }

      if (license.status === "REVOKED") {
        return json({ error: "License revoked" }, 403);
      }

      const now = new Date();
      const expires = new Date(license.expires_at);
      if (expires < now) {
        await supabase.from("licenses").update({ status: "EXPIRED" }).eq("id", license.id);
        return json({ error: "License expired" }, 403);
      }

      // Check if already activated on this machine
      const { data: existing } = await supabase
        .from("license_activations")
        .select("id")
        .eq("license_id", license.id)
        .eq("machine_id", machine_id)
        .maybeSingle();

      if (!existing) {
        if (license.activation_count >= license.max_activations) {
          return json({ error: "Max activations reached" }, 403);
        }
        await supabase.from("license_activations").insert({
          license_id: license.id,
          machine_id,
          machine_info: machine_info || null,
        });
        await supabase
          .from("licenses")
          .update({
            activation_count: license.activation_count + 1,
            status: "ACTIVE",
            machine_id,
          })
          .eq("id", license.id);
      }

      return json({
        success: true,
        edition: license.edition,
        expires_at: license.expires_at,
        customer_name: license.customer_name,
      });
    }

    // POST /generate — Generate a new license (admin only)
    if (req.method === "POST" && segments[0] === "generate") {
      const body = await req.json();
      const { customer_name, customer_email, edition, expires_at, max_activations, machine_id } = body;

      if (!edition || !expires_at) {
        return json({ error: "Missing required fields: edition, expires_at" }, 400);
      }

      if (!["STANDARD", "PREMIUM"].includes(edition.toUpperCase())) {
        return json({ error: "Edition must be STANDARD or PREMIUM" }, 400);
      }

      const mid = machine_id || "ANY";
      const dateStr = expires_at.replace(/-/g, "");
      const licenseKey = generateLicenseKey(mid, dateStr, edition.toUpperCase());

      const { data, error } = await supabase
        .from("licenses")
        .insert({
          license_key: licenseKey,
          edition: edition.toUpperCase(),
          customer_name: customer_name || null,
          customer_email: customer_email || null,
          expires_at,
          max_activations: max_activations || 1,
          status: "PENDING",
          machine_id: machine_id || null,
        })
        .select()
        .single();

      if (error) {
        return json({ error: error.message }, 500);
      }

      return json({ success: true, license: data });
    }

    // GET /list — List all licenses (admin)
    if (req.method === "GET" && segments[0] === "list") {
      const page = parseInt(url.searchParams.get("page") || "1");
      const perPage = parseInt(url.searchParams.get("per_page") || "50");
      const search = url.searchParams.get("search") || "";
      const statusFilter = url.searchParams.get("status") || "";

      let query = supabase.from("licenses").select("*", { count: "exact" });
      if (search) {
        query = query.or(`customer_name.ilike.%${search}%,customer_email.ilike.%${search}%,license_key.ilike.%${search}%`);
      }
      if (statusFilter) {
        query = query.eq("status", statusFilter.toUpperCase());
      }
      query = query.order("created_at", { ascending: false }).range((page - 1) * perPage, page * perPage - 1);

      const { data, error, count } = await query;

      if (error) {
        return json({ error: error.message }, 500);
      }

      return json({ licenses: data, total: count, page, per_page: perPage });
    }

    // GET /activations/:licenseId — List activations for a license
    if (req.method === "GET" && segments[0] === "activations") {
      const licenseId = segments[1];
      if (!licenseId) {
        return json({ error: "Missing license id" }, 400);
      }
      const { data, error } = await supabase
        .from("license_activations")
        .select("*")
        .eq("license_id", licenseId)
        .order("activated_at", { ascending: false });

      if (error) {
        return json({ error: error.message }, 500);
      }

      return json({ activations: data });
    }

    // PUT /revoke/:licenseId — Revoke a license
    if (req.method === "PUT" && segments[0] === "revoke") {
      const licenseId = segments[1];
      if (!licenseId) {
        return json({ error: "Missing license id" }, 400);
      }
      const { data, error } = await supabase
        .from("licenses")
        .update({ status: "REVOKED" })
        .eq("id", licenseId)
        .select()
        .single();

      if (error) {
        return json({ error: error.message }, 500);
      }

      return json({ success: true, license: data });
    }

    // PUT /extend/:licenseId — Extend license expiration
    if (req.method === "PUT" && segments[0] === "extend") {
      const licenseId = segments[1];
      if (!licenseId) {
        return json({ error: "Missing license id" }, 400);
      }
      const body = await req.json();
      const { expires_at } = body;
      if (!expires_at) {
        return json({ error: "Missing expires_at" }, 400);
      }
      const { data, error } = await supabase
        .from("licenses")
        .update({ expires_at, status: "ACTIVE" })
        .eq("id", licenseId)
        .select()
        .single();

      if (error) {
        return json({ error: error.message }, 500);
      }

      return json({ success: true, license: data });
    }

    // DELETE /:licenseId — Delete a license
    if (req.method === "DELETE" && segments.length === 1) {
      const licenseId = segments[0];
      const { error } = await supabase.from("licenses").delete().eq("id", licenseId);
      if (error) {
        return json({ error: error.message }, 500);
      }
      return json({ success: true });
    }

    // GET /stats — Dashboard statistics
    if (req.method === "GET" && segments[0] === "stats") {
      const { data: all } = await supabase.from("licenses").select("status, edition, activation_count, expires_at");
      const now = new Date();
      const stats = {
        total: all?.length || 0,
        active: all?.filter((l: any) => l.status === "ACTIVE").length || 0,
        pending: all?.filter((l: any) => l.status === "PENDING").length || 0,
        expired: all?.filter((l: any) => l.status === "EXPIRED" || new Date(l.expires_at) < now).length || 0,
        revoked: all?.filter((l: any) => l.status === "REVOKED").length || 0,
        premium: all?.filter((l: any) => l.edition === "PREMIUM").length || 0,
        standard: all?.filter((l: any) => l.edition === "STANDARD").length || 0,
        total_activations: all?.reduce((sum: number, l: any) => sum + l.activation_count, 0) || 0,
      };
      return json({ stats });
    }

    return json({ error: "Not found" }, 404);
  } catch (err) {
    return json({ error: err.message }, 500);
  }
});

function json(data: any, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { ...corsHeaders, "Content-Type": "application/json" },
  });
}
