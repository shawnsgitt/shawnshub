import { getStore } from "@netlify/blobs";

// ───────── CATEGORY LEARNING API ─────────
// Stores user-confirmed category mappings so the system auto-categorizes
// similar transactions in future uploads.
//
// POST /api/learn-categories
// Body: { action: "learn", mappings: [{ description: "...", category: "..." }] }
// Body: { action: "get" }
// Body: { action: "forget", key: "normalized-key" }
// Body: { action: "reset" }

function normalizeForLearning(desc) {
  var s = (desc || "").toLowerCase().trim();
  s = s.replace(/\b(ref|txn|transaction|card|payment)\s*[:# -]?\s*\d+\b/gi, "");
  s = s.replace(/\d{1,2}[\/\-\.]\d{1,2}([\/\-\.]\d{2,4})?/g, "");
  s = s.replace(/\b\d{6,}\b/g, "");
  s = s.replace(/[£$€]\s*[\d,.]+/g, "");
  s = s.replace(/\s+/g, " ").trim();
  if (s.length < 3) return (desc || "").toLowerCase().trim();
  return s;
}

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

export default async (req) => {
  if (req.method !== "POST") {
    return json({ error: "POST required" }, 405);
  }

  try {
    const store = getStore({ name: "finance-hub", consistency: "strong" });
    const body = await req.json();
    const action = body.action;

    if (action === "get") {
      const data = await store.get("learned-categories", { type: "json" });
      return json(data || { mappings: {} });
    }

    if (action === "reset") {
      await store.setJSON("learned-categories", { mappings: {} });
      return json({ ok: true, message: "All learned categories cleared" });
    }

    if (action === "forget" && body.key) {
      const data = (await store.get("learned-categories", { type: "json" })) || { mappings: {} };
      delete data.mappings[body.key];
      await store.setJSON("learned-categories", data);
      return json({ ok: true });
    }

    if (action === "learn" && body.mappings && Array.isArray(body.mappings)) {
      const data = (await store.get("learned-categories", { type: "json" })) || { mappings: {} };
      let learned = 0;

      for (const m of body.mappings) {
        if (!m.description || !m.category) continue;
        const key = normalizeForLearning(m.description);
        if (!key) continue;

        if (!data.mappings[key]) {
          data.mappings[key] = {
            category: m.category,
            example: m.description,
            count: 1,
            lastSeen: new Date().toISOString(),
          };
        } else {
          // Update existing — increment count and update category if changed
          data.mappings[key].category = m.category;
          data.mappings[key].count = (data.mappings[key].count || 0) + 1;
          data.mappings[key].lastSeen = new Date().toISOString();
        }
        learned++;
      }

      await store.setJSON("learned-categories", data);
      return json({
        ok: true,
        learned,
        totalMappings: Object.keys(data.mappings).length,
      });
    }

    return json({ error: "Invalid action. Use: learn, get, forget, reset" }, 400);
  } catch (err) {
    return json({ error: err.message }, 500);
  }
};

export const config = {
  path: "/api/learn-categories",
};
