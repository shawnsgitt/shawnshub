import { getStore } from "@netlify/blobs";

// Generic CRUD for all financial hub data collections
// Stores: income, budget, savings-goals, debt, investments, bills, emergency, financial-goals
const VALID_STORES = [
  "income", "budget", "savings-goals", "debt",
  "investments", "bills", "emergency", "financial-goals"
];

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

export default async (req) => {
  try {
    const url = new URL(req.url);
    const collection = url.searchParams.get("collection");

    if (!collection || !VALID_STORES.includes(collection)) {
      return json({ error: "Invalid collection. Valid: " + VALID_STORES.join(", ") }, 400);
    }

    const store = getStore({ name: "finance-hub", consistency: "strong" });
    const key = collection;

    if (req.method === "GET") {
      const data = await store.get(key, { type: "json" });
      return json(data || { items: [] });
    }

    if (req.method === "POST") {
      const body = await req.json();
      const action = body.action; // "set" | "add" | "update" | "delete"

      if (action === "set") {
        // Replace entire collection
        await store.setJSON(key, body.data);
        return json({ ok: true });
      }

      const existing = (await store.get(key, { type: "json" })) || { items: [] };

      if (action === "add") {
        const item = body.item;
        item.id = "item_" + Date.now() + "_" + Math.random().toString(36).slice(2, 8);
        existing.items.push(item);
        await store.setJSON(key, existing);
        return json({ ok: true, item });
      }

      if (action === "update") {
        const idx = existing.items.findIndex(i => i.id === body.item.id);
        if (idx === -1) return json({ error: "Item not found" }, 404);
        existing.items[idx] = body.item;
        await store.setJSON(key, existing);
        return json({ ok: true });
      }

      if (action === "delete") {
        existing.items = existing.items.filter(i => i.id !== body.id);
        await store.setJSON(key, existing);
        return json({ ok: true });
      }

      return json({ error: "Invalid action" }, 400);
    }

    return json({ error: "Method not allowed" }, 405);
  } catch (err) {
    return json({ error: err.message }, 500);
  }
};

export const config = {
  path: "/api/hub-data",
};
