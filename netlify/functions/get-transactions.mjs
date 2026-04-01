import { getStore } from "@netlify/blobs";

export default async (req) => {
  try {
    const store = getStore({ name: "finance-data", consistency: "strong" });
    const data = (await store.get("transactions", { type: "json" })) || { accounts: {} };

    return new Response(JSON.stringify(data), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: err.message }), {
      status: 500,
      headers: { "Content-Type": "application/json" },
    });
  }
};

export const config = {
  path: "/api/transactions",
};
