import { getStore } from "@netlify/blobs";

export default async (req) => {
  if (req.method !== "POST") {
    return new Response(JSON.stringify({ error: "POST required" }), {
      status: 405,
      headers: { "Content-Type": "application/json" },
    });
  }

  try {
    const body = await req.json();
    const { accountLabel } = body;

    const store = getStore({ name: "finance-data", consistency: "strong" });
    const data = (await store.get("transactions", { type: "json" })) || { accounts: {} };

    if (accountLabel) {
      delete data.accounts[accountLabel];
    } else {
      data.accounts = {};
    }

    await store.setJSON("transactions", data);

    return new Response(
      JSON.stringify({ success: true }),
      { status: 200, headers: { "Content-Type": "application/json" } }
    );
  } catch (err) {
    return new Response(JSON.stringify({ error: err.message }), {
      status: 500,
      headers: { "Content-Type": "application/json" },
    });
  }
};

export const config = {
  path: "/api/clear-data",
};
