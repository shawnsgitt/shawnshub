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
    const { accountLabel, transactionId, newCategory } = body;

    if (!accountLabel || !transactionId || !newCategory) {
      return new Response(
        JSON.stringify({ error: "accountLabel, transactionId, and newCategory required" }),
        { status: 400, headers: { "Content-Type": "application/json" } }
      );
    }

    const store = getStore({ name: "finance-data", consistency: "strong" });
    const data = (await store.get("transactions", { type: "json" })) || { accounts: {} };

    if (!data.accounts[accountLabel]) {
      return new Response(JSON.stringify({ error: "Account not found" }), {
        status: 404,
        headers: { "Content-Type": "application/json" },
      });
    }

    const txn = data.accounts[accountLabel].find((t) => t.id === transactionId);
    if (!txn) {
      return new Response(JSON.stringify({ error: "Transaction not found" }), {
        status: 404,
        headers: { "Content-Type": "application/json" },
      });
    }

    // Update the category
    txn.category = newCategory;
    txn.manualCategory = true;

    // Update type based on new category
    if (newCategory === "Income") {
      txn.type = "Income";
    } else if (newCategory === "Savings") {
      txn.type = "Savings";
    } else {
      txn.type = "Expense";
    }

    await store.setJSON("transactions", data);

    return new Response(
      JSON.stringify({ success: true, transaction: txn }),
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
  path: "/api/update-category",
};
