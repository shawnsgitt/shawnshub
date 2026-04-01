import { getStore } from "@netlify/blobs";
import * as XLSX from "xlsx";

// Category keywords for auto-categorization
const CATEGORY_RULES = [
  { category: "Groceries", keywords: ["tesco", "sainsbury", "asda", "aldi", "lidl", "morrisons", "waitrose", "co-op", "coop", "ocado", "m&s food", "iceland", "spar", "costco", "grocery", "supermarket", "farm foods"] },
  { category: "Eating Out", keywords: ["mcdonald", "burger king", "kfc", "nando", "pizza", "domino", "uber eats", "deliveroo", "just eat", "starbucks", "costa", "greggs", "pret", "subway", "restaurant", "cafe", "coffee", "takeaway", "wetherspoon", "wagamama", "five guys", "zizzi", "gourmet"] },
  { category: "Transport", keywords: ["tfl", "transport for london", "uber", "bolt", "lyft", "bus", "train", "rail", "fuel", "petrol", "shell", "bp", "esso", "texaco", "parking", "congestion", "dart charge", "taxi", "national rail", "oyster", "go-ahead"] },
  { category: "Shopping", keywords: ["amazon", "ebay", "asos", "zara", "h&m", "primark", "next", "argos", "john lewis", "currys", "ikea", "tk maxx", "sports direct", "nike", "adidas", "new look", "river island", "shein", "boohoo", "apple store", "google store"] },
  { category: "Subscriptions", keywords: ["netflix", "spotify", "disney", "youtube premium", "apple music", "amazon prime", "hulu", "now tv", "sky", "virgin media", "bt broadband", "audible", "adobe", "microsoft 365", "icloud", "google one", "playstation", "xbox", "crunchyroll", "patreon", "chatgpt", "openai"] },
  { category: "Bills & Utilities", keywords: ["electric", "gas", "water", "council tax", "tv licence", "broadband", "internet", "phone bill", "mobile", "ee ", "vodafone", "three", "o2 ", "giffgaff", "insurance", "rent", "mortgage", "british gas", "edf", "eon", "octopus energy", "thames water", "scottish power", "bulb"] },
  { category: "Health & Fitness", keywords: ["gym", "puregym", "the gym", "david lloyd", "fitness first", "pharmacy", "boots", "superdrug", "doctor", "dentist", "hospital", "health", "vitamin", "myprotein", "holland & barrett", "nuffield"] },
  { category: "Entertainment", keywords: ["cinema", "odeon", "cineworld", "vue", "theatre", "concert", "ticket", "ticketmaster", "eventbrite", "gaming", "steam", "playstation store", "nintendo", "bowling", "museum", "zoo", "theme park"] },
  { category: "Education", keywords: ["udemy", "coursera", "skillshare", "book", "waterstones", "wh smith", "tuition", "school", "university", "student", "course"] },
  { category: "Personal Care", keywords: ["barber", "hairdresser", "salon", "spa", "beauty", "nail", "lush", "the body shop", "perfume"] },
  { category: "Family Support", keywords: ["transfer to", "family", "gift", "charity", "donation"] },
  { category: "Income", keywords: ["salary", "wages", "payroll", "refund", "cashback", "interest earned", "dividend", "freelance", "invoice paid", "pension", "benefit", "tax refund", "hmrc"] },
  { category: "Savings", keywords: ["savings", "save", "investment", "isa", "premium bond", "vanguard", "trading 212", "freetrade", "nutmeg", "moneybox"] },
];

function categorize(description) {
  const lower = (description || "").toLowerCase();
  for (const rule of CATEGORY_RULES) {
    for (const keyword of rule.keywords) {
      if (lower.includes(keyword)) {
        return rule.category;
      }
    }
  }
  return "Uncategorized";
}

function detectAccountType(rows) {
  const sample = rows.slice(0, 10).map((r) => Object.values(r).join(" ")).join(" ").toLowerCase();
  if (sample.includes("saving")) return "Savings";
  if (sample.includes("current") || sample.includes("checking")) return "Current";
  if (sample.includes("credit card") || sample.includes("credit")) return "Credit Card";
  return "Unknown";
}

function parseAmount(val) {
  if (val == null) return 0;
  if (typeof val === "number") return val;
  const cleaned = String(val).replace(/[£$€,\s]/g, "").trim();
  if (!cleaned) return 0;
  return parseFloat(cleaned) || 0;
}

function findColumn(headers, tests) {
  return headers.findIndex((h) => {
    const lower = h.toLowerCase();
    return tests.some((t) => lower.includes(t));
  });
}

function parseExcelSheet(rows) {
  if (!rows || rows.length < 2) throw new Error("Sheet has no data rows");

  const accountType = detectAccountType(rows);
  const headers = Object.keys(rows[0]).map((h) => String(h).trim());
  const headersLower = headers.map((h) => h.toLowerCase());

  const dateCol = findColumn(headers, ["date"]);
  const descCol = findColumn(headers, ["description", "memo", "narrative", "details", "transaction", "reference", "payee"]);
  const amountCol = findColumn(headers, ["amount", "value"]);
  const debitCol = findColumn(headers, ["debit", "money out", "paid out"]);
  const creditCol = findColumn(headers, ["credit", "money in", "paid in"]);
  const balanceCol = findColumn(headers, ["balance"]);

  if (dateCol === -1) {
    throw new Error("Could not find a Date column. Please ensure your Excel file has a Date header.");
  }

  const transactions = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const values = headers.map((h) => row[h]);

    const rawDate = values[dateCol];
    let date = "";
    if (rawDate instanceof Date) {
      const d = rawDate;
      date = `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`;
    } else if (rawDate != null) {
      date = String(rawDate).trim();
    }

    const description = values[descCol >= 0 ? descCol : 1] != null ? String(values[descCol >= 0 ? descCol : 1]).trim() : "";
    if (!date || !description) continue;

    let amount = 0;
    if (amountCol >= 0) {
      amount = parseAmount(values[amountCol]);
    } else if (debitCol >= 0 || creditCol >= 0) {
      const debit = debitCol >= 0 ? parseAmount(values[debitCol]) : 0;
      const credit = creditCol >= 0 ? parseAmount(values[creditCol]) : 0;
      amount = credit > 0 ? credit : -Math.abs(debit);
    }

    const balance = balanceCol >= 0 ? parseAmount(values[balanceCol]) : null;
    const category = categorize(description);
    const type =
      category === "Income"
        ? "Income"
        : category === "Savings"
          ? "Savings"
          : amount > 0
            ? "Income"
            : "Expense";

    transactions.push({
      id: `txn_${Date.now()}_${i}`,
      date,
      description,
      amount,
      absAmount: Math.abs(amount),
      balance,
      category,
      type,
      manualCategory: false,
    });
  }

  return { transactions, accountType };
}

function parseExcel(buffer) {
  const workbook = XLSX.read(buffer, { type: "buffer", cellDates: true });
  const allTransactions = [];
  let accountType = "Unknown";

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    if (rows.length < 2) continue;

    try {
      const result = parseExcelSheet(rows);
      allTransactions.push(...result.transactions);
      if (result.accountType !== "Unknown") {
        accountType = result.accountType;
      }
    } catch (e) {
      // Skip sheets that don't have valid transaction data
      continue;
    }
  }

  if (allTransactions.length === 0) {
    throw new Error("No transactions found in any sheet. Please ensure your Excel file has Date and Description columns.");
  }

  return { transactions: allTransactions, accountType };
}

function summarizeCategories(transactions) {
  const cats = {};
  for (const t of transactions) {
    cats[t.category] = (cats[t.category] || 0) + 1;
  }
  return cats;
}

export default async (req) => {
  if (req.method !== "POST") {
    return new Response(JSON.stringify({ error: "POST required" }), {
      status: 405,
      headers: { "Content-Type": "application/json" },
    });
  }

  try {
    const formData = await req.formData();
    const file = formData.get("file");
    const accountLabel = formData.get("accountLabel") || "";

    if (!file) {
      return new Response(JSON.stringify({ error: "No file uploaded" }), {
        status: 400,
        headers: { "Content-Type": "application/json" },
      });
    }

    const fileName = (file.name || "").toLowerCase();
    const isExcel = fileName.endsWith(".xlsx") || fileName.endsWith(".xls") ||
      file.type === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
      file.type === "application/vnd.ms-excel";

    if (!isExcel) {
      return new Response(
        JSON.stringify({ error: "Only Excel files (.xlsx, .xls) are supported. Please upload an Excel document." }),
        { status: 400, headers: { "Content-Type": "application/json" } }
      );
    }

    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    const { transactions, accountType } = parseExcel(buffer);

    if (transactions.length === 0) {
      return new Response(
        JSON.stringify({ error: "No transactions found in file. Please check the file format." }),
        { status: 400, headers: { "Content-Type": "application/json" } }
      );
    }

    const store = getStore({ name: "finance-data", consistency: "strong" });

    const existing = (await store.get("transactions", { type: "json" })) || { accounts: {} };

    const label = accountLabel || `${accountType} Account`;
    if (!existing.accounts[label]) {
      existing.accounts[label] = [];
    }

    const existingSet = new Set(
      existing.accounts[label].map(
        (t) => `${t.date}|${t.description}|${t.amount}`
      )
    );

    let added = 0;
    for (const txn of transactions) {
      const key = `${txn.date}|${txn.description}|${txn.amount}`;
      if (!existingSet.has(key)) {
        existing.accounts[label].push(txn);
        existingSet.add(key);
        added++;
      }
    }

    await store.setJSON("transactions", existing);

    return new Response(
      JSON.stringify({
        success: true,
        accountType,
        accountLabel: label,
        totalParsed: transactions.length,
        newAdded: added,
        duplicatesSkipped: transactions.length - added,
        sampleCategories: summarizeCategories(transactions),
      }),
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
  path: "/api/upload",
};
