import { getStore } from "@netlify/blobs";
import { PDFParse } from "pdf-parse";

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

function detectAccountType(lines) {
  const joined = lines.slice(0, 10).join(" ").toLowerCase();
  if (joined.includes("saving")) return "Savings";
  if (joined.includes("current") || joined.includes("checking")) return "Current";
  if (joined.includes("credit card") || joined.includes("credit")) return "Credit Card";
  return "Unknown";
}

function splitCSVLine(line) {
  const result = [];
  let current = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      inQuotes = !inQuotes;
    } else if (ch === "," && !inQuotes) {
      result.push(current);
      current = "";
    } else {
      current += ch;
    }
  }
  result.push(current);
  return result;
}

function parseAmount(str) {
  if (!str) return 0;
  const cleaned = str.replace(/[£$€,\s]/g, "").trim();
  if (!cleaned) return 0;
  return parseFloat(cleaned) || 0;
}

function parseCSV(text) {
  const lines = text.split(/\r?\n/).filter((l) => l.trim());
  if (lines.length < 2) throw new Error("File has no data rows");

  const accountType = detectAccountType(lines);

  let headerIdx = -1;
  let headers = [];
  for (let i = 0; i < Math.min(lines.length, 10); i++) {
    const cols = splitCSVLine(lines[i]);
    const lower = cols.map((c) => c.toLowerCase().trim());
    if (
      lower.some((h) => h.includes("date")) &&
      lower.some(
        (h) =>
          h.includes("description") ||
          h.includes("memo") ||
          h.includes("narrative") ||
          h.includes("details") ||
          h.includes("transaction") ||
          h.includes("reference") ||
          h.includes("payee")
      )
    ) {
      headerIdx = i;
      headers = lower;
      break;
    }
  }

  if (headerIdx === -1) {
    headerIdx = 0;
    headers = splitCSVLine(lines[0]).map((c) => c.toLowerCase().trim());
  }

  const dateCol = headers.findIndex((h) => h.includes("date"));
  const descCol = headers.findIndex(
    (h) =>
      h.includes("description") ||
      h.includes("memo") ||
      h.includes("narrative") ||
      h.includes("details") ||
      h.includes("transaction") ||
      h.includes("reference") ||
      h.includes("payee")
  );
  const amountCol = headers.findIndex(
    (h) => h === "amount" || h.includes("amount") || h === "value"
  );
  const debitCol = headers.findIndex(
    (h) => h.includes("debit") || h.includes("money out") || h.includes("paid out")
  );
  const creditCol = headers.findIndex(
    (h) => h.includes("credit") || h.includes("money in") || h.includes("paid in")
  );
  const balanceCol = headers.findIndex((h) => h.includes("balance"));

  if (dateCol === -1) {
    throw new Error("Could not find a Date column. Please ensure your CSV has a Date header.");
  }

  const transactions = [];

  for (let i = headerIdx + 1; i < lines.length; i++) {
    const cols = splitCSVLine(lines[i]);
    if (cols.length < 2) continue;

    const date = cols[dateCol]?.trim();
    const description = cols[descCol >= 0 ? descCol : 1]?.trim();
    if (!date || !description) continue;

    let amount = 0;
    if (amountCol >= 0) {
      amount = parseAmount(cols[amountCol]);
    } else if (debitCol >= 0 || creditCol >= 0) {
      const debit = debitCol >= 0 ? parseAmount(cols[debitCol]) : 0;
      const credit = creditCol >= 0 ? parseAmount(cols[creditCol]) : 0;
      amount = credit > 0 ? credit : -Math.abs(debit);
    }

    const balance = balanceCol >= 0 ? parseAmount(cols[balanceCol]) : null;
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

async function parsePDF(buffer) {
  const parser = new PDFParse({ data: buffer });
  const result = await parser.getText();
  const text = result.text;
  const lines = text.split(/\n/).map((l) => l.trim()).filter(Boolean);

  const accountType = detectAccountType(lines);

  const datePatterns = [
    /(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/,
    /(\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2})/,
    /(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{2,4})/i,
  ];

  const amountPattern = /[£$€]?\s*-?\d{1,3}(?:,\d{3})*(?:\.\d{2})\b/g;

  const transactions = [];
  let lineIndex = 0;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    let dateMatch = null;
    for (const dp of datePatterns) {
      const m = line.match(dp);
      if (m && line.indexOf(m[1]) < 20) {
        dateMatch = m[1];
        break;
      }
    }
    if (!dateMatch) continue;

    let fullText = line;
    if (i + 1 < lines.length) {
      const nextLine = lines[i + 1];
      let hasDate = false;
      for (const dp of datePatterns) {
        const nm = nextLine.match(dp);
        if (nm && nextLine.indexOf(nm[1]) < 20) {
          hasDate = true;
          break;
        }
      }
      if (!hasDate) {
        fullText += " " + nextLine;
      }
    }

    const amounts = fullText.match(amountPattern);
    if (!amounts || amounts.length === 0) continue;

    const dateEnd = fullText.indexOf(dateMatch) + dateMatch.length;
    const firstAmtIdx = fullText.indexOf(amounts[0]);
    let description = fullText.substring(dateEnd, firstAmtIdx).trim();
    description = description.replace(/^[\s,\-|]+/, "").replace(/[\s,\-|]+$/, "").trim();
    if (!description || description.length < 2) {
      description = fullText.substring(dateEnd).replace(/[£$€\d,.\-\s]+$/, "").trim();
    }
    if (!description || description.length < 2) continue;

    let amount = parseAmount(amounts[0]);
    let balance = null;
    if (amounts.length >= 3) {
      const debit = parseAmount(amounts[0]);
      const credit = parseAmount(amounts[1]);
      balance = parseAmount(amounts[amounts.length - 1]);
      amount = credit > 0 ? credit : -Math.abs(debit);
    } else if (amounts.length === 2) {
      amount = parseAmount(amounts[0]);
      balance = parseAmount(amounts[1]);
    }

    if (amount === 0) continue;

    const category = categorize(description);
    const type =
      category === "Income"
        ? "Income"
        : category === "Savings"
          ? "Savings"
          : amount > 0
            ? "Income"
            : "Expense";

    lineIndex++;
    transactions.push({
      id: `txn_${Date.now()}_${lineIndex}`,
      date: dateMatch,
      description,
      amount,
      absAmount: Math.abs(amount),
      balance,
      category,
      type,
      manualCategory: false,
    });
  }

  await parser.destroy();
  return { transactions, accountType };
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
    const isPDF = fileName.endsWith(".pdf") || file.type === "application/pdf";

    let transactions, accountType;
    if (isPDF) {
      const arrayBuffer = await file.arrayBuffer();
      const buffer = Buffer.from(arrayBuffer);
      ({ transactions, accountType } = await parsePDF(buffer));
    } else {
      const text = await file.text();
      ({ transactions, accountType } = parseCSV(text));
    }

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
