import * as XLSX from "xlsx";

// Category keywords for auto-categorization — extensive list for maximum auto-match
const CATEGORY_RULES = [
  { category: "Groceries", keywords: ["tesco", "sainsbury", "asda", "aldi", "lidl", "morrisons", "waitrose", "co-op", "coop", "ocado", "m&s food", "marks & spencer food", "iceland", "spar", "costco", "grocery", "supermarket", "farm foods", "whole foods", "wholefoods", "shoprite", "pick n pay", "checkers", "woolworths food", "food lover", "fruit & veg", "butcher", "bakery", "market", "fresh", "farmfoods", "heron foods", "jack's", "booths", "nisa"] },
  { category: "Eating Out", keywords: ["mcdonald", "burger king", "kfc", "nando", "pizza", "domino", "uber eats", "deliveroo", "just eat", "starbucks", "costa", "greggs", "pret", "subway", "restaurant", "cafe", "coffee", "takeaway", "wetherspoon", "wagamama", "five guys", "zizzi", "gourmet", "eat ", "dine", "dining", "food delivery", "grubhub", "doordash", "postmates", "chipotle", "taco bell", "wendy", "chick-fil-a", "popeyes", "panera", "dunkin", "tim horton", "sushi", "ramen", "kebab", "chicken", "grill", "steers", "wimpy", "debonairs", "roman's", "fishaways", "ocean basket", "spur", "mugg & bean", "vida e", "seattle coffee", "nero"] },
  { category: "Transport", keywords: ["tfl", "transport for london", "uber trip", "bolt", "lyft", "bus", "train", "rail", "fuel", "petrol", "diesel", "shell", "bp", "esso", "texaco", "parking", "congestion", "dart charge", "taxi", "national rail", "oyster", "go-ahead", "engen", "caltex", "sasol", "total garage", "garage", "tollgate", "toll", "e-toll", "etoll", "gautrain", "metrorail", "rea vaya", "myciti", "golden arrow", "car wash", "car service", "tyres", "motor", "vehicle", "aa ", "rac ", "mot test", "breakdown"] },
  { category: "Shopping", keywords: ["amazon", "ebay", "asos", "zara", "h&m", "primark", "next", "argos", "john lewis", "currys", "ikea", "tk maxx", "sports direct", "nike", "adidas", "new look", "river island", "shein", "boohoo", "apple store", "google store", "takealot", "mr price", "jet ", "edgars", "game ", "incredible connect", "makro", "builders", "clothing", "fashion", "shoes", "retail", "store", "shop", "mall", "outlet", "pep ", "ackermans", "truworths", "cotton on", "superdry", "uniqlo", "gap ", "mango", "forever 21", "pull & bear", "bershka", "massimo dutti", "matalan", "george ", "tu clothing", "decathlon"] },
  { category: "Subscriptions", keywords: ["netflix", "spotify", "disney", "youtube premium", "apple music", "amazon prime", "hulu", "now tv", "sky ", "virgin media", "bt broadband", "audible", "adobe", "microsoft 365", "icloud", "google one", "playstation plus", "ps plus", "xbox game pass", "crunchyroll", "patreon", "chatgpt", "openai", "dstv", "showmax", "multichoice", "monthly sub", "subscription", "membership", "renewal", "recurring", "annual fee", "apple tv", "hbo", "paramount", "peacock", "deezer", "tidal", "canva", "notion", "dropbox", "github", "linkedin premium", "twitch"] },
  { category: "Bills & Utilities", keywords: ["electric", "gas ", "water", "council tax", "tv licence", "broadband", "internet", "phone bill", "mobile", "ee ", "vodafone", "three ", "o2 ", "giffgaff", "insurance", "rent ", "mortgage", "british gas", "edf", "eon", "octopus energy", "thames water", "scottish power", "bulb", "eskom", "city power", "city of johannesburg", "city of cape town", "rates", "levy", "body corporate", "municipal", "telkom", "mtn ", "cell c", "vodacom", "fibre", "wifi", "rain ", "afrihost", "webafrica", "nbn", "comcast", "at&t", "verizon", "spectrum", "strata", "property management", "maintenance fee", "service charge"] },
  { category: "Health & Fitness", keywords: ["gym", "puregym", "the gym", "david lloyd", "fitness first", "pharmacy", "boots", "superdrug", "doctor", "dentist", "hospital", "health", "vitamin", "myprotein", "holland & barrett", "nuffield", "clicks", "dischem", "dis-chem", "planet fitness", "virgin active", "medical", "clinic", "optometrist", "optician", "physio", "therapist", "counsell", "psychology", "chiropract", "pathology", "lancet", "ampath", "discovery health", "medical aid", "bupa", "vitality", "wellness", "supplement", "protein"] },
  { category: "Entertainment", keywords: ["cinema", "odeon", "cineworld", "vue", "theatre", "concert", "ticket", "ticketmaster", "eventbrite", "gaming", "steam", "playstation store", "nintendo", "bowling", "museum", "zoo", "theme park", "ster-kinekor", "nu metro", "computicket", "webtickets", "festival", "event", "arcade", "laser", "escape room", "comedy", "show", "performance", "gallery", "aquarium", "funfair", "amusement"] },
  { category: "Education", keywords: ["udemy", "coursera", "skillshare", "book", "waterstones", "wh smith", "tuition", "school", "university", "student", "course", "college", "academy", "training", "workshop", "seminar", "exam", "certification", "study", "learning", "lecture", "textbook", "stationery", "cna ", "exclusive books", "loot.co", "kindle", "amazon book"] },
  { category: "Personal Care", keywords: ["barber", "hairdresser", "salon", "spa", "beauty", "nail", "lush", "the body shop", "perfume", "cosmetic", "makeup", "skincare", "grooming", "massage", "facial", "wax", "sorbet", "rain salon", "reed & barton", "dermatolog", "aesthetic"] },
  { category: "Family Support", keywords: ["transfer to", "family", "gift", "charity", "donation", "send money", "remittance", "allowance", "pocket money", "church", "tithe", "offering", "zakat", "support", "maintenance"] },
  { category: "Income", keywords: ["salary", "wages", "payroll", "refund", "cashback", "interest earned", "dividend", "freelance", "invoice paid", "pension", "benefit", "tax refund", "hmrc", "sars", "income", "commission", "bonus", "stipend", "bursary", "grant", "payout", "deposit from", "payment received", "credit received", "reversal", "reward"] },
  { category: "Savings", keywords: ["savings", "save", "investment", "isa", "premium bond", "vanguard", "trading 212", "freetrade", "nutmeg", "moneybox", "unit trust", "money market", "fixed deposit", "notice deposit", "capitec save", "fnb save", "easy equities", "etf", "satrix", "sygnia", "allan gray", "coronation", "stanlib"] },
  { category: "Cash Withdrawal", keywords: ["atm", "cash withdrawal", "cashback", "cash back at", "withdraw", "cash send", "cardless"] },
  { category: "Bank Fees", keywords: ["bank charge", "bank fee", "service fee", "transaction fee", "admin fee", "card fee", "account fee", "monthly fee", "annual fee", "interest charged", "overdraft", "debit order fee", "eft fee", "penalty", "late fee"] },
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

function parseAmount(val) {
  if (val == null) return 0;
  if (typeof val === "number") return val;
  const cleaned = String(val).replace(/[£$€,\s]/g, "").trim();
  if (!cleaned) return 0;
  return parseFloat(cleaned) || 0;
}

function findColumn(headers, tests) {
  return headers.findIndex((h) => {
    const lower = h.toLowerCase().trim();
    return tests.some((t) => lower.includes(t));
  });
}

// Smart suggestion: guess a likely category from description even if no keyword matches
function suggestCategory(description) {
  const lower = (description || "").toLowerCase();
  // Patterns that hint at common categories
  if (/\b(pay|pymt|payment|eft|debit order)\b/i.test(lower) && /\b(rent|landlord|property|estate)\b/i.test(lower)) return "Bills & Utilities";
  if (/\b(transfer|trfr?|eft|zelle|venmo|cashapp)\b/i.test(lower) && !/\b(saving|invest)\b/i.test(lower)) return null; // could be anything
  if (/pos\b|point of sale|card purchase|purchase/i.test(lower)) return null; // too generic
  return null;
}

function isHeaderRow(values) {
  const joined = values.map((v) => String(v).toLowerCase()).join(" ");
  return joined.includes("date") && (joined.includes("description") || joined.includes("memo") || joined.includes("narrative") || joined.includes("details") || joined.includes("transaction") || joined.includes("reference") || joined.includes("payee") || joined.includes("seller") || joined.includes("merchant") || joined.includes("beneficiary") || joined.includes("particulars") || joined.includes("amount") || joined.includes("debit") || joined.includes("credit") || joined.includes("balance"));
}

function findHeaderRow(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
  for (let r = range.s.r; r <= Math.min(range.e.r, 20); r++) {
    const values = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = sheet[addr];
      values.push(cell ? String(cell.v || "").trim() : "");
    }
    if (isHeaderRow(values)) {
      return r;
    }
  }
  return -1;
}

function parseExcelSheet(rows, headers) {
  if (!rows || rows.length < 1) throw new Error("Sheet has no data rows");

  const dateCol = findColumn(headers, ["date"]);
  const descCol = findColumn(headers, ["description", "memo", "narrative", "details", "particulars", "payee", "beneficiary", "merchant", "seller"]);
  const amountCol = findColumn(headers, ["amount", "value"]);
  const debitCol = findColumn(headers, ["debit", "money out", "paid out", "withdrawal", "dr"]);
  const creditCol = findColumn(headers, ["credit", "money in", "paid in", "deposit", "cr"]);
  const balanceCol = findColumn(headers, ["balance", "running balance", "closing balance"]);
  const typeCol = findColumn(headers, ["type", "transaction type", "dr/cr", "dr cr", "entry type"]);

  // Also look for a "reference" or "seller/payee" column as secondary description
  const refCol = findColumn(headers, ["reference", "ref", "transaction ref"]);
  const sellerCol = findColumn(headers, ["seller", "merchant", "payee", "beneficiary"]);

  if (dateCol === -1) {
    throw new Error("Could not find a Date column. Please ensure your Excel file has a Date header.");
  }

  // Determine column semantics from header names for better accuracy
  const hasDebitCredit = debitCol >= 0 || creditCol >= 0;

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

    const rawDesc = values[descCol >= 0 ? descCol : 1] != null ? String(values[descCol >= 0 ? descCol : 1]).trim() : "";
    // If there's a separate seller/merchant column, prefer it or combine
    const sellerVal = sellerCol >= 0 && sellerCol !== descCol ? String(values[sellerCol] || "").trim() : "";
    const refVal = refCol >= 0 && refCol !== descCol ? String(values[refCol] || "").trim() : "";
    // Build the best description: seller > description > reference
    let description = sellerVal || rawDesc || refVal || "";
    // If seller is separate and description adds context, combine them
    if (sellerVal && rawDesc && sellerVal.toLowerCase() !== rawDesc.toLowerCase()) {
      description = sellerVal + " - " + rawDesc;
    }
    if (!date || !description) continue;

    let amount = 0;

    // Strategy: Use the most specific columns available
    // Debit = money OUT (expense/negative), Credit = money IN (income/positive)
    if (hasDebitCredit) {
      const debitVal = debitCol >= 0 ? parseAmount(values[debitCol]) : 0;
      const creditVal = creditCol >= 0 ? parseAmount(values[creditCol]) : 0;

      if (creditVal !== 0 && debitVal !== 0) {
        // Both filled — net them (unusual but some banks do this)
        amount = Math.abs(creditVal) - Math.abs(debitVal);
      } else if (creditVal !== 0) {
        // Credit column: money IN → positive
        amount = Math.abs(creditVal);
      } else if (debitVal !== 0) {
        // Debit column: money OUT → negative
        amount = -Math.abs(debitVal);
      }
    } else if (amountCol >= 0) {
      amount = parseAmount(values[amountCol]);
      // Check for a type column that indicates debit/credit
      if (typeCol >= 0) {
        const typeVal = String(values[typeCol] || "").toLowerCase().trim();
        if (typeVal === "dr" || typeVal === "debit" || typeVal === "d" || typeVal === "expense" || typeVal === "withdrawal") {
          amount = -Math.abs(amount);
        } else if (typeVal === "cr" || typeVal === "credit" || typeVal === "c" || typeVal === "income" || typeVal === "deposit") {
          amount = Math.abs(amount);
        }
        // Otherwise trust the sign
      }
      // If no type column, trust the sign from the amount column as-is
    }

    const balance = balanceCol >= 0 ? parseAmount(values[balanceCol]) : null;
    const category = categorize(description);

    // Determine transaction type from amount direction
    // Positive amount = money coming IN (Income)
    // Negative amount = money going OUT (Expense)
    let type;
    if (category === "Income" || amount > 0) {
      type = "Income";
    } else if (category === "Savings") {
      type = "Savings";
    } else {
      type = "Expense";
    }

    transactions.push({
      id: `txn_${i}`,
      date,
      description,
      seller: sellerVal || "",
      amount,
      absAmount: Math.abs(amount),
      balance,
      category,
      type,
    });
  }

  return { transactions };
}

// Post-parse: verify debit/credit orientation using balance progression
// If balance goes UP when we say expense, signs are likely inverted
function verifyAndFixOrientation(transactions) {
  // Check consecutive rows where balance is available
  let correctCount = 0;
  let invertedCount = 0;
  for (let i = 1; i < transactions.length; i++) {
    const prev = transactions[i - 1];
    const curr = transactions[i];
    if (prev.balance != null && curr.balance != null) {
      const balanceDiff = curr.balance - prev.balance;
      if (Math.abs(balanceDiff) < 0.01) continue; // skip zero diffs
      // If amount sign matches balance movement direction, orientation is correct
      if ((curr.amount > 0 && balanceDiff > 0) || (curr.amount < 0 && balanceDiff < 0)) {
        correctCount++;
      } else if ((curr.amount > 0 && balanceDiff < 0) || (curr.amount < 0 && balanceDiff > 0)) {
        invertedCount++;
      }
    }
  }

  // If inverted signals outnumber correct ones by 2:1, flip all signs
  if (invertedCount > correctCount * 2 && invertedCount >= 3) {
    transactions.forEach(t => {
      t.amount = -t.amount;
      // Recategorize type based on new amount
      if (t.category === "Income" || t.amount > 0) {
        t.type = "Income";
      } else if (t.category === "Savings") {
        t.type = "Savings";
      } else {
        t.type = "Expense";
      }
    });
  }
  return transactions;
}

function parseExcel(buffer) {
  const workbook = XLSX.read(buffer, { type: "buffer", cellDates: true });
  const allTransactions = [];
  const sheetErrors = [];

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet["!ref"]) continue;

    const headerRowIdx = findHeaderRow(sheet);
    let rows;

    if (headerRowIdx >= 0) {
      const range = XLSX.utils.decode_range(sheet["!ref"]);
      range.s.r = headerRowIdx;
      const newRef = XLSX.utils.encode_range(range);
      rows = XLSX.utils.sheet_to_json(sheet, { defval: "", range: newRef });
    } else {
      rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    }

    if (!rows || rows.length < 1) continue;

    const headers = Object.keys(rows[0]).map((h) => String(h).trim());

    try {
      const result = parseExcelSheet(rows, headers);
      allTransactions.push(...result.transactions);
    } catch (e) {
      sheetErrors.push(`${sheetName}: ${e.message}`);
      continue;
    }
  }

  if (allTransactions.length === 0) {
    const detail = sheetErrors.length > 0 ? " Sheet issues: " + sheetErrors.join("; ") : "";
    throw new Error("No transactions found in any sheet." + detail);
  }

  // Verify and fix orientation using balance column if available
  verifyAndFixOrientation(allTransactions);

  return allTransactions;
}

function generateAnalysis(transactions) {
  const totalIncome = transactions.filter(t => t.type === "Income").reduce((s, t) => s + t.absAmount, 0);
  const totalExpenses = transactions.filter(t => t.type === "Expense").reduce((s, t) => s + t.absAmount, 0);
  const totalSavings = transactions.filter(t => t.type === "Savings").reduce((s, t) => s + t.absAmount, 0);
  const net = totalIncome - totalExpenses - totalSavings;

  // Category breakdown for expenses only
  const categoryBreakdown = {};
  transactions.filter(t => t.type === "Expense").forEach(t => {
    if (!categoryBreakdown[t.category]) {
      categoryBreakdown[t.category] = { total: 0, count: 0, transactions: [] };
    }
    categoryBreakdown[t.category].total += t.absAmount;
    categoryBreakdown[t.category].count++;
  });

  // Sort by total descending
  const sortedCategories = Object.entries(categoryBreakdown)
    .map(([name, data]) => ({
      name,
      total: Math.round(data.total * 100) / 100,
      count: data.count,
      percentage: totalExpenses > 0 ? Math.round(data.total / totalExpenses * 1000) / 10 : 0,
    }))
    .sort((a, b) => b.total - a.total);

  // Generate recommendations
  const recommendations = [];

  if (totalExpenses > totalIncome * 0.9) {
    recommendations.push({
      type: "warning",
      title: "Spending exceeds 90% of income",
      detail: `You're spending £${Math.round(totalExpenses).toLocaleString()} out of £${Math.round(totalIncome).toLocaleString()} income (${totalIncome > 0 ? Math.round(totalExpenses / totalIncome * 100) : 0}%). Try to keep spending below 70-80% of your income to build a safety net.`,
    });
  }

  if (totalExpenses > totalIncome) {
    recommendations.push({
      type: "critical",
      title: "You are spending more than you earn",
      detail: `Your expenses (£${Math.round(totalExpenses).toLocaleString()}) exceed your income (£${Math.round(totalIncome).toLocaleString()}) by £${Math.round(totalExpenses - totalIncome).toLocaleString()}. This is unsustainable and needs immediate attention.`,
    });
  }

  // Subscriptions check
  const subCat = sortedCategories.find(c => c.name === "Subscriptions");
  if (subCat && subCat.percentage > 5) {
    recommendations.push({
      type: "suggestion",
      title: "Review your subscriptions",
      detail: `Subscriptions account for ${subCat.percentage}% of your spending (£${Math.round(subCat.total).toLocaleString()}). Review each subscription and cancel any you don't actively use. Even cutting 1-2 can save you hundreds per year.`,
    });
  }

  // Eating out vs groceries
  const eatingOut = sortedCategories.find(c => c.name === "Eating Out");
  const groceries = sortedCategories.find(c => c.name === "Groceries");
  if (eatingOut && groceries && eatingOut.total > groceries.total) {
    recommendations.push({
      type: "suggestion",
      title: "Eating out costs more than groceries",
      detail: `You spend £${Math.round(eatingOut.total).toLocaleString()} eating out vs £${Math.round(groceries.total).toLocaleString()} on groceries. Cooking more at home could significantly reduce your food spend.`,
    });
  }
  if (eatingOut && eatingOut.percentage > 15) {
    recommendations.push({
      type: "suggestion",
      title: "High eating out spend",
      detail: `Eating out accounts for ${eatingOut.percentage}% of your spending. Consider meal prepping or limiting takeaways to weekends.`,
    });
  }

  // Top spending category advice
  if (sortedCategories.length > 0) {
    const top = sortedCategories[0];
    recommendations.push({
      type: "info",
      title: `${top.name} is your biggest expense`,
      detail: `${top.name} takes up ${top.percentage}% of your spending at £${Math.round(top.total).toLocaleString()} across ${top.count} transactions. This is where reducing spend would have the most impact.`,
    });
  }

  // Shopping advice
  const shopping = sortedCategories.find(c => c.name === "Shopping");
  if (shopping && shopping.percentage > 15) {
    recommendations.push({
      type: "suggestion",
      title: "Consider reducing shopping spend",
      detail: `Shopping accounts for ${shopping.percentage}% of your expenses (£${Math.round(shopping.total).toLocaleString()}). Try implementing a 24-hour rule before non-essential purchases.`,
    });
  }

  // Transport advice
  const transport = sortedCategories.find(c => c.name === "Transport");
  if (transport && transport.percentage > 15) {
    recommendations.push({
      type: "suggestion",
      title: "Transport costs are high",
      detail: `Transport is ${transport.percentage}% of your spending (£${Math.round(transport.total).toLocaleString()}). Consider if public transport, carpooling, or cycling could reduce these costs.`,
    });
  }

  // Savings praise or advice
  if (totalSavings > totalIncome * 0.2) {
    recommendations.push({
      type: "positive",
      title: "Great saving habits!",
      detail: `You're saving ${totalIncome > 0 ? Math.round(totalSavings / totalIncome * 100) : 0}% of your income. That's above the recommended 20%. Keep it up!`,
    });
  } else if (totalSavings < totalIncome * 0.1 && totalIncome > 0) {
    recommendations.push({
      type: "suggestion",
      title: "Try to save more",
      detail: `You're only saving ${Math.round(totalSavings / totalIncome * 100)}% of your income. The recommended minimum is 10-20%. Consider setting up an automatic transfer to savings after payday.`,
    });
  }

  // Uncategorized warning
  const uncatCount = transactions.filter(t => t.category === "Uncategorized").length;
  if (uncatCount > 0) {
    recommendations.push({
      type: "info",
      title: `${uncatCount} uncategorized transaction${uncatCount !== 1 ? "s" : ""}`,
      detail: "Categorize these to get more accurate insights. The system will prompt you to assign categories.",
    });
  }

  return {
    totalIncome: Math.round(totalIncome * 100) / 100,
    totalExpenses: Math.round(totalExpenses * 100) / 100,
    totalSavings: Math.round(totalSavings * 100) / 100,
    net: Math.round(net * 100) / 100,
    transactionCount: transactions.length,
    categoryBreakdown: sortedCategories,
    recommendations,
  };
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
        JSON.stringify({ error: "Only Excel files (.xlsx, .xls) are supported." }),
        { status: 400, headers: { "Content-Type": "application/json" } }
      );
    }

    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    const transactions = parseExcel(buffer);

    if (transactions.length === 0) {
      return new Response(
        JSON.stringify({ error: "No transactions found in file." }),
        { status: 400, headers: { "Content-Type": "application/json" } }
      );
    }

    const analysis = generateAnalysis(transactions);

    return new Response(
      JSON.stringify({
        success: true,
        transactions,
        analysis,
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
