import * as XLSX from "xlsx";
import { getStore } from "@netlify/blobs";

// ───────── LEARNED CATEGORY MAPPINGS (Netlify Blobs) ─────────
async function loadLearnedMappings() {
  try {
    const store = getStore({ name: "finance-hub", consistency: "strong" });
    const data = await store.get("learned-categories", { type: "json" });
    return data || { mappings: {} };
  } catch { return { mappings: {} }; }
}

async function saveLearnedMappings(data) {
  try {
    const store = getStore({ name: "finance-hub", consistency: "strong" });
    await store.setJSON("learned-categories", data);
  } catch { /* silently fail — learning is best-effort */ }
}

// Normalize a description to a stable key for learning
function normalizeForLearning(desc) {
  var s = (desc || "").toLowerCase().trim();
  // Remove reference numbers, dates, card numbers, amounts
  s = s.replace(/\b(ref|txn|transaction|card|payment)\s*[:# -]?\s*\d+\b/gi, "");
  s = s.replace(/\d{1,2}[\/\-\.]\d{1,2}([\/\-\.]\d{2,4})?/g, "");
  s = s.replace(/\b\d{6,}\b/g, "");
  s = s.replace(/[£$€]\s*[\d,.]+/g, "");
  s = s.replace(/\s+/g, " ").trim();
  if (s.length < 3) return (desc || "").toLowerCase().trim();
  return s;
}

// ───────── BALANCE CARRIED FORWARD DETECTION ─────────
// These are not real transactions — they are opening/closing balance entries
const BALANCE_CF_PATTERNS = [
  "balance carried forward", "balance brought forward",
  "balance c/f", "balance b/f", "balance c/d", "balance b/d",
  "carried forward", "brought forward", "carried down", "brought down",
  "opening balance", "closing balance", "ob ", "cb ",
  "balance from previous", "previous balance", "last balance",
];

function isBalanceCarriedForward(description) {
  const lower = (description || "").toLowerCase();
  return BALANCE_CF_PATTERNS.some(p => lower.includes(p));
}

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
  { category: "Cash Transfer", keywords: ["transfer between", "inter account", "interaccount", "internal transfer", "own account", "between accounts", "acc transfer", "account transfer", "move money", "move funds", "sweep", "self transfer", "same name transfer"] },
  { category: "Income", keywords: ["salary", "wages", "payroll", "refund", "cashback", "interest earned", "dividend", "freelance", "invoice paid", "pension", "benefit", "tax refund", "hmrc", "sars", "income", "commission", "bonus", "stipend", "bursary", "grant", "payout", "deposit from", "payment received", "credit received", "reversal", "reward"] },
  { category: "Savings", keywords: ["savings", "save", "investment", "isa", "premium bond", "vanguard", "trading 212", "freetrade", "nutmeg", "moneybox", "unit trust", "money market", "fixed deposit", "notice deposit", "capitec save", "fnb save", "easy equities", "etf", "satrix", "sygnia", "allan gray", "coronation", "stanlib"] },
  { category: "Cash Withdrawal", keywords: ["atm", "cash withdrawal", "withdraw", "cash send", "cardless"] },
  { category: "Cash In", keywords: ["cash deposit", "cash in", "cash payment in", "cash credit", "cash received", "counter deposit", "counter credit", "branch deposit", "cash at branch", "cash lodgement", "lodgement"] },
  { category: "Bank Fees", keywords: ["bank charge", "bank fee", "service fee", "transaction fee", "admin fee", "card fee", "account fee", "monthly fee", "annual fee", "interest charged", "overdraft", "debit order fee", "eft fee", "penalty", "late fee"] },
];

// Non-transactional types — these are not expenses or income
const NON_TRANSACTIONAL_CATEGORIES = new Set(["Balance Carried Forward", "Cash Transfer"]);

function categorize(description, learnedMappings) {
  // First check if this is a balance carried forward entry
  if (isBalanceCarriedForward(description)) {
    return "Balance Carried Forward";
  }

  const lower = (description || "").toLowerCase();

  // Check learned mappings first (user-trained categories — highest confidence)
  if (learnedMappings && learnedMappings.mappings) {
    const normalized = normalizeForLearning(description);
    const learned = learnedMappings.mappings[normalized];
    if (learned && learned.category && learned.count >= 1) {
      return learned.category;
    }
  }

  // Check keyword rules (high confidence)
  for (const rule of CATEGORY_RULES) {
    for (const keyword of rule.keywords) {
      if (lower.includes(keyword)) {
        return rule.category;
      }
    }
  }

  // Smart suggestion — pattern-based heuristic categorization
  const suggested = suggestCategory(description);
  if (suggested) {
    return suggested;
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

// Smart suggestion: guess a likely category from description using pattern analysis
function suggestCategory(description) {
  const lower = (description || "").toLowerCase();

  // ── Balance / Non-transactional ──
  if (/\b(balance carried|balance brought|balance c\/f|balance b\/f|opening balance|closing balance|carried forward|brought forward)\b/.test(lower)) return "Balance Carried Forward";
  if (/\b(transfer between|inter account|interaccount|internal transfer|own account|between accounts|acc transfer|self transfer|same name transfer|sweep|move money|move funds)\b/.test(lower)) return "Cash Transfer";

  // ── Cash In / Deposits ──
  if (/\b(cash deposit|cash in|cash payment in|counter deposit|counter credit|branch deposit|cash at branch|cash lodgement|lodgement)\b/.test(lower)) return "Cash In";

  // ── Cash Withdrawal ──
  if (/\b(atm|cash withdrawal|withdraw|cash send|cardless|cash back at)\b/.test(lower)) return "Cash Withdrawal";

  // ── Income patterns ──
  if (/\b(salary|wages?|payroll|pay\s+from|income|commission|bonus|stipend|bursary|grant|payout|pension|benefit|annuity)\b/.test(lower)) return "Income";
  if (/\b(refund|reversal|cashback|reimburse|rebate|credit note|returned|money back)\b/.test(lower)) return "Income";
  if (/\b(sars|hmrc|tax refund|tax return|tax credit|dividend|interest earned|interest received|investment return)\b/.test(lower)) return "Income";
  if (/\b(freelance|invoice paid?|payment received|deposit from|credit received|reward)\b/.test(lower)) return "Income";

  // ── Bills & Utilities — very common, catch broadly ──
  if (/\b(rent|landlord|property|estate agent|letting|tenant|lease)\b/.test(lower)) return "Bills & Utilities";
  if (/\b(electric|electricity|gas\s|water|sewage|council|rates|levy|municipal|waste|refuse)\b/.test(lower)) return "Bills & Utilities";
  if (/\b(insurance|insure|cover|policy|premium|assurance|old mutual|sanlam|liberty|discovery|hollard|outsurance|momentum|1st for women|santam|miway|auto & general|budget ins)\b/.test(lower)) return "Bills & Utilities";
  if (/\b(vodafone|mtn|vodacom|cell\s?c|ee\b|o2\b|giffgaff|three\b|telkom|airtel|rain\b|afrihost|fibre|broadband|internet|wifi|dstv|multichoice|gotv)\b/.test(lower)) return "Bills & Utilities";
  if (/\b(mortgage|bond repayment|home loan|strata|body corporate|property management|maintenance fee|service charge)\b/.test(lower)) return "Bills & Utilities";
  if (/\b(tv licence|license fee|tv license)\b/.test(lower)) return "Bills & Utilities";
  if (/\b(debit order|d\/o|do\s+|direct debit|recurring payment|standing order)\b/.test(lower) && !/\b(gym|fitness|netflix|spotify|amazon|disney)\b/.test(lower)) return "Bills & Utilities";

  // ── Transport ──
  if (/\b(uber|bolt|lyft|taxi|cab|ride|didi|indriver)\b/.test(lower) && !/\b(eats|eat|food|delivery)\b/.test(lower)) return "Transport";
  if (/\b(fuel|petrol|diesel|garage|shell|bp\b|caltex|engen|sasol|esso|texaco|total\s+garage|filling station|gas station|service station)\b/.test(lower)) return "Transport";
  if (/\b(parking|park\s|ncp|q-park|easy park|paypoint|justpark|ring-?go)\b/.test(lower)) return "Transport";
  if (/\b(tfl|oyster|bus\b|train|rail|metro|tube|underground|gautrain|rea vaya|myciti|golden arrow)\b/.test(lower)) return "Transport";
  if (/\b(toll|e-?toll|dart charge|congestion|road\s+tax|vehicle|mot\s+test|car\s+service|tyre|tire|mechanic|auto\s+repair|car\s+wash)\b/.test(lower)) return "Transport";

  // ── Eating Out ──
  if (/\b(restaurant|cafe|coffee|bistro|brasserie|diner|eatery|canteen|food\s+court|kitchen)\b/.test(lower)) return "Eating Out";
  if (/\b(takeaway|take\s+away|delivery|uber\s*eats|deliveroo|just\s*eat|mr\s*delivery|door\s*dash|grub\s*hub|postmates|food\s+delivery|order\s+food)\b/.test(lower)) return "Eating Out";
  if (/\b(pizza|burger|chicken|sushi|ramen|kebab|curry|noodle|wings|grill|steakhouse|bbq|barbecue)\b/.test(lower)) return "Eating Out";
  if (/\b(bar\b|pub\b|tavern|lounge|cocktail|wine\s+bar|brewery|taproom|beer)\b/.test(lower)) return "Eating Out";

  // ── Groceries ──
  if (/\b(grocery|grocer|supermarket|fresh\s+market|food\s+market|butcher|bakery|greengrocer|fruit|veg|organic|farm\s+stall|deli)\b/.test(lower)) return "Groceries";

  // ── Shopping ──
  if (/\b(shop|store|mart|retail|buy|purchase|outlet|boutique|mall|plaza|centre|center)\b/.test(lower) && !/\b(coffee|food|eat|restaurant|grocery|body\s+shop)\b/.test(lower)) return "Shopping";
  if (/\b(online|click|e-?commerce|order|parcel|dispatch|shipped)\b/.test(lower) && !/\b(food|eat|delivery|uber|just\s*eat)\b/.test(lower)) return "Shopping";
  if (/\b(clothing|fashion|shoes|sneakers|apparel|wear|outfit|accessories|jewel|watch)\b/.test(lower)) return "Shopping";
  if (/\b(furniture|decor|home\s+goods|hardware|tools|diy|garden|appliance)\b/.test(lower)) return "Shopping";

  // ── Subscriptions ──
  if (/\b(subscri|member|renewal|recurring|monthly\s+fee|annual\s+fee|premium\s+plan|pro\s+plan)\b/.test(lower)) return "Subscriptions";
  if (/\b(netflix|spotify|disney|youtube|apple\s+music|hulu|hbo|amazon\s+prime|audible|adobe|microsoft|icloud|google\s+one|playstation|xbox|crunchyroll|patreon|chatgpt|openai|canva|notion|dropbox|github|linkedin|deezer|tidal|showmax|dstv)\b/.test(lower)) return "Subscriptions";

  // ── Health & Fitness ──
  if (/\b(gym|fitness|sport|active|exercise|workout|yoga|pilates|crossfit|martial)\b/.test(lower)) return "Health & Fitness";
  if (/\b(pharmacy|chemist|medical|doctor|dr\s|clinic|hospital|dental|dentist|optom|optician|physio|therapist|counsell|psychology|chiropract|pathology|radiology|x-?ray|scan|lab|blood\s+test)\b/.test(lower)) return "Health & Fitness";
  if (/\b(vitamin|supplement|protein|health\s+food|wellness|remedy|medicine|prescription|dispens)\b/.test(lower)) return "Health & Fitness";
  if (/\b(medical\s+aid|health\s+insurance|bupa|vitality|discovery\s+health|bonitas|gems|medshield)\b/.test(lower)) return "Health & Fitness";

  // ── Entertainment ──
  if (/\b(cinema|movie|theatre|theater|concert|ticket|event|show|performance|gallery|museum|zoo|aquarium|theme\s+park|amusement|funfair|arcade|bowling|escape\s+room|comedy|festival|laser|mini\s+golf)\b/.test(lower)) return "Entertainment";
  if (/\b(gaming|steam|playstation\s+store|nintendo|xbox\s+store|epic\s+games|riot|twitch|gaming)\b/.test(lower)) return "Entertainment";

  // ── Education ──
  if (/\b(school|university|college|tuition|course|academy|training|workshop|seminar|exam|certification|study|learning|lecture|textbook|stationery|education|student|tutorial|bootcamp)\b/.test(lower)) return "Education";
  if (/\b(udemy|coursera|skillshare|book|kindle|library|academic|research)\b/.test(lower)) return "Education";

  // ── Personal Care ──
  if (/\b(salon|barber|hairdresser|hair|beauty|nail|spa|wax|facial|massage|grooming|cosmetic|makeup|skincare|perfume|fragrance|aesthetic|dermatol|laser\s+treatment)\b/.test(lower)) return "Personal Care";

  // ── Savings ──
  if (/\b(sav(e|ing)|invest|isa\b|premium\s+bond|unit\s+trust|money\s+market|fixed\s+deposit|notice\s+deposit|etf|index\s+fund|retirement|provident|pension\s+fund|annuity\s+contrib)\b/.test(lower)) return "Savings";

  // ── Family Support ──
  if (/\b(transfer|trfr|eft|send|remit|allowance|pocket\s+money|support|maintenance)\b/.test(lower) && /\b(to\s|family|child|parent|mother|father|wife|husband|spouse|brother|sister|son|daughter|gift|church|charity|donat|tithe|offering|zakat|mosque|temple)\b/.test(lower)) return "Family Support";
  if (/\b(church|charity|donat|tithe|offering|zakat|ngo|foundation|fundrais|sponsor)\b/.test(lower)) return "Family Support";

  // ── Bank Fees ──
  if (/\b(fee|charge|admin|penalty|interest\s+charge|overdraft|insufficient|nsf|rejected|unpaid|bounce|ledger\s+fee|service\s+fee|transaction\s+fee|card\s+fee|annual\s+card|account\s+fee|monthly\s+fee|eft\s+fee|debit\s+order\s+fee|notification\s+fee|sms\s+fee|excess\s+fee)\b/.test(lower)) return "Bank Fees";

  // ── Fallback patterns based on transaction structure ──
  // POS / card purchases are likely shopping or groceries
  if (/\b(pos|point\s+of\s+sale|card\s+purchase|purchase|buy)\b/.test(lower)) return "Shopping";
  // Generic payments — if "pay" with no other context, likely bills
  if (/\b(pay|pymt|payment|debit\s+order|d\/o)\b/.test(lower)) return "Bills & Utilities";
  // Transfers without clear destination — family support as best guess
  if (/\b(transfer|trfr|eft|send|remit)\b/.test(lower)) return "Family Support";

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

function detectSignConvention(rawEntries) {
  // Use balance changes to detect if positive amounts mean expense (inverted)
  // Check consecutive rows where balance is available
  let balanceVotes = { normal: 0, inverted: 0 };
  for (let i = 1; i < rawEntries.length; i++) {
    const prev = rawEntries[i - 1];
    const curr = rawEntries[i];
    if (prev.balance != null && curr.balance != null && curr.amount !== 0) {
      const balanceDelta = curr.balance - prev.balance;
      // If amount sign matches balance change direction, sign convention is normal
      // (positive amount = balance goes up = income)
      if (Math.abs(balanceDelta - curr.amount) < 0.02) {
        balanceVotes.normal++;
      } else if (Math.abs(balanceDelta + curr.amount) < 0.02) {
        balanceVotes.inverted++;
      }
    }
  }
  if (balanceVotes.normal + balanceVotes.inverted >= 2) {
    return balanceVotes.inverted > balanceVotes.normal ? "inverted" : "normal";
  }

  // Fallback: use category-based heuristics
  // If known expense categories (Groceries, Bills, etc.) have positive amounts, sign is inverted
  let expensePositive = 0, expenseNegative = 0;
  let incomePositive = 0, incomeNegative = 0;
  const expenseCategories = new Set([
    "Groceries", "Eating Out", "Transport", "Shopping", "Subscriptions",
    "Bills & Utilities", "Health & Fitness", "Entertainment", "Education", "Personal Care"
  ]);

  for (const entry of rawEntries) {
    const cat = categorize(entry.description, null);
    if (expenseCategories.has(cat)) {
      if (entry.amount > 0) expensePositive++;
      else if (entry.amount < 0) expenseNegative++;
    } else if (cat === "Income") {
      if (entry.amount > 0) incomePositive++;
      else if (entry.amount < 0) incomeNegative++;
    }
  }

  // If most recognizable expenses are positive, the sign convention is inverted
  if (expensePositive > expenseNegative && expensePositive >= 2) return "inverted";
  if (incomeNegative > incomePositive && incomeNegative >= 2) return "inverted";
  // If most recognizable income is positive, convention is normal
  if (incomePositive > incomeNegative && incomePositive >= 2) return "normal";
  if (expenseNegative > expensePositive && expenseNegative >= 2) return "normal";

  return "normal";
}

function parseExcelSheet(rows, headers, learnedMappings) {
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

  // Track whether we're using separate debit/credit columns
  const hasSeparateDebitCredit = (debitCol >= 0 || creditCol >= 0) && amountCol < 0;

  // First pass: parse raw amounts and metadata
  const rawEntries = [];

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
    if (hasSeparateDebitCredit) {
      // Separate debit/credit columns: debit = money out (expense), credit = money in (income)
      const debit = debitCol >= 0 ? Math.abs(parseAmount(values[debitCol])) : 0;
      const credit = creditCol >= 0 ? Math.abs(parseAmount(values[creditCol])) : 0;
      // Credit (money in) is positive, Debit (money out) is negative
      amount = credit - debit;
    } else if (amountCol >= 0) {
      amount = parseAmount(values[amountCol]);
    }

    const balance = balanceCol >= 0 ? parseAmount(values[balanceCol]) : null;

    rawEntries.push({ index: i, date, description, amount, balance });
  }

  // For single amount column, detect if sign convention is inverted
  // (positive = expense instead of positive = income)
  if (!hasSeparateDebitCredit && amountCol >= 0) {
    const convention = detectSignConvention(rawEntries);
    if (convention === "inverted") {
      for (const entry of rawEntries) {
        entry.amount = -entry.amount;
      }
    }
  }

  // Second pass: build final transactions with corrected amounts
  const transactions = [];
  for (const entry of rawEntries) {
    const category = categorize(entry.description, learnedMappings);

    // Determine transaction type
    let type;
    if (NON_TRANSACTIONAL_CATEGORIES.has(category)) {
      type = "Non-Transactional";
    } else if (category === "Income" || category === "Cash In") {
      type = "Income";
    } else if (category === "Savings") {
      type = "Savings";
    } else if (category === "Cash Withdrawal") {
      type = "Expense";
    } else if (entry.amount > 0) {
      type = "Income";
    } else {
      type = "Expense";
    }

    transactions.push({
      id: `txn_${Date.now()}_${entry.index}`,
      date: entry.date,
      description: entry.description,
      amount: entry.amount,
      absAmount: Math.abs(entry.amount),
      balance: entry.balance,
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
      if (t.type === "Non-Transactional") return; // don't flip non-transactional entries
      t.amount = -t.amount;
      // Recategorize type based on new amount
      if (t.category === "Income" || t.category === "Cash In" || t.amount > 0) {
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

async function parseExcel(buffer) {
  const workbook = XLSX.read(buffer, { type: "buffer", cellDates: true });
  const allTransactions = [];
  const sheetErrors = [];

  // Load learned category mappings for smart auto-categorization
  const learnedMappings = await loadLearnedMappings();

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
      const result = parseExcelSheet(rows, headers, learnedMappings);
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
  // Exclude non-transactional entries (balance carried forward, cash transfers) from totals
  const realTransactions = transactions.filter(t => t.type !== "Non-Transactional");
  const totalIncome = realTransactions.filter(t => t.type === "Income").reduce((s, t) => s + t.absAmount, 0);
  const totalExpenses = realTransactions.filter(t => t.type === "Expense").reduce((s, t) => s + t.absAmount, 0);
  const totalSavings = realTransactions.filter(t => t.type === "Savings").reduce((s, t) => s + t.absAmount, 0);
  const net = totalIncome - totalExpenses - totalSavings;

  // Category breakdown for expenses only (excluding non-transactional)
  const categoryBreakdown = {};
  realTransactions.filter(t => t.type === "Expense").forEach(t => {
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
  const uncatCount = realTransactions.filter(t => t.category === "Uncategorized").length;

  // Non-transactional info
  const nonTxnCount = transactions.filter(t => t.type === "Non-Transactional").length;
  if (nonTxnCount > 0) {
    recommendations.push({
      type: "info",
      title: `${nonTxnCount} non-transactional entr${nonTxnCount !== 1 ? "ies" : "y"} detected`,
      detail: "Balance carried forward and cash transfer entries were automatically excluded from income/expense calculations.",
    });
  }
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
    const transactions = await parseExcel(buffer);

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
