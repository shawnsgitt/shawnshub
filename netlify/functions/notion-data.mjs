export default async (req, context) => {
  const NOTION_TOKEN = Netlify.env.get("NOTION_TOKEN");
  const SPENDING_DB = Netlify.env.get("NOTION_SPENDING_DB");
  const GOALS_DB = Netlify.env.get("NOTION_GOALS_DB");

  if (!NOTION_TOKEN || !SPENDING_DB || !GOALS_DB) {
    return new Response(JSON.stringify({ error: "Missing environment variables" }), {
      status: 500,
      headers: { "Content-Type": "application/json" },
    });
  }

  const headers = {
    Authorization: "Bearer " + NOTION_TOKEN,
    "Notion-Version": "2022-06-28",
    "Content-Type": "application/json",
  };

  async function queryDB(dbId) {
    const res = await fetch(
      "https://api.notion.com/v1/databases/" + dbId + "/query",
      { method: "POST", headers: headers, body: JSON.stringify({ page_size: 100 }) }
    );
    if (!res.ok) {
      const err = await res.text();
      throw new Error("Notion API error " + res.status + ": " + err);
    }
    return res.json();
  }

  function extractProp(prop) {
    if (!prop) return null;
    if (prop.type === "title") return (prop.title || []).map(function(t) { return t.plain_text; }).join("");
    if (prop.type === "rich_text") return (prop.rich_text || []).map(function(t) { return t.plain_text; }).join("");
    if (prop.type === "number") return prop.number;
    if (prop.type === "select") return prop.select ? prop.select.name : null;
    if (prop.type === "checkbox") return prop.checkbox;
    return null;
  }

  function parsePages(results) {
    return results.map(function(page) {
      var parsed = { id: page.id };
      var keys = Object.keys(page.properties);
      for (var i = 0; i < keys.length; i++) {
        parsed[keys[i]] = extractProp(page.properties[keys[i]]);
      }
      return parsed;
    });
  }

  try {
    var results = await Promise.all([queryDB(SPENDING_DB), queryDB(GOALS_DB)]);
    var spending = parsePages(results[0].results);
    var goals = parsePages(results[1].results);

    return new Response(JSON.stringify({ spending: spending, goals: goals }), {
      status: 200,
      headers: {
        "Content-Type": "application/json",
        "Cache-Control": "public, max-age=60",
      },
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: err.message }), {
      status: 500,
      headers: { "Content-Type": "application/json" },
    });
  }
};

export const config = {
  path: "/api/finance",
};
