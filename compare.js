// Vercel serverless function.
// Keeps your Anthropic API key on the server — it never reaches the browser.
// Set ANTHROPIC_API_KEY as an environment variable in your Vercel project settings.

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const { contacts } = req.body;
  if (!Array.isArray(contacts) || contacts.length === 0) {
    return res.status(400).json({ error: "contacts array is required" });
  }

  const prompt = `Compare each contact's old and new title/company. For each, decide if it changed.
Return ONLY a JSON array, no other text, one object per contact in the same order given:
[{"status": "changed"|"no_change"|"not_found", "resolved_title": string, "resolved_company": string}]
Use "not_found" only if new title and new company are both empty.

Contacts:
${contacts.map((c, idx) => `${idx}. old_title="${c.old_title}", old_company="${c.old_company}", new_title="${c.new_title}", new_company="${c.new_company}"`).join("\n")}`;

  try {
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": process.env.ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: 1000,
        messages: [{ role: "user", content: prompt }],
      }),
    });

    const data = await response.json();
    const text = (data.content || []).map((b) => b.text || "").join("");
    const cleaned = text.replace(/```json|```/g, "").trim();

    let results;
    try {
      results = JSON.parse(cleaned);
    } catch {
      results = contacts.map(() => ({ status: "error", resolved_title: "", resolved_company: "" }));
    }

    res.status(200).json({ results });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
