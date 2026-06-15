export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const { question } = req.body || {};
  if (!question || typeof question !== "string" || !question.trim()) {
    return res.status(400).json({ error: "Pitanje je obavezno." });
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: "Server nije konfiguriran (nedostaje API ključ)." });
  }

  try {
    const anthropicRes = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: 1000,
        system: "Ti si stručnjak za hrvatski mirovinski sustav. Odgovaraj na pitanja korisnika jasno, točno i sažeto na hrvatskom jeziku. Citiraj relevantne zakone (ZOMO, ZOMF) i službene izvore (HZMO, HANFA, mss.gov.hr). AVM od 1.7.2025. = 14,45 EUR. Nikad ne izmišljaj podatke — ako ne znaš, reci to jasno.",
        messages: [{ role: "user", content: question.slice(0, 2000) }],
      }),
    });

    const data = await anthropicRes.json();

    if (!anthropicRes.ok) {
      return res.status(anthropicRes.status).json({ error: data.error?.message || "Greška API-ja." });
    }

    return res.status(200).json({ answer: data.content?.[0]?.text || "Greška u odgovoru." });
  } catch (e) {
    return res.status(502).json({ error: "Nije moguće dohvatiti odgovor s AI servisa." });
  }
}
