export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).end();

  // בדיקת סיסמה
  const { password, ...bodyWithoutPassword } = req.body;
  if (password !== 'moshe16') {
    return res.status(401).json({ error: 'סיסמה שגויה' });
  }

  // API Key מה-Environment Variable
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: 'API Key לא מוגדר בשרת' });
  }

  try {
    const body = { ...bodyWithoutPassword, stream: false };

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify(body)
    });

    const data = await response.json();
    res.status(response.status).json(data);
  } catch(err) {
    res.status(500).json({ error: err.message });
  }
}
