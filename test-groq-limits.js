import Groq from 'groq-sdk';
const groq = new Groq({apiKey: process.env.AnalizadordeDocumentos});
async function run() {
  const models = ["openai/gpt-oss-120b", "openai/gpt-oss-20b", "qwen/qwen3.8-27b", "groq/compound"];
  for (const m of models) {
    try {
      const res = await groq.chat.completions.create({
        model: m,
        messages: [{role: "user", content: "test"}]
      });
      console.log(m, "SUCCESS");
    } catch (e) {
      console.log(m, "ERROR:", e.message);
    }
  }
}
run().catch(console.error);
