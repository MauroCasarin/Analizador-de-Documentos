import Groq from 'groq-sdk';
const groq = new Groq({apiKey: process.env.AnalizadordeDocumentos});
async function run() {
  const models = ["openai/gpt-oss-120b", "openai/gpt-oss-20b", "groq/compound"];
  // generate large prompt
  const text = "hola ".repeat(15000);
  for (const m of models) {
    try {
      const res = await groq.chat.completions.create({
        model: m,
        messages: [{role: "user", content: text}]
      });
      console.log(m, "SUCCESS");
    } catch (e) {
      console.log(m, "ERROR:", e.message);
    }
  }
}
run().catch(console.error);
