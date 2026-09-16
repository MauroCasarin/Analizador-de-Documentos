import Groq from 'groq-sdk';
const groq = new Groq({apiKey: process.env.AnalizadordeDocumentos});
async function run() {
  const modelsResponse = await groq.models.list();
  const availableModels = modelsResponse.data || [];
  const textModels = availableModels.filter(m => m.input_modalities && m.input_modalities.includes("text") && !m.id.includes("prompt-guard"));
  const preferred = textModels.find(m => m.id.includes("120b") || m.id.includes("70b"));
  console.log("Preferred:", preferred ? preferred.id : null);
  console.log("Text[0]:", textModels[0] ? textModels[0].id : null);
}
run().catch(console.error);
