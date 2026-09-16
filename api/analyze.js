import { GoogleGenAI } from "@google/genai";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Método no permitido" });
  }

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    return res.status(500).json({
      error: "GEMINI_API_KEY no está configurada en las variables de entorno de Vercel.",
    });
  }

  try {
    const ai = new GoogleGenAI({ apiKey });
    const { file, prompt, fieldsToExtract } = req.body;

    const parts = [];

    if (file && file.data) {
      const mimeType = file.mimeType || "application/octet-stream";
      if (
        mimeType.startsWith("text/") ||
        mimeType === "application/json" ||
        mimeType === "text/csv" ||
        mimeType === "text/markdown"
      ) {
        try {
          const textContent = Buffer.from(file.data, "base64").toString("utf-8");
          parts.push({
            text: `--- DOCUMENTO ADJUNTO: ${file.name || "Archivo"} (${mimeType}) ---\n${textContent}\n--- FIN DEL DOCUMENTO ---`,
          });
        } catch {
          parts.push({
            inlineData: { mimeType, data: file.data },
          });
        }
      } else {
        parts.push({
          inlineData: { mimeType, data: file.data },
        });
      }
    }

    let userQuery = `SOLICITUD DEL USUARIO:\n${prompt || "Extrae toda la información relevante, datos clave y requerimientos del archivo adjunto."}`;

    if (fieldsToExtract && fieldsToExtract.trim()) {
      userQuery += `\n\nCAMPOS ESPECÍFICOS QUE DEBES RECAUDAR:\n${fieldsToExtract.trim()}`;
    }

    parts.push({ text: userQuery });

    const systemInstruction = `Eres un asistente experto en extracción y análisis de datos en documentos (PDF, imágenes, planos y textos).
Tu objetivo es extraer con máxima precisión los datos solicitados por el usuario.

REGLAS DE RESPUESTA:
1. Sé estrictamente conciso, directo y sin redundancias ni introducciones innecesarias (para optimizar consumo y lectura rápida).
2. Extrae explícitamente los campos requeridos (ej. Nombre del evento, Fechas/Días, Ubicación/Lugar, Medidas de lote/stand según empresa o número de lote, Horario de armado/desmontaje, Tipo de seguro exigido para el ingreso, Dónde enviar la información - mail o teléfono, etc.).
3. Presenta la información estructurada con tablas breves o viñetas limpias.
4. Si un dato solicitado no está presente en el documento, indica explícitamente: "No especificado en el archivo".
5. Si encuentras algún dato crítico adicional imprescindible, inclúyelo de forma breve en "### 📌 Notas clave" (máximo 2 a 4 líneas).
6. Verifica minuciosamente la información para evitar confusiones o mala información.`;

    // Intentar con gemini-flash-latest y fallback a gemini-3.8-flash / gemini-3.1-flash-lite
    const modelsToTry = ["gemini-flash-latest", "gemini-3.8-flash", "gemini-3.1-flash-lite"];
    let lastError = null;
    let resultText = "";
    let usedModel = "";

    for (const modelName of modelsToTry) {
      try {
        const response = await ai.models.generateContent({
          model: modelName,
          contents: { parts },
          config: {
            systemInstruction,
            temperature: 0.2,
          },
        });

        if (response && response.text) {
          resultText = response.text;
          usedModel = modelName;
          break;
        }
      } catch (err) {
        lastError = err;
        console.warn(`Aviso: Error con modelo ${modelName}, intentando siguiente fallback...`, err?.message);
      }
    }

    if (!resultText) {
      throw lastError || new Error("No se pudo obtener respuesta del modelo.");
    }

    return res.status(200).json({
      success: true,
      result: resultText,
      model: usedModel,
    });
  } catch (error) {
    console.error("Error en endpoint Vercel:", error);
    return res.status(500).json({
      error: error?.message || "Error al procesar con la API de Gemini.",
    });
  }
}
