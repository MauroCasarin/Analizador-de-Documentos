import express from "express";
import path from "path";
import { fileURLToPath } from "url";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI } from "@google/genai";
import dotenv from "dotenv";

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(express.json({ limit: "50mb" }));
  app.use(express.urlencoded({ limit: "50mb", extended: true }));

  // Initialize Gemini Client
  const apiKey = process.env.GEMINI_API_KEY;
  let ai: GoogleGenAI | null = null;
  if (apiKey) {
    ai = new GoogleGenAI({
      apiKey,
      httpOptions: {
        headers: {
          "User-Agent": "aistudio-build",
        },
      },
    });
  }

  // Health check endpoint
  app.get("/api/health", (_req, res) => {
    res.json({
      status: "ok",
      hasApiKey: !!process.env.GEMINI_API_KEY,
    });
  });

  // Document analysis endpoint
  app.post("/api/analyze", async (req, res) => {
    try {
      const activeApiKey = process.env.GEMINI_API_KEY;
      if (!activeApiKey) {
        return res.status(500).json({
          error: "La clave GEMINI_API_KEY no está configurada en las variables de entorno del servidor.",
        });
      }

      if (!ai) {
        ai = new GoogleGenAI({
          apiKey: activeApiKey,
          httpOptions: {
            headers: {
              "User-Agent": "aistudio-build",
            },
          },
        });
      }

      const { file, prompt, fieldsToExtract } = req.body;

      if (!file && !prompt) {
        return res.status(400).json({
          error: "Debes proporcionar un archivo para analizar o una consulta.",
        });
      }

      const parts: any[] = [];

      if (file && file.data) {
        const mimeType = file.mimeType || "application/octet-stream";
        // If it's a plain text/markdown/csv/json file, decode to text
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
              inlineData: {
                mimeType,
                data: file.data,
              },
            });
          }
        } else {
          // PDF or Images
          parts.push({
            inlineData: {
              mimeType,
              data: file.data,
            },
          });
        }
      }

      // Build structured user instruction
      let userQuery = `SOLICITUD DEL USUARIO:\n${prompt || "Extrae toda la información relevante, datos clave y requerimientos del archivo adjunto."}`;

      if (fieldsToExtract && fieldsToExtract.trim()) {
        userQuery += `\n\nCAMPOS ESPECÍFICOS QUE DEBES RECAUDAR:\n${fieldsToExtract.trim()}`;
      }

      parts.push({
        text: userQuery,
      });

      const systemInstruction = `Eres un asistente experto de alto nivel en extracción y análisis minucioso de datos en documentos (PDF, imágenes, planos, contratos y textos).
Tu objetivo es analizar el archivo adjunto junto con la consulta y campos solicitados por el usuario y extraer con la máxima precisión los datos requeridos.

REGLAS DE RESPUESTA:
1. Extrae explícitamente los campos requeridos (ej. Nombre del evento, Fechas/Días, Ubicación/Lugar, Medidas de lote/stand según empresa o número de lote, Horario de armado/desmontaje, Qué tipo de seguro solicitan para el ingreso, Dónde enviar la información que solicitan - mail o teléfono, etc.).
2. Presenta la información extraída estructurada en un formato claro (usa Markdown enriquecido, tablas comparativas y listas con viñetas claras).
3. Si un dato solicitado no está presente en el documento, indica explícitamente: "No especificado en el archivo".
4. Si encuentras detalles relevantes adicionales vinculados a la solicitud, agrégalos en una sección llamada "### 📌 Notas clave".
5. Mantén un tono profesional, preciso, riguroso y directo.
6. Verifica minuciosamente la información para evitar confusiones o mala información.`;

      // Intentar con gemini-flash-latest y fallback a gemini-3.8-flash / gemini-3.1-flash-lite
      const modelsToTry = ["gemini-flash-latest", "gemini-3.8-flash", "gemini-3.1-flash-lite"];
      let lastError: any = null;
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
        } catch (err: any) {
          lastError = err;
          console.warn(`Aviso: Error con modelo ${modelName}, intentando siguiente fallback...`, err?.message);
        }
      }

      if (!resultText) {
        throw lastError || new Error("No se pudo obtener respuesta del modelo.");
      }

      return res.json({
        success: true,
        result: resultText,
        model: usedModel,
      });
    } catch (error: any) {
      console.error("Error al procesar el análisis con Gemini:", error);
      return res.status(500).json({
        error: error?.message || "Ocurrió un error al procesar el documento con la API de Gemini.",
      });
    }
  });

  // Vite middleware for development or static serving for production
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (_req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Servidor iniciado y escuchando en http://0.0.0.0:${PORT}`);
  });
}

startServer().catch((err) => {
  console.error("Error al iniciar el servidor:", err);
});
