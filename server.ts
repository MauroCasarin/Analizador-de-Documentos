import express from "express";
import path from "path";
import { fileURLToPath } from "url";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI } from "@google/genai";
import dotenv from "dotenv";
import * as XLSX from "xlsx";
import mammoth from "mammoth";

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
        const fileName = (file.name || "Archivo").toLowerCase();
        const buffer = Buffer.from(file.data, "base64");

        // 1. Detección y procesamiento de hojas de cálculo Excel (.xlsx, .xls, .csv)
        if (
          fileName.endsWith(".xlsx") ||
          fileName.endsWith(".xls") ||
          mimeType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
          mimeType === "application/vnd.ms-excel"
        ) {
          try {
            const workbook = XLSX.read(buffer, { type: "buffer" });
            let sheetText = `--- PLANILLA EXCEL: ${file.name || "Archivo.xlsx"} ---\n`;
            workbook.SheetNames.forEach((sheetName) => {
              const worksheet = workbook.Sheets[sheetName];
              const csvData = XLSX.utils.sheet_to_csv(worksheet);
              sheetText += `\n[HOJA: ${sheetName}]\n${csvData}\n`;
            });
            sheetText += `--- FIN DE PLANILLA EXCEL ---`;
            parts.push({ text: sheetText });
          } catch (e: any) {
            console.warn("Fallo al leer Excel con XLSX:", e?.message);
            parts.push({
              inlineData: { mimeType: "application/octet-stream", data: file.data },
            });
          }
        }
        // 2. Detección y procesamiento de documentos Word (.docx)
        else if (
          fileName.endsWith(".docx") ||
          mimeType === "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ) {
          try {
            const docResult = await mammoth.extractRawText({ buffer });
            const docText = `--- DOCUMENTO WORD: ${file.name || "Documento.docx"} ---\n${docResult.value}\n--- FIN DEL DOCUMENTO WORD ---`;
            parts.push({ text: docText });
          } catch (e: any) {
            console.warn("Fallo al leer Word con mammoth:", e?.message);
            parts.push({
              inlineData: { mimeType: "application/octet-stream", data: file.data },
            });
          }
        }
        // 3. Documentos de texto plano / markdown / csv / json
        else if (
          mimeType.startsWith("text/") ||
          mimeType === "application/json" ||
          mimeType === "text/csv" ||
          mimeType === "text/markdown" ||
          fileName.endsWith(".csv") ||
          fileName.endsWith(".txt") ||
          fileName.endsWith(".json") ||
          fileName.endsWith(".md")
        ) {
          try {
            const textContent = buffer.toString("utf-8");
            parts.push({
              text: `--- DOCUMENTO: ${file.name || "Archivo"} (${mimeType}) ---\n${textContent}\n--- FIN DEL DOCUMENTO ---`,
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
          // 4. PDF o Imágenes (PNG, JPG, WEBP, etc.)
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

      const systemInstruction = `Eres un asistente experto en extracción y análisis de datos en documentos (PDF, Excel, Word, imágenes, planos y textos).
Tu objetivo es extraer con máxima precisión los datos solicitados por el usuario a máxima velocidad de procesamiento.

REGLAS DE RESPUESTA:
1. Sé estrictamente conciso, directo y sin redundancias ni introducciones innecesarias (para optimizar consumo, acelerar el análisis y permitir lectura rápida).
2. Extrae explícitamente los campos requeridos (ej. CUIT, Razón Social / Nombre de empresas, Nombre del evento, Fechas/Días, Ubicación/Lugar, Medidas de lote/stand según empresa o número de lote, Horario de armado/desmontaje, Tipo de seguro exigido para el ingreso, Dónde enviar la información - mail o teléfono, etc.).
3. Presenta la información estructurada con tablas breves o viñetas limpias.
4. IMPORTANTE FORMATO EN ROJO: Si un dato solicitado no está presente o no se encontró en el documento, indícalo explícitamente con la frase exacta: "<span class=\"text-red-600 font-bold bg-red-50 px-1.5 py-0.5 rounded border border-red-200\">No especificado en el archivo</span>" (o "No encontrado").
5. Si encuentras algún dato crítico adicional imprescindible, inclúyelo de forma breve en "### 📌 Notas clave" (máximo 2 a 4 líneas).
6. Verifica minuciosamente la información para evitar confusiones o mala información.`;

      // Modelos optimizados para máxima velocidad de respuesta y precisión
      const modelsToTry = ["gemini-2.5-flash", "gemini-flash-latest", "gemini-3.8-flash", "gemini-3.1-flash-lite"];
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
              temperature: 0.1,
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
