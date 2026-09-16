import { GoogleGenAI } from "@google/genai";
import * as XLSX from "xlsx";
import mammoth from "mammoth";

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
      const fileName = (file.name || "Archivo").toLowerCase();
      const buffer = Buffer.from(file.data, "base64");

      // 1. Excel (.xlsx, .xls, .csv)
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
        } catch (e) {
          console.warn("Fallo al leer Excel con XLSX:", e?.message);
          parts.push({ inlineData: { mimeType: "application/octet-stream", data: file.data } });
        }
      }
      // 2. Word (.docx)
      else if (
        fileName.endsWith(".docx") ||
        mimeType === "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      ) {
        try {
          const docResult = await mammoth.extractRawText({ buffer });
          const docText = `--- DOCUMENTO WORD: ${file.name || "Documento.docx"} ---\n${docResult.value}\n--- FIN DEL DOCUMENTO WORD ---`;
          parts.push({ text: docText });
        } catch (e) {
          console.warn("Fallo al leer Word con mammoth:", e?.message);
          parts.push({ inlineData: { mimeType: "application/octet-stream", data: file.data } });
        }
      }
      // 3. Documentos de texto
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
          parts.push({ inlineData: { mimeType, data: file.data } });
        }
      } else {
        // 4. PDF o imágenes
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
            temperature: 0.1,
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
