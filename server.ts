import express from "express";
import path from "path";
import { fileURLToPath } from "url";
import { createServer as createViteServer } from "vite";
import Groq from "groq-sdk";
import dotenv from "dotenv";
import * as XLSX from "xlsx";
import mammoth from "mammoth";
import { extractText } from "unpdf";

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(express.json({ limit: "50mb" }));
  app.use(express.urlencoded({ limit: "50mb", extended: true }));

  // Initialize Groq Client
  const apiKey = process.env.GROQ_API_KEY || process.env.AnalizadordeDocumentos;
  let groq: Groq | null = null;
  if (apiKey) {
    groq = new Groq({ apiKey });
  }

  // Health check endpoint
  app.get("/api/health", (_req, res) => {
    res.json({
      status: "ok",
      hasApiKey: !!apiKey,
    });
  });

  // Document analysis endpoint
  app.post("/api/analyze", async (req, res) => {
    try {
      const activeApiKey = process.env.GROQ_API_KEY || process.env.AnalizadordeDocumentos;
      if (!activeApiKey) {
        return res.status(500).json({
          error: "La clave de API de Groq no está configurada en las variables de entorno del servidor.",
        });
      }

      if (!groq) {
        groq = new Groq({ apiKey: activeApiKey });
      }

      const { file, prompt, fieldsToExtract } = req.body;

      if (!file && !prompt) {
        return res.status(400).json({
          error: "Debes proporcionar un archivo para analizar o una consulta.",
        });
      }

      const parts: any[] = [];
      let extractedDocText = "";

      if (file && file.data) {
        const fileName = (file.name || "Archivo").toLowerCase();
        let mimeType = file.mimeType || file.type || "";
        
        // Corrección estricta de MIME type para evitar 'application/octet-stream'
        if (!mimeType || mimeType === "application/octet-stream") {
          if (fileName.endsWith(".pdf")) mimeType = "application/pdf";
          else if (fileName.endsWith(".xlsx")) mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
          else if (fileName.endsWith(".xls")) mimeType = "application/vnd.ms-excel";
          else if (fileName.endsWith(".docx")) mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
          else if (fileName.endsWith(".png")) mimeType = "image/png";
          else if (fileName.endsWith(".jpg") || fileName.endsWith(".jpeg")) mimeType = "image/jpeg";
          else if (fileName.endsWith(".webp")) mimeType = "image/webp";
          else if (fileName.endsWith(".csv")) mimeType = "text/csv";
          else if (fileName.endsWith(".txt")) mimeType = "text/plain";
          else mimeType = "application/pdf";
        }

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
            extractedDocText = sheetText;
            parts.push({ text: sheetText });
          } catch (e: any) {
            console.warn("Fallo al leer Excel con XLSX:", e?.message);
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
            extractedDocText = docResult.value;
            parts.push({ text: docText });
          } catch (e: any) {
            console.warn("Fallo al leer Word con mammoth:", e?.message);
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
            extractedDocText = textContent;
            parts.push({
              text: `--- DOCUMENTO: ${file.name || "Archivo"} (${mimeType}) ---\n${textContent}\n--- FIN DEL DOCUMENTO ---`,
            });
          } catch {
            // Ignorar
          }
        }
        // 4. Documentos PDF (extracción de texto digital multipágina con unpdf)
        else if (fileName.endsWith(".pdf") || mimeType === "application/pdf") {
          try {
            const uint8 = new Uint8Array(buffer);
            const pdfData = await extractText(uint8, { mergePages: true });
            if (pdfData && pdfData.text && typeof pdfData.text === "string" && pdfData.text.trim()) {
              extractedDocText = pdfData.text;
              parts.push({
                text: `--- TEXTO DIGITAL EXTRAÍDO DE TODAS LAS PÁGINAS DEL PDF (${file.name || "documento.pdf"} - ${pdfData.totalPages || 1} páginas) ---\n${pdfData.text}\n--- FIN DEL TEXTO DIGITAL DEL PDF ---`,
              });
            }
          } catch (pdfErr: any) {
            console.warn("Aviso al extraer texto del PDF:", pdfErr?.message);
          }
        } else if (mimeType.startsWith("image/")) {
          // 5. Imágenes (PNG, JPG, WEBP, etc.)
          parts.push({
            inlineData: {
              mimeType,
              data: file.data,
            },
          });
        }
      }

      // Parser inteligente de CUITs y empresas (soporta copiado directo de Excel con comillas, tabulaciones y saltos de línea)
      interface CuitItem {
        cuit11: string;
        cuitFormatted: string;
        company: string;
        found?: boolean;
        matchedSnippet?: string;
      }

      function extractCuitItemsFromText(input: string): CuitItem[] {
        if (!input) return [];
        const lines = input.split(/[\r\n]+/);
        const list: CuitItem[] = [];
        const seen = new Set<string>();

        for (const rawLine of lines) {
          const line = rawLine.trim();
          if (!line) continue;

          // Buscar CUIT argentino de 11 dígitos (con o sin guiones/espacios)
          const cuitMatch =
            line.match(/\b(\d{2})[-.\s]?(\d{8})[-.\s]?(\d{1})\b/) ||
            line.match(/\b(\d{11})\b/);

          if (cuitMatch) {
            let cuit11 = "";
            let cuitFormatted = "";
            if (cuitMatch.length >= 4 && cuitMatch[1] && cuitMatch[2] && cuitMatch[3]) {
              cuit11 = `${cuitMatch[1]}${cuitMatch[2]}${cuitMatch[3]}`;
              cuitFormatted = `${cuitMatch[1]}-${cuitMatch[2]}-${cuitMatch[3]}`;
            } else {
              const d = cuitMatch[1] || cuitMatch[0].replace(/[^0-9]/g, "");
              if (d.length === 11) {
                cuit11 = d;
                cuitFormatted = `${d.slice(0, 2)}-${d.slice(2, 10)}-${d.slice(10, 11)}`;
              }
            }

            if (cuit11 && cuit11.length === 11 && !seen.has(cuit11)) {
              seen.add(cuit11);

              // Extraer la razón social eliminando el CUIT y todas las comillas generadas por Excel
              let company = line
                .replace(cuitMatch[0], "")
                .replace(/["'“”«»]/g, "") // limpiar comillas de Excel
                .replace(/^[\s,;\t\-:]+|[\s,;\t\-:]+$/g, "") // limpiar delimitadores externos
                .replace(/\s+/g, " ")
                .trim();

              list.push({
                cuit11,
                cuitFormatted,
                company: company || "(Razón Social no indicada)",
              });
            }
          }
        }

        // Si no se encontraron por salto de línea, intentar buscar por comas o punto y coma
        if (list.length === 0) {
          const parts = input.split(/[,;]+/);
          for (const part of parts) {
            const p = part.trim().replace(/["'“”«»]/g, "");
            const m =
              p.match(/\b(\d{2})[-.\s]?(\d{8})[-.\s]?(\d{1})\b/) ||
              p.match(/\b(\d{11})\b/);
            if (m) {
              const d = m[0].replace(/[^0-9]/g, "");
              if (d.length === 11 && !seen.has(d)) {
                seen.add(d);
                const formatted = `${d.slice(0, 2)}-${d.slice(2, 10)}-${d.slice(10, 11)}`;
                const company = p.replace(m[0], "").trim();
                list.push({
                  cuit11: d,
                  cuitFormatted: formatted,
                  company: company || "(Razón Social no indicada)",
                });
              }
            }
          }
        }

        return list;
      }

      // Cotejar los CUITs ingresados contra el texto digital del documento si está disponible
      const userRawInput = `${fieldsToExtract || ""} \n ${prompt || ""}`;
      const cuitItems = extractCuitItemsFromText(userRawInput);

      let verificationAuditText = "";
      if (extractedDocText && cuitItems.length > 0) {
        const docLines = extractedDocText.split(/\r?\n/);
        
        for (const item of cuitItems) {
          // Búsqueda flexible de los 11 dígitos numéricos
          const regexFlexible = new RegExp(item.cuit11.split("").join("[-.\\s]?"), "i");
          let matchedLine = "";

          for (const line of docLines) {
            const l = line.trim();
            if (!l) continue;
            const digitsOnly = l.replace(/[^0-9]/g, "");
            if (digitsOnly.includes(item.cuit11) || regexFlexible.test(l) || l.includes(item.cuitFormatted)) {
              matchedLine = l;
              break;
            }
          }

          // Si aún no se encontró por CUIT directo, cotejar si la razón social figura asociada a un CUIT
          if (!matchedLine && item.company && item.company !== "(Razón Social no indicada)") {
            const cleanName = item.company
              .toUpperCase()
              .replace(/[^A-Z0-9\s]/g, "")
              .replace(/\b(SA|SRL|SAS|SOCIEDAD|ANONIMA|DE|Y|EN|DEL|LA|LOS|GRUPO|GROUP)\b/g, "")
              .trim();
            const words = cleanName.split(/\s+/).filter((w) => w.length >= 4);
            if (words.length > 0) {
              for (const line of docLines) {
                const upperLine = line.toUpperCase();
                const matchedWords = words.filter((w) => upperLine.includes(w));
                if (matchedWords.length >= Math.min(2, words.length)) {
                  matchedLine = line.trim();
                  break;
                }
              }
            }
          }

          if (matchedLine) {
            item.found = true;
            item.matchedSnippet = matchedLine;
          } else {
            item.found = false;
          }
        }

        verificationAuditText = `\n\n### RESULTADO DEL COTEJO TÉCNICO EXACTO EN EL DOCUMENTO (Texto digital de las páginas):\n`;
        for (const item of cuitItems) {
          if (item.found) {
            verificationAuditText += `- CUIT ${item.cuitFormatted} (${item.cuit11}) | Empresa: "${item.company}" -> ¡SÍ FIGURA EN EL DOCUMENTO! Coincidencia hallada: "${item.matchedSnippet}"\n`;
          } else {
            verificationAuditText += `- CUIT ${item.cuitFormatted} (${item.cuit11}) | Empresa: "${item.company}" -> NO figura en el texto del documento.\n`;
          }
        }
      }

      // Build structured user instruction
      let userQuery = `SOLICITUD DEL USUARIO:\n${prompt || "Extrae toda la información relevante, datos clave y requerimientos del archivo adjunto."}`;

      if (fieldsToExtract && fieldsToExtract.trim()) {
        userQuery += `\n\nDATOS O CAMPOS INGRESADOS POR EL USUARIO (Desde Excel / formulario):\n${fieldsToExtract.trim()}`;
      }

      if (verificationAuditText) {
        userQuery += `\n${verificationAuditText}\n\nIMPORTANTE PARA LA TABLA:\nPara cada uno de los CUITs donde se indicó "¡SÍ FIGURA EN EL DOCUMENTO!", debes responder obligatoriamente "**Sí**" en la columna "¿Figura en el documento?" y detallar la coincidencia hallada. Para los que indiquen "NO figura", debes responder con: <span class="text-red-600 font-bold bg-red-50 px-1.5 py-0.5 rounded border border-red-200">No figura en el documento</span>.`;
      }

      parts.push({
        text: userQuery,
      });

      const systemInstruction = `Eres un asistente experto en extracción y análisis de datos en documentos (PDF, Excel, Word, imágenes, planos y textos).
Tu objetivo es extraer y verificar con máxima precisión los datos solicitados por el usuario.

REGLAS GENERALES:
1. Sé conciso, directo y sin introducciones innecesarias.
2. Presenta los resultados organizados en tablas legibles o listas ordenadas.
3. Si un dato solicitado no está presente en el documento, indícalo explícitamente como "No figura en el documento" o "No encontrado".

REGLAS DE VERIFICACIÓN DE CUIT / CUIL Y EMPRESAS (RAZÓN SOCIAL):
- Si la solicitud pide verificar CUIT / CUIL y Razón Social / Empresas:
  A. REVISIÓN EXHAUSTIVA: Analiza minuciosamente TODO el contenido del documento adjunto (tablas, columnas, párrafos, notas al pie, encabezados, etc.).
  B. NORMALIZACIÓN DE CUIT: Los CUITs en Argentina constan de 11 dígitos. En el documento pueden aparecer con guiones (ej. 30-12345678-9), sin guiones (30123456789), con espacios o con puntos. Considera coincidencia si los 11 dígitos numéricos coinciden, independientemente del formato de puntuación.
  C. NORMALIZACIÓN DE EMPRESA: Coincidencia válida si el nombre o razón social coincide, independientemente de variaciones societarias (S.A., SA, S.R.L., SRL, S.A.S., etc.), mayúsculas/minúsculas o nombres de fantasía.
  D. TABLA DE RESULTADOS:
     Presenta una tabla clara con las columnas:
     | CUIT / CUIL | Razón Social / Empresa | ¿Figura en el documento? | Coincidencia y Ubicación en el documento |
     - Si figura: En la columna "¿Figura en el documento?", indica **Sí**, y en la columna de coincidencia indica exactamente el texto encontrado y la página o sección.
     - Si NO figura: En la columna "¿Figura en el documento?", indica <span class="text-red-600 font-bold bg-red-50 px-1.5 py-0.5 rounded border border-red-200">No figura en el documento</span>.
  E. SI EL USUARIO NO ESPECIFICÓ UNA LISTA PREVIA DE CUITS A BUSCAR:
     No digas que no figuran. Extrae y lista TODOS los CUIT / CUIL y Empresas / Razones Sociales que efectivamente figuren dentro del documento adjunto indicando "Sí" para cada uno.
  F. En este modo de verificación, NO incluyas información de stands, módulos, medidas, horarios ni datos secundarios ajenos a CUIT y Razón Social.`;

      // Preparar los mensajes para Groq
      let hasImage = false;
      const contentArray: any[] = [];

      for (const part of parts) {
        if (part.text) {
          contentArray.push({ type: "text", text: part.text });
        } else if (part.inlineData) {
          hasImage = true;
          contentArray.push({
            type: "image_url",
            image_url: {
              url: `data:${part.inlineData.mimeType};base64,${part.inlineData.data}`
            }
          });
        }
      }

      let modelName = "llama-3.3-70b-versatile";
      try {
        const modelsResponse = await groq.models.list();
        const availableModels = modelsResponse.data || [];
        if (hasImage) {
          const visionModels = availableModels.filter((m: any) => m.input_modalities?.includes("image"));
          modelName = visionModels.length > 0 ? visionModels[0].id : "llama-3.2-90b-vision-preview";
        } else {
          const textModels = availableModels.filter((m: any) => m.input_modalities?.includes("text") && !m.id.includes("prompt-guard"));
          const preferred = textModels.find((m: any) => m.id.includes("120b") || m.id.includes("70b"));
          modelName = preferred ? preferred.id : (textModels[0]?.id || "llama-3.3-70b-versatile");
        }
      } catch (e) {
        console.error("GROQ LIST ERROR", e); modelName = hasImage ? "llama-3.2-90b-vision-preview" : "llama-3.3-70b-versatile";
      }

      const response = await groq.chat.completions.create({
        model: modelName,
        messages: [
          { role: "system", content: systemInstruction },
          { role: "user", content: contentArray }
        ],
        temperature: 0.1,
      });

      const resultText = response.choices[0]?.message?.content;

      if (!resultText) {
        throw new Error("No se pudo obtener respuesta del modelo.");
      }

      return res.json({
        success: true,
        result: resultText,
        model: modelName,
      });
    } catch (error: any) {
      console.error("Error al procesar el análisis con Groq:", error);
      return res.status(500).json({
        error: error?.message || "Ocurrió un error al procesar el documento con la API de Groq.",
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

