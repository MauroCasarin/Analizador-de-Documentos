import { GoogleGenAI } from "@google/genai";
import * as XLSX from "xlsx";
import mammoth from "mammoth";
import { PDFParse } from "pdf-parse";

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
          extractedDocText = sheetText;
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
          extractedDocText = docResult.value;
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
          extractedDocText = textContent;
          parts.push({
            text: `--- DOCUMENTO: ${file.name || "Archivo"} (${mimeType}) ---\n${textContent}\n--- FIN DEL DOCUMENTO ---`,
          });
        } catch {
          parts.push({ inlineData: { mimeType, data: file.data } });
        }
      }
      // 4. PDF (extracción digital multipágina + visión multimodal)
      else if (fileName.endsWith(".pdf") || mimeType === "application/pdf") {
        try {
          const parser = new PDFParse({ data: buffer });
          const pdfData = await parser.getText();
          if (pdfData && pdfData.text && pdfData.text.trim()) {
            extractedDocText = pdfData.text;
            parts.push({
              text: `--- TEXTO DIGITAL EXTRAÍDO DE TODAS LAS PÁGINAS DEL PDF (${file.name || "documento.pdf"} - ${pdfData.total || 1} páginas) ---\n${pdfData.text}\n--- FIN DEL TEXTO DIGITAL DEL PDF ---`,
            });
          }
        } catch (pdfErr) {
          console.warn("Aviso al extraer texto con pdf-parse:", pdfErr?.message);
        }

        parts.push({
          inlineData: { mimeType: "application/pdf", data: file.data },
        });
      } else {
        // 5. Imágenes u otros formatos
        parts.push({
          inlineData: { mimeType, data: file.data },
        });
      }
    }

    // Parser de CUITs y empresas (soporta copiado directo de Excel con comillas y tabulaciones)
    function extractCuitItemsFromText(input) {
      if (!input) return [];
      const lines = input.split(/[\r\n]+/);
      const list = [];
      const seen = new Set();

      for (const rawLine of lines) {
        const line = rawLine.trim();
        if (!line) continue;

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

            let company = line
              .replace(cuitMatch[0], "")
              .replace(/["'“”«»]/g, "")
              .replace(/^[\s,;\t\-:]+|[\s,;\t\-:]+$/g, "")
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

    const userRawInput = `${fieldsToExtract || ""} \n ${prompt || ""}`;
    const cuitItems = extractCuitItemsFromText(userRawInput);

    let verificationAuditText = "";
    if (extractedDocText && cuitItems.length > 0) {
      const docLines = extractedDocText.split(/\r?\n/);
      
      for (const item of cuitItems) {
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

    let userQuery = `SOLICITUD DEL USUARIO:\n${prompt || "Extrae toda la información relevante, datos clave y requerimientos del archivo adjunto."}`;

    if (fieldsToExtract && fieldsToExtract.trim()) {
      userQuery += `\n\nDATOS O CAMPOS INGRESADOS POR EL USUARIO (Desde Excel / formulario):\n${fieldsToExtract.trim()}`;
    }

    if (verificationAuditText) {
      userQuery += `\n${verificationAuditText}\n\nIMPORTANTE PARA LA TABLA:\nPara cada uno de los CUITs donde se indicó "¡SÍ FIGURA EN EL DOCUMENTO!", debes responder obligatoriamente "**Sí**" en la columna "¿Figura en el documento?" y detallar la coincidencia hallada. Para los que indiquen "NO figura", debes responder con: <span class="text-red-600 font-bold bg-red-50 px-1.5 py-0.5 rounded border border-red-200">No figura en el documento</span>.`;
    }

    parts.push({ text: userQuery });

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

    // Modelos soportados según SDK @google/genai, priorizando disponibilidad y precisión
    const modelsToTry = [
      "gemini-3.8-flash",
      "gemini-flash-latest",
      "gemini-3.1-flash-lite",
    ];
    let lastError = null;
    let resultText = "";
    let usedModel = "";

    for (const modelName of modelsToTry) {
      // Reintentos automáticos para mitigar picos temporales de demanda (503 / 429)
      for (let attempt = 0; attempt < 2; attempt++) {
        try {
          if (attempt > 0) {
            await new Promise((resolve) => setTimeout(resolve, 1000));
          }
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
          console.warn(`Aviso: Error con modelo ${modelName} (intento ${attempt + 1}):`, err?.message);
          // Si el modelo da 404 (no existe), no reintentar ese modelo
          if (err?.status === 404 || err?.message?.includes("404")) {
            break;
          }
        }
      }
      if (resultText) break;
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
