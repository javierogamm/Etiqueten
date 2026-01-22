const csv1Input = document.getElementById("csv1");
const csv2Input = document.getElementById("csv2");
const csv1Status = document.getElementById("csv1Status");
const csv2Status = document.getElementById("csv2Status");
const csv1Summary = document.getElementById("csv1Summary");
const csv2Summary = document.getElementById("csv2Summary");
const resultSummary = document.getElementById("resultSummary");
const filterBtn = document.getElementById("filterBtn");
const downloadBtn = document.getElementById("downloadDoc");
const previewBtn = document.getElementById("previewBtn");
const previewSection = document.getElementById("previewSection");
const labelsPreview = document.getElementById("labelsPreview");

const state = {
  csv1: null,
  csv2: null,
  matches: [],
  headers: [],
};

const appVersion = "0.1.1";
const versionLabel = document.getElementById("appVersion");
if (versionLabel) {
  versionLabel.textContent = appVersion;
}

const readFile = (file) =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("No se pudo leer el archivo"));
    reader.readAsText(file, "utf-8");
  });

const parseCSV = (text) => {
  const rows = [];
  let current = "";
  let row = [];
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];

    if (char === '"') {
      if (inQuotes && next === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (char === "," && !inQuotes) {
      row.push(current.trim());
      current = "";
      continue;
    }

    if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") {
        i += 1;
      }
      row.push(current.trim());
      if (row.some((cell) => cell.length > 0)) {
        rows.push(row);
      }
      row = [];
      current = "";
      continue;
    }

    current += char;
  }

  if (current.length > 0 || row.length > 0) {
    row.push(current.trim());
    if (row.some((cell) => cell.length > 0)) {
      rows.push(row);
    }
  }

  const headers = rows.shift() || [];
  const data = rows.filter((r) => r.length > 0);
  return { headers, rows: data };
};

const detectNifColumn = (headers) => {
  const normalized = headers.map((header) => header.toLowerCase().trim());
  const index = normalized.findIndex((header) => header === "nif" || header.includes("nif"));
  return index >= 0 ? index : null;
};

const updateSummary = (element, rows, headers, nifIndex) => {
  element.innerHTML = `
    <li>Filas: ${rows.length}</li>
    <li>Columna NIF: ${nifIndex !== null ? headers[nifIndex] : "No detectada"}</li>
  `;
};

const updateResultSummary = (matches, headers) => {
  resultSummary.innerHTML = `
    <li>Coincidencias: ${matches.length}</li>
    <li>Columnas disponibles: ${headers.join(", ") || "—"}</li>
  `;
};

const setStatus = (element, text, isReady) => {
  element.textContent = text;
  element.classList.toggle("ready", Boolean(isReady));
};

const validateReady = () => {
  filterBtn.disabled = !(state.csv1 && state.csv2);
};

const handleCSVUpload = async (file, target) => {
  try {
    const text = await readFile(file);
    const parsed = parseCSV(text);
    const nifIndex = detectNifColumn(parsed.headers);

    if (nifIndex === null) {
      throw new Error("No se ha encontrado una columna con NIF.");
    }

    return { ...parsed, nifIndex, name: file.name };
  } catch (error) {
    setStatus(target, error.message, false);
    return null;
  }
};

csv1Input.addEventListener("change", async (event) => {
  const [file] = event.target.files;
  if (!file) return;

  setStatus(csv1Status, `Cargando ${file.name}...`, false);
  const parsed = await handleCSVUpload(file, csv1Status);
  if (!parsed) return;

  state.csv1 = parsed;
  setStatus(csv1Status, `CSV1 cargado: ${parsed.name}`, true);
  updateSummary(csv1Summary, parsed.rows, parsed.headers, parsed.nifIndex);
  validateReady();
});

csv2Input.addEventListener("change", async (event) => {
  const [file] = event.target.files;
  if (!file) return;

  setStatus(csv2Status, `Cargando ${file.name}...`, false);
  const parsed = await handleCSVUpload(file, csv2Status);
  if (!parsed) return;

  state.csv2 = parsed;
  setStatus(csv2Status, `CSV2 cargado: ${parsed.name}`, true);
  updateSummary(csv2Summary, parsed.rows, parsed.headers, parsed.nifIndex);
  validateReady();
});

const buildMatches = () => {
  if (!state.csv1 || !state.csv2) return [];
  const nifSet = new Set(
    state.csv1.rows
      .map((row) => row[state.csv1.nifIndex])
      .filter((value) => value && value.length > 0)
      .map((value) => value.toUpperCase())
  );

  return state.csv2.rows.filter((row) => {
    const nifValue = row[state.csv2.nifIndex];
    if (!nifValue) return false;
    return nifSet.has(nifValue.toUpperCase());
  });
};

const buildLabelHTML = (headers, rows) => {
  const labels = rows
    .map((row) => {
      const lines = headers
        .map((header, index) => {
          const value = row[index] || "";
          if (!value) return null;
          return `<div><strong>${header}:</strong> ${value}</div>`;
        })
        .filter(Boolean)
        .join("");
      return `<div class="label">${lines}</div>`;
    })
    .join("");

  return `
    <div class="sheet">
      <div class="label-grid">
        ${labels}
      </div>
    </div>
  `;
};

const buildWordDocument = (headers, rows) => {
  const labels = rows
    .map((row) => {
      const lines = headers
        .map((header, index) => {
          const value = row[index] || "";
          if (!value) return "";
          return `<div><strong>${header}:</strong> ${value}</div>`;
        })
        .join("");
      return `<div class="label">${lines}</div>`;
    })
    .join("");

  return `
  <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
    <head>
      <meta charset="utf-8">
      <title>Etiquetas Apli 01273</title>
      <style>
        @page { size: A4; margin: 0mm; }
        body { font-family: Arial, sans-serif; font-size: 10pt; margin: 0; }
        .label-grid {
          display: grid;
          grid-template-columns: repeat(3, 70mm);
          grid-auto-rows: 37mm;
          gap: 0;
        }
        .label {
          border: 0.2mm solid #999;
          padding: 3mm 3mm;
          box-sizing: border-box;
          display: flex;
          flex-direction: column;
          justify-content: center;
          overflow: hidden;
        }
        .label div { line-height: 1.2; }
      </style>
    </head>
    <body>
      <div class="label-grid">
        ${labels}
      </div>
    </body>
  </html>
  `;
};

filterBtn.addEventListener("click", () => {
  const matches = buildMatches();
  state.matches = matches;
  state.headers = state.csv2 ? state.csv2.headers : [];
  updateResultSummary(matches, state.headers);
  downloadBtn.disabled = matches.length === 0;
  previewBtn.disabled = matches.length === 0;
  previewSection.hidden = matches.length === 0;
});

previewBtn.addEventListener("click", () => {
  if (!state.matches.length) return;
  labelsPreview.innerHTML = buildLabelHTML(state.headers, state.matches);
  previewSection.hidden = false;
  previewSection.scrollIntoView({ behavior: "smooth", block: "start" });
});

downloadBtn.addEventListener("click", () => {
  if (!state.matches.length) return;
  const html = buildWordDocument(state.headers, state.matches);
  const blob = new Blob([html], { type: "application/msword" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "etiquetas-apli-01273.doc";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
});
