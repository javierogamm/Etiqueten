const csvPrimaryInput = document.getElementById("csvPrimary");
const csvSecondaryInput = document.getElementById("csvSecondary");
const csvPrimaryStatus = document.getElementById("csvPrimaryStatus");
const csvSecondaryStatus = document.getElementById("csvSecondaryStatus");
const primaryTable = document.getElementById("primaryTable");
const primaryData = document.getElementById("primaryData");
const primaryCount = document.getElementById("primaryCount");
const openModalBtn = document.getElementById("openModal");
const labelsModal = document.getElementById("labelsModal");
const closeModalBtn = document.getElementById("closeModal");
const matchSummary = document.getElementById("matchSummary");
const matchCount = document.getElementById("matchCount");
const mismatchCount = document.getElementById("mismatchCount");
const matchesTable = document.getElementById("matchesTable");
const mismatchesTable = document.getElementById("mismatchesTable");
const matchResults = document.getElementById("matchResults");
const generateWordBtn = document.getElementById("generateWord");
const downloadPreviewBtn = document.getElementById("downloadPreview");
const previewSection = document.getElementById("previewSection");
const labelsPreview = document.getElementById("labelsPreview");

const state = {
  primary: null,
  secondary: null,
  matches: [],
  mismatches: [],
  labelRows: [],
};

const appVersion = "0.3.0";
const versionLabel = document.getElementById("appVersion");
if (versionLabel) {
  versionLabel.textContent = appVersion;
}

const requiredPrimaryHeaders = [
  "promoción cag",
  "promoción caz",
  "promoción cgd",
  "alumno",
  "estado expediente",
  "nif",
  "entidad",
  "cif entidad",
  "provincia",
  "dirección completa",
  "envío revista",
  "cag",
  "caz",
  "cgd",
];

const requiredSecondaryHeaders = ["alumno", "estado expediente", "nif"];

const readFile = (file) =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("No se pudo leer el archivo"));
    reader.readAsText(file, "utf-8");
  });

const readFileAsArrayBuffer = (file) =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("No se pudo leer el archivo"));
    reader.readAsArrayBuffer(file);
  });

const detectDelimiter = (line) => {
  const delimiters = [",", ";", "\t"];
  let best = delimiters[0];
  let maxCount = 0;
  delimiters.forEach((delimiter) => {
    const count = line.split(delimiter).length - 1;
    if (count > maxCount) {
      maxCount = count;
      best = delimiter;
    }
  });
  return best;
};

const parseCSV = (text) => {
  const rows = [];
  let current = "";
  let row = [];
  let inQuotes = false;
  const firstLine = text.split(/\r?\n/)[0] || "";
  const delimiter = detectDelimiter(firstLine);

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

    if (char === delimiter && !inQuotes) {
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

const normalizeCells = (row) =>
  row.map((cell) => {
    if (cell === null || cell === undefined) return "";
    return String(cell).trim();
  });

const parseXLSX = (arrayBuffer) => {
  if (!window.XLSX) {
    throw new Error("No se encontró la librería para leer XLSX.");
  }
  const workbook = window.XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    return { headers: [], rows: [] };
  }
  const sheet = workbook.Sheets[sheetName];
  const rows = window.XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const headers = normalizeCells(rows.shift() || []);
  const data = rows
    .map((row) => normalizeCells(row))
    .filter((row) => row.some((cell) => cell.length > 0));
  return { headers, rows: data };
};

const readSpreadsheet = async (file) => {
  const extension = file.name.split(".").pop()?.toLowerCase() || "";
  if (extension === "csv") {
    const text = await readFile(file);
    return parseCSV(text);
  }
  if (extension === "xlsx") {
    const buffer = await readFileAsArrayBuffer(file);
    return parseXLSX(buffer);
  }
  throw new Error("Formato no compatible. Usa CSV o XLSX.");
};

const normalizeHeader = (header) => header.toLowerCase().trim();

const findMissingHeaders = (headers, required) => {
  const normalized = headers.map(normalizeHeader);
  return required.filter((field) => !normalized.includes(field));
};

const findHeaderIndex = (headers, name) => {
  const normalized = headers.map(normalizeHeader);
  return normalized.findIndex((header) => header === name);
};

const renderTable = (table, headers, rows, options = {}) => {
  const { highlightNif = false, nifIndex = null, limit = 50 } = options;
  const shownRows = rows.slice(0, limit);
  const thead = `
    <thead>
      <tr>
        ${headers.map((header) => `<th>${header}</th>`).join("")}
      </tr>
    </thead>
  `;

  const tbody = `
    <tbody>
      ${shownRows
        .map((row) =>
          `<tr>${headers
            .map((header, index) => {
              const value = row[index] || "";
              const highlight = highlightNif && index === nifIndex ? "match" : "";
              return `<td class="${highlight}">${value}</td>`;
            })
            .join("")}</tr>`
        )
        .join("")}
    </tbody>
  `;

  table.innerHTML = `${thead}${tbody}`;
};

const setStatus = (element, text, isReady) => {
  element.textContent = text;
  element.classList.toggle("ready", Boolean(isReady));
};

const resetModalState = () => {
  csvSecondaryInput.value = "";
  setStatus(csvSecondaryStatus, "Sin archivo cargado.", false);
  matchSummary.hidden = true;
  matchResults.hidden = true;
  previewSection.hidden = true;
  generateWordBtn.disabled = true;
  downloadPreviewBtn.disabled = true;
  matchesTable.innerHTML = "";
  mismatchesTable.innerHTML = "";
  labelsPreview.innerHTML = "";
  state.secondary = null;
  state.matches = [];
  state.mismatches = [];
  state.labelRows = [];
};

const openModal = () => {
  labelsModal.hidden = false;
  labelsModal.setAttribute("aria-hidden", "false");
  document.body.classList.add("modal-open");
};

const closeModal = () => {
  labelsModal.hidden = true;
  labelsModal.setAttribute("aria-hidden", "true");
  document.body.classList.remove("modal-open");
  resetModalState();
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
      <title>Etiquetas alumnos</title>
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

csvPrimaryInput.addEventListener("change", async (event) => {
  const [file] = event.target.files;
  if (!file) return;

  setStatus(csvPrimaryStatus, `Cargando ${file.name}...`, false);

  try {
    const parsed = await readSpreadsheet(file);
    const missing = findMissingHeaders(parsed.headers, requiredPrimaryHeaders);

    if (missing.length) {
      throw new Error(`Faltan columnas: ${missing.join(", ")}`);
    }

    const nifIndex = findHeaderIndex(parsed.headers, "nif");
    const normalizedRows = parsed.rows.map((row) => {
      const copy = [...row];
      if (nifIndex >= 0) {
        copy[nifIndex] = (copy[nifIndex] || "").toUpperCase();
      }
      return copy;
    });

    state.primary = { ...parsed, rows: normalizedRows, nifIndex, name: file.name };
    setStatus(csvPrimaryStatus, `Archivo cargado: ${file.name}`, true);
    primaryData.hidden = false;
    primaryCount.textContent = `${parsed.rows.length} filas`;
    renderTable(primaryTable, parsed.headers, normalizedRows, { limit: 50 });
    openModalBtn.disabled = false;
  } catch (error) {
    setStatus(csvPrimaryStatus, error.message, false);
    primaryData.hidden = true;
    openModalBtn.disabled = true;
  }
});

openModalBtn.addEventListener("click", () => {
  if (!state.primary) return;
  openModal();
});

closeModalBtn.addEventListener("click", closeModal);
labelsModal.addEventListener("click", (event) => {
  if (event.target === labelsModal) {
    closeModal();
  }
});

csvSecondaryInput.addEventListener("change", async (event) => {
  const [file] = event.target.files;
  if (!file) return;

  setStatus(csvSecondaryStatus, `Cargando ${file.name}...`, false);

  try {
    const parsed = await readSpreadsheet(file);
    const missing = findMissingHeaders(parsed.headers, requiredSecondaryHeaders);

    if (missing.length) {
      throw new Error(`Faltan columnas: ${missing.join(", ")}`);
    }

    const nifIndex = findHeaderIndex(parsed.headers, "nif");
    const normalizedRows = parsed.rows.map((row) => {
      const copy = [...row];
      copy[nifIndex] = (copy[nifIndex] || "").toUpperCase();
      return copy;
    });

    const primaryNifIndex = state.primary.nifIndex;
    const primaryNifSet = new Set(
      state.primary.rows
        .map((row) => row[primaryNifIndex])
        .filter((value) => value && value.length > 0)
    );

    const matches = [];
    const mismatches = [];

    normalizedRows.forEach((row) => {
      const nif = row[nifIndex];
      if (nif && primaryNifSet.has(nif)) {
        matches.push(row);
      } else {
        mismatches.push(row);
      }
    });

    state.secondary = { ...parsed, rows: normalizedRows, nifIndex, name: file.name };
    state.matches = matches;
    state.mismatches = mismatches;

    const matchNifSet = new Set(matches.map((row) => row[nifIndex]).filter(Boolean));
    state.labelRows = state.primary.rows.filter((row) => matchNifSet.has(row[primaryNifIndex]));

    setStatus(csvSecondaryStatus, `Archivo cargado: ${file.name}`, true);
    matchSummary.hidden = false;
    matchResults.hidden = false;
    matchCount.textContent = matches.length.toString();
    mismatchCount.textContent = mismatches.length.toString();

    renderTable(matchesTable, parsed.headers, matches, {
      highlightNif: true,
      nifIndex,
      limit: 50,
    });
    renderTable(mismatchesTable, parsed.headers, mismatches, { limit: 50 });

    generateWordBtn.disabled = matches.length === 0;
    downloadPreviewBtn.disabled = matches.length === 0;
  } catch (error) {
    setStatus(csvSecondaryStatus, error.message, false);
    matchSummary.hidden = true;
    matchResults.hidden = true;
    generateWordBtn.disabled = true;
    downloadPreviewBtn.disabled = true;
  }
});

downloadPreviewBtn.addEventListener("click", () => {
  if (!state.labelRows.length) return;
  labelsPreview.innerHTML = buildLabelHTML(state.primary.headers, state.labelRows);
  previewSection.hidden = false;
  previewSection.scrollIntoView({ behavior: "smooth", block: "start" });
});

generateWordBtn.addEventListener("click", () => {
  if (!state.labelRows.length) return;
  const html = buildWordDocument(state.primary.headers, state.labelRows);
  const blob = new Blob([html], { type: "application/msword" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "etiquetas-alumnos.doc";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
});
