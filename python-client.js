/* Connects the existing web interface to the Python cleaning/enrichment API. */

function convertPythonRecord(record) {
  const review = record._review || { attributes: [], confidence: 0, reasons: [] };

  // The table renderer expects attributes as [label, value, UOM] tuples.
  record._meta = {
    attrs: review.attributes.map((attribute) => [
      attribute.label,
      attribute.value,
      attribute.uom,
    ]),
    confidence: review.confidence,
    needs: review.reasons.length > 0,
  };

  return record;
}

async function processRowsWithPython(rows, fileName) {
  const response = await fetch("/api/enrich", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ rows }),
  });

  const result = await response.json();
  if (!response.ok) {
    throw new Error(result.error || "The Python API could not process this file.");
  }

  return result;
}

// Override the old browser-only loading path. Every uploaded file now reaches Python.
load = async function loadWithPython(rows, fileName) {
  try {
    const result = await processRowsWithPython(rows, fileName);

    enriched = result.records.map(convertPythonRecord);
    duplicateTotal = result.duplicates;
    currentName = fileName;

    $("#fileName").textContent = fileName;
    $("#batchMeta").textContent =
      `${rows.length.toLocaleString()} source rows · ${result.duplicates} duplicates merged · ${result.schemaLength} delivery fields`;
    $("#workspace").hidden = false;
    $(".upload-card").hidden = true;

    render();
    $("#workspace").scrollIntoView({ behavior: "smooth" });
  } catch (error) {
    alert(
      `Python processing error: ${error.message}\n\nStart the app with: python server.py\nThen open: http://127.0.0.1:8000`,
    );
  }
};

// A reliable UTF-8 CSV download. It also exposes a fallback link if a browser blocks it.
exportCSV = function downloadEnrichedCsv() {
  if (!enriched.length) {
    alert("Upload and process a file before downloading.");
    return;
  }

  const csvRows = [
    schema.join(","),
    ...enriched.map((record) =>
      schema.map((column) => escapeCSV(record[column])).join(","),
    ),
  ];
  const csv = csvRows.join("\r\n");
  const downloadName = `${currentName.replace(/\.(csv|xlsx|xls)$/i, "") || "enriched"}_delivery.csv`;
  const blobUrl = URL.createObjectURL(
    new Blob([`\uFEFF${csv}`], { type: "text/csv;charset=utf-8" }),
  );

  const downloadLink = document.createElement("a");
  downloadLink.href = blobUrl;
  downloadLink.download = downloadName;
  document.body.appendChild(downloadLink);
  downloadLink.click();
  downloadLink.remove();

  const fallbackLink = $("#downloadFallback");
  fallbackLink.href = blobUrl;
  fallbackLink.download = downloadName;
  fallbackLink.hidden = false;
  $("#exportBtn").textContent = "Download ready ✓";
  setTimeout(() => { $("#exportBtn").textContent = "Download enriched CSV"; }, 1800);
};

$("#exportBtn").onclick = exportCSV;

async function readReferenceFile(file) {
  if (/\.xlsx?$/i.test(file.name)) {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    return XLSX.utils.sheet_to_json(firstSheet, { defval: "" });
  }

  return csvParse(await file.text());
}

async function loadUniCatMaster(file) {
  try {
    const rows = await readReferenceFile(file);
    const response = await fetch("/api/unicat-master", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ rows }),
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "Could not read the UniCat file.");

    $("#unicatBtn").textContent = `UniCat loaded (${result.loaded.toLocaleString()})`;
  } catch (error) {
    alert(`UniCat master could not be loaded: ${error.message}`);
  }
}

$("#unicatBtn").onclick = () => $("#unicatInput").click();
$("#unicatInput").onchange = (event) => {
  const file = event.target.files[0];
  if (file) loadUniCatMaster(file);
};
