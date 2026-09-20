// WH-347 Builder — Under the Sun Solar
// Single-page intake. Drop a QBO pay-stub PDF (or paste stub text),
// review/edit the parsed data, export a JSON config that build_wh347.py
// consumes to produce the WH-347 PDF.

const $ = (s) => document.querySelector(s);
const $$ = (s) => Array.from(document.querySelectorAll(s));

// pdf.js worker config
pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

// ---------- Row management ----------
const rowsEl = $("#rows");
const rowTpl = $("#rowTemplate");

function newRow(seed) {
  const node = rowTpl.content.firstElementChild.cloneNode(true);
  if (seed) {
    node.querySelector(".r_name").value = seed.name ?? "";
    node.querySelector(".r_dbHrs").value = seed.dbHrs ?? "";
    node.querySelector(".r_dbFringeHrs").value = seed.dbFringeHrs ?? 0;
    node.querySelector(".r_stubWages").value = seed.stubWages ?? "";
    node.querySelector(".r_fica").value = seed.fica ?? 0;
    node.querySelector(".r_wh").value = seed.wh ?? 0;
    node.querySelector(".r_other").value = seed.other ?? 0;
    node.querySelector(".r_daily").value = seed.daily ?? "";
  }
  node.querySelector(".removeRow").addEventListener("click", () => {
    node.remove();
    refreshRowTitles();
  });
  // recompute on any change
  node.querySelectorAll("input").forEach((inp) => {
    inp.addEventListener("input", () => {
      updateRowComputed(node);
      updateJsonPreview();
    });
  });
  rowsEl.appendChild(node);
  updateRowComputed(node);
  refreshRowTitles();
}

function refreshRowTitles() {
  $$("#rows .row").forEach((row, i) => {
    const name = row.querySelector(".r_name").value || `Employee ${i + 1}`;
    row.querySelector(".rowTitle").textContent = `${i + 1}. ${name}`;
  });
}

function updateRowComputed(row) {
  const out = row.querySelector(".rowComputed");
  const dbHrs = parseFloat(row.querySelector(".r_dbHrs").value) || 0;
  const fringeHrs = parseFloat(row.querySelector(".r_dbFringeHrs").value) || 0;
  const baseRate = parseFloat($("#wd_baseRate").value) || 0;
  const fringeRate = parseFloat($("#wd_fringeRate").value) || 0;
  const stubWages = parseFloat(row.querySelector(".r_stubWages").value) || 0;
  const fica = parseFloat(row.querySelector(".r_fica").value) || 0;
  const wh = parseFloat(row.querySelector(".r_wh").value) || 0;
  const other = parseFloat(row.querySelector(".r_other").value) || 0;
  const daily = (row.querySelector(".r_daily").value || "")
    .split(/[,\s]+/).map(parseFloat).filter((n) => !isNaN(n));

  if (!stubWages || !dbHrs) { out.textContent = ""; return; }
  const baseGross = round2(baseRate * dbHrs);
  const dbWages = baseRate * dbHrs + fringeRate * fringeHrs;
  const ratio = dbWages / stubWages;
  const dFica = round2(fica * ratio);
  const dWh   = round2(wh * ratio);
  const dOther = round2(other * ratio);
  const totDed = round2(dFica + dWh + dOther);
  const net = round2(baseGross - totDed);
  const dailySum = round2(daily.reduce((a, b) => a + b, 0));
  const dailyOk = Math.abs(dailySum - dbHrs) < 0.005 ? "✓" : "✗ expected " + dbHrs;
  out.innerHTML = `
    <span class="num">base gross = $${baseGross.toFixed(2)}</span> ·
    <span class="num">ratio = ${ratio.toFixed(4)}</span> ·
    <span class="num">FICA $${dFica.toFixed(2)} + W/H $${dWh.toFixed(2)} + Other $${dOther.toFixed(2)} = total $${totDed.toFixed(2)}</span> ·
    <span class="num"><b>net = $${net.toFixed(2)}</b></span> ·
    <span class="num">daily sum $${dailySum.toFixed(2)} ${dailyOk}</span>
  `;
}

function round2(n) { return Math.round(n * 100) / 100; }

// ---------- File drop / PDF parse ----------
const dz = $("#dropzone");
const fi = $("#fileinput");
dz.addEventListener("click", () => fi.click());
dz.addEventListener("dragover", (e) => { e.preventDefault(); dz.classList.add("dragover"); });
dz.addEventListener("dragleave", () => dz.classList.remove("dragover"));
dz.addEventListener("drop", (e) => {
  e.preventDefault();
  dz.classList.remove("dragover");
  handleFiles(e.dataTransfer.files);
});
fi.addEventListener("change", (e) => handleFiles(e.target.files));

async function handleFiles(fileList) {
  for (const file of fileList) {
    const name = file.name || "";
    const ext = (name.split(".").pop() || "").toLowerCase();
    const isPdf =
      file.type === "application/pdf" || ext === "pdf";
    const isSheet =
      ["xls", "xlsx", "xlsm", "csv"].includes(ext) ||
      file.type === "application/vnd.ms-excel" ||
      file.type === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
      file.type === "text/csv";

    if (isPdf) {
      const buf = await file.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
      let text = "";
      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const content = await page.getTextContent();
        text += content.items.map((it) => it.str).join(" ") + "\n";
      }
      parseStubText(text);
    } else if (isSheet) {
      const buf = await file.arrayBuffer();
      parseSpreadsheet(buf, name);
    } else {
      alert(`Unsupported file type: ${name} (${file.type || "unknown"}). Drop a PDF or XLS/XLSX/CSV.`);
    }
  }
}

// ---------- Spreadsheet parser (XLS/XLSX/CSV) ----------
function parseSpreadsheet(arrayBuffer, fileName) {
  let workbook;
  try {
    workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: false });
  } catch (e) {
    alert(`Couldn't read ${fileName}: ${e.message}`);
    return;
  }

  // Pick the first sheet (QBO reports are usually single-sheet).
  // Future: let user pick if multiple.
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    alert("Workbook has no sheets.");
    return;
  }
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", blankrows: false });
  if (!rows.length) {
    alert("Sheet is empty.");
    return;
  }

  // Dump everything to the raw preview so the user can eyeball it
  $("#rawDump").classList.remove("hidden");
  $("#rawText").textContent =
    `File: ${fileName}\nSheet: ${sheetName}\nRows: ${rows.length}\n\n` +
    rows.map((r, i) => `R${i + 1}: ${r.join("\t")}`).join("\n");

  // ---- Wide-format detection: find a header row containing 'employee' ----
  const norm = (s) => String(s || "").trim().toLowerCase();
  let headerRowIdx = -1;
  for (let i = 0; i < Math.min(rows.length, 25); i++) {
    const cells = rows[i].map(norm);
    if (cells.some((c) => /^(employee(\s+name)?|payee|name)$/i.test(c)) ||
        cells.some((c) => c.includes("employee") && !c.includes("ssn"))) {
      headerRowIdx = i;
      break;
    }
  }

  if (headerRowIdx === -1) {
    alert(
      "Couldn't find an 'Employee' header row in this sheet. " +
      "Inspect the raw dump below, then paste the stub text manually or " +
      "export the QBO report with the standard 'Employee' column."
    );
    return;
  }

  const headers = rows[headerRowIdx].map(norm);

  // Column resolver: pick the FIRST header whose normalized text matches any of the regex patterns
  const findCol = (patterns) => {
    for (let i = 0; i < headers.length; i++) {
      const h = headers[i];
      if (patterns.some((p) => p.test(h))) return i;
    }
    return -1;
  };

  const colName       = findCol([/^employee(\s+name)?$/, /^name$/, /^payee$/]);
  const colPeriod     = findCol([/pay\s*period/, /check\s*date/, /^date$/, /period\s+end/]);
  // Earnings: regular pay + DB base (Greene County / Davis-Bacon / project labor) + DB fringe
  const colRegHours   = findCol([/^regular\s+hours?$/, /^hours$/]);
  const colRegAmt     = findCol([/^regular(\s+pay|\s+earnings?)?$/, /^reg(\s+pay)?$/]);
  const colDbBaseAmt  = findCol([/new\s+electrician/, /greene\s+county/, /\bpw\b/, /davis[\s_-]*bacon/, /db\s*base/]);
  const colDbBaseHrs  = findCol([/electrician.*\bhours?\b/, /db\s*base.*hours/, /project\s*labor.*hours/]);
  const colFringeAmt  = findCol([/\bfringe\b/, /db\s*fringe/]);
  const colFringeHrs  = findCol([/fringe.*\bhours?\b/, /db\s*fringe.*hours/]);
  const colGross      = findCol([/^gross(\s+pay)?$/, /^total\s+earnings?$/, /^gross\s+earnings?$/]);
  // Deductions
  const colSsTax      = findCol([/social\s+security/, /\bss\s+tax\b/, /\boasdi\b/]);
  const colMedicare   = findCol([/medicare/, /\bmed\s+tax\b/]);
  const colFedTax     = findCol([/federal\s+income\s+tax/, /federal\s+withhold/, /\bfed\s+tax\b/, /\bfit\b/]);
  const colStateTax   = findCol([/il\s+income\s+tax/, /il\s+withhold/, /state\s+income\s+tax/, /state\s+withhold/]);
  const colNet        = findCol([/^net\s+pay$/]);

  if (colName === -1) {
    alert(
      "Found a header row but couldn't locate an 'Employee' column. " +
      "Rename the column to 'Employee' and re-export, or paste stub text manually."
    );
    return;
  }

  const num = (v) => {
    if (typeof v === "number") return v;
    if (v == null) return 0;
    const s = String(v).replace(/[$,\s]/g, "").replace(/[()]/g, "-");
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  };

  let addedCount = 0;
  const seenNames = new Set();

  for (let r = headerRowIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    const name = String(row[colName] || "").trim();
    if (!name) continue;
    // Skip totals/footer rows (often have "Total" in name or no numeric data)
    if (/^(total|grand\s+total|net\s+total)/i.test(name)) continue;
    // De-dup: if QBO lists each pay-check item as a separate row, skip repeats of the same name+period
    const period = colPeriod !== -1 ? String(row[colPeriod] || "").trim() : "";
    const key = `${name}|${period}`;
    if (seenNames.has(key)) continue;
    seenNames.add(key);

    const regAmt     = colRegAmt    !== -1 ? num(row[colRegAmt])    : 0;
    const regHrs     = colRegHours  !== -1 ? num(row[colRegHours])  : 0;
    const dbBaseAmt  = colDbBaseAmt !== -1 ? num(row[colDbBaseAmt]) : 0;
    const dbBaseHrs  = colDbBaseHrs !== -1 ? num(row[colDbBaseHrs]) : 0;
    const fringeAmt  = colFringeAmt !== -1 ? num(row[colFringeAmt]) : 0;
    const fringeHrs  = colFringeHrs !== -1 ? num(row[colFringeHrs]) : 0;
    const gross      = colGross     !== -1 ? num(row[colGross])     : 0;
    const ssTax      = colSsTax     !== -1 ? num(row[colSsTax])     : 0;
    const medicare   = colMedicare  !== -1 ? num(row[colMedicare])  : 0;
    const fedTax     = colFedTax    !== -1 ? num(row[colFedTax])    : 0;
    const stateTax   = colStateTax  !== -1 ? num(row[colStateTax])  : 0;

    // If we have a gross column, prefer it as the stub wage total
    const stubWages = gross > 0 ? gross : round2(regAmt + dbBaseAmt + fringeAmt);
    // DB base hours: prefer dedicated column, else derive from rate*amount if rate is in there,
    // else use regular hours as the best guess (matches "the whole DB project IS the regular shift" case)
    const dbHrs = dbBaseHrs > 0
      ? dbBaseHrs
      : (regHrs > 0 ? regHrs : (dbBaseAmt > 0 ? Math.round((dbBaseAmt / 30.49) * 100) / 100 : 0));
    const dbFringeHrs = fringeHrs > 0 ? fringeHrs : dbHrs;
    const fica = round2(ssTax + medicare);
    const wh = round2(fedTax + stateTax);
    // "Other" = residual after fica + wh, only if we have a Net Pay reference
    let other = 0;
    if (colNet !== -1) {
      const net = num(row[colNet]);
      // Net + fica + wh should equal gross (approximately). Diff = other.
      const otherResidual = stubWages - fica - wh - net;
      other = Math.max(0, round2(otherResidual));
    }

    newRow({
      name: name,
      dbHrs: dbHrs,
      dbFringeHrs: dbFringeHrs,
      stubWages: stubWages,
      fica: fica,
      wh: wh,
      other: other,
      daily: distributeDefault(dbHrs),
    });
    addedCount++;
  }

  // First row's pay period → week ending if empty
  if (colPeriod !== -1 && !$("#h_weekEnding").value) {
    const firstPeriod = String(rows[headerRowIdx + 1][colPeriod] || "").trim();
    const dateMatch = firstPeriod.match(/(\d{1,2}\/\d{1,2}\/\d{4})/);
    if (dateMatch) $("#h_weekEnding").value = dateMatch[1];
  }

  if (addedCount === 0) {
    alert(
      "Found the header row but no data rows matched. " +
      "Inspect the raw dump and paste stub text manually, or check the report's row layout."
    );
  } else {
    updateJsonPreview();
  }
}

$("#parseText").addEventListener("click", () => {
  const t = $("#stubtext").value;
  if (t.trim()) parseStubText(t);
});

function parseStubText(text) {
  $("#rawDump").classList.remove("hidden");
  $("#rawText").textContent = text;

  // Pull the employee name
  const nameMatch = text.match(/PAY TO\s+([A-Z][A-Za-z]+(?:\s+[A-Z]\.)?\s+[A-Z][A-Za-z]+)/);
  // Pull the week ending
  const weekMatch = text.match(/Week ending\s+(\d{2}\/\d{2}\/\d{4})/i) || text.match(/Pay period\s+\d{2}\/\d{2}\/\d{4} to (\d{2}\/\d{2}\/\d{4})/i);
  if (weekMatch && !$("#h_weekEnding").value) $("#h_weekEnding").value = weekMatch[1];

  // Pull line items — look for DB base, DB fringe, regular pay
  const dbBaseM = text.match(/New Electrician Greene County PW\s+(\d+\.\d+)\s+\$?(\d+\.\d+)\s+\$?(\d+\.\d+)/);
  const dbFringeM = text.match(/New Electrician Greene County PW Fringe\s+(\d+\.\d+)\s+\$?(\d+\.\d+)\s+\$?(\d+\.\d+)/);
  const regPayM = text.match(/Regular Pay\s+(\d+\.\d+)\s+\$?(\d+\.\d+)\s+\$?(\d+\.\d+)/);
  // Deductions
  const ssM  = text.match(/Social Security\s+\$?(\d+\.\d+)/);
  const medM = text.match(/Medicare\s+\$?(\d+\.\d+)/);
  const fitM = text.match(/Federal Income Tax\s+\$?(\d+\.\d+)/);
  const ilM  = text.match(/IL Income Tax\s+\$?(\d+\.\d+)/);
  // Total
  const totalM = text.match(/Total\s+\$?(\d{3,5}\.\d{2})/);

  if (dbBaseM) {
    // Found DB data — make a new row
    const dbHrs = parseFloat(dbBaseM[1]);
    const dbWages = parseFloat(dbBaseM[3]);
    const dbFringeHrs = dbFringeM ? parseFloat(dbFringeM[1]) : dbHrs;
    const dbFringe = dbFringeM ? parseFloat(dbFringeM[3]) : 0;

    // Stub total wages = DB base + DB fringe + regular pay
    const regWages = regPayM ? parseFloat(regPayM[3]) : 0;
    const stubWages = round2(dbWages + dbFringe + regWages);

    const fica = (ssM ? parseFloat(ssM[1]) : 0) + (medM ? parseFloat(medM[1]) : 0);
    const wh = (fitM ? parseFloat(fitM[1]) : 0) + (ilM ? parseFloat(ilM[1]) : 0);
    // Other = total - fica - wh  (residual)
    const tot = totalM ? parseFloat(totalM[1]) : 0;
    const other = Math.max(0, round2(tot - fica - wh));

    newRow({
      name: nameMatch ? nameMatch[1] : "",
      dbHrs: dbHrs,
      dbFringeHrs: dbFringeHrs,
      stubWages: stubWages,
      fica: round2(fica),
      wh: round2(wh),
      other: other,
      daily: distributeDefault(dbHrs),
    });
  } else {
    alert("Couldn't find 'New Electrician Greene County PW' line in the stub. Fill the row manually or check the raw text below.");
  }
}

function distributeDefault(total) {
  // Default: spread equally across 5 DB days (Sun, Mon, Tue, Wed, Sat),
  // last day absorbs the rounding.
  const days = [0, 1, 2, 3, 6]; // indices 0,1,2,3,6 of [S,M,T,W,T,F,S]
  const per = Math.floor((total / 5) * 100) / 100;
  const remainder = round2(total - per * 5);
  const out = [0, 0, 0, 0, 0, 0, 0];
  days.forEach((d, i) => { out[d] = per; });
  out[days[days.length - 1]] = round2(per + remainder);
  return out.join(", ");
}

// ---------- JSON export ----------
$("#exportJson").addEventListener("click", () => {
  const cfg = buildConfig();
  download("wh347_config.json", JSON.stringify(cfg, null, 2));
});
$("#copyJson").addEventListener("click", async () => {
  const cfg = buildConfig();
  await navigator.clipboard.writeText(JSON.stringify(cfg, null, 2));
  alert("Copied JSON to clipboard.");
});
$("#exportRules").addEventListener("click", () => {
  // Embed the rules text inline since we don't fetch RULES.md from disk in a static page.
  // (Mavis will have the full RULES.md; this is a courtesy export.)
  const rules = `WH-347 RULES — quick summary

Wage determination: SAM.gov WD for project county. We use 4(b) cash fringe.

Per row:
  ratio = (base × dbHrs + fringe × dbFringeHrs) / stubTotalWages
  gross = base × dbHrs
  dbFica = round2(stubFica × ratio)
  dbWh   = round2(stubWh   × ratio)
  dbOther = round2(stubOther × ratio)
  totalDed = dbFica + dbWh + dbOther
  net = gross - totalDed

Daily hours: 7 day cells (S..S), 0 for non-DB days. Sum of cells = dbHrs.
Excluded from WH-347: per diem, training, bonuses, drive time.

Page 2: 4(b) checked, name James M. Carpenter, title Owner / Subcontractor, signature.png.
`;
  download("wh347_rules.txt", rules);
});

// ---------- WH-347 PDF build + email ----------
$("#buildPdf").addEventListener("click", () => buildWh347Pdf());
$("#emailPdf").addEventListener("click", () => emailWh347Pdf());

function buildWh347Pdf() {
  const cfg = buildConfig();
  if (!cfg.employees.length) {
    alert("Add at least one employee row before building the WH-347 PDF.");
    return;
  }
  if (!cfg.header.week_ending &&
      !confirm("Week ending is empty — the PDF header will show a blank date. Continue?")) {
    return;
  }

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "pt", format: "letter" });
  const PAGE_W = 612;
  const MARGIN = 40;

  // =================== PAGE 1 ===================
  doc.setFontSize(9);
  doc.setFont("helvetica", "normal");
  doc.text("Form WH-347 (Rev. 02/2013)", MARGIN, 36);

  doc.setFontSize(15);
  doc.setFont("helvetica", "bold");
  doc.text("PAYROLL  (Construction Industry)", MARGIN, 56);

  doc.setFontSize(10);
  doc.setFont("helvetica", "italic");
  doc.text("U.S. Department of Labor", MARGIN, 72);

  // Week-ending bar
  doc.setFillColor(225, 225, 225);
  doc.rect(MARGIN, 84, PAGE_W - 2 * MARGIN, 22, "F");
  doc.setFontSize(11);
  doc.setFont("helvetica", "bold");
  doc.text(`Payroll For Week Ending:    ${cfg.header.week_ending || "—"}`, MARGIN + 6, 99);

  // Header info
  doc.setFontSize(10);
  doc.setFont("helvetica", "normal");
  let y = 124;
  const hdrLine = (label, value) => {
    doc.setFont("helvetica", "bold");
    doc.text(label, MARGIN, y);
    doc.setFont("helvetica", "normal");
    doc.text(String(value || "—"), MARGIN + doc.getTextWidth(label) + 6, y);
    y += 14;
  };
  hdrLine("Project and Location: ", cfg.header.project_location);
  hdrLine("Project or Contract No.: ", cfg.header.contract_no);
  hdrLine("Payroll No.: ", cfg.header.payroll_no);
  hdrLine("Trade / Classification: ", cfg.wage_determination.trade);
  hdrLine("Base rate ($/hr): ", `$${cfg.wage_determination.base_rate.toFixed(2)}`);
  hdrLine("Fringe rate ($/hr): ", `$${cfg.wage_determination.fringe_rate.toFixed(2)}`);

  // Compute per-row WH-347 columns
  const baseRate = cfg.wage_determination.base_rate;
  const fringeRate = cfg.wage_determination.fringe_rate;
  const rows = cfg.employees.map((emp) => {
    const dbHrs = emp.db_hours || 0;
    const fringeHrs = emp.db_fringe_hours || 0;
    const baseGross = round2(baseRate * dbHrs);
    const dbWages = baseRate * dbHrs + fringeRate * fringeHrs;
    const ratio = emp.stub_wages > 0 ? dbWages / emp.stub_wages : 0;
    const dFica = round2(emp.stub_fica * ratio);
    const dWh = round2(emp.stub_wh * ratio);
    const dOther = round2(emp.stub_other * ratio);
    const totDed = round2(dFica + dWh + dOther);
    const net = round2(baseGross - totDed);
    return {
      name: emp.name || "(unnamed)",
      ssn: "",          // not captured by intake
      straight: dbHrs,  // all DB hours treated as straight-time
      ot: 0,
      totalHrs: dbHrs,
      gross: baseGross,
      fica: dFica,
      wh: dWh,
      other: dOther,
      net: net,
      checkNo: "",      // not captured by intake
    };
  });

  const sumCol = (key) => round2(rows.reduce((a, r) => a + (r[key] || 0), 0));
  const totalsRow = {
    name: "TOTAL",
    straight: sumCol("straight"),
    ot: sumCol("ot"),
    totalHrs: sumCol("totalHrs"),
    gross: sumCol("gross"),
    fica: sumCol("fica"),
    wh: sumCol("wh"),
    other: sumCol("other"),
    net: sumCol("net"),
  };

  // ---- Manual table render (no AutoTable dependency) ----
  // Column geometry: x positions and widths
  const cols = [
    { label: "NAME",           x:  40, w: 110, align: "left"  },
    { label: "SSN\n(last 4)",  x: 150, w:  35, align: "right" },
    { label: "STRAIGHT\nHRS",  x: 185, w:  38, align: "right" },
    { label: "OVERTIME\nHRS",  x: 223, w:  38, align: "right" },
    { label: "TOTAL\nHRS",     x: 261, w:  38, align: "right" },
    { label: "GROSS\nAMOUNT",  x: 299, w:  52, align: "right" },
    { label: "FICA",           x: 351, w:  40, align: "right" },
    { label: "WITH-\nHOLDING", x: 391, w:  44, align: "right" },
    { label: "OTHER\nDED",     x: 435, w:  44, align: "right" },
    { label: "NET\nWAGES",     x: 479, w:  53, align: "right" },
    { label: "CHECK\nNO.",     x: 532, w:  40, align: "right" },
  ];
  const tableLeft = cols[0].x;
  const tableRight = cols[cols.length - 1].x + cols[cols.length - 1].w;
  const headH = 22; // 2-line header
  const rowH  = 16;
  let ty = y + 6;

  // Header
  doc.setFillColor(60, 60, 60);
  doc.rect(tableLeft, ty, tableRight - tableLeft, headH, "F");
  doc.setFontSize(7);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(255, 255, 255);
  cols.forEach((c) => {
    const lines = c.label.split("\n");
    const lineH = 8;
    const blockH = lines.length * lineH;
    const top = ty + (headH - blockH) / 2 + lineH - 1;
    doc.text(lines[0], c.align === "left" ? c.x + 4 : c.x + c.w - 4, top, { align: c.align });
    if (lines[1]) {
      doc.text(lines[1], c.align === "left" ? c.x + 4 : c.x + c.w - 4, top + lineH, { align: c.align });
    }
  });
  ty += headH;

  // Body rows
  doc.setTextColor(0, 0, 0);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);

  const drawRow = (vals, isTotal) => {
    if (isTotal) {
      doc.setFillColor(240, 240, 240);
      doc.rect(tableLeft, ty, tableRight - tableLeft, rowH, "F");
      doc.setFont("helvetica", "bold");
    }
    cols.forEach((c, i) => {
      const txt = String(vals[i] ?? "");
      doc.text(
        txt,
        c.align === "left" ? c.x + 4 : c.x + c.w - 4,
        ty + rowH - 4,
        { align: c.align },
      );
    });
    if (isTotal) doc.setFont("helvetica", "normal");
    // horizontal rule under the row
    doc.setDrawColor(180, 180, 180);
    doc.setLineWidth(0.3);
    doc.line(tableLeft, ty + rowH, tableRight, ty + rowH);
    ty += rowH;
  };

  rows.forEach((r) => {
    drawRow([
      r.name, r.ssn,
      r.straight.toFixed(2), r.ot.toFixed(2), r.totalHrs.toFixed(2),
      r.gross.toFixed(2), r.fica.toFixed(2), r.wh.toFixed(2),
      r.other.toFixed(2), r.net.toFixed(2), r.checkNo,
    ], false);
  });
  drawRow([
    totalsRow.name, "",
    totalsRow.straight.toFixed(2), totalsRow.ot.toFixed(2), totalsRow.totalHrs.toFixed(2),
    totalsRow.gross.toFixed(2), totalsRow.fica.toFixed(2), totalsRow.wh.toFixed(2),
    totalsRow.other.toFixed(2), totalsRow.net.toFixed(2), "",
  ], true);

  // Vertical column rules
  doc.setDrawColor(180, 180, 180);
  doc.setLineWidth(0.3);
  doc.line(tableLeft, y + 6, tableLeft, ty);                          // left edge
  doc.line(tableRight, y + 6, tableRight, ty);                        // right edge
  for (let i = 1; i < cols.length; i++) {
    doc.line(cols[i].x, y + 6, cols[i].x, ty);
  }
  doc.setLineWidth(0.2);

  // =================== PAGE 2 ===================
  doc.addPage();
  y = 36;

  doc.setFontSize(9);
  doc.setFont("helvetica", "normal");
  doc.text("Form WH-347 (Rev. 02/2013)", MARGIN, y);
  y += 22;

  doc.setFontSize(14);
  doc.setFont("helvetica", "bold");
  doc.text("Statement of Compliance", MARGIN + 180, y);
  y += 22;

  doc.setFontSize(10);
  doc.setFont("helvetica", "normal");
  const statement =
    "I, ____________________________________________________ (Name)\n" +
    "        ____________________________________________________ (Title)\n\n" +
    "do hereby state under penalty of perjury that the payroll data above is correct and\n" +
    "complete; that the wage rates paid to the laborers and mechanics listed above are\n" +
    "not less than the applicable wage determinations issued by the U.S. Department of\n" +
    "Labor; and that the classifications set forth for each laborer or mechanic conform\n" +
    "with the work performed.";
  doc.text(doc.splitTextToSize(statement, PAGE_W - 2 * MARGIN), MARGIN, y);
  y += 110;

  // Fringe checkboxes
  doc.setFontSize(10);
  doc.setFont("helvetica", "bold");
  doc.text("Fringe Benefits paid:", MARGIN, y);
  y += 16;
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);

  const isCashFringe  = fringeRate > 0;
  const isPlan        = false;          // not modeled in intake
  const isBoth        = false;

  const checkBox = (label, checked, indent = 0) => {
    const x = MARGIN + indent;
    doc.rect(x, y - 7, 9, 9);
    if (checked) {
      doc.setLineWidth(1.2);
      doc.line(x + 1.2, y - 4.5, x + 3.6, y - 2);
      doc.line(x + 3.6, y - 2,   x + 8,   y - 7.5);
      doc.setLineWidth(0.3);
    }
    const wrapped = doc.splitTextToSize(label, PAGE_W - 2 * MARGIN - 30 - indent);
    doc.text(wrapped, x + 14, y);
    y += Math.max(14, wrapped.length * 11);
  };

  checkBox(
    "4(a) CONTRIBUTIONS to fringe benefit plans (plans approved by the Department of Labor)",
    isPlan,
  );
  checkBox(
    "4(b) CASH PAYMENTS in lieu of fringe benefits (paid in addition to the basic hourly wage rate)",
    isCashFringe,
  );
  checkBox(
    "4(c) BOTH — fringe benefits paid partially in plans and partially in cash",
    isBoth,
  );

  y += 14;
  // Signature block
  doc.setFontSize(10);
  doc.setFont("helvetica", "bold");
  doc.text("Signature:", MARGIN, y);
  doc.setFont("helvetica", "normal");
  doc.line(MARGIN + 70, y, MARGIN + 320, y);
  y += 28;

  const sigLine = (label, value) => {
    doc.setFont("helvetica", "bold");
    doc.text(label, MARGIN, y);
    doc.setFont("helvetica", "normal");
    doc.text(String(value || "—"), MARGIN + doc.getTextWidth(label) + 6, y);
    y += 16;
  };
  sigLine("Name (printed): ", cfg.page2.name);
  sigLine("Title: ", cfg.page2.title);
  sigLine("Date signed: ", cfg.page2.date);

  // Footer
  doc.setFontSize(8);
  doc.setFont("helvetica", "italic");
  doc.text(
    "Generated by WH-347 Builder — Under the Sun Solar · FEIN 99-1689536",
    MARGIN,
    770,
  );

  const fname = `WH-347_${(cfg.header.week_ending || "draft").replace(/[^\d]/g, "_")}_p${cfg.header.payroll_no}.pdf`;
  // jsPDF 4.x removed doc.save(); trigger the download manually via a blob.
  const blob = doc.output("blob");
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fname;
  a.style.display = "none";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  // Defer revoke so the browser has time to start the download.
  setTimeout(() => URL.revokeObjectURL(url), 1000);
  return fname;
}

function emailWh347Pdf() {
  const cfg = buildConfig();
  if (!cfg.employees.length) {
    alert("Add at least one employee row before emailing the WH-347 PDF.");
    return;
  }
  const subject =
    `WH-347 — ${cfg.header.project_location || "Davis-Bacon payroll"} — week ending ${cfg.header.week_ending || "(date)"}`;
  const body =
    `Project: ${cfg.header.project_location || "—"}\n` +
    `Contract No.: ${cfg.header.contract_no || "—"}\n` +
    `Payroll No.: ${cfg.header.payroll_no}\n` +
    `Week ending: ${cfg.header.week_ending || "—"}\n` +
    `Employees: ${cfg.employees.length}\n\n` +
    `1. Save the WH-347 PDF first (click "📄 Build WH-347 PDF").\n` +
    `2. Attach the saved PDF to this email.\n` +
    `3. Send.\n\n` +
    `— Under the Sun Solar · FEIN 99-1689536`;
  window.location.href =
    `mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
}

function buildConfig() {
  const employees = $$("#rows .row").map((row) => ({
    name: row.querySelector(".r_name").value.trim(),
    db_hours: parseFloat(row.querySelector(".r_dbHrs").value) || 0,
    db_fringe_hours: parseFloat(row.querySelector(".r_dbFringeHrs").value) || 0,
    stub_wages: parseFloat(row.querySelector(".r_stubWages").value) || 0,
    stub_fica: parseFloat(row.querySelector(".r_fica").value) || 0,
    stub_wh: parseFloat(row.querySelector(".r_wh").value) || 0,
    stub_other: parseFloat(row.querySelector(".r_other").value) || 0,
    daily_hours: (row.querySelector(".r_daily").value || "")
      .split(/[,\s]+/).map(parseFloat).filter((n) => !isNaN(n)),
  }));
  return {
    header: {
      payroll_no: parseInt($("#h_payrollNo").value, 10) || 1,
      week_ending: $("#h_weekEnding").value.trim(),
      project_location: $("#h_projectLocation").value.trim(),
      contract_no: $("#h_contractNo").value.trim(),
    },
    wage_determination: {
      trade: $("#wd_trade").value.trim(),
      base_rate: parseFloat($("#wd_baseRate").value) || 0,
      fringe_rate: parseFloat($("#wd_fringeRate").value) || 0,
    },
    employees: employees,
    page2: {
      date: $("#p2_date").value.trim(),
      name: $("#p2_name").value.trim(),
      title: $("#p2_title").value.trim(),
    },
  };
}

function updateJsonPreview() {
  const cfg = buildConfig();
  $("#jsonOut").textContent = JSON.stringify(cfg, null, 2);
  $("#jsonOut").classList.remove("hidden");
}

function download(filename, contents) {
  const blob = new Blob([contents], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ---------- Sample loader ----------
$("#loadSample").addEventListener("click", () => {
  $("#h_payrollNo").value = 1;
  $("#h_weekEnding").value = "08/29/2026";
  $("#h_projectLocation").value = "white hall general | White Hall, IL";
  $("#h_contractNo").value = "[Federal contract #]";
  $("#wd_trade").value = "Electrician";
  $("#wd_baseRate").value = 30.49;
  $("#wd_fringeRate").value = 8.58;
  $("#p2_date").value = "08/28/2026";
  $("#p2_name").value = "James M. Carpenter";
  $("#p2_title").value = "Owner / Subcontractor";
  rowsEl.innerHTML = "";
  [
    { name: "Melton, Garrett | 1036",     dbHrs: 21.55, dbFringeHrs: 21.55, stubWages: 1165.56, fica: 89.17, wh: 97.94,  other: 193.18, daily: "0, 5.39, 5.39, 5.39, 0, 0, 5.38" },
    { name: "Hancock, Austin | 0692",     dbHrs: 21.55, dbFringeHrs: 21.55, stubWages: 1173.16, fica: 89.74, wh: 98.86,  other: 58.07,  daily: "0, 5.39, 5.39, 5.39, 0, 0, 5.38" },
    { name: "Wallace, Charles D. | 4653", dbHrs: 21.55, dbFringeHrs: 21.55, stubWages: 1157.96, fica: 88.58, wh: 97.03,  other: 57.32,  daily: "0, 5.39, 5.39, 5.39, 0, 0, 5.38" },
    { name: "Gaffner, Kamron A. | 9287",  dbHrs: 21.55, dbFringeHrs: 21.55, stubWages: 1257.96, fica: 96.24, wh: 109.03, other: 62.27,  daily: "0, 5.39, 5.39, 5.39, 0, 0, 5.38" },
  ].forEach(newRow);
  updateJsonPreview();
});

// ---------- Clear-all (wipe employees + header + source for a new week) ----------
$("#clearAll").addEventListener("click", () => {
  const hasContent =
    rowsEl.children.length > 0 ||
    $("#h_weekEnding").value ||
    $("#h_projectLocation").value ||
    $("#h_contractNo").value ||
    $("#stubtext").value;
  if (hasContent && !confirm("Clear all employees, header fields, and source data? This cannot be undone.")) return;
  // Wipe employees
  rowsEl.innerHTML = "";
  refreshRowTitles();
  // Wipe header (reset payroll no to 1, clear the rest)
  $("#h_payrollNo").value = 1;
  $("#h_weekEnding").value = "";
  $("#h_projectLocation").value = "";
  $("#h_contractNo").value = "";
  // Wipe source data
  $("#stubtext").value = "";
  $("#rawDump").classList.add("hidden");
  $("#rawText").textContent = "";
  // Wipe JSON preview (will re-render as empty config)
  updateJsonPreview();
});

// ---------- Wire up ----------
$("#addRow").addEventListener("click", () => newRow());
$("#h_payrollNo").addEventListener("input", updateJsonPreview);
$("#h_weekEnding").addEventListener("input", updateJsonPreview);
$("#h_projectLocation").addEventListener("input", updateJsonPreview);
$("#h_contractNo").addEventListener("input", updateJsonPreview);
$("#wd_baseRate").addEventListener("input", () => { $$("#rows .row").forEach(updateRowComputed); updateJsonPreview(); });
$("#wd_fringeRate").addEventListener("input", () => { $$("#rows .row").forEach(updateRowComputed); updateJsonPreview(); });

// Load the sample on first visit so the user sees a working example
$("#loadSample").click();
