let dbFiles = [];
let dbSelected = [];
let exportHistory = [];
let searchTimeout = null;

function switchPage(pageId, btnElement) {
  document
    .querySelectorAll(".page")
    .forEach((p) => p.classList.remove("active"));
  document
    .querySelectorAll(".nav-item")
    .forEach((b) => b.classList.remove("active"));

  document.getElementById("page-" + pageId).classList.add("active");
  btnElement.classList.add("active");

  if (pageId === "pencarian") {
    setTimeout(() => document.getElementById("searchInput").focus(), 100);
  }
  if (window.innerWidth <= 1024) {
    document.getElementById("sidebar").classList.remove("show");
  }
}

function showToast(message, type = "success") {
  const container = document.getElementById("toast-container");
  const toast = document.createElement("div");
  toast.className = type === "success" ? "toast" : "toast error";

  const iconSvg =
    type === "success"
      ? `<svg fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>`
      : `<svg fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"></path></svg>`;

  toast.innerHTML = `${iconSvg} <span>${message}</span>`;
  container.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = "0";
    toast.style.transform = "translateX(100%)";
    toast.style.transition = "all 0.3s ease";
    setTimeout(() => toast.remove(), 300);
  }, 3500);
}

window.onload = async function () {
  const savedFiles = await localforage.getItem("dbFilesMultiSheet");
  const savedSelected = await localforage.getItem("dbSelectedMultiSheet");
  const savedHistory = await localforage.getItem("appExportHistory");

  if (savedFiles) dbFiles = savedFiles;
  if (savedSelected) dbSelected = savedSelected;
  if (savedHistory) exportHistory = savedHistory;

  updateUI();
};

function updateUI() {
  renderFileList();
  updateSheetDropdownOptions();
  renderSelectedTables();
  renderExportHistory();
  updateDashboardStats();

  const badge = document.getElementById("count-badge");
  badge.innerText = dbSelected.length;
  badge.style.display = dbSelected.length > 0 ? "inline-block" : "none";
}

function updateDashboardStats() {
  let totalData = dbFiles.reduce((acc, f) => acc + f.rows.length, 0);
  document.getElementById("stat-total-data").innerText =
    totalData.toLocaleString("id-ID");
  document.getElementById("stat-file-db").innerText = dbFiles.length;
  document.getElementById("stat-total-export").innerText = exportHistory.length;

  if (exportHistory.length > 0) {
    document.getElementById("stat-last-export").innerText =
      exportHistory[0].date.split(" ")[0];
  } else {
    document.getElementById("stat-last-export").innerText = "-";
  }
}

function updateSheetDropdownOptions() {
  const select = document.getElementById("sheetFilterSelect");
  const currentVal = select.value;
  let sheetsSet = new Set();
  dbFiles.forEach((f) => {
    if (f.sheetName) sheetsSet.add(f.sheetName);
  });

  let html = '<option value="ALL">📁 Semua Sheet</option>';
  sheetsSet.forEach((s) => {
    html += `<option value="${s}">📑 ${s}</option>`;
  });
  select.innerHTML = html;
  if (sheetsSet.has(currentVal)) select.value = currentVal;
}

function openQRModal(encodedText, encodedTitle) {
  const rawText = decodeURIComponent(encodedText);
  document.getElementById("modalTitle").innerText =
    decodeURIComponent(encodedTitle);

  const formattedHTML = rawText
    .split("\n")
    .map((line) => {
      let parts = line.split(": ");
      if (parts.length > 1) {
        let key = parts.shift();
        let val = parts.join(": ");
        return `<span style="color:#94a3b8; font-weight:600;">${key}:</span> <span style="color:var(--accent-cyan); font-weight:800; font-size:15px;">${val}</span>`;
      }
      return line;
    })
    .join("<br>");

  document.getElementById("modalDataText").innerHTML = formattedHTML;
  new QRious({
    element: document.getElementById("modalCanvas"),
    value: rawText,
    size: 280,
    level: "M",
  });
  document.getElementById("qrModal").classList.add("show");
}

function closeModal() {
  document.getElementById("qrModal").classList.remove("show");
}

document.getElementById("qrModal").addEventListener("click", function (e) {
  if (e.target === this) closeModal();
});

function generateQRSpecificText(rowObj) {
  let textParts = [];
  for (let key in rowObj) {
    if (key.startsWith("_") || key === "SYSTEM_ID") continue;
    if (rowObj[key]) textParts.push(`${key}: ${rowObj[key]}`);
  }
  return textParts.length > 0 ? textParts.join("\n") : "Data Kosong";
}

document.getElementById("fileUpload").addEventListener("change", function (e) {
  const files = e.target.files;
  if (!files || files.length === 0) return;

  Array.from(files).forEach((file) => {
    const reader = new FileReader();
    reader.onload = async function (ev) {
      try {
        showToast("Memproses file Excel...", "success");
        const data = new Uint8Array(ev.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        let totalNewRows = 0;

        workbook.SheetNames.forEach((sheetName) => {
          let rawData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
            header: 1,
            raw: false,
            defval: "",
          });
          if (!rawData || rawData.length === 0) return;

          let headerRowIndex = 0,
            maxFilled = 0;
          for (let i = 0; i < Math.min(15, rawData.length); i++) {
            let filledCount = rawData[i].filter(
              (c) => c && c.toString().trim() !== "",
            ).length;
            if (filledCount > maxFilled) {
              maxFilled = filledCount;
              headerRowIndex = i;
            }
          }

          let rawHeaders = rawData[headerRowIndex];
          let validColumns = [];
          let headers = [];

          rawHeaders.forEach((h, idx) => {
            if (h && h.toString().trim() !== "") {
              validColumns.push(idx);
              headers.push(h.toString().trim());
            }
          });

          let parsedSheetRows = [];
          for (let i = headerRowIndex + 1; i < rawData.length; i++) {
            let rowArr = rawData[i];
            if (!rowArr || rowArr.length === 0) continue;

            let isEmptyRow = true;
            let rowObj = {};
            validColumns.forEach((colIdx, arrIdx) => {
              let val =
                rowArr[colIdx] !== undefined && rowArr[colIdx] !== null
                  ? rowArr[colIdx].toString().trim()
                  : "";
              rowObj[headers[arrIdx]] = val;
              if (val !== "") isEmptyRow = false;
            });

            if (isEmptyRow) continue;

            rowObj["SYSTEM_ID"] =
              "SYS_" +
              Date.now() +
              "_" +
              Math.random().toString(36).substr(2, 5);
            rowObj["_sourceHeaders"] = headers;
            rowObj["_sheetName"] = sheetName;
            rowObj["_fileName"] = file.name;
            rowObj["_QR_SPECIFIC_TEXT"] = generateQRSpecificText(rowObj);
            parsedSheetRows.push(rowObj);
          }

          if (parsedSheetRows.length > 0) {
            totalNewRows += parsedSheetRows.length;
            dbFiles.push({
              id:
                "FILE_" +
                Date.now() +
                "_" +
                Math.random().toString(36).substr(2, 4),
              fileName: file.name,
              sheetName: sheetName,
              uploadTime: new Date().toLocaleString("id-ID"),
              headers: headers,
              rows: parsedSheetRows,
            });
          }
        });

        if (totalNewRows === 0)
          throw new Error("File kosong atau format salah.");
        await localforage.setItem("dbFilesMultiSheet", dbFiles);
        updateUI();
        showToast(`Upload sukses: ${file.name} (${totalNewRows} baris)`);
      } catch (err) {
        showToast(err.message, "error");
      }
    };
    reader.readAsArrayBuffer(file);
  });
  e.target.value = "";
});

document.getElementById("searchInput").addEventListener("input", function () {
  clearTimeout(searchTimeout);
  searchTimeout = setTimeout(() => triggerSearchNow(), 250);
});

function triggerSearchNow() {
  const query = document
    .getElementById("searchInput")
    .value.trim()
    .toLowerCase();
  executeSearch(query);
}

function executeSearch(query) {
  const desktopBody = document.getElementById("searchBody");
  const desktopHeader = document.getElementById("searchHeader");
  const selectedSheet = document.getElementById("sheetFilterSelect").value;

  if (!query && selectedSheet === "ALL") {
    desktopHeader.innerHTML = "<th>Daftar Pencarian</th>";
    desktopBody.innerHTML =
      '<tr><td style="text-align:center; padding:60px; color:var(--text-muted); font-size:15px;">Ketikkan kata kunci untuk memunculkan data.</td></tr>';
    return;
  }

  let matchedRows = [];
  dbFiles.forEach((file) => {
    if (selectedSheet !== "ALL" && file.sheetName !== selectedSheet) return;
    file.rows.forEach((row) => {
      let rowStr = (
        Object.values(row).join(" ") +
        " " +
        file.sheetName
      ).toLowerCase();
      if (!query || rowStr.includes(query)) matchedRows.push(row);
    });
  });

  if (matchedRows.length === 0) {
    desktopHeader.innerHTML = "<th>Status</th>";
    desktopBody.innerHTML =
      '<tr><td style="text-align:center; padding:60px; color:var(--danger); font-size:15px; font-weight:600;">Data tidak ditemukan.</td></tr>';
    return;
  }

  let headersOrdered =
    matchedRows[0]._sourceHeaders ||
    Object.keys(matchedRows[0]).filter(
      (k) => !k.startsWith("_") && k !== "SYSTEM_ID",
    );

  let headHtml =
    '<th class="fit-col">Aksi</th><th class="fit-col">Barcode</th><th class="fit-col">Sheet</th>';
  headersOrdered.forEach((h) => (headHtml += `<th>${h}</th>`));
  desktopHeader.innerHTML = headHtml;

  let bodyHtml = "";
  matchedRows.forEach((r) => {
    let isSelected = dbSelected.some((s) => s.SYSTEM_ID === r.SYSTEM_ID);
    let btnClass = isSelected ? "btn-danger-outline" : "btn-primary";
    let btnText = isSelected ? "Batalkan" : "Pilih";

    bodyHtml += `<tr>
              <td class="fit-col"><button class="btn ${btnClass} btn-sm" onclick="toggleSelectRow('${r.SYSTEM_ID}')">${btnText}</button></td>
              <td class="fit-col"><div class="qr-wrapper" onclick="openQRModal('${encodeURIComponent(r._QR_SPECIFIC_TEXT)}', '${encodeURIComponent(r[headersOrdered[0]] || "Item")}')">
                  <canvas id="qr_search_${r.SYSTEM_ID}"></canvas></div></td>
              <td class="fit-col"><span class="sheet-badge">${r._sheetName}</span></td>`;
    headersOrdered.forEach((h) => (bodyHtml += `<td>${r[h] || "-"}</td>`));
    bodyHtml += "</tr>";
  });
  desktopBody.innerHTML = bodyHtml;

  matchedRows.forEach((r) => {
    let el = document.getElementById(`qr_search_${r.SYSTEM_ID}`);
    if (el)
      new QRious({
        element: el,
        value: r._QR_SPECIFIC_TEXT,
        size: 100,
        level: "M",
      });
  });
}

async function toggleSelectRow(systemId) {
  let foundRow = null;
  for (let file of dbFiles) {
    let match = file.rows.find((r) => r.SYSTEM_ID === systemId);
    if (match) {
      foundRow = match;
      break;
    }
  }
  if (!foundRow) return;

  let existingIndex = dbSelected.findIndex((s) => s.SYSTEM_ID === systemId);
  if (existingIndex >= 0) {
    dbSelected.splice(existingIndex, 1);
  } else {
    dbSelected.push(foundRow);
  }

  await localforage.setItem("dbSelectedMultiSheet", dbSelected);
  updateUI();
  triggerSearchNow();
}

async function clearSelected() {
  if (dbSelected.length === 0) return;
  if (!confirm("Hapus seluruh data di daftar Export?")) return;
  dbSelected = [];
  await localforage.setItem("dbSelectedMultiSheet", dbSelected);
  updateUI();
  triggerSearchNow();
  showToast("Daftar export telah dibersihkan.");
}

function renderSelectedTables() {
  const body = document.getElementById("selectedBody");
  const header = document.getElementById("selectedHeader");

  if (dbSelected.length === 0) {
    header.innerHTML = "<th>Data Kosong</th>";
    body.innerHTML =
      '<tr><td style="text-align:center; padding:60px; color:var(--text-muted); font-size:15px;">Belum ada data yang dipilih untuk di-export.</td></tr>';
    return;
  }

  let headersOrdered =
    dbSelected[0]._sourceHeaders ||
    Object.keys(dbSelected[0]).filter(
      (k) => !k.startsWith("_") && k !== "SYSTEM_ID",
    );

  let headHtml =
    '<th class="fit-col">Aksi</th><th class="fit-col">Barcode</th><th class="fit-col">Sheet</th>';
  headersOrdered.forEach((h) => (headHtml += `<th>${h}</th>`));
  header.innerHTML = headHtml;

  let bodyHtml = "";
  dbSelected.forEach((r) => {
    bodyHtml += `<tr>
              <td class="fit-col"><button class="btn btn-danger-outline btn-sm" onclick="toggleSelectRow('${r.SYSTEM_ID}')">Hapus</button></td>
              <td class="fit-col"><div class="qr-wrapper" onclick="openQRModal('${encodeURIComponent(r._QR_SPECIFIC_TEXT)}', 'Preview Barcode')">
                  <canvas id="qr_sel_${r.SYSTEM_ID}"></canvas></div></td>
              <td class="fit-col"><span class="sheet-badge">${r._sheetName}</span></td>`;
    headersOrdered.forEach((h) => (bodyHtml += `<td>${r[h] || "-"}</td>`));
    bodyHtml += "</tr>";
  });
  body.innerHTML = bodyHtml;

  dbSelected.forEach((r) => {
    let el = document.getElementById(`qr_sel_${r.SYSTEM_ID}`);
    if (el)
      new QRious({
        element: el,
        value: r._QR_SPECIFIC_TEXT,
        size: 100,
        level: "M",
      });
  });
}

function renderFileList() {
  const container = document.getElementById("fileListContainer");
  if (dbFiles.length === 0) {
    container.innerHTML =
      '<div style="color:var(--text-muted); font-size:15px; text-align:center; padding: 40px;">Database saat ini kosong.</div>';
    return;
  }
  let html = "";
  dbFiles.forEach((f) => {
    html += `<div style="display:flex; justify-content:space-between; align-items:center; padding:22px 26px; background:rgba(0,0,0,0.5); border-radius:16px; border:1px solid rgba(255,255,255,0.1); transition: all 0.3s; box-shadow: inset 0 2px 10px rgba(0,0,0,0.3);">
              <div>
                  <div style="font-weight:800; font-size:16px; letter-spacing:0.5px;">📊 ${f.fileName}</div>
                  <div style="font-size:13px; color:var(--accent-cyan); margin-top:8px; font-weight: 700;">Sheet: ${f.sheetName} &nbsp;•&nbsp; ${f.rows.length} Baris</div>
              </div>
              <button class="btn btn-danger-outline btn-sm" onclick="removeFile('${f.id}')">Hapus</button>
          </div>`;
  });
  container.innerHTML = html;
}

function renderExportHistory() {
  const container = document.getElementById("exportHistoryContainer");
  if (exportHistory.length === 0) {
    container.innerHTML =
      '<div style="color:var(--text-muted); font-size:15px; text-align:center; padding: 40px;">Tidak ada riwayat export.</div>';
    return;
  }
  let html = "";
  exportHistory.forEach((exp) => {
    html += `<div style="padding:22px 26px; background:rgba(0,0,0,0.5); border-radius:16px; border:1px solid rgba(255,255,255,0.1); display:flex; justify-content:space-between; align-items:center; flex-wrap: wrap; gap:15px; transition: all 0.3s; box-shadow: inset 0 2px 10px rgba(0,0,0,0.3);">
              <div>
                  <div style="font-size:13px; color:var(--text-muted); font-weight:700;">${exp.date}</div>
                  <div style="font-size:16px; font-weight:800; color:var(--accent-green); margin-top:8px;">${exp.count} Barcode Diexport</div>
              </div>
              <div style="display:flex; gap:12px;">
                  <button class="btn btn-primary btn-sm" onclick="loadHistoryData('${exp.id}')">Lihat Data</button>
                  <button class="btn btn-danger-outline btn-sm" onclick="deleteHistory('${exp.id}')">✕</button>
              </div>
          </div>`;
  });
  container.innerHTML = html;
}

async function removeFile(fileId) {
  if (!confirm("Hapus file dan data ini dari database?")) return;
  let fileObj = dbFiles.find((f) => f.id === fileId);
  let removedIds = new Set(fileObj.rows.map((r) => r.SYSTEM_ID));
  dbFiles = dbFiles.filter((f) => f.id !== fileId);
  dbSelected = dbSelected.filter((s) => !removedIds.has(s.SYSTEM_ID));

  await localforage.setItem("dbFilesMultiSheet", dbFiles);
  await localforage.setItem("dbSelectedMultiSheet", dbSelected);
  updateUI();
  triggerSearchNow();
  showToast("File berhasil dihapus.");
}

async function clearDatabase() {
  if (!confirm("Peringatan: Seluruh data akan terhapus. Lanjutkan?")) return;
  dbFiles = [];
  dbSelected = [];
  await localforage.setItem("dbFilesMultiSheet", dbFiles);
  await localforage.setItem("dbSelectedMultiSheet", dbSelected);
  updateUI();
  triggerSearchNow();
  showToast("Database telah dikosongkan.");
}

async function loadHistoryData(historyId) {
  const target = exportHistory.find((e) => e.id === historyId);
  if (target) {
    dbSelected = JSON.parse(JSON.stringify(target.items));
    await localforage.setItem("dbSelectedMultiSheet", dbSelected);
    updateUI();
    switchPage("terpilih", document.querySelectorAll(".nav-item")[2]);
    showToast("Data riwayat berhasil dimuat ulang.");
  }
}

async function deleteHistory(historyId) {
  exportHistory = exportHistory.filter((e) => e.id !== historyId);
  await localforage.setItem("appExportHistory", exportHistory);
  renderExportHistory();
  updateDashboardStats();
}

async function downloadExcel() {
  if (dbSelected.length === 0) {
    showToast("Pilih data yang ingin di-export terlebih dahulu!", "error");
    return;
  }
  try {
    showToast("Memproses pembuatan Excel...");
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Data Export");

    let headersOrdered =
      dbSelected[0]._sourceHeaders ||
      Object.keys(dbSelected[0]).filter(
        (k) => !k.startsWith("_") && k !== "SYSTEM_ID",
      );

    worksheet.columns = [
      { header: "Barcode / QR", key: "qr_code", width: 24 },
      { header: "Sumber Sheet", key: "_sheetName", width: 24 },
      ...headersOrdered.map((h) => ({ header: h, key: h, width: 30 })),
    ];

    const headerRow = worksheet.getRow(1);
    headerRow.height = 40;
    headerRow.font = { bold: true, color: { argb: "FFFFFF" }, size: 13 };
    headerRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "00F2FE" },
    };
    headerRow.alignment = { vertical: "middle", horizontal: "center" };

    const canvas = document.getElementById("hiddenCanvas");

    for (let i = 0; i < dbSelected.length; i++) {
      const rowData = dbSelected[i];
      let rowObj = { _sheetName: rowData._sheetName || "-" };
      headersOrdered.forEach((h) => (rowObj[h] = rowData[h] || "-"));

      const addedRow = worksheet.addRow(rowObj);
      addedRow.height = 120;
      addedRow.eachCell({ includeEmpty: true }, (cell) => {
        cell.alignment = {
          vertical: "middle",
          wrapText: true,
          horizontal: "left",
        };
        cell.border = {
          top: { style: "thin", color: { argb: "E2E8F0" } },
          bottom: { style: "thin", color: { argb: "E2E8F0" } },
          left: { style: "thin", color: { argb: "E2E8F0" } },
          right: { style: "thin", color: { argb: "E2E8F0" } },
        };
      });

      new QRious({
        element: canvas,
        value: rowData._QR_SPECIFIC_TEXT,
        size: 256,
        level: "M",
      });
      const base64Image = canvas.toDataURL("image/png");
      const imageId = workbook.addImage({
        base64: base64Image,
        extension: "png",
      });

      worksheet.addImage(imageId, {
        tl: { col: 0.15, row: i + 1 + 0.1 },
        ext: { width: 130, height: 130 },
        editAs: "oneCell",
      });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(
      new Blob([buffer]),
      `Export_Barcode_PT_DayaSaktiNiaga_${Date.now()}.xlsx`,
    );

    const exportId = "EXP_" + Date.now();
    const exportDate = new Date().toLocaleString("id-ID");

    exportHistory.unshift({
      id: exportId,
      date: exportDate,
      count: dbSelected.length,
      items: JSON.parse(JSON.stringify(dbSelected)),
    });
    await localforage.setItem("appExportHistory", exportHistory);
    updateUI();

    const payloadServer = {
      kode_export: exportId,
      jumlah_data: dbSelected.length,
      detail_data: dbSelected,
    };

    try {
      const response = await fetch("simpan_riwayat.php", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payloadServer),
      });

      const result = await response.json();
      if (result.status === "success") {
        showToast("File Excel diunduh & Riwayat ter-backup di Server!");
      } else {
        showToast("File diunduh, tapi gagal backup server.", "error");
      }
    } catch (error) {
      console.error("Error Fetch API:", error);
      showToast("Gagal menghubungi server database.", "error");
    }
  } catch (err) {
    showToast("Gagal melakukan export: " + err.message, "error");
  }
}
