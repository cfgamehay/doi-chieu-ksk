(function () {
  const DB_NAME = "doi-chieu-ksk", STORE = "dashboard", KEY = "imported";
  const text = value => value == null ? "" : String(value).trim();
  const fold = value => text(value).toUpperCase().replace(/Đ/g, "D").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^A-Z0-9]+/g, " ").trim();
  const cccd = value => {
    const digits = text(value).replace(/\D/g, "");
    return digits.length >= 9 && digits.length < 12 ? digits.padStart(12, "0") : digits;
  };
  const dateKey = value => {
    if (value instanceof Date && !Number.isNaN(value.valueOf())) {
      return `${String(value.getFullYear()).padStart(4, "0")}-${String(value.getMonth() + 1).padStart(2, "0")}-${String(value.getDate()).padStart(2, "0")}`;
    }
    const raw = text(value).split(" ")[0];
    if (!raw) return "";
    let match = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (match) {
      const valid = (year, month, day) => {
        const date = new Date(Date.UTC(year, month - 1, day));
        return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day;
      };
      const first = Number(match[1]), second = Number(match[2]), year = Number(match[3]);
      if (valid(year, second, first)) return `${year}-${match[2].padStart(2, "0")}-${match[1].padStart(2, "0")}`;
      if (valid(year, first, second)) return `${year}-${match[1].padStart(2, "0")}-${match[2].padStart(2, "0")}`;
      return raw;
    }
    match = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    return match ? `${match[1]}-${match[2].padStart(2, "0")}-${match[3].padStart(2, "0")}` : raw;
  };
  const count = values => values.reduce((result, value) => (result[value] = (result[value] || 0) + 1, result), {});
  const openDb = () => new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => request.result.createObjectStore(STORE);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
  const stateRequest = async (mode, action) => {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction(STORE, mode), store = transaction.objectStore(STORE), request = action(store);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
      transaction.oncomplete = () => db.close();
    });
  };
  const loadState = () => stateRequest("readonly", store => store.get(KEY));
  const saveState = state => stateRequest("readwrite", store => store.put(state, KEY));
  const clearState = () => stateRequest("readwrite", store => store.delete(KEY));

  function workbookRows(workbook, required, preferredSheet, fileName) {
    const names = preferredSheet && workbook.SheetNames.includes(preferredSheet)
      ? [preferredSheet, ...workbook.SheetNames.filter(name => name !== preferredSheet)]
      : workbook.SheetNames;
    for (const sheetName of names) {
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {header: 1, raw: true, defval: null});
      for (let headerIndex = 0; headerIndex < Math.min(rows.length, 8); headerIndex++) {
        const headers = rows[headerIndex].map(fold);
        if (required.every(name => headers.includes(name))) {
          return {sheetName, rows, headerIndex, indexes: Object.fromEntries(headers.map((name, index) => [name, index]).filter(([name]) => name))};
        }
      }
    }
    throw new Error(`${fileName}: không tìm thấy đủ cột ${required.join(", ")} trong 8 dòng đầu.`);
  }

  async function readWorkbook(file) {
    if (file.name.startsWith("~$")) return null;
    try {
      return XLSX.read(await file.arrayBuffer(), {type: "array", cellDates: true});
    } catch (error) {
      throw new Error(`${file.name}: không đọc được file Excel (${error.message}).`);
    }
  }

  async function readVnpt(file) {
    const workbook = await readWorkbook(file);
    const table = workbookRows(workbook, ["TEN BENH NHAN", "CCCD", "NGAY SINH", "TG TIEP NHAN"], null, file.name);
    const source = `VNPT · ${file.name}`;
    const records = table.rows.slice(table.headerIndex + 1).map((row, index) => ({
      row: table.headerIndex + index + 2,
      name: text(row[table.indexes["TEN BENH NHAN"]]),
      dob: dateKey(row[table.indexes["NGAY SINH"]]),
      cccd: cccd(row[table.indexes.CCCD]),
      date: dateKey(row[table.indexes["TG TIEP NHAN"]]),
      source
    })).filter(row => row.name || row.cccd);
    return {file, records, sheetName: table.sheetName, source};
  }

  async function readSystem(file) {
    const workbook = await readWorkbook(file);
    const table = workbookRows(workbook, ["HO TEN", "SO CCCD", "NGAY SINH", "NGAY VAO"], "Duoi 6", file.name);
    const source = `Hệ thống y tế · ${file.name} · ${table.sheetName}`;
    const records = table.rows.slice(table.headerIndex + 1).map((row, index) => ({
      row: table.headerIndex + index + 2,
      name: text(row[table.indexes["HO TEN"]]),
      dob: dateKey(row[table.indexes["NGAY SINH"]]),
      cccd: cccd(row[table.indexes["SO CCCD"]]),
      date: dateKey(row[table.indexes["NGAY VAO"]]),
      instruction: fold(row[table.indexes["HO TEN"]]).startsWith("NHAP HO TEN") || fold(row[table.indexes["SO CCCD"]]).startsWith("NHAP"),
      source
    })).filter(row => !row.instruction && (row.name || row.cccd));
    return {file, records, sheetName: table.sheetName, source};
  }

  async function readSchool(file) {
    const workbook = await readWorkbook(file);
    if (!workbook) return null;
    const table = workbookRows(workbook, ["HO TEN", "NGAY SINH", "CCCD"], null, file.name);
    const records = table.rows.slice(table.headerIndex + 1).map((row, index) => ({
      row: table.headerIndex + index + 2,
      name: text(row[table.indexes["HO TEN"]]),
      dob: dateKey(row[table.indexes["NGAY SINH"]]),
      cccd: cccd(row[table.indexes.CCCD]),
      status: table.indexes["TRANG THAI"] == null ? "" : text(row[table.indexes["TRANG THAI"]]),
      source: file.name
    })).filter(row => row.name || row.cccd);
    return {file, records, sheetName: table.sheetName};
  }

  const identityKey = record => `${fold(record.name)}|${record.dob}`;
  function indexes(records) {
    const byCccd = new Map(), byIdentity = new Map();
    for (const record of records) {
      if (record.cccd) byCccd.set(record.cccd, [...(byCccd.get(record.cccd) || []), record]);
      if (record.name && record.dob) {
        const key = identityKey(record);
        byIdentity.set(key, [...(byIdentity.get(key) || []), record]);
      }
    }
    return {byCccd, byIdentity};
  }
  function classify(record, target) {
    const sameCccd = record.cccd && target.byCccd.get(record.cccd);
    if (sameCccd) {
      const sameIdentity = sameCccd.some(item => (!record.name || fold(record.name) === fold(item.name)) && (!record.dob || record.dob === item.dob));
      return sameIdentity ? "exact" : "conflict";
    }
    return (target.byIdentity.get(identityKey(record)) || []).length === 1 ? "probable" : "missing";
  }
  const publicRecord = (record, result) => ({
    row: record.row, name: record.name, dob: record.dob, cccd: record.cccd || "Không có", result,
    source: record.source, status: record.status || "", date: record.date || ""
  });
  const statusKind = value => {
    const normalized = fold(value);
    return normalized.startsWith("TRUNG VOI DONG") ? "duplicate" : normalized === "DA KHAM" ? "examined" : normalized === "CHUA KHAM" ? "notExamined" : "other";
  };
  const rawSchoolName = fileName => fileName.replace(/\.[^.]+$/, "").replace(/\s*\([^)]*\)\s*$/, "").replace(/[\s-]+(?:GV|HS\d*)\s*$/i, "").replace(/^[\s-]+|[\s-]+$/g, "");
  function editDistance(left, right) {
    const row = Array.from({length: right.length + 1}, (_, index) => index);
    for (let i = 1; i <= left.length; i++) {
      let previous = row[0]; row[0] = i;
      for (let j = 1; j <= right.length; j++) {
        const old = row[j];
        row[j] = Math.min(row[j] + 1, row[j - 1] + 1, previous + (left[i - 1] === right[j - 1] ? 0 : 1));
        previous = old;
      }
    }
    return row[right.length];
  }
  function schoolNames(files) {
    const labels = files.map(file => rawSchoolName(file.name)), keys = labels.map(label => fold(label).replace(/[^A-Z0-9]/g, "")), parent = files.map((_, index) => index);
    const find = index => parent[index] === index ? index : (parent[index] = find(parent[index]));
    for (let left = 0; left < files.length; left++) for (let right = left + 1; right < files.length; right++) {
      const ratio = 1 - editDistance(keys[left], keys[right]) / Math.max(keys[left].length, keys[right].length, 1);
      if (keys[left] === keys[right] || ratio >= .92) parent[find(right)] = find(left);
    }
    const groups = new Map();
    labels.forEach((label, index) => groups.set(find(index), [...(groups.get(find(index)) || []), label]));
    const canonical = new Map([...groups].map(([root, items]) => [root, items.sort((a, b) => (b.split(" ").length - a.split(" ").length) || a.length - b.length)[0]]));
    return Object.fromEntries(files.map((file, index) => [file.name, canonical.get(find(index))]));
  }
  const sourceRecord = item => ({source: item.source, row: item.row, name: item.name, dob: item.dob, cccd: item.cccd || "Không có"});

  function duplicateData(recordsByFile, schoolByFile, vnptIndex, systemIndex) {
    const issues = [], fileCounts = {};
    for (const [fileName, records] of Object.entries(recordsByFile)) {
      const ownIndex = indexes(records).byCccd;
      for (const record of records.filter(item => statusKind(item.status) === "duplicate")) {
        const candidates = (ownIndex.get(record.cccd) || []).filter(item => item.row !== record.row);
        const conflict = candidates.some(item => identityKey(item) !== identityKey(record));
        issues.push({
          school: schoolByFile[fileName], file: fileName, row: record.row, reference: record.status,
          cccd: record.cccd || "Không có", name: record.name, dob: record.dob, conflict,
          reason: !candidates.length ? "Không tìm thấy dòng dùng cùng CCCD trong file" : conflict ? "Cùng CCCD nhưng khác họ tên/ngày sinh" : "Lặp lại cùng CCCD và thông tin cá nhân",
          candidates: candidates.map(item => ({...sourceRecord(item), status: item.status})),
          vnpt: (vnptIndex.get(record.cccd) || []).map(sourceRecord), system: (systemIndex.get(record.cccd) || []).map(sourceRecord)
        });
        fileCounts[fileName] = (fileCounts[fileName] || 0) + 1;
      }
    }
    return {
      total: issues.length, files: Object.keys(fileCounts).length, conflicts: issues.filter(item => item.conflict).length,
      samePerson: issues.filter(item => !item.conflict).length,
      byFile: Object.entries(fileCounts).sort((a, b) => b[1] - a[1]).map(([file, total]) => ({file, school: schoolByFile[file], count: total})), issues
    };
  }

  async function buildData(vnptFile, systemFile, schoolFileList, onProgress = () => {}) {
    if (!globalThis.XLSX) throw new Error("Không tải được thư viện đọc Excel.");
    if (!vnptFile || !systemFile || !schoolFileList.length) throw new Error("Cần chọn đủ file VNPT, hệ thống y tế và ít nhất một file trường.");
    onProgress("Đang đọc file VNPT…");
    const vnptData = await readVnpt(vnptFile);
    onProgress("Đang đọc file hệ thống y tế…");
    const systemData = await readSystem(systemFile);
    const schoolFiles = [...schoolFileList].filter(file => !file.name.startsWith("~$"));
    const schoolsData = [];
    for (let index = 0; index < schoolFiles.length; index++) {
      onProgress(`Đang đọc file trường ${index + 1}/${schoolFiles.length}: ${schoolFiles[index].name}`);
      schoolsData.push(await readSchool(schoolFiles[index]));
    }
    const vnpt = vnptData.records, system = systemData.records, systemTarget = indexes(system), vnptTarget = indexes(vnpt);
    const vnptResults = vnpt.map(record => ({record, result: classify(record, systemTarget)}));
    const vnptCccd = new Set(vnpt.filter(record => record.cccd).map(record => record.cccd));
    const systemCccd = new Set(system.filter(record => record.cccd).map(record => record.cccd));
    const schoolByFile = schoolNames(schoolFiles), recordsByFile = {}, schoolRecords = {}, schoolFilesSummary = [], allSchoolRows = [], schoolResults = [];
    for (const data of schoolsData) {
      const records = data.records, resultRows = records.map(record => ({record, result: classify(record, systemTarget)}));
      recordsByFile[data.file.name] = records;
      allSchoolRows.push(...records); schoolResults.push(...resultRows);
      const results = count(resultRows.map(item => item.result)), statuses = count(records.map(record => statusKind(record.status)));
      schoolFilesSummary.push({file: data.file.name, school: schoolByFile[data.file.name], total: records.length, exact: results.exact || 0, probable: results.probable || 0, conflict: results.conflict || 0, missing: results.missing || 0, examined: statuses.examined || 0, notExamined: statuses.notExamined || 0, duplicate: statuses.duplicate || 0, other: statuses.other || 0});
      schoolRecords[data.file.name] = resultRows.map(({record, result}) => {
        const systemMatches = systemTarget.byCccd.get(record.cccd) || systemTarget.byIdentity.get(identityKey(record)) || [];
        return {...publicRecord(record, result), statusKind: statusKind(record.status), vnpt: (vnptTarget.byCccd.get(record.cccd) || []).map(sourceRecord), system: systemMatches.map(sourceRecord)};
      });
    }
    const vnptCounts = count(vnptResults.map(item => item.result)), schoolCounts = count(schoolResults.map(item => item.result)), schoolStatuses = count(allSchoolRows.map(record => statusKind(record.status)));
    const reportedDuplicates = duplicateData(recordsByFile, schoolByFile, vnptTarget.byCccd, systemTarget.byCccd);
    const vnptRecords = vnptResults.map(({record, result}) => ({...publicRecord(record, result), system: (systemTarget.byCccd.get(record.cccd) || systemTarget.byIdentity.get(identityKey(record)) || []).map(sourceRecord)}));
    const uniqueMatched = [...vnptCccd].filter(value => systemCccd.has(value)).length;
    const data = {
      generatedAt: new Date().toISOString(), method: "Đối chiếu chính theo CCCD. Họ tên + ngày sinh chỉ dùng để gợi ý khi CCCD không khớp.",
      vnpt: {total: vnpt.length, uniqueCccd: vnptCccd.size, matchedByCccd: (vnptCounts.exact || 0) + (vnptCounts.conflict || 0), notMatchedByCccd: (vnptCounts.probable || 0) + (vnptCounts.missing || 0), uniqueMatchedByCccd: uniqueMatched, uniqueNotMatchedByCccd: vnptCccd.size - uniqueMatched, exact: vnptCounts.exact || 0, probable: vnptCounts.probable || 0, conflict: vnptCounts.conflict || 0, missing: vnptCounts.missing || 0},
      system: {total: system.length, uniqueCccd: systemCccd.size},
      schools: {files: schoolFilesSummary.length, total: allSchoolRows.length, matchedByCccd: (schoolCounts.exact || 0) + (schoolCounts.conflict || 0), notMatchedByCccd: (schoolCounts.probable || 0) + (schoolCounts.missing || 0), examined: schoolStatuses.examined || 0, notExamined: schoolStatuses.notExamined || 0, reportedDuplicate: schoolStatuses.duplicate || 0, otherStatus: schoolStatuses.other || 0, exact: schoolCounts.exact || 0, probable: schoolCounts.probable || 0, conflict: schoolCounts.conflict || 0, missing: schoolCounts.missing || 0},
      schoolFiles: schoolFilesSummary, schoolNames: [...new Set(Object.values(schoolByFile))].sort((a, b) => a.localeCompare(b, "vi")), schoolRecords,
      vnptRecords, reportedDuplicates, exceptions: vnptResults.filter(item => item.result !== "exact").map(({record, result}) => publicRecord(record, result)),
      sources: [
        {name: vnptFile.name, rows: vnpt.length, updated: new Date(vnptFile.lastModified).toISOString()},
        {name: systemFile.name, rows: system.length, updated: new Date(systemFile.lastModified).toISOString(), sheet: systemData.sheetName},
        {name: `${schoolFiles.length} file trường đã chọn`, rows: allSchoolRows.length, files: schoolFiles.length}
      ]
    };
    console.assert(data.vnpt.total === data.vnpt.exact + data.vnpt.probable + data.vnpt.conflict + data.vnpt.missing, "Cân đối VNPT không bằng tổng");
    console.assert(data.schools.total === data.schools.exact + data.schools.probable + data.schools.conflict + data.schools.missing, "Cân đối trường không bằng tổng");
    return data;
  }

  globalThis.KSK_IMPORT = {buildData, loadState, saveState, clearState};
})();
