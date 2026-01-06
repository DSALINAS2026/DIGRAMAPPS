/**
 * Diagrama Mensual de Conductores - WebApp (Google Apps Script)
 * Reemplazá TODO tu Code.gs por este archivo.
 */

const SPREADSHEET_ID = "1yfLRNOvpd4v-G5lTdmUOtWSSTSF4Jg43aXVyGddw4Sk";

const SH_SERVICIOS    = "Servicios";
const SH_FERIADOS     = "Feriados";
const SH_CONDUCTORES  = "Conductores";
const SH_DIAGRAMAS    = "Diagramas";
const SH_ASIGNACIONES = "Asignaciones";

/************* WEBAPP *************/
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Diagrama Mensual de Conductores");
}

/************* HELPERS *************/
function ss_() { return SpreadsheetApp.openById(SPREADSHEET_ID); }

function sh_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error(`No existe la hoja: ${name}`);
  return sh;
}

function getTable_(sheetName) {
  const sh = sh_(sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h || "").trim());
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const obj = {};
    headers.forEach((h, idx) => obj[h || `COL_${idx + 1}`] = row[idx]);
    const hasAny = Object.values(obj).some(v => v !== "" && v !== null && v !== undefined);
    if (hasAny) out.push(obj);
  }
  return out;
}

function normalizeDate_(d) {
  const dd = new Date(d);
  dd.setHours(0, 0, 0, 0);
  return dd;
}

function toKey_(date) {
  const d = normalizeDate_(date);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const da = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${da}`;
}

function parseTimeToMinutes_(t) {
  if (!t) return null;
  if (t instanceof Date) return t.getHours() * 60 + t.getMinutes();
  const s = String(t).trim();
  const m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  return (+m[1]) * 60 + (+m[2]);
}

function minutesToHHMM_(min) {
  if (min === null || min === undefined) return "";
  const h = Math.floor(min / 60);
  const m = min % 60;
  return String(h).padStart(2, "0") + ":" + String(m).padStart(2, "0");
}

function calcHours_(servRow, horaIni, horaFin) {
  // Prioridad 1: HorasPactadas del servicio
  const hp = Number(servRow?.HorasPactadas);
  if (!isNaN(hp) && hp > 0) return hp;

  // Prioridad 2: horas ingresadas en la asignación
  const mi = parseTimeToMinutes_(horaIni);
  const mf = parseTimeToMinutes_(horaFin);
  if (mi !== null && mf !== null) {
    let diff = mf - mi;
    if (diff < 0) diff += 24 * 60; // cruza medianoche
    return diff / 60;
  }

  // Prioridad 3: HoraInicio/HoraFin del servicio
  const mi2 = parseTimeToMinutes_(servRow?.HoraInicio);
  const mf2 = parseTimeToMinutes_(servRow?.HoraFin);
  if (mi2 !== null && mf2 !== null) {
    let diff = mf2 - mi2;
    if (diff < 0) diff += 24 * 60;
    return diff / 60;
  }

  return 0;
}

/************* API *************/
function api_bootstrap() {
  const diagramas = getTable_(SH_DIAGRAMAS)
    .filter(d => String(d.Activo || "").toUpperCase() === "SI")
    .map(d => ({
      IdDiagrama: String(d.IdDiagrama || "").trim(),
      Nombre: String(d.Nombre || "").trim(),
      Activo: String(d.Activo || "").trim()
    }))
    .filter(d => d.IdDiagrama);

  return { diagramas };
}

function api_getServicios() {
  return getTable_(SH_SERVICIOS)
    .map(s => ({
      servicio: String(s.Servicio || "").trim(),
      horaIni: s.HoraInicio || "",
      horaFin: s.HoraFin || "",
      horasPactadas: s.HorasPactadas || "",
      color: String(s.Color || "").trim()
    }))
    .filter(s => s.servicio);
}

function api_getDiagramaMensual(payload) {
  const { idDiagrama } = payload || {};
  let { year, month } = payload || {};
  year = Number(year);
  month = Number(month); // 1-12

  if (!idDiagrama) throw new Error("Falta idDiagrama");
  if (!year || !month) throw new Error("Falta year/month");

  const servicios = getTable_(SH_SERVICIOS);
  const serviciosMap = new Map(servicios.map(s => [String(s.Servicio || "").trim(), s]));

  const feriados = new Set(
    getTable_(SH_FERIADOS)
      .filter(r => r.Fecha)
      .map(r => toKey_(r.Fecha))
  );

  const conductoresAll = getTable_(SH_CONDUCTORES)
    .filter(c => String(c.Activo || "").toUpperCase() === "SI");

  // Para arrancar simple: todos los conductores activos participan del diagrama
  const conductores = conductoresAll
    .map(c => ({ legajo: String(c.Legajo || "").trim(), nombre: String(c.Conductor || "").trim() }))
    .filter(c => c.legajo);

  const asign = getTable_(SH_ASIGNACIONES).filter(a =>
    String(a.IdDiagrama || "").trim() === String(idDiagrama).trim()
  );

  // Filtrar por mes/año
  const from = new Date(year, month - 1, 1);
  const to = new Date(year, month, 1);

  const asignMes = asign.filter(a => {
    if (!a.Fecha) return false;
    const d = normalizeDate_(a.Fecha);
    return d >= from && d < to;
  });

  // Index por (fechaKey|legajo)
  const idx = new Map();
  for (const a of asignMes) {
    const key = `${toKey_(a.Fecha)}|${String(a.Legajo || "").trim()}`;
    idx.set(key, a);
  }

  // Armar grilla de días
  const daysInMonth = new Date(year, month, 0).getDate();
  const days = [];
  for (let day = 1; day <= daysInMonth; day++) {
    const date = new Date(year, month - 1, day);
    const k = toKey_(date);
    const isFer = feriados.has(k);

    const row = {
      dateKey: k,
      day,
      dow: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"][date.getDay()],
      feriado: isFer,
      cells: {}
    };

    for (const c of conductores) {
      const a = idx.get(`${k}|${c.legajo}`);
      if (!a) {
        row.cells[c.legajo] = { servicio: "", hora: "", color: "", hours: 0, feriado: isFer };
        continue;
      }

      const servicio = String(a.Servicio || "").trim();
      const servRow = serviciosMap.get(servicio) || null;

      const hIni = a.HoraInicio || servRow?.HoraInicio || "";
      const hFin = a.HoraFin || servRow?.HoraFin || "";

      const mi = parseTimeToMinutes_(hIni);
      const mf = parseTimeToMinutes_(hFin);
      const hora = (mi !== null && mf !== null)
        ? `${minutesToHHMM_(mi)} - ${minutesToHHMM_(mf)}`
        : "";

      const hours = calcHours_(servRow, a.HoraInicio, a.HoraFin);
      const color = String(servRow?.Color || "").trim();

      row.cells[c.legajo] = { servicio, hora, color, hours, feriado: isFer };
    }

    days.push(row);
  }

  // Totales por conductor
  const totals = {};
  for (const c of conductores) {
    let total = 0, fer = 0;
    for (const d of days) {
      const h = d.cells[c.legajo]?.hours || 0;
      if (d.feriado) fer += h;
      else total += h;
    }
    totals[c.legajo] = { totalHoras: total, totalFeriado: fer, totalGeneral: total + fer };
  }

  return { year, month, idDiagrama, conductores, days, totals };
}

/************* CRUD ASIGNACIONES *************/
function api_saveAsignacion(data) {
  if (!data) throw new Error("Falta data");
  const sh = sh_(SH_ASIGNACIONES);
  const values = sh.getDataRange().getValues();
  if (!values.length) throw new Error("La hoja Asignaciones no tiene encabezados.");

  const headers = values[0];
  const col = {};
  headers.forEach((h, i) => col[String(h).trim()] = i);

  // Validar columnas mínimas
  const required = ["IdDiagrama", "Fecha", "Legajo", "Servicio"];
  required.forEach(r => { if (col[r] === undefined) throw new Error(`Falta columna en Asignaciones: ${r}`); });

  const fecha = new Date(data.Fecha);
  if (isNaN(fecha.getTime())) throw new Error("Fecha inválida");
  const keyFecha = toKey_(fecha);

  let rowIndex = -1; // 1-based en sheet
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    const match =
      String(r[col.IdDiagrama]).trim() === String(data.IdDiagrama).trim() &&
      String(r[col.Legajo]).trim() === String(data.Legajo).trim() &&
      r[col.Fecha] && toKey_(r[col.Fecha]) === keyFecha;
    if (match) { rowIndex = i + 1; break; }
  }

  const rowData = Array(headers.length).fill("");
  rowData[col.IdDiagrama] = String(data.IdDiagrama || "").trim();
  rowData[col.Fecha] = fecha;
  rowData[col.Legajo] = String(data.Legajo || "").trim();
  rowData[col.Servicio] = String(data.Servicio || "").trim();

  if (col.HoraInicio !== undefined) rowData[col.HoraInicio] = String(data.HoraInicio || "").trim();
  if (col.HoraFin !== undefined) rowData[col.HoraFin] = String(data.HoraFin || "").trim();
  if (col.Notas !== undefined) rowData[col.Notas] = String(data.Notas || "").trim();

  if (rowIndex === -1) sh.appendRow(rowData);
  else sh.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);

  return { ok: true };
}

/************* CRUD DIAGRAMAS *************/
function api_saveDiagrama(data) {
  if (!data) throw new Error("Falta data");
  const sh = sh_(SH_DIAGRAMAS);
  const values = sh.getDataRange().getValues();
  if (!values.length) throw new Error("La hoja Diagramas no tiene encabezados.");

  const headers = values[0];
  const col = {};
  headers.forEach((h, i) => col[String(h).trim()] = i);

  const required = ["IdDiagrama", "Nombre", "Activo"];
  required.forEach(r => { if (col[r] === undefined) throw new Error(`Falta columna en Diagramas: ${r}`); });

  const id = String(data.IdDiagrama || "").trim();
  if (!id) throw new Error("IdDiagrama vacío");
  const nombre = String(data.Nombre || "").trim();
  const activo = String(data.Activo || "SI").trim();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][col.IdDiagrama]).trim() === id) {
      sh.getRange(i + 1, col.Nombre + 1).setValue(nombre);
      sh.getRange(i + 1, col.Activo + 1).setValue(activo);
      return { ok: true, updated: true };
    }
  }

  sh.appendRow([id, nombre, activo]);
  return { ok: true, created: true };
}
