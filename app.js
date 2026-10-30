/* Entrega 64. Complemento Formulario Contreras para Excel (Office.js).
   Basado en el script 41 suministrado por Agus. NO pegar en Automatizar.
   La identidad se consulta en cada operación y nunca se lee de una celda.
*/
const ec49 = { busy: false, pending: null, saved: null, searchPending: null, deletePending: null, newPending: null, clearBasePending: null };
const formProtection49 = {
  allowFormatRows: true, allowFormatColumns: true,
  allowFormatCells: true, allowEditObjects: true
};

Office.onReady(() => {
  const button = document.getElementById("ec49-save");
  if (!button) { console.log("Entrega 64: no se encontró la interfaz del complemento. Revise taskpane.html y actualice el panel."); return; }
  button.addEventListener("click", () => task49(start49));
  document.getElementById("ec49-confirm").addEventListener("click", () => task49(confirm49));
  document.getElementById("ec49-cancel").addEventListener("click", () => {
    ec49.pending = null; section49("ec49-confirm-box", false);
    message49("Actualización cancelada. Base y Auditoría no se modificaron.");
  });
  document.getElementById("ec49-clean").addEventListener("click", () => task49(clean49, false));
  document.getElementById("ec49-keep").addEventListener("click", () => {
    ec49.saved = null; section49("ec49-clean-box", false);
    message49("Se conservaron los datos del formulario. La solicitud ya está guardada.", true);
  });
  document.getElementById("ec52-search").addEventListener("click", () => task49(() => search52(false), false));
  document.getElementById("ec52-search-confirm").addEventListener("click", () => task49(() => search52(true), false));
  document.getElementById("ec52-search-cancel").addEventListener("click", () => {
    ec49.searchPending = null; section49("ec52-search-confirm-box", false);
    message49("Búsqueda cancelada. Se conservaron los datos del formulario.");
  });
  document.getElementById("ec53-delete").addEventListener("click", () => task49(() => delete53(false)));
  document.getElementById("ec53-delete-confirm").addEventListener("click", () => task49(() => delete53(true)));
  document.getElementById("ec53-delete-cancel").addEventListener("click", () => {
    cancelDelete53(); message49("Eliminación cancelada. Base, Auditoría y el formulario se conservaron.");
  });
  document.getElementById("ec54-new").addEventListener("click", () => task49(() => new54(false), false));
  document.getElementById("ec54-new-confirm").addEventListener("click", () => task49(() => new54(true), false));
  document.getElementById("ec54-new-cancel").addEventListener("click", () => {
    cancelNew54(); message49("Nueva carga cancelada. Se conservaron los datos del formulario.");
  });
  document.getElementById("ec55-tab-form").addEventListener("click", () => selectTab55("form"));
  document.getElementById("ec55-tab-base").addEventListener("click", () => selectTab55("base"));
  document.getElementById("ec55-show-all").addEventListener("click", () => task49(showAll55));
  for (const [id, target] of [["ec55-tab-form", "base"], ["ec55-tab-base", "form"]]) {
    document.getElementById(id).addEventListener("keydown", event => {
      if (["ArrowLeft", "ArrowRight"].includes(event.key)) {
        event.preventDefault(); selectTab55(target);
        document.getElementById("ec55-tab-" + target).focus();
      }
    });
  }
  document.getElementById("ec56-clear-base").addEventListener("click", () => task49(() => clearBase56(false)));
  document.getElementById("ec56-clear-confirm").addEventListener("click", () => task49(() => clearBase56(true)));
  document.getElementById("ec56-clear-cancel").addEventListener("click", () => {
    cancelClearBase56(); message49("Limpieza de Base cancelada. Todos los registros se conservaron.");
  });
  document.getElementById("ec58-pdf-horizontal").addEventListener("click", () => task49(pdfHorizontal58, false));
  document.getElementById("ec59-pdf-vertical").addEventListener("click", () => task49(pdfVertical59, false));
  selectTab55("form");
  message49("Listo. Complete el formulario de Excel y pulse Cargar / Actualizar aquí.");
});
function section49(id, show) { document.getElementById(id).hidden = !show; }
function message49(text, ok = false) {
  const box = document.getElementById("ec49-message");
  box.textContent = text; box.className = ok ? "message success" : "message";
  console.log("ENTREGA 64: " + text);
}
async function task49(action, editBaseAudit = true) {
  if (ec49.busy) return;
  ec49.busy = true;
  document.querySelectorAll("#ec49 button").forEach(button => button.disabled = true);
  try { if (editBaseAudit) await withProtectedSheets61(action); else await action(); }
  catch (error) {
    cancelClearBase56();
    cancelNew54();
    cancelDelete53();
    ec49.pending = null; ec49.searchPending = null;
    section49("ec49-confirm-box", false); section49("ec52-search-confirm-box", false);
    message49(String(error && error.message ? error.message : error));
    console.log(error);
  } finally {
    ec49.busy = false;
    document.querySelectorAll("#ec49 button").forEach(button => button.disabled = false);
  }
}
async function identity49() {
  if (!Office.auth || typeof Office.auth.getAuthContext !== "function") throw new Error("No se puede consultar el usuario Microsoft en esta sesión.");
  const auth = await Office.auth.getAuthContext();
  const user = String(auth.userPrincipalName || "").trim().toLowerCase();
  const objectId = String(auth.userObjectId || "").trim();
  if (!user || !objectId) throw new Error("Microsoft no devolvió el usuario y su identificador. No se guardó nada.");
  document.getElementById("ec49-user").textContent = user;
  return { user, objectId, tenantId: String(auth.tenantId || "") };
}
function fieldIndex49(s, label) {
  const i = s.fields.findIndex(f => String(f[3]).trim() === label);
  if (i < 0) throw new Error("Falta el campo " + label + " en EC_Campos.");
  return i;
}
function replica49(f) { return String(f[7] || "").trim() !== ""; }
function replicate49(fields, values) {
  const result = values.slice();
  for (let i = 0; i < fields.length; i++) {
    if (!replica49(fields[i])) continue;
    const source = fields.findIndex(f => String(f[0]) === String(fields[i][7]));
    if (source < 0 || replica49(fields[source])) throw new Error("Revise el origen de la réplica " + fields[i][0] + ".");
    result[i] = result[source];
  }
  return result;
}
function stateRule49(block, label, motive) {
  block = block.trim(); label = label.trim(); motive = motive.trim();
  const identification = ["CUIL", "Nombre", "Apellidos"].includes(label);
  const auto = ["Usuario Carga", "Fecha/Hora Carga"].includes(label);
  const enabled = block !== "Persona" || motive === "Ingreso" || motive === "Actualización organizativa" || identification;
  const addressOptional = block === "Persona" && ["Piso", "Departamento", "Lote/ Manzana", "Lote/Manzana"].includes(label);
  const required = !auto && !addressOptional && (
    (block === "Solicitud" && label !== "Notas") || identification ||
    (block === "Persona" && motive === "Ingreso") ||
    (block === "Posición/Especialidad" && motive !== "Actualización organizativa")
  );
  return { enabled, required, auto };
}
function dateField49(label) { return label.trim().startsWith("Fecha") && label.trim() !== "Fecha/Hora Carga"; }
function dateSerial49(value) {
  if (typeof value === "number") return Number.isFinite(value) && value >= 1 && value <= 2958465 ? Math.floor(value) : NaN;
  const text = String(value).trim();
  if (/^\d{5,7}$/.test(text)) { const n = Number(text); return n >= 1 && n <= 2958465 ? n : NaN; }
  const local = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!local && !iso) return NaN;
  const y = Number(local ? local[3] : iso[1]);
  const m = Number(local ? local[2] : iso[2]);
  const d = Number(local ? local[1] : iso[3]);
  if (y < 1900 || y > 9999 || m < 1 || m > 12 || d < 1 || d > 31) return NaN;
  const date = new Date(Date.UTC(y, m - 1, d));
  if (date.getUTCFullYear() !== y || date.getUTCMonth() !== m - 1 || date.getUTCDate() !== d) return NaN;
  return Math.round((date.getTime() - Date.UTC(1899, 11, 30)) / 86400000);
}
function text49(value, label) {
  const text = String(value).trim();
  if (!text || !dateField49(label)) return text;
  const serial = dateSerial49(value);
  if (!Number.isFinite(serial)) return text;
  const date = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
  return String(date.getUTCDate()).padStart(2, "0") + "/" + String(date.getUTCMonth() + 1).padStart(2, "0") + "/" + date.getUTCFullYear();
}
function literal49(value) { return typeof value === "string" && /^[=+\-@]/.test(value) ? "'" + value : value; }
function formats49(rows, columns, format = "@") { return Array.from({ length: rows }, () => Array(columns).fill(format)); }
function now49() { return new Date(Date.now() - 10800000).toISOString().slice(0, 19).replace("T", " ") + " UTC-03:00"; }

// Se cargan bloques completos en pocas sincronizaciones, no una lectura por celda.
async function read49(context, searchOnly = false) {
  const sheets = context.workbook.worksheets;
  const form = sheets.getItemOrNullObject("EC_Formulario");
  const meta = sheets.getItemOrNullObject("EC_Campos");
  const base = sheets.getItemOrNullObject("EC_Base");
  const audit = sheets.getItemOrNullObject("EC_Auditoria");
  for (const sheet of [form, meta, base, audit]) sheet.load("name,id");
  await context.sync();
  if ([form, meta, base, audit].some(sheet => sheet.isNullObject)) throw new Error("Faltan hojas de la instalación.");
  const map = meta.getUsedRangeOrNullObject(true); map.load("values");
  const settings = meta.getRange("P2:Q2"); settings.load("values");
  const baseline = meta.getRange("J2"); baseline.load("values");
  const selected = meta.getRange("K2:L2"); selected.load("values");
  const baseUsed = base.getUsedRangeOrNullObject(true); baseUsed.load("rowIndex,rowCount");
  const auditUsed = audit.getUsedRangeOrNullObject(true); auditUsed.load("rowIndex,rowCount");
  const auditHeads = audit.getRange("A1:E1"); auditHeads.load("text");
  for (const sheet of [form, meta, base, audit]) sheet.protection.load("protected");
  await context.sync();
  if (map.isNullObject) throw new Error("El mapa de campos está vacío.");
  const fields = map.values.slice(1).filter(row => String(row[0]).trim() !== "").map(row => row.slice(0, 8));
  const start = Number(settings.values[0][0]) || 4;
  if (!fields.length || ![4, 6].includes(start)) throw new Error("Revise el mapa de campos y el inicio de Base.");
  if (meta.protection.protected || (!searchOnly && (base.protection.protected || audit.protection.protected))) throw new Error("Desproteja EC_Campos, Base y Auditoría antes de guardar, como en el botón actual.");
  const formRows = fields.filter(f => !replica49(f)).map(f => Number(f[4]));
  if (formRows.some(r => !Number.isInteger(r) || r < 7 || r > 10000)) throw new Error("El mapa contiene una fila de formulario inválida.");
  const maxFormRow = Math.max(...formRows);
  const formData = form.getRangeByIndexes(0, 2, maxFormRow, 1); formData.load("values");
  const headings = base.getRangeByIndexes(start - 2, 0, 1, fields.length); headings.load("text");
  const end = Math.max(start - 1, baseUsed.isNullObject ? start - 1 : baseUsed.rowIndex + baseUsed.rowCount);
  if ((end - start + 1) * fields.length > 500000) throw new Error("El rango usado de Base es demasiado grande para esta prueba. Revise las filas con contenido.");
  let baseData = null;
  if (end >= start) { baseData = base.getRangeByIndexes(start - 1, 0, end - start + 1, fields.length); baseData.load("values"); }
  const names = searchOnly ? [] : Array.from(new Set(fields.map(f => String(f[5] || "").trim()).filter(v => v !== "")));
  const named = names.map(name => { const item = context.workbook.names.getItemOrNullObject(name); item.load("type"); return { name, item }; });
  await context.sync();
  if (headings.text[0].some((v, i) => v !== String(fields[i][3]))) throw new Error("Base no coincide con EC_Campos. No se guardó nada.");
  const expected = ["Fecha/Hora Carga", "Usuario Carga", "Área Solicitante", "Motivo de Solicitud", "Detalle de campos cargados o modificados"];
  if (!searchOnly && auditHeads.text[0].some((v, i) => v !== expected[i])) throw new Error("Auditoría tiene otra estructura.");
  const catalogRanges = named.map(entry => {
    if (entry.item.isNullObject || String(entry.item.type).toLowerCase() !== "range") throw new Error("No se encontró el catálogo " + entry.name + ".");
    const range = entry.item.getRange(); range.load("text"); return { name: entry.name, range };
  });
  if (catalogRanges.length) await context.sync();
  const catalogs = {};
  catalogRanges.forEach(entry => catalogs[entry.name] = entry.range.text.map(row => row[0]));
  const rawBase = baseData ? baseData.values.map(row => row.slice()) : [];
  const raw = replicate49(fields, fields.map(f => replica49(f) ? "" : text49(formData.values[Number(f[4]) - 1][0], String(f[3]))));
  const rows = rawBase.map(row => replicate49(fields, row.map((v, i) => text49(v, String(fields[i][3])))));
  return {
    fields, start, end, rawBase, rows, raw, catalogs, baseline: String(baseline.values[0][0]),
    auditNext: Math.max(1, auditUsed.isNullObject ? 1 : auditUsed.rowIndex + auditUsed.rowCount),
    formProtected: form.protection.protected, selectedRow: Number(selected.values[0][0]), selectedRecord: String(selected.values[0][1]),
    fixedG: Number(settings.values[0][1]),
    baseId: base.id, metaId: meta.id, formId: form.id, auditId: audit.id
  };
}
function plan49(s, identity) {
  const raw = s.raw.slice();
  raw[fieldIndex49(s, "ID Solicitante")] = identity.user;
  const motive = raw[fieldIndex49(s, "Motivo de la solicitud")];
  if (!["Ingreso", "Reingreso", "Transferencia", "Actualización organizativa"].includes(motive)) throw new Error("Seleccione un motivo válido en el formulario.");
  const ci = fieldIndex49(s, "CUIL");
  const errors = []; let changes = 0;
  const candidate = s.fields.map((f, i) => {
    const label = String(f[3]).trim(); const v = raw[i];
    const rule = stateRule49(String(f[1]), label, motive);
    if (rule.auto) return "";
    if (replica49(f)) return v;
    if (rule.required && !v) errors.push(label + ": obligatorio.");
    if (v && String(f[1]) !== "Solicitud" && !["CUIL", "Nombre", "Apellidos"].includes(label)) changes++;
    if (v && label === "CUIL" && !/^\d{11}$/.test(v)) errors.push("CUIL: ingrese 11 dígitos.");
    if (v && label.startsWith("DNI") && !/^\d{8}$/.test(v)) errors.push("DNI: ingrese 8 dígitos.");
    if (v && label === "CBU" && !/^\d{22}$/.test(v)) errors.push("CBU: ingrese 22 dígitos.");
    if (v && dateField49(label) && !Number.isFinite(dateSerial49(v))) errors.push(label + ": fecha inválida.");
    const catalog = String(f[5] || "").trim();
    if (v && catalog && (!s.catalogs[catalog] || !s.catalogs[catalog].includes(v))) errors.push(label + ": valor fuera del desplegable.");
    return v;
  });
  if (motive === "Actualización organizativa" && !changes) errors.push("Complete al menos un campo que quiera modificar.");
  if (errors.length) throw new Error("No se guardó. Revise:\n" + errors.join("\n"));
  const matches = s.rows.map((r, i) => r[ci] === candidate[ci] ? i : -1).filter(i => i >= 0);
  if (matches.length > 1) throw new Error("Hay más de un registro para ese CUIL. Revise Base antes de actualizar.");
  const exists = matches.length === 1;
  const before = exists ? s.rows[matches[0]] : s.fields.map(() => "");
  const row = exists ? s.start + matches[0] : s.end + 1;
  const autos = s.fields.map((f, i) => ["Usuario Carga", "Fecha/Hora Carga"].includes(String(f[3])) ? i : -1).filter(i => i >= 0);
  const diff = s.fields.map((f, i) => !autos.includes(i) && before[i] !== candidate[i]
    ? f[1] + " / " + f[3] + " [" + f[0] + "]: " + (before[i] || "(vacío)") + " → " + (candidate[i] || "(vacío)") : "").filter(v => v !== "");
  return { raw, candidate, motive, ci, exists, before, row, autos, diff };
}
function signature49(s) {
  // Excluir identidad y marcas automáticas de la comparación del formulario.
  const raw = s.raw.map((v, i) => ["ID Solicitante", "Usuario Carga", "Fecha/Hora Carga"].includes(String(s.fields[i][3])) ? "" : v);
  return JSON.stringify([s.baseId, s.metaId, s.formId, s.auditId, s.fields, s.start, s.end, s.rawBase, raw, s.catalogs]);
}
async function start49() {
  cancelClearBase56();
  cancelNew54();
  cancelDelete53();
  ec49.searchPending = null; section49("ec52-search-confirm-box", false);
  ec49.pending = null; ec49.saved = null;
  section49("ec49-confirm-box", false); section49("ec49-clean-box", false);
  message49("Consultando usuario y validando el formulario…");
  const identity = await identity49();
  const s = await Excel.run(context => read49(context));
  const plan = plan49(s, identity);
  if (plan.exists && !plan.diff.length) { message49("No hay cambios para actualizar.", true); return; }
  if (plan.exists) {
    ec49.pending = { signature: signature49(s), identity, created: Date.now() };
    document.getElementById("ec49-confirm-text").textContent = "El CUIL " + plan.candidate[plan.ci] + " ya está cargado. ¿Desea reemplazar TODOS sus campos con lo que está en el formulario? Los campos vacíos también borrarán los valores anteriores.";
    section49("ec49-confirm-box", true);
    message49("Solicitud existente. Confirme la actualización o cancele para buscar y revisar el registro.");
    return;
  }
  await save49(s, identity, false);
}
async function confirm49() {
  const pending = ec49.pending;
  if (!pending || Date.now() - pending.created > 900000) throw new Error("La confirmación venció. Pulse Cargar / Actualizar para revisar otra vez.");
  const identity = await identity49();
  if (identity.objectId !== pending.identity.objectId || identity.user !== pending.identity.user || identity.tenantId !== pending.identity.tenantId) throw new Error("Cambió la cuenta Microsoft. Vuelva a iniciar la operación.");
  message49("Verificando que el formulario y Base no hayan cambiado…");
  const s = await Excel.run(context => read49(context));
  if (signature49(s) !== pending.signature) throw new Error("Cambió el formulario, el mapa, un catálogo o Base desde el aviso. No se sobrescribió nada. Pulse Cargar / Actualizar para revisar y confirmar otra vez.");
  await save49(s, identity, true);
}
async function save49(previous, identity, allowUpdate) {
  message49("Guardando solicitud y registrando el usuario…");
  const outcome = await Excel.run(async context => {
    const s = await read49(context);
    if (signature49(s) !== signature49(previous)) throw new Error("Cambió el formulario o Base durante la operación. Revise y vuelva a cargar.");
    const plan = plan49(s, identity);
    if (plan.exists && !allowUpdate) throw new Error("El CUIL acaba de ser cargado por otra operación. Vuelva a pulsar Cargar / Actualizar para revisarlo.");
    if (plan.exists && !plan.diff.length) throw new Error("No hay cambios para actualizar.");
    const values = plan.candidate.slice(); const stamp = now49();
    values[fieldIndex49(s, "Usuario Carga")] = identity.user;
    values[fieldIndex49(s, "Fecha/Hora Carga")] = stamp;
    const detail = (plan.exists ? "ACTUALIZAR" : "AGREGAR") + " SOLICITUD | CUIL " + values[plan.ci] + " | Cuenta Microsoft " + identity.objectId + " | " + plan.diff.join("; ");
    if (detail.length > 32000) throw new Error("No se guardó: reduzca las notas extensas.");
    const base = context.workbook.worksheets.getItem("EC_Base");
    const audit = context.workbook.worksheets.getItem("EC_Auditoria");
    const target = base.getRangeByIndexes(plan.row - 1, 0, 1, s.fields.length);
    const auditTarget = audit.getRangeByIndexes(s.auditNext, 0, 1, 5);
    target.load("values"); auditTarget.load("values");
    await context.sync();
    const beforeRaw = target.values.map(row => row.slice());
    const auditBefore = auditTarget.values.map(row => row.slice());
    if (plan.exists && JSON.stringify(beforeRaw[0]) !== JSON.stringify(s.rawBase[plan.row - s.start])) throw new Error("El registro cambió durante la operación. Vuelva a cargar y confirmar.");
    if (!plan.exists && beforeRaw[0].some(v => String(v).trim() !== "")) throw new Error("La fila de destino ya fue ocupada. Vuelva a cargar.");
    if (auditBefore[0].some(v => String(v).trim() !== "")) throw new Error("Cambió Auditoría durante la operación. Vuelva a cargar.");
    let attempted = false;
    try {
      attempted = true;
      target.numberFormat = formats49(1, s.fields.length);
      target.values = [values.map(v => literal49(v))];
      auditTarget.numberFormat = formats49(1, 5);
      auditTarget.values = [[stamp, identity.user, plan.raw[fieldIndex49(s, "Área solicitante")], plan.motive, detail].map(v => literal49(v))];
      await context.sync();
      target.load("values"); auditTarget.load("values");
      await context.sync();
      const actual = target.values[0].map((v, i) => text49(v, String(s.fields[i][3])));
      if (JSON.stringify(actual) !== JSON.stringify(values) || auditTarget.values[0][1] !== identity.user || auditTarget.values[0][4] !== detail) throw new Error("No se pudo verificar lo escrito en Base y Auditoría.");
    } catch (error) {
      if (attempted) {
        try {
          // Restaurar solo si las celdas siguen mostrando esta operación o su valor previo.
          target.load("values"); auditTarget.load("values"); await context.sync();
          const ownBase = JSON.stringify(target.values[0].map((v, i) => text49(v, String(s.fields[i][3])))) === JSON.stringify(values);
          const oldBase = JSON.stringify(target.values) === JSON.stringify(beforeRaw);
          const ownAudit = auditTarget.values[0][4] === detail;
          const oldAudit = JSON.stringify(auditTarget.values) === JSON.stringify(auditBefore);
          if ((!ownBase && !oldBase) || (!ownAudit && !oldAudit)) throw new Error("Se detectaron cambios ajenos; no se restauró automáticamente.");
          if (ownBase) target.values = beforeRaw.map(row => row.map(v => literal49(v)));
          if (ownAudit) auditTarget.values = auditBefore.map(row => row.map(v => literal49(v)));
          await context.sync();
        } catch (rollback) {
          throw new Error("No se confirmó el guardado ni se pudo restaurar con seguridad. Revise Base y Auditoría antes de repetir. " + String(rollback));
        }
      }
      throw new Error("No se completó el guardado; se restauraron los valores previos. " + String(error));
    }
    // Base y Auditoría ya están verificadas. Los ajustes visuales no repiten la escritura.
    const warnings = [];
    try { view49(context, s, plan, values); await context.sync(); }
    catch (error) { warnings.push("La solicitud YA está guardada; no se pudo ajustar la vista. Use Mostrar todos los campos. " + String(error)); }
    let formReady = false;
    try { await fillForm49(context, s, values, plan); formReady = true; }
    catch (error) { warnings.push("La solicitud YA está guardada; no se pudo preparar el formulario. Revise Base antes de repetir. " + String(error)); }
    return { s, plan, values, warnings, formReady, identity };
  });
  ec49.pending = null; section49("ec49-confirm-box", false);
  message49((outcome.plan.exists ? "Solicitud actualizada" : "Solicitud cargada") + " correctamente en Base, fila " + outcome.plan.row + ". Usuario: " + identity.user + "." + (outcome.warnings.length ? "\n" + outcome.warnings.join("\n") : ""), true);
  if (outcome.formReady) {
    // Huella que permite limpiar únicamente la carga recién guardada.
    ec49.saved = { identity, signature: JSON.stringify([outcome.s.fields, outcome.s.start, outcome.s.baseId, outcome.s.formId]), values: outcome.values, row: outcome.plan.row };
    section49("ec49-clean-box", true);
  }
}

// La vista de Base conserva el ajuste de la entrega 41: unión de campos con contenido.
function view49(context, s, plan, values, showAll = false) {
  const base = context.workbook.worksheets.getItem("EC_Base");
  const rows = s.rawBase.map(row => row.slice());
  if (plan) rows[plan.row - s.start] = values;
  const visible = s.fields.map((f, c) => showAll || (s.start === 6 && c < 4) || rows.some(row => row && String(row[c]).trim() !== ""));
  const titles = base.getRangeByIndexes(s.start - 4, 0, 2, s.fields.length);
  titles.unmerge(); titles.clear("Contents");
  for (let level = 0; level < 2; level++) {
    let first = 0;
    while (first < s.fields.length) {
      const block = String(s.fields[first][1]);
      const label = String(s.fields[first][level === 0 ? 1 : 2]);
      let end = first + 1;
      while (end < s.fields.length && String(s.fields[end][1]) === block && (level === 0 || String(s.fields[end][2]) === label)) end++;
      const positions = [];
      for (let c = first; c < end; c++) if (visible[c]) positions.push(c);
      if (positions.length && label.trim() !== "") {
        const left = positions[0]; const right = positions[positions.length - 1];
        if (right > left) base.getRangeByIndexes(s.start - 4 + level, left, 1, right - left + 1).merge(false);
        base.getRangeByIndexes(s.start - 4 + level, left, 1, 1).values = [[literal49(label)]];
      }
      first = end;
    }
  }
  [16, 14, 11].forEach((size, level) => {
    const font = base.getRangeByIndexes(s.start - 4 + level, 0, 1, s.fields.length).format.font;
    font.size = size; font.bold = true;
  });
  let first = 0;
  while (first < s.fields.length) {
    const block = String(s.fields[first][1]).trim(); let end = first + 1;
    while (end < s.fields.length && String(s.fields[end][1]).trim() === block) end++;
    if (["Persona", "Posición/Especialidad"].includes(block)) {
      const range = base.getRangeByIndexes(s.start - 4, first, 3, end - first);
      range.format.fill.color = block === "Persona" ? "#595959" : "#922D3D";
      range.format.font.color = "#FFFFFF";
    }
    first = end;
  }
  const last = Math.max(s.end, plan ? plan.row : s.start - 1);
  for (let c = 0; c < visible.length; c++) {
    base.getRangeByIndexes(0, c, 1, 1).columnHidden = !visible[c];
    if (visible[c] && c !== 6) base.getRangeByIndexes(s.start - 2, c, last - s.start + 2, 1).format.autofitColumns();
  }
  if (Number.isFinite(s.fixedG) && s.fixedG > 0) base.getRange("G:G").format.columnWidth = s.fixedG;
  if (last < s.start) return;
  const data = base.getRangeByIndexes(s.start - 1, 0, last - s.start + 1, s.fields.length);
  data.format.fill.color = "#FFFFFF";
  for (const side of ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"]) {
    const border = data.format.borders.getItem(side); border.style = "Continuous"; border.weight = "Thin"; border.color = "#000000";
  }
  data.format.borders.getItem("DiagonalDown").style = "None";
  data.format.borders.getItem("DiagonalUp").style = "None";
}
async function withForm49(context, action) {
  const form = context.workbook.worksheets.getItem("EC_Formulario");
  form.protection.load("protected"); await context.sync();
  const wasProtected = form.protection.protected;
  if (wasProtected) { form.protection.unprotect(); await context.sync(); }
  try { await action(form); }
  finally { if (wasProtected) { form.protection.protect(formProtection49); await context.sync(); } }
}
async function fillForm49(context, s, values, plan, notice = "") {
  await withForm49(context, async form => {
    for (let i = 0; i < s.fields.length; i++) {
      const f = s.fields[i]; if (replica49(f)) continue;
      const label = String(f[3]); const cell = form.getRange("C" + String(f[4]));
      if (dateField49(label)) {
        cell.numberFormat = [["dd/mm/yyyy"]];
        const serial = dateSerial49(values[i]);
        cell.values = [[values[i] && Number.isFinite(serial) ? serial : literal49(values[i])]];
      } else { cell.numberFormat = [["@"]]; cell.values = [[literal49(values[i])]]; }
      const state = stateRule49(String(f[1]), label, plan.motive);
      cell.format.protection.locked = state.auto;
      cell.format.fill.color = state.auto || !state.enabled ? "#FFFFFF" : state.required ? "#D9D9D9" : "#FFF2CC";
    }
    const meta = context.workbook.worksheets.getItem("EC_Campos");
    meta.getRange("K2").values = [[plan.row]];
    meta.getRange("L2").values = [[JSON.stringify(values)]];
    meta.getRange("J2").values = [[JSON.stringify(plan.candidate.map((v, i) => plan.autos.includes(i) ? "" : v))]];
    // Se conserva la zona de aviso usada por los botones existentes.
    const box = form.getRange("B5:C6"); box.unmerge(); box.merge(false);
    form.getRange("B5").values = [[notice || "Solicitud " + (plan.exists ? "actualizada" : "cargada") + " en Base, fila " + plan.row + ". Usuario Microsoft: " + values[fieldIndex49(s, "ID Solicitante")] + "."]];
    box.format.wrapText = true; box.format.font.size = 12; box.format.font.bold = true;
    box.format.font.color = "#375623"; box.format.fill.color = "#E2F0D9";
    form.getRange("5:5").format.rowHeight = 75;
    form.getRange("6:6").format.rowHeight = 12;
    await context.sync();
  });
}
async function clean49() {
  cancelClearBase56();
  cancelNew54();
  cancelDelete53();
  const saved = ec49.saved;
  if (!saved) throw new Error("No hay una carga recién guardada para limpiar.");
  const identity = await identity49();
  if (identity.objectId !== saved.identity.objectId || identity.tenantId !== saved.identity.tenantId) throw new Error("Cambió la cuenta. No se limpió el formulario.");
  await Excel.run(async context => {
    // Solo se consulta Base; no exigir desproteger Base ni Auditoría para limpiar el formulario.
    const s = await read49(context, true);
    if (JSON.stringify([s.fields, s.start, s.baseId, s.formId]) !== saved.signature || JSON.stringify(s.raw) !== JSON.stringify(saved.values)) throw new Error("El formulario cambió después del guardado. No se limpió para conservar esos cambios. Guarde o use Nueva Carga.");
    await withForm49(context, async form => {
      const keep = ["Motivo de la solicitud", "Fecha de solicitud", "Área solicitante", "ID Solicitante"];
      for (const f of s.fields) if (!replica49(f) && !keep.includes(String(f[3]))) form.getRange("C" + String(f[4])).clear("Contents");
      const meta = context.workbook.worksheets.getItem("EC_Campos");
      meta.getRange("J2:L2").clear("Contents");
      form.getRange("B5").values = [["Formulario listo para una nueva carga. Se conservaron los datos principales de Solicitud."]];
      await context.sync();
    });
  });
  ec49.saved = null; section49("ec49-clean-box", false);
  message49("Formulario listo para otra persona. Se conservaron Motivo, Fecha de solicitud, Área solicitante e ID Solicitante.", true);
}


// Entrega 64: Buscar conserva las identidades históricas; no escribe Base ni Auditoría.
function searchSignature52(s) {
  return JSON.stringify([s.baseId, s.formId, s.metaId, s.fields, s.start, s.rawBase, s.raw, s.baseline]);
}
function searchPlan52(s, cuil) {
  const ci = fieldIndex49(s, "CUIL");
  const matches = s.rows.map((r, i) => r[ci] === cuil ? i : -1).filter(i => i >= 0);
  if (!matches.length) throw new Error(s.rows.length
    ? "No se encontró una solicitud para el CUIL " + cuil + " en esta entrega. Verifique el número o cargue una nueva solicitud."
    : "Esta entrega todavía no tiene registros cargados.");
  if (matches.length > 1) throw new Error("Hay más de una solicitud para ese CUIL. Revise Base; no se seleccionó ningún registro.");
  const current = s.raw.map((v, i) => ["Usuario Carga", "Fecha/Hora Carga"].includes(String(s.fields[i][3])) ? "" : v);
  const hasData = s.fields.some((f, i) => (String(f[1]) !== "Solicitud" || String(f[3]) === "Notas") && current[i] !== "");
  const record = s.rows[matches[0]].slice();
  const autos = s.fields.map((f, i) => ["Usuario Carga", "Fecha/Hora Carga"].includes(String(f[3])) ? i : -1).filter(i => i >= 0);
  return { record, row: s.start + matches[0], motive: record[fieldIndex49(s, "Motivo de la solicitud")],
    candidate: record, autos, exists: true, unsaved: hasData && JSON.stringify(current) !== s.baseline };
}
async function search52(confirmDiscard = false) {
  const pending = ec49.searchPending;
  const cuil = String(document.getElementById("ec52-cuil").value || "").trim();
  if (!/^\d{11}$/.test(cuil)) throw new Error("Ingrese un CUIL de 11 dígitos, sin guiones ni espacios, en el panel.");
  if (!confirmDiscard) {
    cancelClearBase56();
    cancelNew54();
    cancelDelete53();
    ec49.searchPending = null;
    ec49.pending = null; ec49.saved = null;
    section49("ec52-search-confirm-box", false);
    section49("ec49-confirm-box", false); section49("ec49-clean-box", false);
  }
  message49("Buscando el CUIL en Base…");
  const identity = await identity49();
  await Excel.run(async context => {
    const s = await read49(context, true);
    const plan = searchPlan52(s, cuil);
    if (confirmDiscard) {
      if (!pending || pending.cuil !== cuil || Date.now() - pending.created > 900000)
        throw new Error("La confirmación venció o cambió el CUIL. Pulse Buscar CUIL nuevamente.");
      if (identity.user !== pending.identity.user || identity.objectId !== pending.identity.objectId || identity.tenantId !== pending.identity.tenantId)
        throw new Error("Cambió la cuenta Microsoft. Pulse Buscar CUIL nuevamente.");
      if (searchSignature52(s) !== pending.signature)
        throw new Error("Cambió el formulario, el mapa o Base desde el aviso. No se reemplazaron los datos. Pulse Buscar CUIL nuevamente.");
    } else if (plan.unsaved) {
      ec49.searchPending = { cuil, identity, created: Date.now(), signature: searchSignature52(s) };
      document.getElementById("ec52-search-confirm-text").textContent = "Hay cambios sin guardar. ¿Desea descartarlos y traer la solicitud del CUIL " + cuil + " desde Base?";
      section49("ec52-search-confirm-box", true);
      message49("Se encontró la solicitud. Confirme si desea reemplazar los cambios sin guardar del formulario.");
      return;
    }
    const message = "Solicitud recuperada de Base, fila " + plan.row + ". Revise o modifique los datos y pulse Cargar / Actualizar para guardar. Base y Auditoría se conservaron.";
    await fillForm49(context, s, plan.record, plan, message);
    context.workbook.worksheets.getItem("EC_Formulario").activate();
    await context.sync();
    ec49.searchPending = null;
    section49("ec52-search-confirm-box", false);
    message49(message, true);
  });
}


// Entrega 64: eliminación de un registro seleccionado, con cuenta Microsoft y confirmación.
function cancelDelete53() {
  ec49.deletePending = null; section49("ec53-delete-confirm-box", false);
}
function deleteSignature53(s) {
  return JSON.stringify([searchSignature52(s), s.auditId, s.auditNext, s.selectedRow, s.selectedRecord]);
}
function deletePlan53(s) {
  const ci = fieldIndex49(s, "CUIL");
  const cuil = s.raw[ci];
  if (!/^\d{11}$/.test(cuil)) throw new Error("No se eliminó nada. Busque un CUIL de 11 dígitos para recuperar la solicitud que desea eliminar.");
  const matches = s.rows.map((r, i) => r[ci] === cuil ? i : -1).filter(i => i >= 0);
  if (matches.length !== 1) throw new Error(matches.length > 1
    ? "No se eliminó nada: hay varios registros para ese CUIL. Revise Base."
    : "No se encontró el CUIL " + cuil + " en esta entrega. No se eliminó ningún registro.");
  const index = matches[0], row = s.start + index, before = s.rows[index];
  if (s.selectedRow !== row || s.selectedRecord !== JSON.stringify(before))
    throw new Error("No se eliminó nada. Utilice Buscar CUIL para recuperar el registro que desea eliminar. Si cambió el CUIL o Base fue modificada después de buscar, búsquelo nuevamente.");
  const part = label => { const i = s.fields.findIndex(f => String(f[3]).trim() === label); return i < 0 ? "" : before[i]; };
  const detail = "ELIMINAR SOLICITUD | CUIL " + cuil + " | " + s.fields.map((f, i) => before[i] !== ""
    ? f[1] + " / " + f[3] + " [" + f[0] + "]: " + before[i] + " → (eliminado)" : "").filter(v => v !== "").join("; ");
  return { cuil, index, row, before, person: (part("Nombre") + " " + part("Apellidos")).trim(),
    area: part("Área solicitante"), motive: part("Motivo de la solicitud"), detail };
}
async function checkDeleteAccess53(context) {
  const sheets = context.workbook.worksheets;
  const meta = sheets.getItem("EC_Campos"), base = sheets.getItem("EC_Base"), audit = sheets.getItem("EC_Auditoria");
  for (const sheet of [meta, base, audit]) sheet.protection.load("protected");
  const heads = audit.getRange("A1:E1"); heads.load("text");
  await context.sync();
  if ([meta, base, audit].some(sheet => sheet.protection.protected))
    throw new Error("No se eliminó nada: desproteja Base, Auditoría y EC_Campos antes de continuar.");
  const expected = ["Fecha/Hora Carga", "Usuario Carga", "Área Solicitante", "Motivo de Solicitud", "Detalle de campos cargados o modificados"];
  if (heads.text[0].some((v, i) => v !== expected[i])) throw new Error("No se eliminó nada: Auditoría tiene una estructura diferente.");
}
async function delete53(confirm = false) {
  const pending = ec49.deletePending;
  if (!confirm) {
    cancelClearBase56();
    cancelNew54();
    cancelDelete53(); ec49.pending = null; ec49.searchPending = null; ec49.saved = null;
    section49("ec49-confirm-box", false); section49("ec52-search-confirm-box", false); section49("ec49-clean-box", false);
  }
  const identity = await identity49();
  message49(confirm ? "Verificando y eliminando la solicitud…" : "Revisando la solicitud seleccionada…");
  const outcome = await Excel.run(async context => {
    const s = await read49(context, true);
    const plan = deletePlan53(s);
    await checkDeleteAccess53(context);
    const detail = plan.detail + " | Cuenta Microsoft " + identity.objectId;
    if (detail.length > 32000) throw new Error("No se eliminó nada: el registro es demasiado extenso para guardar su detalle completo en Auditoría.");
    if (!confirm) {
      ec49.deletePending = { identity, signature: deleteSignature53(s), created: Date.now() };
      document.getElementById("ec53-delete-confirm-text").textContent = "¿Desea eliminar la solicitud del CUIL " + plan.cuil + " · " + plan.person + "? Se quitará de esta entrega y su contenido quedará registrado en Auditoría con su usuario Microsoft. Después se limpiará el formulario conservando Motivo, Fecha de solicitud, Área e ID Solicitante. Los cambios sin guardar del formulario también se descartarán.";
      section49("ec53-delete-confirm-box", true);
      message49("No se eliminó nada todavía. Revise el CUIL y confirme o cancele.");
      return null;
    }
    if (!pending || Date.now() - pending.created > 900000) throw new Error("La confirmación venció. Pulse Eliminar solicitud nuevamente.");
    if (identity.user !== pending.identity.user || identity.objectId !== pending.identity.objectId || identity.tenantId !== pending.identity.tenantId)
      throw new Error("Cambió la cuenta Microsoft. No se eliminó nada; vuelva a iniciar la operación.");
    if (deleteSignature53(s) !== pending.signature)
      throw new Error("Cambió el formulario, Base, la selección o Auditoría desde el aviso. No se eliminó nada. Busque la solicitud y revise nuevamente.");
    const base = context.workbook.worksheets.getItem("EC_Base"), audit = context.workbook.worksheets.getItem("EC_Auditoria");
    const tail = s.rawBase.slice(plan.index).map(r => r.slice());
    const remaining = tail.slice(1); remaining.push(s.fields.map(() => ""));
    const target = base.getRangeByIndexes(plan.row - 1, 0, tail.length, s.fields.length);
    const auditTarget = audit.getRangeByIndexes(s.auditNext, 0, 1, 5);
    target.load("values"); auditTarget.load("values");
    await context.sync();
    const auditBefore = auditTarget.values.map(r => r.slice());
    if (JSON.stringify(target.values) !== JSON.stringify(tail) || auditBefore[0].some(v => String(v).trim() !== ""))
      throw new Error("Base o Auditoría cambiaron durante la revisión. No se eliminó nada; vuelva a buscar la solicitud.");
    const stamp = now49(), auditValues = [[stamp, identity.user, plan.area, plan.motive, detail]];
    try {
      auditTarget.numberFormat = formats49(1, 5);
      auditTarget.values = auditValues.map(r => r.map(v => literal49(v)));
      target.values = remaining.map(r => r.map(v => literal49(v)));
      await context.sync();
      target.load("values"); auditTarget.load("values"); await context.sync();
      if (JSON.stringify(target.values) !== JSON.stringify(remaining) || JSON.stringify(auditTarget.values) !== JSON.stringify(auditValues))
        throw new Error("No se pudo verificar la eliminación y su entrada de Auditoría.");
    } catch (error) {
      try {
        target.load("values"); auditTarget.load("values"); await context.sync();
        const ownBase = JSON.stringify(target.values) === JSON.stringify(remaining);
        const oldBase = JSON.stringify(target.values) === JSON.stringify(tail);
        const ownAudit = JSON.stringify(auditTarget.values) === JSON.stringify(auditValues);
        const oldAudit = JSON.stringify(auditTarget.values) === JSON.stringify(auditBefore);
        if ((!ownBase && !oldBase) || (!ownAudit && !oldAudit)) throw new Error("Se detectaron cambios ajenos y no se restauró automáticamente.");
        if (ownBase) target.values = tail.map(r => r.map(v => literal49(v)));
        if (ownAudit) auditTarget.values = auditBefore.map(r => r.map(v => literal49(v)));
        await context.sync();
      } catch (rollback) {
        throw new Error("No se pudo confirmar ni restaurar la eliminación con seguridad. Revise Base y Auditoría antes de repetir. " + String(rollback));
      }
      throw new Error("No se completó la eliminación; se restauraron los valores anteriores. " + String(error));
    }
    // Borrado y auditoría verificados: un fallo posterior no debe repetir la eliminación.
    let warning = "";
    try {
      const latest = await read49(context, true);
      const meta = context.workbook.worksheets.getItem("EC_Campos");
      meta.getRange("J2:L2").clear("Contents"); await context.sync();
      if (JSON.stringify(latest.fields) !== JSON.stringify(s.fields) || JSON.stringify(latest.raw) !== JSON.stringify(s.raw))
        throw new Error("El formulario cambió durante la operación; se conservaron esos datos.");
      await withForm49(context, async form => {
        const keep = ["Motivo de la solicitud", "Fecha de solicitud", "Área solicitante", "ID Solicitante"];
        for (const f of s.fields) if (!replica49(f) && !keep.includes(String(f[3]))) form.getRange("C" + String(f[4])).clear("Contents");
        form.getRange("B5").values = [["Solicitud del CUIL " + plan.cuil + " eliminada. Usuario Microsoft: " + identity.user + ". Formulario listo para una nueva carga."]];
        const box = form.getRange("B5:C6"); box.format.font.color = "#375623"; box.format.fill.color = "#E2F0D9";
        const oldPending = context.workbook.worksheets.getItemOrNullObject("EC_Confirmacion");
        oldPending.load("name"); await context.sync();
        if (!oldPending.isNullObject) {
          oldPending.protection.load("protected"); await context.sync();
          if (!oldPending.protection.protected) {
            const used = oldPending.getUsedRangeOrNullObject(true); used.load("address"); await context.sync();
            if (!used.isNullObject) used.clear("Contents");
          }
        }
        await context.sync();
      });
    } catch (error) {
      warning = " La solicitud YA FUE ELIMINADA y registrada en Auditoría; no se pudo completar la limpieza del formulario. No repita la eliminación. " + String(error && error.message ? error.message : error);
    }
    return { cuil: plan.cuil, warning };
  });
  if (!outcome) return;
  cancelDelete53(); ec49.pending = null; ec49.searchPending = null; ec49.saved = null;
  section49("ec49-confirm-box", false); section49("ec52-search-confirm-box", false); section49("ec49-clean-box", false);
  message49("Solicitud del CUIL " + outcome.cuil + " eliminada de esta entrega. Auditoría registró el contenido eliminado y el usuario " + identity.user + "." + (outcome.warning || " Formulario listo para una nueva carga."), true);
}


// Entrega 64: Nueva carga depende solo del formulario y su mapa, como el script 15.
function cancelNew54() {
  ec49.newPending = null; section49("ec54-new-confirm-box", false);
}
async function readNew54(context) {
  const sheets = context.workbook.worksheets;
  const form = sheets.getItemOrNullObject("EC_Formulario"), meta = sheets.getItemOrNullObject("EC_Campos");
  form.load("name,id"); meta.load("name,id"); await context.sync();
  if (form.isNullObject || meta.isNullObject) throw new Error("No se encontró el formulario o el mapa de campos.");
  meta.protection.load("protected");
  const map = meta.getUsedRangeOrNullObject(true); map.load("values");
  const state = meta.getRange("J2:L2"); state.load("values"); await context.sync();
  if (meta.protection.protected) throw new Error("No se limpió el formulario: desproteja EC_Campos antes de continuar.");
  if (map.isNullObject) throw new Error("El mapa de campos está vacío.");
  const fields = map.values.slice(1).filter(f => String(f[0]).trim() !== "").map(f => f.slice(0, 8));
  if (!fields.length) throw new Error("No se encontraron campos en el mapa.");
  const rows = fields.filter(f => !replica49(f)).map(f => Number(f[4]));
  if (!rows.length || rows.some(r => !Number.isInteger(r) || r < 7 || r > 10000)) throw new Error("El mapa contiene una fila de formulario inválida.");
  const data = form.getRangeByIndexes(0, 2, Math.max(...rows), 1); data.load("values"); await context.sync();
  const cells = fields.map(f => replica49(f) ? "" : data.values[Number(f[4]) - 1][0]);
  const raw = replicate49(fields, fields.map((f, i) => text49(cells[i], String(f[3]))));
  const current = raw.map((v, i) => ["Usuario Carga", "Fecha/Hora Carga"].includes(String(fields[i][3])) ? "" : v);
  const hasData = fields.some((f, i) => (String(f[1]) !== "Solicitud" || String(f[3]) === "Notas") && current[i] !== "");
  return { fields, cells, raw, current, metaState: state.values.map(r => r.slice()), formId: form.id, metaId: meta.id,
    unsaved: hasData && JSON.stringify(current) !== String(state.values[0][0]) };
}
function newSignature54(s) { return JSON.stringify([s.formId, s.metaId, s.fields, s.cells, s.metaState]); }
async function clearOldConfirmation54(context) {
  const sheet = context.workbook.worksheets.getItemOrNullObject("EC_Confirmacion");
  sheet.load("name"); await context.sync();
  if (sheet.isNullObject) return;
  sheet.protection.load("protected"); await context.sync();
  if (sheet.protection.protected) return;
  const range = sheet.getUsedRangeOrNullObject(true); range.load("address"); await context.sync();
  if (!range.isNullObject) { range.clear("Contents"); await context.sync(); }
}
async function new54(confirmDiscard = false) {
  const pending = ec49.newPending;
  if (!confirmDiscard) {
    cancelClearBase56();
    cancelNew54(); cancelDelete53(); ec49.pending = null; ec49.searchPending = null; ec49.saved = null;
    section49("ec49-confirm-box", false); section49("ec52-search-confirm-box", false); section49("ec49-clean-box", false);
  }
  const identity = await identity49();
  message49("Revisando los datos del formulario…");
  await Excel.run(async context => {
    const s = await readNew54(context);
    if (confirmDiscard) {
      if (!pending || Date.now() - pending.created > 900000) throw new Error("La confirmación venció. Pulse Nueva carga nuevamente.");
      if (identity.user !== pending.identity.user || identity.objectId !== pending.identity.objectId || identity.tenantId !== pending.identity.tenantId)
        throw new Error("Cambió la cuenta Microsoft. No se limpió el formulario. Pulse Nueva carga nuevamente.");
      if (newSignature54(s) !== pending.signature) throw new Error("El formulario, el mapa o la selección cambiaron desde el aviso. No se limpió nada. Pulse Nueva carga para revisar y confirmar nuevamente.");
    } else if (s.unsaved) {
      ec49.newPending = { identity, created: Date.now(), signature: newSignature54(s) };
      section49("ec54-new-confirm-box", true);
      message49("Hay cambios sin guardar. Confirme si desea descartarlos para comenzar una nueva carga.");
      return;
    }
    const keep = ["Motivo de la solicitud", "Fecha de solicitud", "Área solicitante", "ID Solicitante"];
    const clearIndexes = s.fields.map((f, i) => !replica49(f) && !keep.includes(String(f[3])) ? i : -1).filter(i => i >= 0);
    let cleared = false;
    try {
      await withForm49(context, async form => {
        // Volver a comprobar después de desproteger: el usuario puede haber seguido editando.
        const latest = await readNew54(context);
        if (newSignature54(latest) !== newSignature54(s)) throw new Error("El formulario cambió durante la revisión. No se limpió nada; vuelva a pulsar Nueva carga.");
        const meta = context.workbook.worksheets.getItem("EC_Campos");
        try {
          for (const i of clearIndexes) form.getRange("C" + String(s.fields[i][4])).clear("Contents");
          meta.getRange("J2:L2").clear("Contents");
          await context.sync();
          const actual = await readNew54(context);
          if (clearIndexes.some(i => String(actual.cells[i]) !== "") || actual.metaState[0].some(v => String(v) !== ""))
            throw new Error("No se pudo verificar la limpieza del formulario.");
          cleared = true;
        } catch (error) {
          try {
            const actual = await readNew54(context);
            if (JSON.stringify(actual.fields) !== JSON.stringify(s.fields)) throw new Error("Cambió el mapa de campos; no se restauró automáticamente.");
            if (clearIndexes.some(i => String(actual.cells[i]) !== "" && JSON.stringify(actual.cells[i]) !== JSON.stringify(s.cells[i])))
              throw new Error("Hay nuevos datos ajenos a esta limpieza; no se sobrescribieron.");
            if (JSON.stringify(actual.metaState) !== JSON.stringify(s.metaState) && actual.metaState[0].some(v => String(v) !== ""))
              throw new Error("Cambió la selección; no se restauró automáticamente.");
            for (const i of clearIndexes) if (String(actual.cells[i]) === "") form.getRange("C" + String(s.fields[i][4])).values = [[literal49(s.cells[i])]];
            meta.getRange("J2:L2").values = s.metaState.map(r => r.map(v => literal49(v)));
            await context.sync();
          } catch (rollback) {
            throw new Error("No se pudo completar ni restaurar la limpieza con seguridad. Revise los datos del formulario antes de continuar. " + String(rollback && rollback.message ? rollback.message : rollback));
          }
          throw new Error("No se completó la limpieza; se restauraron los valores anteriores. " + String(error && error.message ? error.message : error));
        }
        const box = form.getRange("B5:C6"); box.unmerge(); box.merge(false);
        form.getRange("B5").values = [["Formulario listo para una nueva carga. Se conservaron Motivo, Fecha de solicitud, Área solicitante e ID Solicitante. Base y Auditoría se conservaron."]];
        box.format.wrapText = true; box.format.font.size = 12; box.format.font.bold = true;
        box.format.font.color = "#375623"; box.format.fill.color = "#E2F0D9";
        form.getRange("5:5").format.rowHeight = 90; form.getRange("6:6").format.rowHeight = 12;
        await context.sync();
      });
    } catch (error) {
      cancelNew54();
      if (!cleared) throw error;
      message49("El formulario YA se limpió. Se conservaron los datos principales de Solicitud, Base y Auditoría. No se pudo completar el aviso o restablecer la protección: " + String(error && error.message ? error.message : error));
      return;
    }
    cancelNew54();
    let warning = "";
    try { await clearOldConfirmation54(context); }
    catch { warning = " No se pudo limpiar la confirmación del botón antiguo; utilice el panel actualizado para continuar."; }
    document.getElementById("ec52-cuil").value = "";
    message49("Formulario listo para una nueva carga. Se conservaron Motivo, Fecha de solicitud, Área solicitante e ID Solicitante. Base y Auditoría se conservaron." + warning, true);
  });
}


// Entrega 64: navegación visual del panel. Cambiar pestaña no modifica el libro.
function selectTab55(tab) {
  if (ec49.busy || !["form", "base"].includes(tab)) return;
  for (const name of ["form", "base"]) {
    const button = document.getElementById("ec55-tab-" + name);
    const panel = document.getElementById("ec55-panel-" + name);
    button.setAttribute("aria-selected", String(name === tab));
    button.tabIndex = name === tab ? 0 : -1;
    panel.hidden = name !== tab;
  }
}
// Este botón solo lee el esquema y ajusta la vista. No requiere editar Auditoría ni el formulario.
async function readBaseView55(context) {
  const sheets = context.workbook.worksheets;
  const base = sheets.getItemOrNullObject("EC_Base"), meta = sheets.getItemOrNullObject("EC_Campos");
  base.load("name,id"); meta.load("name,id"); await context.sync();
  if (base.isNullObject || meta.isNullObject) throw new Error("Faltan EC_Base o EC_Campos.");
  const map = meta.getUsedRangeOrNullObject(true); map.load("values");
  const settings = meta.getRange("P2:Q2"); settings.load("values");
  const used = base.getUsedRangeOrNullObject(true); used.load("rowIndex,rowCount");
  base.protection.load("protected"); await context.sync();
  if (base.protection.protected) throw new Error("Desproteja EC_Base para mostrar todos sus campos.");
  if (map.isNullObject) throw new Error("EC_Campos está vacío.");
  const fields = map.values.slice(1).filter(f => String(f[0]).trim() !== "").map(f => f.slice(0, 8));
  const start = Number(settings.values[0][0]) || 4;
  if (!fields.length || ![4, 6].includes(start)) throw new Error("Revise el mapa de campos y el inicio de Base.");
  const headers = base.getRangeByIndexes(start - 2, 0, 1, fields.length); headers.load("text"); await context.sync();
  if (headers.text[0].some((v, i) => v !== String(fields[i][3]))) throw new Error("Base no coincide con EC_Campos. No se ajustó la vista.");
  const end = Math.max(start - 1, used.isNullObject ? start - 1 : used.rowIndex + used.rowCount);
  if ((end - start + 1) * fields.length > 500000) throw new Error("El rango usado de Base es demasiado grande. Revise las filas con contenido antes de ajustar la vista.");
  return { fields, start, end, fixedG: Number(settings.values[0][1]), rawBase: [], baseId: base.id, metaId: meta.id };
}
async function showAll55() {
  message49("Mostrando todos los campos de Base…");
  const total = await Excel.run(async context => {
    const s = await readBaseView55(context);
    view49(context, s, null, null, true);
    await context.sync();
    context.workbook.worksheets.getItem("EC_Base").activate(); await context.sync();
    return s.fields.length;
  });
  message49("Base muestra sus " + total + " campos. La próxima carga o actualización volverá a ocultar las columnas vacías en toda la entrega. Los registros y Auditoría se conservaron.", true);
}


// Entrega 64: Limpiar Base conserva el formulario y audita cada registro eliminado.
function cancelClearBase56() {
  ec49.clearBasePending = null; section49("ec56-clear-confirm-box", false);
}
async function readClearBase56(context) {
  const s = await readBaseView55(context);
  const audit = context.workbook.worksheets.getItemOrNullObject("EC_Auditoria");
  audit.load("name,id"); await context.sync();
  if (audit.isNullObject) throw new Error("Falta EC_Auditoria. No se eliminó nada.");
  await checkDeleteAccess53(context);
  const used = audit.getUsedRangeOrNullObject(true); used.load("rowIndex,rowCount");
  let data = null;
  if (s.end >= s.start) {
    data = context.workbook.worksheets.getItem("EC_Base").getRangeByIndexes(s.start - 1, 0, s.end - s.start + 1, s.fields.length);
    data.load("values");
  }
  await context.sync();
  const raw = data ? data.values.map(r => r.slice()) : [];
  let count = raw.length;
  while (count > 0 && !raw[count - 1].some(v => String(v).trim() !== "")) count--;
  const snapshot = raw.slice(0, count);
  const records = snapshot.filter(r => r.some(v => String(v).trim() !== ""));
  // Reutilizar los ID leídos con load("name,id") y context.sync en readBaseView55.
  return { ...s, snapshot, records, auditId: audit.id, auditNext: Math.max(1, used.isNullObject ? 1 : used.rowIndex + used.rowCount) };
}
function clearSignature56(s) {
  return JSON.stringify([s.baseId, s.metaId, s.auditId, s.fields, s.start, s.end, s.snapshot, s.auditNext]);
}
function auditRows56(s, identity, stamp) {
  const find = label => s.fields.findIndex(f => String(f[3]).trim() === label);
  const ci = find("CUIL"), ai = find("Área solicitante"), mi = find("Motivo de la solicitud");
  return s.records.map(record => {
    const detail = "LIMPIAR BASE | CUIL " + (ci >= 0 ? String(record[ci]) : "sin campo CUIL")
      + " | Cuenta Microsoft " + identity.objectId + " | "
      + s.fields.map((f, i) => String(record[i]).trim() !== ""
        ? f[1] + " / " + f[3] + " [" + f[0] + "]: " + String(record[i]) + " → (eliminado)" : "").filter(v => v !== "").join("; ");
    if (detail.length > 32000) throw new Error("No se eliminó nada: uno de los registros supera el tamaño permitido para guardar su detalle completo en Auditoría.");
    return [stamp, identity.user, ai >= 0 ? record[ai] : "", mi >= 0 ? record[mi] : "", detail];
  });
}
async function clearBase56(confirm = false) {
  const pending = ec49.clearBasePending;
  if (!confirm) {
    cancelClearBase56(); cancelNew54(); cancelDelete53();
    ec49.pending = null; ec49.searchPending = null; ec49.saved = null;
    section49("ec49-confirm-box", false); section49("ec52-search-confirm-box", false); section49("ec49-clean-box", false);
  }
  const identity = await identity49();
  message49(confirm ? "Verificando la entrega y registrando su limpieza…" : "Revisando los registros de Base…");
  const outcome = await Excel.run(async context => {
    const s = await readClearBase56(context);
    if (!s.records.length) {
      cancelClearBase56();
      message49("Base ya está vacía. Se conservaron los encabezados, el formulario y la Auditoría.", true);
      return null;
    }
    // Comprobar longitud de los detalles antes de ofrecer o ejecutar la eliminación.
    const auditValues = auditRows56(s, identity, now49());
    if (!confirm) {
      ec49.clearBasePending = { identity, signature: clearSignature56(s), created: Date.now() };
      document.getElementById("ec56-clear-confirm-text").textContent = "La entrega contiene " + s.records.length
        + " registros. ¿Desea eliminarlos TODOS de Base? Si necesita conservar el informe, genérelo primero. Se registrará cada eliminación en Auditoría con su usuario Microsoft. Se conservarán encabezados, formatos, botones y los datos actuales del formulario.";
      section49("ec56-clear-confirm-box", true);
      message49("No se eliminó nada todavía. La entrega contiene " + s.records.length + " registros; confirme la limpieza o cancele.");
      return null;
    }
    if (!pending || Date.now() - pending.created > 900000) throw new Error("La confirmación venció. Pulse Limpiar Base para revisar la entrega nuevamente.");
    if (identity.user !== pending.identity.user || identity.objectId !== pending.identity.objectId || identity.tenantId !== pending.identity.tenantId)
      throw new Error("Cambió la cuenta Microsoft. No se eliminó nada; vuelva a iniciar Limpiar Base.");
    if (clearSignature56(s) !== pending.signature)
      throw new Error("Base, el mapa o Auditoría cambiaron desde el aviso. No se eliminó nada. Pulse Limpiar Base para revisar nuevamente la cantidad y confirmar.");
    const sheets = context.workbook.worksheets;
    const base = sheets.getItem("EC_Base"), audit = sheets.getItem("EC_Auditoria");
    const target = base.getRangeByIndexes(s.start - 1, 0, s.snapshot.length, s.fields.length);
    const auditTarget = audit.getRangeByIndexes(s.auditNext, 0, auditValues.length, 5);
    target.load("values"); auditTarget.load("values"); await context.sync();
    const auditBefore = auditTarget.values.map(r => r.slice());
    if (JSON.stringify(target.values) !== JSON.stringify(s.snapshot) || auditBefore.some(r => r.some(v => String(v).trim() !== "")))
      throw new Error("Base o Auditoría cambiaron durante la revisión. No se limpió la entrega; vuelva a revisar.");
    const blank = s.snapshot.map(() => s.fields.map(() => ""));
    let baseAttempted = false;
    try {
      auditTarget.numberFormat = formats49(auditValues.length, 5);
      auditTarget.values = auditValues.map(r => r.map(v => literal49(v)));
      await context.sync();
      auditTarget.load("values"); target.load("values");
      const used = base.getUsedRangeOrNullObject(true); used.load("rowIndex,rowCount");
      await context.sync();
      if (JSON.stringify(auditTarget.values) !== JSON.stringify(auditValues)) throw new Error("No se verificó el detalle completo en Auditoría.");
      const currentEnd = Math.max(s.start - 1, used.isNullObject ? s.start - 1 : used.rowIndex + used.rowCount);
      if (JSON.stringify(target.values) !== JSON.stringify(s.snapshot) || currentEnd !== s.end)
        throw new Error("Base cambió antes del vaciado; se detuvo la limpieza.");
      baseAttempted = true;
      target.clear("Contents"); await context.sync();
      target.load("values"); auditTarget.load("values"); await context.sync();
      if (JSON.stringify(target.values) !== JSON.stringify(blank) || JSON.stringify(auditTarget.values) !== JSON.stringify(auditValues))
        throw new Error("No se verificó el vaciado de Base y su detalle de Auditoría.");
    } catch (error) {
      try {
        target.load("values"); auditTarget.load("values"); await context.sync();
        const ownBase = JSON.stringify(target.values) === JSON.stringify(blank), oldBase = JSON.stringify(target.values) === JSON.stringify(s.snapshot);
        const ownAudit = JSON.stringify(auditTarget.values) === JSON.stringify(auditValues), oldAudit = JSON.stringify(auditTarget.values) === JSON.stringify(auditBefore);
        if ((baseAttempted && !ownBase && !oldBase) || (!ownAudit && !oldAudit)) throw new Error("Hay cambios ajenos; no se sobrescribieron automáticamente.");
        if (baseAttempted && ownBase) target.values = s.snapshot.map(r => r.map(v => literal49(v)));
        if (ownAudit) auditTarget.values = auditBefore.map(r => r.map(v => literal49(v)));
        await context.sync(); target.load("values"); auditTarget.load("values"); await context.sync();
        if ((baseAttempted && JSON.stringify(target.values) !== JSON.stringify(s.snapshot)) || JSON.stringify(auditTarget.values) !== JSON.stringify(auditBefore))
          throw new Error("No se pudo verificar la restauración completa.");
      } catch (rollback) {
        throw new Error("No se pudo completar ni restaurar la limpieza con seguridad. Revise Base y Auditoría antes de repetir. " + String(rollback && rollback.message ? rollback.message : rollback));
      }
      throw new Error((baseAttempted ? "No se completó la limpieza; se restauraron los registros anteriores de Base y Auditoría. " : "No se vació Base; se retiró el intento de limpieza de Auditoría y se conservaron los datos actuales. ") + String(error && error.message ? error.message : error));
    }
    let warning = "";
    try {
      sheets.getItem("EC_Campos").getRange("J2:L2").clear("Contents"); await context.sync();
      await clearOldConfirmation54(context);
    } catch (error) {
      warning = " Base YA FUE LIMPIADA y registrada en Auditoría; no se pudieron reiniciar todas las referencias auxiliares. Use Nueva carga antes de continuar. " + String(error && error.message ? error.message : error);
    }
    return { count: s.records.length, warning };
  });
  if (!outcome) return;
  cancelClearBase56(); cancelNew54(); cancelDelete53(); ec49.pending = null; ec49.searchPending = null; ec49.saved = null;
  section49("ec49-confirm-box", false); section49("ec52-search-confirm-box", false); section49("ec49-clean-box", false);
  message49("Base limpia: " + outcome.count + " registros eliminados y registrados en Auditoría con el usuario " + identity.user
    + ". Se conservaron título, instrucciones, encabezados, formatos, botones y el historial de Auditoría. El formulario conserva sus datos; use Nueva carga para comenzar otra persona." + outcome.warning, true);
}

// Entrega 64: PDF horizontal local. La hoja temporal no modifica Base ni Auditoría.
let pdfDownload58 = null;
function pdfPlan58(fields, raw) {
  const records = raw.filter(row => row.some(v => String(v).trim() !== ""));
  if (!records.length) throw new Error("Base está vacía. Cargue al menos una solicitud antes de generar el PDF.");
  const columns = fields.map((_, c) => c).filter(c => records.some(row => String(row[c]).trim() !== ""));
  if (records.some(row => columns.some(c => /[\r\n]/.test(String(row[c])))))
    throw new Error("Hay un dato con saltos de línea en Base. Corríjalo antes de generar el PDF para conservar una fila por registro al copiarlo a Excel.");
  const selected = columns.map(c => fields[c]);
  const data = records.map(row => columns.map(c => String(row[c] ?? "")));
  const widths = selected.map((f, c) => Math.max(65, Math.min(600,
    data.reduce((max, row) => Math.max(max, row[c].length), String(f[3]).length) * 7 + 24)));
  const scale = Math.min(1, 6500 / widths.reduce((a, b) => a + b, 0));
  return { fields: selected, data, widths: widths.map(w => w * scale), count: records.length };
}
async function readPdf58(context) {
  const sheets = context.workbook.worksheets;
  const base = sheets.getItemOrNullObject("EC_Base"), meta = sheets.getItemOrNullObject("EC_Campos");
  base.load("name,id"); meta.load("name,id"); await context.sync();
  if (base.isNullObject || meta.isNullObject) throw new Error("Faltan EC_Base o EC_Campos.");
  const map = meta.getUsedRangeOrNullObject(true); map.load("values");
  const setting = meta.getRange("P2"); setting.load("values");
  const used = base.getUsedRangeOrNullObject(true); used.load("rowIndex,rowCount");
  await context.sync();
  if (map.isNullObject) throw new Error("EC_Campos está vacío.");
  const fields = map.values.slice(1).filter(row => String(row[0]).trim() !== "").map(row => row.slice(0, 8));
  const start = Number(setting.values[0][0]) || 4;
  if (!fields.length || ![4, 6].includes(start)) throw new Error("Revise el mapa de campos y la primera fila de datos de Base.");
  const headers = base.getRangeByIndexes(start - 2, 0, 1, fields.length); headers.load("text");
  const end = used.isNullObject ? start - 1 : used.rowIndex + used.rowCount;
  if (end < start) throw new Error("Base está vacía. No se generó ningún informe.");
  if ((end - start + 1) * fields.length > 500000) throw new Error("Base tiene un rango usado demasiado grande. Revise las filas con contenido antes de generar el informe.");
  const data = base.getRangeByIndexes(start - 1, 0, end - start + 1, fields.length); data.load("text");
  await context.sync();
  if (headers.text[0].some((v, i) => v !== String(fields[i][3]))) throw new Error("Los encabezados de Base no coinciden con EC_Campos. No se generó el PDF.");
  return pdfPlan58(fields, data.text);
}
function buildPdfSheet58(sheet, plan) {
  const n = plan.fields.length, height = plan.data.length + 4;
  const table = sheet.getRangeByIndexes(0, 0, height, n);
  table.numberFormat = formats49(height, n);
  table.format.font.name = "Arial"; table.format.font.size = 11;
  table.format.font.color = "#000000"; table.format.fill.color = "#FFFFFF";
  table.format.wrapText = false; table.format.shrinkToFit = true;
  table.format.verticalAlignment = "Center";
  table.format.font.strikethrough = false;
  const grid = sheet.getRangeByIndexes(1, 0, height - 1, n);
  for (const edge of ["InsideHorizontal", "InsideVertical"]) {
    const border = grid.format.borders.getItem(edge);
    border.style = "Continuous"; border.weight = "Thin"; border.color = "#000000";
  }
  for (const edge of ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "DiagonalDown", "DiagonalUp"])
    grid.format.borders.getItem(edge).style = "None";
  sheet.getRangeByIndexes(3, 0, 1, n).values = [plan.fields.map(f => literal49(String(f[3])))];
  sheet.getRangeByIndexes(4, 0, plan.data.length, n).values = plan.data.map(row => row.map(literal49));
  for (let level = 0; level < 2; level++) {
    let first = 0;
    while (first < n) {
      const block = String(plan.fields[first][1]);
      const label = String(plan.fields[first][level === 0 ? 1 : 2]);
      let end = first + 1;
      while (end < n && String(plan.fields[end][1]) === block &&
        (level === 0 || String(plan.fields[end][2]) === label)) end++;
      if (label.trim()) {
        const range = sheet.getRangeByIndexes(level + 1, first, 1, end - first);
        if (end - first > 1) range.merge();
        sheet.getRangeByIndexes(level + 1, first, 1, 1).values = [[literal49(label)]];
      }
      first = end;
    }
  }
  // Colorear el bloque completo después de combinar, incluyendo sus títulos y subgrupos.
  let firstBlock = 0;
  while (firstBlock < n) {
    const block = String(plan.fields[firstBlock][1]).trim();
    let end = firstBlock + 1;
    while (end < n && String(plan.fields[end][1]).trim() === block) end++;
    const header = sheet.getRangeByIndexes(1, firstBlock, 3, end - firstBlock);
    header.format.fill.color = block === "Persona" ? "#595959" : block === "Posición/Especialidad" ? "#922D3D" : "#FFFFFF";
    header.format.font.color = ["Persona", "Posición/Especialidad"].includes(block) ? "#FFFFFF" : "#000000";
    firstBlock = end;
  }
  for (let c = 0; c < n; c++) sheet.getRangeByIndexes(0, c, height, 1).format.columnWidth = plan.widths[c];
  for (let level = 0; level < 3; level++) {
    const row = sheet.getRangeByIndexes(level + 1, 0, 1, n);
    row.format.font.size = [16, 14, 11][level]; row.format.font.bold = true;
    row.format.rowHeight = [38, 32, 30][level];
  }
  sheet.getRangeByIndexes(4, 0, plan.data.length, n).format.rowHeight = 26;
  const legend = sheet.getRangeByIndexes(0, 0, 1, n);
  if (n > 1) legend.merge();
  sheet.getRangeByIndexes(0, 0, 1, 1).values = [["Conformidad de entrega: Quien genera este informe declara haber revisado la información incluida y presta conformidad con los datos entregados."]];
  legend.format.font.bold = true; legend.format.rowHeight = 28;
  bordersPdf59(table);
  const page = sheet.pageLayout;
  page.orientation = "Landscape"; page.paperSize = "A3";
  page.zoom = { horizontalFitToPages: 1, verticalFitToPages: 1 };
  page.printGridlines = false; page.printHeadings = false; page.draftMode = false;
  page.blackAndWhite = false;
  page.setPrintArea(table);
  page.setPrintMargins("Points", { left: 8, right: 8, top: 8, bottom: 8 });
}
// Cierra el archivo de Office después de leer todos los fragmentos, incluso si falla la lectura.
async function pdfBytes58() {
  const file = await new Promise((resolve, reject) => {
    let expired = false;
    const timer = setTimeout(() => { expired = true; reject(new Error("Excel no terminó de convertir el PDF en 60 segundos. Intente nuevamente cuando el archivo termine de sincronizar.")); }, 60000);
    try {
      Office.context.document.getFileAsync(Office.FileType.Pdf, { sliceSize: 65536 }, result => {
        if (expired) { if (result.status === Office.AsyncResultStatus.Succeeded) result.value.closeAsync(() => {}); return; }
        clearTimeout(timer);
        if (result.status !== Office.AsyncResultStatus.Succeeded) reject(new Error("Excel no pudo exportar el PDF: " + (result.error?.message || "error de conversión")));
        else resolve(result.value);
      });
    } catch (e) { clearTimeout(timer); reject(e); }
  });
  try {
    if (!file.size || !file.sliceCount) throw new Error("Excel devolvió un PDF vacío. No se descargó ningún archivo. Intente nuevamente después de sincronizar el libro.");
    const chunks = [];
    for (let i = 0; i < file.sliceCount; i++) {
      const data = await new Promise((resolve, reject) => {
        const timer = setTimeout(() => reject(new Error("Se interrumpió la lectura del PDF. Vuelva a intentar.")), 15000);
        try { file.getSliceAsync(i, result => {
          clearTimeout(timer);
          if (result.status !== Office.AsyncResultStatus.Succeeded) reject(new Error(result.error?.message || "No se pudo leer un fragmento del PDF."));
          else resolve(result.value.data);
        }); } catch (e) { clearTimeout(timer); reject(e); }
      });
      chunks.push(new Uint8Array(data));
    }
    const size = chunks.reduce((total, part) => total + part.length, 0);
    if (size !== file.size) throw new Error("El PDF se recibió incompleto. No se descargó ningún archivo.");
    const bytes = new Uint8Array(size); let at = 0;
    for (const part of chunks) { bytes.set(part, at); at += part.length; }
    if (String.fromCharCode(...bytes.slice(0, 5)) !== "%PDF-") throw new Error("Excel devolvió un archivo que no tiene formato PDF.");
    return bytes;
  } finally { file.closeAsync(() => {}); }
}
async function restorePdf58(state) {
  if (!state.originals.length) return;
  await Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    const refs = state.originals.map(saved => ({ saved, sheet: sheets.getItemOrNullObject(saved.id) }));
    const temp = state.tempId ? sheets.getItemOrNullObject(state.tempId) : sheets.getItemOrNullObject(state.tempName);
    refs.forEach(r => r.sheet.load("name,id,visibility")); temp.load("name,id"); await context.sync();
    // Primero reabrir las hojas originalmente visibles; nunca ocultar la última hoja visible.
    refs.filter(r => !r.sheet.isNullObject && r.saved.visibility === "Visible").forEach(r => r.sheet.visibility = "Visible");
    await context.sync();
    const active = refs.find(r => r.saved.id === state.activeId && !r.sheet.isNullObject);
    if (active) active.sheet.activate();
    refs.filter(r => !r.sheet.isNullObject && r.saved.visibility !== "Visible").forEach(r => r.sheet.visibility = r.saved.visibility);
    await context.sync();
    if (!temp.isNullObject) temp.delete();
    await context.sync();
  });
}
async function pdfHorizontal58() { return generatePdf59(false); }
async function generatePdf59(vertical = false) {
  const label = vertical ? "vertical" : "horizontal";
  if (!Office.context?.document?.getFileAsync) throw new Error("Esta versión de Excel no ofrece la exportación PDF desde el complemento.");
  if (Office.context.requirements && !Office.context.requirements.isSetSupported("ExcelApi", "1.9"))
    throw new Error("Actualice Excel para generar el PDF desde este complemento (ExcelApi 1.9).");
  cancelClearBase56(); cancelNew54(); cancelDelete53();
  ec49.pending = null; ec49.searchPending = null; ec49.saved = null;
  section49("ec49-confirm-box", false); section49("ec52-search-confirm-box", false); section49("ec49-clean-box", false);
  section49("ec58-download-box", false);
  if (pdfDownload58) { URL.revokeObjectURL(pdfDownload58); pdfDownload58 = null; }
  message49("Preparando el PDF " + label + "… Espere a que finalice antes de editar el libro.");
  const state = { originals: [], activeId: "", tempId: "", tempName: (vertical ? "EC_PDFV_" : "EC_PDFH_") + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 6) };
  let bytes, plan, error, cleanupError;
  try {
    await Excel.run(async context => {
      const sheets = context.workbook.worksheets;
      const active = sheets.getActiveWorksheet(); active.load("id");
      sheets.load("items/name,items/id,items/visibility");
      context.workbook.protection.load("protected"); await context.sync();
      if (context.workbook.protection.protected) throw new Error("Desproteja la estructura del libro para generar el PDF. Los datos no se modificaron.");
      plan = await readPdf58(context);
      if (vertical && plan.count + 3 > 16384) throw new Error("La entrega supera el límite de columnas del informe vertical de Excel.");
      // Snapshot de identidad y visibilidad de cada hoja antes de crear o esconder ninguna.
      state.originals = sheets.items.map(s => ({ id: s.id, visibility: s.visibility }));
      state.activeId = active.id;
      const temp = sheets.add(state.tempName); temp.load("name,id"); await context.sync(); state.tempId = temp.id;
      if (vertical) buildVerticalSheet59(temp, plan); else buildPdfSheet58(temp, plan);
      temp.activate(); await context.sync();
      state.originals.filter(s => s.visibility === "Visible").forEach(s => sheets.getItem(s.id).visibility = "Hidden");
      await context.sync();
    });
    message49("Excel está convirtiendo el informe de " + plan.count + " registros y " + plan.fields.length + " campos a PDF…");
    bytes = await pdfBytes58();
  } catch (e) { error = e; }
  finally { try { await restorePdf58(state); } catch (e) { cleanupError = e; } }
  if (cleanupError) {
    message49("No se pudo retirar por completo la vista temporal del PDF. Sus registros siguen en EC_Base. No limpie Base. Muestre las hojas originales y revise la hoja " + state.tempName + ". Detalle: " + String(cleanupError));
    return;
  }
  if (error) throw error;
  const name = (vertical ? "Informe_EC_Vertical_" : "Informe_EC_Base_") + new Date().toISOString().slice(0, 19).replace(/[T:]/g, "-") + ".pdf";
  if (pdfDownload58) URL.revokeObjectURL(pdfDownload58);
  pdfDownload58 = URL.createObjectURL(new Blob([bytes], { type: "application/pdf" }));
  const link = document.getElementById("ec58-download");
  link.href = pdfDownload58; link.download = name; link.textContent = "Descargar PDF " + label;
  section49("ec58-download-box", true); link.click();
  message49("PDF " + label + " generado con " + plan.count + " registros y " + plan.fields.length + " campos. Se solicitó la descarga; si no comienza, pulse Descargar PDF " + label + ". Base y Auditoría se conservaron. Revise el archivo antes de limpiar la entrega.", true);
}

// Entrega 64: misma línea fina negra para bordes internos y perímetro.
function bordersPdf59(range) {
  for (const edge of ["InsideHorizontal", "InsideVertical", "EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"]) {
    const border = range.format.borders.getItem(edge);
    border.style = "Continuous"; border.weight = "Thin"; border.color = "#000000";
  }
  for (const edge of ["DiagonalDown", "DiagonalUp"]) range.format.borders.getItem(edge).style = "None";
}
function labelPdf59(value) { return String(value).replace(/\s+/g, " ").trim(); }
function linesPdf59(text, size, width) {
  const capacity = Math.max(1, Math.floor((width - 12) / (size * 0.56)));
  return text.split("\n").reduce((total, paragraph) => {
    let length = 0, lines = 1;
    for (const word of paragraph.split(/\s+/)) {
      if (length > 0 && length + 1 + word.length > capacity) { lines++; length = 0; }
      length += (length ? 1 : 0) + word.length;
    }
    return total + lines;
  }, 0);
}
function buildVerticalSheet59(sheet, plan) {
  const n = plan.fields.length, columns = plan.count + 3;
  const full = sheet.getRangeByIndexes(0, 0, n + 1, columns);
  full.numberFormat = formats49(n + 1, columns);
  full.format.font.name = "Arial"; full.format.font.size = 11;
  full.format.font.color = "#000000"; full.format.font.strikethrough = false;
  full.format.fill.color = "#FFFFFF"; full.format.wrapText = false;
  full.format.shrinkToFit = true; full.format.verticalAlignment = "Top";
  full.format.rowHeight = 25;
  const cells = plan.fields.map((f, i) => ["", "", labelPdf59(f[3]), ...plan.data.map(row => row[i])]);
  sheet.getRangeByIndexes(1, 0, n, columns).values = cells.map(row => row.map(literal49));
  const legend = sheet.getRangeByIndexes(0, 0, 1, columns);
  legend.merge();
  sheet.getRangeByIndexes(0, 0, 1, 1).values = [["Conformidad de entrega: Quien genera este informe declara haber revisado la información incluida y presta conformidad con los datos entregados."]];
  legend.format.font.bold = true; legend.format.rowHeight = 18;
  legend.format.wrapText = false; legend.format.shrinkToFit = true;
  sheet.getRangeByIndexes(1, 2, n, columns - 2).format.autofitColumns();
  sheet.getRangeByIndexes(0, 0, n + 1, 1).format.columnWidth = 150;
  sheet.getRangeByIndexes(0, 1, n + 1, 1).format.columnWidth = 190;
  sheet.getRangeByIndexes(1, 2, n, 1).format.font.bold = true;
  const blocks = plan.fields.map(f => labelPdf59(f[1]));
  const groups = plan.fields.map(f => labelPdf59(f[2]));
  const heights = plan.fields.map(() => 25);
  let first = 0;
  while (first < n) {
    const block = blocks[first]; let end = first + 1;
    while (end < n && blocks[end] === block) end++;
    const blockRange = sheet.getRangeByIndexes(first + 1, 0, end - first, 1);
    if (end - first > 1) blockRange.merge();
    const visual = block === "Detalle de la Organización" ? "Detalle de la\nOrganización" : block;
    sheet.getRangeByIndexes(first + 1, 0, 1, 1).values = [[literal49(visual)]];
    heights[first] = Math.max(25, linesPdf59(visual, 16, 150) * 20 + 4 - (end - first - 1) * 25);
    let groupStart = first;
    while (groupStart < end) {
      let groupEnd = groupStart + 1;
      while (groupEnd < end && groups[groupEnd] === groups[groupStart]) groupEnd++;
      const group = groups[groupStart];
      if (group) {
        const groupRange = sheet.getRangeByIndexes(groupStart + 1, 1, groupEnd - groupStart, 1);
        if (groupEnd - groupStart > 1) groupRange.merge();
        sheet.getRangeByIndexes(groupStart + 1, 1, 1, 1).values = [[literal49(group)]];
        heights[groupStart] = Math.max(heights[groupStart], linesPdf59(group, 14, 190) * 18 + 4 - (groupEnd - groupStart - 1) * 25);
      }
      groupStart = groupEnd;
    }
    first = end;
  }
  // Después de combinar: cubrir también los nombres de bloque y de subgrupo.
  first = 0;
  while (first < n) {
    const block = blocks[first]; let end = first + 1;
    while (end < n && blocks[end] === block) end++;
    const band = sheet.getRangeByIndexes(first + 1, 0, end - first, columns);
    const dark = block === "Persona" || block === "Posición/Especialidad";
    band.format.fill.color = block === "Persona" ? "#595959" : block === "Posición/Especialidad" ? "#922D3D" : "#FFFFFF";
    band.format.font.color = dark ? "#FFFFFF" : "#000000";
    first = end;
  }
  const descriptors = sheet.getRangeByIndexes(1, 0, n, 2);
  descriptors.format.font.bold = true; descriptors.format.wrapText = true;
  descriptors.format.horizontalAlignment = "Left"; descriptors.format.verticalAlignment = "Top";
  sheet.getRangeByIndexes(1, 0, n, 1).format.font.size = 16;
  sheet.getRangeByIndexes(1, 1, n, 1).format.font.size = 14;
  for (let r = 0; r < n; r++) if (heights[r] > 25) sheet.getRangeByIndexes(r + 1, 0, 1, columns).format.rowHeight = heights[r];
  bordersPdf59(full);
  const page = sheet.pageLayout;
  page.orientation = "Portrait"; page.paperSize = "A3";
  page.zoom = { horizontalFitToPages: 1, verticalFitToPages: 1 };
  page.printGridlines = false; page.printHeadings = false;
  page.draftMode = false; page.blackAndWhite = false;
  page.setPrintArea(full);
  page.setPrintMargins("Points", { left: 8, right: 8, top: 8, bottom: 8 });
}
async function pdfVertical59() { return generatePdf59(true); }

// Entrega 64: actualización explícita, una vez. Los botones de trabajo no reescriben las instrucciones.
function instructionParts60() {
  const form = [];
  const base = [];
  const add = (list, text, bold = false) => list.push({ text, font: { bold } });
  add(form, "FORMULARIO UNIFICADO DE SOLICITUD\n", true);
  add(form, "Permite registrar ingresos, reingresos, transferencias y actualizaciones organizativas con datos de persona, posición y organización. Al guardar, valida los campos obligatorios y acumula los registros en Base para generar el PDF y enviarlo al personal correspondiente.\n");
  add(form, "ACCESO Y USO\n", true);
  add(form, "Abra Inicio > Formulario Contreras en Excel. En el complemento, seleccione Formulario para operar las solicitudes o Base e informes para preparar la entrega.\n");
  add(form, "MOTIVOS DE SOLICITUD\n", true);
  add(form, "• Ingreso: ", true); add(form, "la persona ingresa por primera vez a Contreras.\n");
  add(form, "• Reingreso: ", true); add(form, "la persona ya trabajó en Contreras y comienza un nuevo período laboral.\n");
  add(form, "• Transferencia: ", true); add(form, "la persona está activa y debe reubicarse dentro de la empresa.\n");
  add(form, "• Actualización organizativa: ", true); add(form, "la persona está activa y se necesita modificar algún dato de su información.\n");
  add(form, "CAMPOS A COMPLETAR\n", true);
  add(form, "Los campos se identificarán con distintos colores según el motivo seleccionado: gris, obligatorio; amarillo, opcional; blanco, no requerido para ese motivo.\n");
  add(form, "BOTONES DEL COMPLEMENTO\n", true);
  add(form, "• Cargar / Actualizar solicitud: ", true); add(form, "agrega el registro. Si el CUIL ya existe, solicita confirmar su reemplazo completo, incluidos los campos vacíos. Al guardar, pregunta si desea limpiar el formulario.\n");
  add(form, "• Nueva carga / Limpiar formulario: ", true); add(form, "limpia los datos y conserva la información principal de Solicitud. Si hay cambios sin guardar, pide confirmación.\n");
  add(form, "• Buscar CUIL: ", true); add(form, "recupera los datos del CUIL indicado en el complemento; si no existe en esta entrega, lo informa.\n");
  add(form, "• Eliminar solicitud seleccionada: ", true); add(form, "elimina el registro recuperado mediante Buscar CUIL, previa confirmación.\n");
  add(form, "La carga, actualización y eliminación quedan registradas en Auditoría con el usuario de Microsoft que ejecuta la acción.");
  add(base, "BASE DE DATOS DE LA ENTREGA\n", true);
  add(base, "Reúne los registros guardados desde el formulario. Muestra solo los campos con información en al menos un registro; los campos vacíos en toda la entrega se ocultan.\n");
  add(base, "ACCESO A LOS BOTONES\n", true);
  add(base, "En Excel, abra Inicio > Formulario Contreras y seleccione Base e informes en el complemento. Allí encontrará los botones de esta hoja:\n");
  add(base, "• Generar PDF horizontal: ", true); add(base, "descarga los registros y los campos con información, con un registro por fila. Es el formato recomendado si no se indicó otro.\n");
  add(base, "• Generar PDF vertical: ", true); add(base, "descarga los registros y los campos con información en formato vertical, con un registro por columna. Se utiliza para funciones particulares; si no sabe cuál se requiere, genere el PDF horizontal.\n");
  add(base, "• Mostrar todos los campos: ", true); add(base, "muestra el esquema completo, incluidos los campos vacíos. La próxima carga o actualización volverá a mostrar solo los campos con información.\n");
  add(base, "• Limpiar Base / Eliminar entrega: ", true); add(base, "elimina todos los registros de la entrega, previa confirmación. Conserva los encabezados, formatos y el historial de Auditoría. Revise el informe descargado antes de limpiar Base.\n");
  add(base, "CONFORMIDAD DE ENTREGA\n", true);
  add(base, "Quien genera el informe declara haber revisado la información incluida y presta conformidad con los datos entregados.");
  return { form, base };
}
function instructionText60(parts) { return parts.map(p => p.text).join(""); }
function instructionHeight60(parts, width, size) {
  const capacity = Math.max(20, Math.floor((width - 12) / (size * 0.56)));
  let lines = 0;
  for (const paragraph of instructionText60(parts).split("\n")) {
    let length = 0; lines++;
    for (const word of paragraph.split(/\s+/)) {
      if (length > 0 && length + word.length + 1 > capacity) { lines++; length = 0; }
      length += (length ? 1 : 0) + word.length;
    }
  }
  return Math.min(409.5, Math.max(60, lines * size * 1.25 + 12));
}
function setInstructions60(cell, parts, rich, name, size) {
  cell.values = [[instructionText60(parts)]];
  cell.format.font.bold = false;
  if (rich) cell.setCellProperties([[{ textRuns: parts.map(part => ({ text: part.text, font: { name, size, color: "#000000", bold: part.font.bold } })) }]]);
}
async function updateInstructions60() {
  message49("Actualizando únicamente las instrucciones y el formato de las cabeceras…");
  const rich = !!Office.context?.requirements?.isSetSupported("ExcelApi", "1.18");
  let completed = false;
  await Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    const form = sheets.getItemOrNullObject("EC_Formulario"), base = sheets.getItemOrNullObject("EC_Base"), meta = sheets.getItemOrNullObject("EC_Campos");
    form.load("name"); base.load("name"); meta.load("name"); await context.sync();
    if (form.isNullObject || base.isNullObject || meta.isNullObject) throw new Error("Faltan Formulario, Base o EC_Campos.");
    const start = meta.getRange("P2"); start.load("values");
    const map = meta.getUsedRangeOrNullObject(true); map.load("values");
    const headers = base.getRange("A1:H2"); headers.load("values");
    const info = form.getRange("A3:H3"); info.load("values");
    form.protection.load("protected,options"); base.protection.load("protected"); await context.sync();
    if (Number(start.values[0][0]) !== 6 || map.isNullObject) throw new Error("Esta actualización necesita los encabezados de Base en las filas 3 a 5 y datos desde la fila 6. No se modificó nada.");
    const fields = map.values.slice(1).filter(f => String(f[0]).trim() !== "");
    const pinned = ["Motivo de la solicitud", "Fecha de solicitud", "Área solicitante", "ID Solicitante"];
    if (fields.length < 4 || pinned.some((name, i) => String(fields[i][3]).trim() !== name)) throw new Error("Las cuatro primeras columnas de Base no coinciden con Solicitud. No se modificó nada.");
    if (fields.some(f => !String(f[7] || "").trim() && Number(f[4]) < 7)) throw new Error("El mapa contiene campos dentro de la zona de instrucciones. No se modificó nada.");
    if (base.protection.protected) throw new Error("Desproteja EC_Base antes de actualizar sus instrucciones. No se modificó nada.");
    if (headers.values.some(row => row.slice(1, 4).some(v => String(v).trim() !== ""))) throw new Error("Hay texto adicional en B1:D2 de Base. Revíselo antes de combinar A:D para no perderlo.");
    const populated = info.values[0].map((v, i) => String(v).trim() !== "" ? i : -1).filter(i => i >= 0);
    if (populated.length > 1) throw new Error("Hay más de una celda con texto en la fila 3 del formulario. Revísela antes de unificar las instrucciones.");
    const first = populated.length ? populated[0] : 1;
    const input = form.getRangeByIndexes(2, first, 1, 1); input.format.font.load("name,size");
    const baseInfo = base.getRange("A2"); baseInfo.format.font.load("name,size"); await context.sync();
    const formName = input.format.font.name || "Calibri", formSize = input.format.font.size || 12;
    const baseName = baseInfo.format.font.name || "Calibri", baseSize = baseInfo.format.font.size || 11;
    const priorProtected = form.protection.protected, options = form.protection.options;
    if (priorProtected) { form.protection.unprotect(); await context.sync(); }
    try {
      const parts = instructionParts60();
      // Unificar solo cabeceras; nunca insertar/eliminar filas ni desplazar el mapa.
      base.getRange("1:2").unmerge();
      base.getRange("A1:D1").merge(false); base.getRange("A2:D2").merge(false);
      const title = base.getRange("A1:D1"); title.format.font.bold = true; title.format.wrapText = true;
      if (String(headers.values[0][0]).trim() === "") base.getRange("A1").values = [["BASE DE DATOS DE LA ENTREGA"]];
      const baseRange = base.getRange("A2:D2");
      baseRange.format.wrapText = true; baseRange.format.shrinkToFit = false;
      baseRange.format.horizontalAlignment = "Left"; baseRange.format.verticalAlignment = "Top";
      baseRange.format.font.name = baseName; baseRange.format.font.size = baseSize;
      baseRange.format.font.color = "#000000";
      setInstructions60(base.getRange("A2"), parts.base, rich, baseName, baseSize);
      form.getRange("3:3").unmerge();
      const formRange = form.getRangeByIndexes(2, first, 1, 8 - first); formRange.merge(false);
      formRange.format.wrapText = true; formRange.format.shrinkToFit = false;
      formRange.format.horizontalAlignment = "Left"; formRange.format.verticalAlignment = "Top";
      formRange.format.font.name = formName; formRange.format.font.size = formSize;
      formRange.format.font.color = "#000000";
      setInstructions60(form.getRangeByIndexes(2, first, 1, 1), parts.form, rich, formName, formSize);
      form.getRange("4:4").rowHidden = true;
      // Las columnas A:D quedan visibles porque sostienen las dos cabeceras combinadas.
      base.getRange("A:D").columnHidden = false;
      formRange.load("width"); baseRange.load("width"); await context.sync();
      formRange.format.rowHeight = instructionHeight60(parts.form, formRange.width || 800, formSize);
      // Reservar espacio para que el autoajuste de columnas de una carga posterior no corte las instrucciones.
      baseRange.format.rowHeight = Math.max(360, instructionHeight60(parts.base, baseRange.width || 500, baseSize));
      await context.sync(); completed = true;
    } finally {
      if (priorProtected) { form.protection.protect(options); await context.sync(); }
    }
  });
  if (completed) message49("Instrucciones actualizadas en Formulario y Base. A1:D1 y A2:D2 de Base quedaron combinadas. Se conservaron los campos, registros, Auditoría, colores y PDF." + (rich ? " Solo los títulos y nombres de botones están en negrita." : " Esta versión de Excel no permite negrita parcial; se dejó el texto normal con títulos en mayúsculas."), true);
}

// Entrega 64. Protección operativa SIN contraseña. No es un control de confidencialidad.
// La estructura del libro queda libre para que funcionen los PDF temporales.
async function restoreProtection61(states) {
  if (!states.length) return;
  await Excel.run(async context => {
    const refs = states.filter(s => s.protected).map(s => ({ state: s, sheet: context.workbook.worksheets.getItemOrNullObject(s.name) }));
    refs.forEach(r => { r.sheet.load("name"); r.sheet.protection.load("protected"); });
    await context.sync();
    const errors = [];
    // Son solo dos hojas: intentar ambas aunque una falle al recuperar su protección.
    for (const r of refs) {
      if (r.sheet.isNullObject) { errors.push("Falta " + r.state.name); continue; }
      if (r.sheet.protection.protected) continue;
      try { r.sheet.protection.protect(r.state.options); await context.sync(); }
      catch (e) { errors.push(r.state.name + ": " + String(e)); }
    }
    if (errors.length) throw new Error(errors.join("; "));
  });
}
async function withProtectedSheets61(action) {
  let states = [], result, error, completed = false;
  try {
    await Excel.run(async context => {
      const refs = ["EC_Base", "EC_Auditoria"].map(name => ({ name, sheet: context.workbook.worksheets.getItemOrNullObject(name) }));
      refs.forEach(r => r.sheet.load("name")); await context.sync();
      for (const r of refs) if (!r.sheet.isNullObject) r.sheet.protection.load("protected,options");
      await context.sync();
      states = refs.filter(r => !r.sheet.isNullObject).map(r => ({ name: r.name, protected: r.sheet.protection.protected, options: { ...r.sheet.protection.options } }));
      for (const r of refs) if (!r.sheet.isNullObject && r.sheet.protection.protected) r.sheet.protection.unprotect();
      await context.sync();
    });
    result = await action(); completed = true;
  } catch (e) { error = e; }
  try { await restoreProtection61(states); }
  catch (e) {
    throw new Error((completed ? "La operación finalizó, pero no se pudo restablecer la protección de Base o Auditoría. Verifique los datos antes de repetir una carga. " : "La operación se interrumpió y no se pudo restablecer toda la protección. Revise las hojas antes de continuar. ") + String(e) + (error ? " Error inicial: " + String(error) : ""));
  }
  if (error) throw error;
  return result;
}
async function applyProtection61() {
  message49("Aplicando protección sin contraseña a Base y Auditoría…");
  await withProtectedSheets61(async () => {
    await Excel.run(async context => {
      await readBaseView55(context);
      const base = context.workbook.worksheets.getItem("EC_Base");
      const audit = context.workbook.worksheets.getItemOrNullObject("EC_Auditoria");
      audit.load("name"); context.workbook.protection.load("protected"); await context.sync();
      if (audit.isNullObject) throw new Error("Falta EC_Auditoria. No se aplicó la protección.");
      if (context.workbook.protection.protected) throw new Error("Desproteja la estructura del libro antes de continuar. Los PDF necesitan crear y retirar sus hojas temporales.");
      const headers = audit.getRange("A1:E1"); headers.load("text"); await context.sync();
      const expected = ["Fecha/Hora Carga", "Usuario Carga", "Área Solicitante", "Motivo de Solicitud", "Detalle de campos cargados o modificados"];
      if (headers.text[0].some((v, i) => v !== expected[i])) throw new Error("Los encabezados de Auditoría no coinciden. No se aplicó la protección.");
      // Una única asignación de formato para toda la hoja: también bloquea las futuras filas.
      base.getRange().format.protection.locked = true;
      audit.getRange().format.protection.locked = true;
      base.protection.protect({}); audit.protection.protect({});
      base.visibility = "Visible"; base.activate(); await context.sync();
      audit.visibility = "Hidden"; await context.sync();
    });
  });
  message49("Protección aplicada: Base y Auditoría quedaron bloqueadas sin contraseña; Auditoría está oculta. Los botones pueden guardar y actualizar los registros. Formulario, Campos, Catálogos, Confirmación y los datos se conservaron. Para mantenimiento manual: Revisar > Desproteger hoja.", true);
}
