/* Entrega 57. Complemento Formulario Contreras para Excel (Office.js).
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
  if (!button) { console.log("Entrega 57: no se encontró la interfaz del complemento. Revise taskpane.html y actualice el panel."); return; }
  button.addEventListener("click", () => task49(start49));
  document.getElementById("ec49-confirm").addEventListener("click", () => task49(confirm49));
  document.getElementById("ec49-cancel").addEventListener("click", () => {
    ec49.pending = null; section49("ec49-confirm-box", false);
    message49("Actualización cancelada. Base y Auditoría no se modificaron.");
  });
  document.getElementById("ec49-clean").addEventListener("click", () => task49(clean49));
  document.getElementById("ec49-keep").addEventListener("click", () => {
    ec49.saved = null; section49("ec49-clean-box", false);
    message49("Se conservaron los datos del formulario. La solicitud ya está guardada.", true);
  });
  document.getElementById("ec52-search").addEventListener("click", () => task49(() => search52(false)));
  document.getElementById("ec52-search-confirm").addEventListener("click", () => task49(() => search52(true)));
  document.getElementById("ec52-search-cancel").addEventListener("click", () => {
    ec49.searchPending = null; section49("ec52-search-confirm-box", false);
    message49("Búsqueda cancelada. Se conservaron los datos del formulario.");
  });
  document.getElementById("ec53-delete").addEventListener("click", () => task49(() => delete53(false)));
  document.getElementById("ec53-delete-confirm").addEventListener("click", () => task49(() => delete53(true)));
  document.getElementById("ec53-delete-cancel").addEventListener("click", () => {
    cancelDelete53(); message49("Eliminación cancelada. Base, Auditoría y el formulario se conservaron.");
  });
  document.getElementById("ec54-new").addEventListener("click", () => task49(() => new54(false)));
  document.getElementById("ec54-new-confirm").addEventListener("click", () => task49(() => new54(true)));
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
  selectTab55("form");
  message49("Listo. Complete el formulario de Excel y pulse Cargar / Actualizar aquí.");
});
function section49(id, show) { document.getElementById(id).hidden = !show; }
function message49(text, ok = false) {
  const box = document.getElementById("ec49-message");
  box.textContent = text; box.className = ok ? "message success" : "message";
  console.log("ENTREGA 57: " + text);
}
async function task49(action) {
  if (ec49.busy) return;
  ec49.busy = true;
  document.querySelectorAll("#ec49 button").forEach(button => button.disabled = true);
  try { await action(); }
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
  const visible = s.fields.map((f, c) => showAll || rows.some(row => row && String(row[c]).trim() !== ""));
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
    const s = await read49(context);
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


// Entrega 57: Buscar conserva las identidades históricas; no escribe Base ni Auditoría.
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


// Entrega 57: eliminación de un registro seleccionado, con cuenta Microsoft y confirmación.
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


// Entrega 57: Nueva carga depende solo del formulario y su mapa, como el script 15.
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


// Entrega 57: navegación visual del panel. Cambiar pestaña no modifica el libro.
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


// Entrega 57: Limpiar Base conserva el formulario y audita cada registro eliminado.
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
