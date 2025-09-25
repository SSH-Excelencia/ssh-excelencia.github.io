
//import * as XLSX from "xlsx";

const XLSX = window.XLSX || window.xlsx;

let logoBuffer = null; // variable global para el logo

async function cargarImagenComoArrayBuffer(url) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`No se pudo cargar imagen: ${response.status}`);
  return await response.arrayBuffer();
}

document.addEventListener("DOMContentLoaded", async () => {
  
  let graficaTorta = null;
  let graficaBarras = null;
  let workbook = null; // variable global para Excel
  // === OBJETO GLOBAL PARA GUARDAR RESPUESTAS ===
  let respuestasPorTabla = {}; // { nombreHoja: { fila: { opcion, observacion } } }

  // --- ELEMENTOS ---
  const btnContinuar = document.getElementById("btnContinuar");
  const bloque1 = document.getElementById("form-inicial");
  const bloque2 = document.getElementById("panel-inicial");
  const bloque3 = document.getElementById("tabla-excel");
  const selectTablas = document.getElementById("menu-tablas");
  const btnExportWord = document.getElementById("btn-export-word");
  const btnExportPdf = document.getElementById("btn-export-pdf");
  const btnReiniciar = document.getElementById("btnReiniciar");
  const imgHeader = document.getElementById("logo-header");

 try {
    logoBuffer = await cargarImagenComoArrayBuffer("Image/logo1.png");
    console.log("Logo precargado");
  } catch (err) {
    console.warn("No se pudo precargar logo:", err);
    logoBuffer = null;
  }


  imgHeader.addEventListener("click", () => {
    // Rellenar los textbox del bloque 1
    document.getElementById("nombreIps").value = "Auditor Interno ";
    document.getElementById("numeroContacto").value = "123456789";
    document.getElementById("correoElectronico").value = "Diseno@ssh.com";

  });

  
  // --- BLOQUE 2 INICIALMENTE DESHABILITADO ---
  bloque2.style.visibility = "hidden";  // no se ve
  bloque2.style.position = "absolute";  // saca del flujo visual
  bloque2.style.opacity = "0";          // transición suave si quieres
  bloque2.style.pointerEvents = "none"; // no interactúa
  selectTablas.disabled = true;
  btnExportWord.disabled = true;
  //btnExportPdf.disabled = true;

  cargarTextoEjemplo("public/mensaje-ejemplo.txt");

  

  // --- BOTÓN CONTINUAR ---
  btnContinuar.addEventListener("click", () => {
    const nombreIps = document.getElementById("nombreIps").value.trim();
    const numeroContacto = document.getElementById("numeroContacto").value.trim();
    const correoElectronico = document.getElementById("correoElectronico").value.trim();

    if (!nombreIps || !numeroContacto || !correoElectronico) {
      alert("⚠️ Por favor ingresa todos los datos del evaluador.");
      return;
    }

    if (!/^\d+$/.test(numeroContacto)) {
      alert("⚠️ El número de contacto debe contener solo números.");
      return;
    }

    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correoElectronico)) {
      alert("⚠️ Ingresa un correo electrónico válido.");
      return;
    }

    // Mostrar bloque 2
    bloque1.style.display = "none";
    
    bloque2.style.visibility = "visible"; // ahora se ve
    bloque2.style.position = "static";    // vuelve al flujo normal
    bloque2.style.opacity = "1";          // opacidad total
    bloque2.style.pointerEvents = "auto"; // interactivo
    
    selectTablas.disabled = false;
    btnExportWord.disabled = false;
    btnExportPdf.disabled = false;

    // Cargar Excel solo una vez
    if (!bloque2.dataset.loaded) {
      cargarExcel("public/lista-chequeo.xlsx");
      bloque2.dataset.loaded = "true";
    }

    // --- Mostrar opción inicial y cargar mensaje desde txt ---
    selectTablas.innerHTML = '<option value="" selected>Seleccione...</option>';
    //cargarTextoEjemplo("mensaje-ejemplo.txt");
  });

  // --- BOTÓN REINICIAR ---
  btnReiniciar.addEventListener("click", () => {
    if (confirm("⚠️ Todos los datos se eliminarán. ¿Deseas reiniciar la evaluación?")) {
      location.reload();
    }
  });

  // --- BOTÓN SALIR ---
  btnExportPdf.addEventListener("click", () => {
    window.location.href = "https://ssh-excelencia.github.io/";
  });

  // --- FUNCIONES PARA EXCEL ---
  async function cargarExcel(ruta) {
    try {
      const response = await fetch(ruta);
      const data = await response.arrayBuffer();
      workbook = XLSX.read(data, { type: "array" });

      // Llenar select de hojas, ignorando "tabla de contenido"
      const hojasFiltradas = workbook.SheetNames.filter(
        nombreHoja => nombreHoja.toLowerCase() !== "tabla de contenido"
      );

      hojasFiltradas.forEach((nombreHoja) => {
        const opt = document.createElement("option");
        opt.value = nombreHoja;
        opt.textContent = nombreHoja;
        selectTablas.appendChild(opt);
      });

      // Evento change para cargar tabla al seleccionar hoja
      selectTablas.addEventListener("change", (e) => {
        const hojaSeleccionada = e.target.value;

        if (!hojaSeleccionada) {
          // Mostrar mensaje de ejemplo si se selecciona "Seleccione..."
          cargarTextoEjemplo("public/mensaje-ejemplo.txt");
          return;
        }

        mostrarTabla(workbook, hojaSeleccionada);
      });

    } catch (err) {
      console.error("Error cargando Excel:", err);
    }
  }


// === FUNCION MOSTRAR TABLA ===
function mostrarTabla(workbook, hoja) {
  const worksheet = workbook.Sheets[hoja];
  const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, blankrows:false });

  // Limpiar contenedor
  bloque3.innerHTML = "";

  // Crear tabla
  const table = document.createElement("table");
  table.classList.add("tabla-excel");

  const anchos = ["55%", "21%", "24%"];
  const thead = document.createElement("thead");
  const tbody = document.createElement("tbody");

  // --- Encabezado ---
  const headerRow = document.createElement("tr");
  jsonData[0].forEach((celda, j) => {
    const th = document.createElement("th");
    th.textContent = celda ?? "";
    th.style.textAlign = "center";
    th.style.verticalAlign = "middle";
    if (anchos[j]) th.style.width = anchos[j];
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  // --- Filas ---
for (let i = 1; i < jsonData.length; i++) {
  const fila = jsonData[i];
  const tr = document.createElement("tr");

  const esTitulo = fila[0]?.toString().startsWith("TT-");
  const esSubtitulo = fila[0]?.toString().startsWith("T-");

  if (esTitulo) {
    // === FILA DE TÍTULO PRINCIPAL ===
    const td = document.createElement("td");
    td.colSpan = 3;
    td.innerHTML = `<strong>${fila[0].toString().replace(/^TT-/, "")}</strong>`;
    td.style.textAlign = "center";
    tr.appendChild(td);

  } else if (esSubtitulo) {
    // === FILA DE SUBTÍTULO (sin radios, combinando col1+col2) ===
    const td = document.createElement("td");
    td.colSpan = 2;
    td.innerHTML = `<em>${fila[0].toString().replace(/^T-/, "")}</em>`;
    td.style.textAlign="justify";
    td.style.verticalAlign="middle"
        tr.appendChild(td);

    // === Columna 3: Observación con textarea ===
    const tdObs = document.createElement("td");
    tdObs.style.verticalAlign = "middle";
    


    const textarea = document.createElement("textarea");
    textarea.placeholder = "Ingrese observación";
    textarea.style.width = "95%";
    textarea.style.height = "50px";
    textarea.style.backgroundColor="#faefbcff";

    // Restaurar observación si estaba guardada
    const respuestasGuardadas = respuestasPorTabla[hoja]?.[i] || {};
    if (respuestasGuardadas.observacion) {
      textarea.value = respuestasGuardadas.observacion;
    }

    // Guardar al escribir
    textarea.addEventListener("input", () => {
      if (!respuestasPorTabla[hoja]) respuestasPorTabla[hoja] = {};
      if (!respuestasPorTabla[hoja][i]) respuestasPorTabla[hoja][i] = {};
      respuestasPorTabla[hoja][i].observacion = textarea.value;
    });

    tdObs.appendChild(textarea);
    tr.appendChild(tdObs);

} else {
      // === FILA NORMAL DE CRITERIO ===
      for (let j = 0; j < 3; j++) {
        const td = document.createElement("td");
        td.style.verticalAlign = "middle";
        td.style.textAlign="justify";
        // Recuperar respuestas guardadas si existen
        const respuestasGuardadas = respuestasPorTabla[hoja]?.[i] || {};
        if (j === 1) {
          // === Columna Radios ===
          const opciones = ["Cumple", "No Cumple", "No Aplica"];
          const contenedorRadios = document.createElement("div");
          contenedorRadios.classList.add("opciones-radios");
          opciones.forEach(opcion => {
            const label = document.createElement("label");
            label.style.marginRight = "10px";
            const input = document.createElement("input");
            input.type = "radio";
            input.name = `opcion_${i}`;
            input.value = opcion;
            // Restaurar radio si estaba guardado
            if (respuestasGuardadas.opcion === opcion) {
              input.checked = true;
            }
            // Guardar al cambiar
            input.addEventListener("change", () => {
              if (!respuestasPorTabla[hoja]) respuestasPorTabla[hoja] = {};
              if (!respuestasPorTabla[hoja][i]) respuestasPorTabla[hoja][i] = {};
              respuestasPorTabla[hoja][i].opcion = opcion;
              actualizarContadoresCol2();
            });
            label.appendChild(input);
            label.appendChild(document.createTextNode(" " + opcion));
            contenedorRadios.appendChild(label);
          });
          td.appendChild(contenedorRadios);
        } else if (j === 2) {
          // === Columna Observación ===
          const textarea = document.createElement("textarea");
          textarea.placeholder = "Ingrese observación";
          textarea.style.width = "95%";
          textarea.style.height = "50px";
          // Restaurar observación si estaba guardada
          if (respuestasGuardadas.observacion) {
            textarea.value = respuestasGuardadas.observacion;
          }
          // Guardar al escribir
          textarea.addEventListener("input", () => {
            if (!respuestasPorTabla[hoja]) respuestasPorTabla[hoja] = {};
            if (!respuestasPorTabla[hoja][i]) respuestasPorTabla[hoja][i] = {};
            respuestasPorTabla[hoja][i].observacion = textarea.value;
          });
          td.appendChild(textarea);
        } else {
          // === Columna Pregunta ===
          td.textContent = fila[j] ?? "";
        }
        if (anchos[j]) td.style.width = anchos[j];
        tr.appendChild(td);
      }
    }
    tbody.appendChild(tr);
  }

  // Insertar en el DOM
  table.appendChild(thead);
  table.appendChild(tbody);
  bloque3.appendChild(table);
  bloque3.classList.remove("oculto");

  // Reiniciar scroll
  bloque3.scrollTop = 0;

  // Total de criterios
  const totalListas = bloque3.querySelectorAll(
    "table tbody tr td:nth-child(2) .opciones-radios"
  ).length;
  const spanNumCriterios = document.getElementById("num-criterios");
  if (spanNumCriterios) spanNumCriterios.textContent = totalListas;

  // Inicializar contadores
  actualizarContadoresCol2();
}


// --- FUNCION ACTUALIZAR CONTADORES COLUMNA 2 ---

function actualizarContadoresCol2() {
  // Contenedores con radios por fila (solo filas de criterios)
  const radiosPorFila = bloque3.querySelectorAll(
    "table tbody tr td:nth-child(2) .opciones-radios"
  );

  let cumple = 0, noCumple = 0, noAplica = 0;

  radiosPorFila.forEach(contenedor => {
    const seleccionado = contenedor.querySelector("input[type='radio']:checked");
    if (!seleccionado) return;

    const val = (seleccionado.value || "").trim().toLowerCase();
    if (val === "cumple") cumple++;
    else if (val === "no cumple" || val === "nocumple") noCumple++;
    else if (val === "no aplica") noAplica++; // acepta "No Aplica" / "No aplica"
  });

  const total = radiosPorFila.length;
  const seleccionados = cumple + noCumple + noAplica;
  const pendientes = Math.max(total - seleccionados, 0);

  // Helper para actualizar texto si el elemento existe
  const setText = (id, value) => {
    const el = document.getElementById(id);
    if (el) el.textContent = String(value);
  };

  setText("num-criterios", total);
  setText("num-cumple", cumple);
  setText("num-nocumple", noCumple);
  setText("num-noaplica", noAplica);
  setText("num-ptes", pendientes); // id en minúsculas
  

  // Barra de progreso
  const barra = document.getElementById("barra-progreso");
  if (barra) {
    const porcentaje = total > 0 ? Math.ceil((seleccionados / total) * 100) : 0;
    barra.style.width = porcentaje + "%";
    barra.textContent = porcentaje + "%";  // 👈 Texto visible dentro de la barra
  }

  // === ACTUALIZAR GRÁFICAS ===
  const datos = [cumple, noCumple, noAplica];
  const etiquetas = ["Cumple", "No Cumple", "No Aplica"];
  const colores = ["#4CAF50", "#F44336", "#FFC107"]; // verde, rojo, amarillo

  // --- Gráfica Torta ---
  if (graficaTorta) graficaTorta.destroy();
  graficaTorta = new Chart(document.getElementById("graficaTorta"), {
  type: "pie",
  data: {
    labels: etiquetas,
    datasets: [{
      data: datos,
      backgroundColor: colores,
      borderColor: "#fff",
      borderWidth: 2
    }]
  },
  options: {
    responsive: true,
    plugins: {
      legend: { display: false } // leyenda abajo personalizada
    }
  }
  });

  // --- Gráfica Barras ---
  if (graficaBarras) graficaBarras.destroy();
  graficaBarras = new Chart(document.getElementById("graficaBarras"), {
    type: "bar",
    data: {
      labels: etiquetas,
      datasets: [{
        label: "Cantidad",
        data: datos,
        backgroundColor: colores
      }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { display: false }
      },
      scales: {
        y: {
        beginAtZero: true,
        precision: 0
        }
      }
    }
  });

  // === CALCULAR PORCENTAJES ===
  const calcularPorcentaje = (valor) => {
    return seleccionados > 0 ? ((valor / seleccionados) * 100).toFixed(1) : 0;
  };

  const setPorcentaje = (id, valor) => {
    const el = document.getElementById(id);
    if (el) el.textContent = calcularPorcentaje(valor) + "%";
  };

  setPorcentaje("p-cumple", cumple);
  setPorcentaje("p-nocumple", noCumple);
  setPorcentaje("p-noaplica", noAplica);

}
  
  // --- FUNCION CARGAR TEXTO DESDE ARCHIVO ---
  async function cargarTextoEjemplo(rutaTxt) {
    try {
      // 🔹 Si hay gráficas previas, destruirlas
    if (graficaTorta) {
      graficaTorta.destroy();
      graficaTorta = null;
    }
    if (graficaBarras) {
      graficaBarras.destroy();
      graficaBarras = null;
    }
      
      
      
      
      const response = await fetch(rutaTxt);
      if (!response.ok) throw new Error("No se pudo cargar el archivo de texto.");
      const texto = await response.text();
      bloque3.innerHTML = `<div style="padding: 20px; text-align: center;">${texto}</div>`;
      bloque3.classList.remove("oculto");
      
      // Resetear contadores columna 2
      document.getElementById("num-criterios").textContent = "0";
      document.getElementById("num-cumple").textContent = "0";
      document.getElementById("num-nocumple").textContent = "0";
      document.getElementById("num-noaplica").textContent = "0";
      document.getElementById("num-ptes").textContent = "0";
      document.getElementById("p-cumple").textContent = "0%";
      document.getElementById("p-nocumple").textContent = "0%";
      document.getElementById("p-noaplica").textContent = "0%";

      // 🔹 Vaciar barra de progreso
      const barra = document.getElementById("barra-progreso");
      if (barra) {
        barra.style.width = "0%";
        barra.textContent = "0%";
      }


    } catch (err) {
      console.error("Error cargando texto de ejemplo:", err);
      bloque3.innerHTML = `<p style="color:red;">No se pudo cargar el mensaje de ejemplo.</p>`;
    }
  }




   // === EXPORTAR WORD ===

// ====== IMPORTANTE: usar la instancia UMD de docx ======
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  Header,
  Footer,
  ImageRun,
  BorderStyle
} = window.docx;

// --- Utilidad: sanitizar nombre de archivo ---
function sanitizeFileName(str) {
  return String(str || "")
    .trim()
    .replace(/[\\/:*?"<>|]+/g, "")
    .replace(/\s+/g, "_")
    .substring(0, 80);
}

// --- Utilidad: obtener yyyy-mm-dd ---
function fechaISO() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
}

// --- Lee la tabla actual del DOM y retorna una estructura usable ---
function leerTablaDesdeDOM() {
  const tabla = bloque3.querySelector("table.tabla-excel");
  if (!tabla) return { headers: [], filas: [] };

  const headers = Array.from(tabla.querySelectorAll("thead th")).map(th => th.textContent.trim());
  const filas = [];

  Array.from(tabla.querySelectorAll("tbody tr")).forEach(tr => {
    const celdas = Array.from(tr.children);

    // Fila título de sección (celda única con colspan=3)
    if (celdas.length === 1 && celdas[0].hasAttribute("colspan")) {
      filas.push({
        tipo: "seccion",
        titulo: celdas[0].innerText.trim()
      });
      return;
    }
    // 🔹 Fila subtítulo (2 celdas: una con colspan=2 y otra para observación)
    if (celdas.length === 2 && celdas[0].hasAttribute("colspan")) {
      const subtitulo = (celdas[0]?.innerText || "").trim();
      const observaciones =
        (celdas[1]?.querySelector("textarea")?.value || "").trim();

      filas.push({
        tipo: "subtitulo",
        subtitulo,
        observaciones
      });
      return;
    }

    // Fila de criterio (3 columnas)
    const celCriterio = celdas[0];
    const celEval = celdas[1];
    const celObs = celdas[2];

    const criterio = (celCriterio?.innerText || "").trim();

    // Buscar radio seleccionado y tomar SOLO su valor
    let evaluacion = "";
    const checked = celEval?.querySelector("input[type='radio']:checked");
    if (checked) evaluacion = checked.value.trim();

    const observaciones = (celObs?.querySelector("textarea")?.value || "").trim();

    filas.push({
      tipo: "criterio",
      criterio,
      evaluacion,
      observaciones
    });
  });

  return { headers, filas };
}

// --- Crea una celda con configuración común ---
function celda(parrafos, opts = {}) {
  const {
    widthPct, // porcentaje 0-100
    align = AlignmentType.BOTH, // BOTH = justificado
    bold = false,
    colSpan = 1,
    noWrap = false
  } = opts;

  const children = Array.isArray(parrafos) ? parrafos : [parrafos];
  const runs = children.map(txt => new Paragraph({
    children: [new TextRun({ text: txt, bold })],
    alignment: align
  }));

  return new TableCell({
    columnSpan: colSpan,
    children: runs,
    width: widthPct ? { size: Math.round(widthPct * 50), type: WidthType.PERCENTAGE } : undefined, // docx usa base 5000 para %
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
    borders: {
      top:   { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      left:  { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      bottom:{ style: BorderStyle.SINGLE, size: 1, color: "000000" },
      right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
    },
    verticalAlign: "center",
    shading: undefined,
    cantSplit: noWrap
  });
}

// --- Construye la tabla DOCX preservando estructura y alineaciones ---
function construirTablaDocx(headers, filas) {
  const anchoCols = [50, 24, 24]; // % como en el frontend

  // Fila de encabezado
  const rowHeader = new TableRow({
    children: [
      celda(headers[0] || "Criterio", { widthPct: anchoCols[0], align: AlignmentType.CENTER, bold: true }),
      celda(headers[1] || "Evaluación", { widthPct: anchoCols[1], align: AlignmentType.CENTER, bold: true }),
      celda(headers[2] || "Observaciones", { widthPct: anchoCols[2], align: AlignmentType.CENTER, bold: true })
    ],
    tableHeader: true
  });

  const rows = [rowHeader];

  filas.forEach(f => {
    if (f.tipo === "seccion") {
      // Fila de sección (celda fusionada, centrada, negrita)
      rows.push(new TableRow({
        children: [
          celda(f.titulo, { colSpan: 3, widthPct: 100, align: AlignmentType.CENTER, bold: true })
        ]
      }));
      } else if (f.tipo === "subtitulo") {
      // === SUBTÍTULO (T-) ===
      rows.push(
        new TableRow({
          children: [
            celda(f.subtitulo, {
              colSpan: 2,
              widthPct: anchoCols[0] + anchoCols[1],
            }),
            celda(f.observaciones || "", {
              widthPct: anchoCols[2],
            })
          ]
        })
      );

    } else {
      // Fila normal de criterio
      rows.push(new TableRow({
        children: [
          // Criterio: JUSTIFICADO
          celda(f.criterio || "", { widthPct: anchoCols[0], align: AlignmentType.BOTH }),
          // Evaluación: CENTRADO (solo el valor seleccionado)
          celda(f.evaluacion || "", { widthPct: anchoCols[1], align: AlignmentType.CENTER }),
          // Observaciones: JUSTIFICADO
          celda(f.observaciones || "", { widthPct: anchoCols[2], align: AlignmentType.BOTH })
        ]
      }));
    }
  });

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows
  });
}

// === TABLA DE RESUMEN DE CRITERIOS
function construirTablaResumenDocx() {
  const numCriterios = parseInt(document.getElementById("num-criterios")?.textContent || "0");
  const numCumple = parseInt(document.getElementById("num-cumple")?.textContent || "0");
  const numNoCumple = parseInt(document.getElementById("num-nocumple")?.textContent || "0");
  const numNoAplica = parseInt(document.getElementById("num-noaplica")?.textContent || "0");
  const numPendientes = parseInt(document.getElementById("num-ptes")?.textContent || "0");

  const seleccionados = numCumple + numNoCumple + numNoAplica;
  const calcPct = (v) => (numCriterios > 0 ? ((v / numCriterios) * 100).toFixed(1) + "%" : "0%");

  const filasResumen = [
    { criterio: "Cumple", cantidad: numCumple, pct: calcPct(numCumple) },
    { criterio: "No Cumple", cantidad: numNoCumple, pct: calcPct(numNoCumple) },
    { criterio: "No Aplica", cantidad: numNoAplica, pct: calcPct(numNoAplica) },
    { criterio: "Sin evaluar", cantidad: numPendientes, pct: calcPct(numPendientes) },
    { criterio: "TOTAL", cantidad: numCriterios, pct: "100%" }
  ];

  const rowHeader = new TableRow({
    children: [
      celda("CRITERIO", { widthPct: 70, align: AlignmentType.CENTER, bold: true }),
      celda("CANTIDAD", { widthPct: 15, align: AlignmentType.CENTER, bold: true }),
      celda("PORCENTAJE", { widthPct: 15, align: AlignmentType.CENTER, bold: true }),
    ],
    tableHeader: true
  });

  const rows = [rowHeader];

  filasResumen.forEach(f => {
    rows.push(new TableRow({
      children: [
        celda(f.criterio, { widthPct: 70, align: AlignmentType.BOTH, bold: f.criterio === "TOTAL" }),
        celda(String(f.cantidad), { widthPct: 15, align: AlignmentType.CENTER, bold: f.criterio === "TOTAL" }),
        celda(f.pct, { widthPct: 15, align: AlignmentType.CENTER, bold: f.criterio === "TOTAL" }),
      ]
    }));
  });

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows
  });
}

//obtener las graficas para enviar a word

function construirTablaGraficasDocx() {
  try {
    const idsGraficas = [
      { id: "graficaBarras", titulo: "Gráfica de Barras" },
      { id: "graficaTorta", titulo: "Gráfica de Torta" }
    ];

    const celdas = [];

    idsGraficas.forEach(g => {
      const canvas = document.getElementById(g.id);
      if (canvas) {
        const dataUrl = canvas.toDataURL("image/png");
        const imageBuffer = Uint8Array.from(
          atob(dataUrl.split(",")[1]),
          c => c.charCodeAt(0)
        );

        // Cada celda tendrá título + imagen
        celdas.push(
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE }, // 50% del ancho
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: g.titulo, bold: true })]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new ImageRun({
                    data: imageBuffer,
                    transformation: { width: 250, height: 200 } // ajusta tamaño
                  })
                ]
              })
            ]
          })
        );
      } else {
        // Si no existe la gráfica, celda vacía con aviso
        celdas.push(
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: `⚠ ${g.titulo} no disponible`, italics: true })]
              })
            ]
          })
        );
      }
    });

    return new Table({
      rows: [
        new TableRow({
          children: celdas
        })
      ],
      width: { size: 100, type: WidthType.PERCENTAGE }
    });
  } catch (err) {
    console.error("Error construyendo tabla de gráficas:", err);
    return new Paragraph({ text: "⚠ No se pudieron generar las gráficas." });
  }
}








// === Reemplaza el placeholder del botón Exportar Word ===
btnExportWord.addEventListener("click", async () => {
  // Validaciones básicas
  const nombreIps = document.getElementById("nombreIps")?.value?.trim() || "";
  const numeroContacto = document.getElementById("numeroContacto")?.value?.trim() || "";
  const correoElectronico = document.getElementById("correoElectronico")?.value?.trim() || "";
  const nombreTabla = selectTablas?.value || "";

  const { headers, filas } = leerTablaDesdeDOM();
  if (!headers.length || !filas.length) {
    if (!nombreTabla) {
      alert("Selecciona una tabla antes de exportar.");
      return;
    }
  }
  if (!logoBuffer) {
    alert("No se cargó el logo, se exportará sin imagen.");
  }
 
  // Construir documento
  const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: "Arial", size: 22 }, // 11 pt
      },
      paragraph: {
        spacing: { after: 120 } // 6pt
      }
    }
  },
  sections: [
    {
      headers: {
          default: new Header({
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "Resultados Criterios de Evaluación",
                    bold: true
                  }),
                  new TextRun("\t"),
                  ...(logoBuffer ? [
                    new ImageRun({
                      data: logoBuffer,
                      transformation: { width: 80, height: 40 }
                    })
                  ] : [])
                ],
                tabStops: [{ type: "right", position: 9000 }]
              })
            ]
          })
        },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({ text: "SSH Excelencia en Salud - Generado automáticamente", italics: true }),
                new TextRun({ text: " | " }),
                new TextRun({ text: fechaISO(), bold: true })
              ]
            })
          ]
        })
      },
        children: [
          // Título del reporte (en el cuerpo, por si se imprime sin encabezado)
          new Paragraph({
            children: [new TextRun({ text: "Auditoria de Criterios de Habilitación Res 3100", bold: true })],
            alignment: AlignmentType.CENTER
          }),
          
          new Paragraph({ text: "" }),

          // Datos del evaluador
          new Paragraph({
            children: [
              new TextRun({ text: "Nombre IPS / Profesional: ", bold: true }),
              new TextRun({ text: nombreIps || "-" })
            ]
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "Contacto: ", bold: true }),
              new TextRun({ text: numeroContacto || "-" })
            ]
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "Correo: ", bold: true }),
              new TextRun({ text: correoElectronico || "-" })
            ]
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "Tabla evaluada: ", bold: true }),
              new TextRun({ text: nombreTabla || "-" })
            ]
          }),

          // === TABLA DE RESUMEN DE CRITERIOS ===
          new Paragraph({ text: "" }),
          new Paragraph({
          children: [new TextRun({ text: "Resumen de Evaluación", bold: true })],
          alignment: AlignmentType.CENTER
          }),
          construirTablaResumenDocx(),
          new Paragraph({ text: "" }),
          
         // === Gráficas (Punto 3) ===
new Paragraph({
  children: [new TextRun({ text: "Gráficas", bold: true })],
  alignment: AlignmentType.CENTER
}),
construirTablaGraficasDocx(),   // 👈 aquí va la tabla de 2 columnas
new Paragraph({ text: "" }),

          // Espacio para análisis/conclusiones
          new Paragraph({
            children: [new TextRun({ text: "Análisis / Conclusiones", bold: true })]
          }),
          new Paragraph({ text: "" }),
          new Paragraph({ text: "" }),

          // Tabla con la información diligenciada
          construirTablaDocx(headers, filas)
        ]
      }
    ]
  });

  // Descargar
  const blob = await Packer.toBlob(doc);
  const base = `Lista3100_${sanitizeFileName(nombreTabla || "Tabla")}_${sanitizeFileName(nombreIps || "IPS")}_${fechaISO()}`;
  saveAs(blob, `${base}.docx`);
});

  



});
