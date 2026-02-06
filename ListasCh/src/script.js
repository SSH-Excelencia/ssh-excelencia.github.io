
//import * as XLSX from "xlsx";

const XLSX = window.XLSX || window.xlsx;
const { jsPDF } = window.jspdf;

// ================= CONFIGURACIÓN DE ARCHIVOS =================
const RUTAS_ARCHIVOS = {
  MensajeEjemplo: "public/mensaje-ejemplo.txt",
  correosAutorizados: "public/data.b64",
  TablaExcel: "public/lista-chequeo.xlsx"
};


let logoBuffer = null; // variable global para el logo
let overlayTimer = null;

async function cargarImagenComoArrayBuffer(url) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`No se pudo cargar imagen: ${response.status}`);
  return await response.arrayBuffer();
}

document.addEventListener("DOMContentLoaded", async () => {

  let graficaTorta = null;
  let graficaBarras = null;
  let workbook = null; // variable global para Excel

  const hojasEvaluadas = new Set();

  // === OBJETO GLOBAL PARA GUARDAR RESPUESTAS ===
  let respuestasPorTabla = {}; // { nombreHoja: { fila: { opcion, observacion } } }

  // --- ELEMENTOS ---
  const btnContinuar = document.getElementById("btnContinuar");
  const bloque1 = document.getElementById("form-inicial");
  const bloque2 = document.getElementById("panel-inicial");
  const bloque3 = document.getElementById("tabla-excel");
  const selectTablas = document.getElementById("menu-tablas");
  const btnExportWord = document.getElementById("btn-export-word");
  const btnExportTodo = document.getElementById("btn-export-todo");
  const btnExportPdf = document.getElementById("btn-export-pdf");
  const btnReiniciar = document.getElementById("btnReiniciar");
  const imgHeader = document.getElementById("logo-header");

  const btntestOv = document.getElementById("btn-testOv");


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

  cargarTextoEjemplo(RUTAS_ARCHIVOS.MensajeEjemplo);



  // --- BOTÓN CONTINUAR ---
btnContinuar.addEventListener("click", () => {
  const nombreIps = document.getElementById("nombreIps").value.trim();
  const numeroContacto = document.getElementById("numeroContacto").value.trim();
  const correoElectronico = document.getElementById("correoElectronico").value.trim();

  if (!nombreIps || !numeroContacto || !correoElectronico) {
    mostrarOverlay({
      mensaje: "⚠️ Por favor ingresa todos los datos del evaluador.",
      temporal: true,
      autoCerrar: true,
      tiempo: 3000
    });
    return;
  }

  if (!/^\d+$/.test(numeroContacto)) {
    mostrarOverlay({
      mensaje: "⚠️ El número de contacto debe contener solo números.",
      temporal: true,
      autoCerrar: true,
      tiempo: 3000
    });
    return;
  }

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correoElectronico)) {
    mostrarOverlay({
      mensaje: "⚠️ Ingresa un correo electrónico válido.",
      temporal: true,
      autoCerrar: true,
      tiempo: 3000
    });
    return;
  }

  /* =====================================
     ✅ AQUÍ ES DONDE SE OCULTA EL HEADER
     ===================================== */
  document.body.classList.add("sin-header");

  // Ocultar bloque 1
  bloque1.style.display = "none";

  // Mostrar bloque 2
  bloque2.style.visibility = "visible";
  bloque2.style.position = "static";
  bloque2.style.opacity = "1";
  bloque2.style.pointerEvents = "auto";

  selectTablas.disabled = false;
  btnExportWord.disabled = false;
  btnExportPdf.disabled = false;

  // Cargar Excel solo una vez
  if (!bloque2.dataset.loaded) {
    cargarExcel(RUTAS_ARCHIVOS.TablaExcel);
    bloque2.dataset.loaded = "true";
  }

  // Reset menú
  selectTablas.innerHTML = '<option value="" selected>Seleccione...</option>';
});


  // --- BOTÓN REINICIAR ---
  btnReiniciar.addEventListener("click", async () => {
    const ok = await mostrarOverlay({
      mensaje: "⚠️ <strong>Todos los datos se eliminarán.</strong><br><br>¿Deseas reiniciar la evaluación?",
      aceptar: true,
      cancelar: true,
      textoAceptar: "Sí, reiniciar",
      textoCancelar: "Cancelar"
    });

    if (ok) {
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

        // si ya había sido evaluada, marcar de nuevo
        if (hojasEvaluadas.has(nombreHoja)) {
          opt.classList.add("hoja-evaluada");
        }

        selectTablas.appendChild(opt);
      });

      // Evento change para cargar tabla al seleccionar hoja
      selectTablas.onchange = null;
      selectTablas.addEventListener("change", (e) => {
        const hojaSeleccionada = e.target.value;

        if (!hojaSeleccionada) {
          // Mostrar mensaje de ejemplo si se selecciona "Seleccione..."
          cargarTextoEjemplo(RUTAS_ARCHIVOS.MensajeEjemplo);
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
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, blankrows: false });

    // Limpiar contenedor
    bloque3.innerHTML = "";

    // Crear tabla
    const table = document.createElement("table");
    table.classList.add("tabla-excel");

    const anchos = ["55%", "15%", "25%"];
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

      const textoCelda = fila[0]?.toString() || "";

      const esTTT = textoCelda.startsWith("TTT-");
      const esTitulo = textoCelda.startsWith("TT-");
      const esSubtitulo = textoCelda.startsWith("T-");


      if (esTTT) {
        // === FILA TTT (negrilla + justificado, ocupa 3 columnas) ===
        const td = document.createElement("td");
        td.colSpan = 3;
        td.innerHTML = `<strong>${textoCelda.replace(/^TTT-/, "")}</strong>`;
        td.style.textAlign = "justify";
        td.style.verticalAlign = "middle";
        tr.appendChild(td);

      } else if (esTitulo) {
        // === FILA TT (título principal centrado) ===
        const td = document.createElement("td");
        td.colSpan = 3;
        td.innerHTML = `<strong>${textoCelda.replace(/^TT-/, "")}</strong>`;
        td.style.textAlign = "center";
        td.style.verticalAlign = "middle";
        tr.appendChild(td);

      } else if (esSubtitulo) {
        // === FILA T (subtítulo, col1 + col2) ===
        const td = document.createElement("td");
        td.colSpan = 2;
        td.innerHTML = `<em>${textoCelda.replace(/^T-/, "")}</em>`;
        td.style.textAlign = "justify";
        td.style.verticalAlign = "middle";
        tr.appendChild(td);

        // === Columna observación ===
        const tdObs = document.createElement("td");
        tdObs.style.verticalAlign = "middle";
        tr.appendChild(tdObs);



        const textarea = document.createElement("textarea");
        textarea.placeholder = "Ingrese observación";
        textarea.style.width = "95%";
        textarea.style.height = "50px";
        textarea.style.backgroundColor = "#faefbcff";

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
          td.style.textAlign = "justify";
          // Recuperar respuestas guardadas si existen
          const respuestasGuardadas = respuestasPorTabla[hoja]?.[i] || {};
          if (j === 1) {
            // === Columna Radios ===
            const opciones = ["Cumple", "No Cumple", "No Aplica"];
            const contenedorRadios = document.createElement("div");
            contenedorRadios.classList.add("opciones-radios");
            contenedorRadios.style.display = "flex";
            contenedorRadios.style.flexDirection = "column";  // << importante
            contenedorRadios.style.gap = "4px";
            contenedorRadios.style.alignItems = "flex-start";

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


              let estabaMarcado = false;

              input.addEventListener("mousedown", () => {
                estabaMarcado = input.checked;
              });

              input.addEventListener("click", () => {
                if (estabaMarcado) {
                  // 🔁 DESMARCAR
                  input.checked = false;

                  if (respuestasPorTabla[hoja]?.[i]) {
                    delete respuestasPorTabla[hoja][i].opcion;
                  }
                } else {
                  // ✅ MARCAR NORMAL
                  if (!respuestasPorTabla[hoja]) respuestasPorTabla[hoja] = {};
                  if (!respuestasPorTabla[hoja][i]) respuestasPorTabla[hoja][i] = {};
                  respuestasPorTabla[hoja][i].opcion = opcion;

                  registrarHojaEvaluada(hoja);
                }

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

  function registrarHojaEvaluada(hoja) {
    hojasEvaluadas.add(hoja);

    const select = document.getElementById("menu-tablas");
    const opt = select.querySelector(`option[value="${hoja}"]`);

    if (opt) {
      opt.classList.add("hoja-evaluada");
    }
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
          legend: {
            display: false,
            position: "top",
            labels: {
              padding: 15,
              font: { size: 12 }
            }
          },
          datalabels: {
            formatter: (value, ctx) => {
              const total = datos.reduce((a, b) => a + b, 0);
              if (total === 0) return "0%";
              return ((value / total) * 100).toFixed(1) + "%";
            },
            color: "#fff",
            font: {
              weight: "bold",
              size: 10
            }
          }
        }
      },
      plugins: [ChartDataLabels]
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
          backgroundColor: colores,
          borderColor: "#333",
          borderWidth: 1,
          hoverBackgroundColor: colores.map(c => c + "cc")
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          datalabels: {
            anchor: "end",
            align: "top",
            color: "#000",
            font: { size: 10, weight: "bold" },
            formatter: v => v,
            offset: 2
          }
        },
        layout: {
          padding: {
            left: 10,
            right: 10,
            top: 20,
            bottom: 5
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            precision: 0,
            suggestedMax: Math.max(...datos) + 1,
            ticks: {
              stepSize: 1
            }
          }
        }
      },
      plugins: [ChartDataLabels]
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

      // 🔹 Resetear select a "Seleccione..."
      const selectTablas = document.getElementById("menu-tablas");
      if (selectTablas) {
        selectTablas.value = "";
      }


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
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
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
        top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      },
      verticalAlign: "center",
      shading: undefined,
      cantSplit: noWrap
    });
  }

  // --- Construye la tabla DOCX preservando estructura y alineaciones ---
  function construirTablaDocx(headers, filas) {
    const anchoCols = [60, 15, 20]; // % como en el frontend

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


  async function exportarHojaWord() {
    // Validaciones básicas
    const nombreIps = document.getElementById("nombreIps")?.value?.trim() || "";
    const numeroContacto = document.getElementById("numeroContacto")?.value?.trim() || "";
    const correoElectronico = document.getElementById("correoElectronico")?.value?.trim() || "";
    const nombreTabla = selectTablas?.value || "";

    const { headers, filas } = leerTablaDesdeDOM();
    if (!headers.length || !filas.length) {
      if (!nombreTabla) {
        mostrarOverlay({
          mensaje: "Selecciona una tabla antes de exportar.",
          aceptar: false,
          cancelar: false,
          temporal: true,
          autoCerrar: true,
          tiempo: 3000,
          textoAceptar: "Aceptar",
          textoCancelar: "Cancelar"
        });
        //alert("Selecciona una tabla antes de exportar.");

        return;
      }
    }
    if (!logoBuffer) {
      mostrarOverlay({
        mensaje: "No se cargó el logo, se exportará sin imagen.",
        aceptar: false,
        cancelar: false,
        temporal: true,
        autoCerrar: true,
        tiempo: 3000,
        textoAceptar: "Aceptar",
        textoCancelar: "Cancelar"
      });
      //alert("No se cargó el logo, se exportará sin imagen.");

    }
    mostrarOverlay({
      mensaje: "Generando archivo Word...",
      aceptar: false,
      cancelar: false,
      temporal: true,
      autoCerrar: false,
      tiempo: 3000,
      textoAceptar: "Aceptar",
      textoCancelar: "Cancelar"
    });

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
    setTimeout(() => {
      ocultarOverlay();
    }, 3000);

  };

  async function obtenerHojasExportables(workbook) {
    const hojas = workbook.SheetNames.filter(
      h => h.toLowerCase() !== "tabla de contenido"
    );

    const menuTablas = document.getElementById("menu-tablas");
    const exportables = [];

    for (const hoja of hojas) {
      mostrarTabla(workbook, hoja);
      menuTablas.value = hoja;

      // ⏳ dar un frame al DOM para reflejar radios
      await new Promise(r => requestAnimationFrame(r));

      const radiosMarcados = bloque3.querySelectorAll(
        "input[type='radio']:checked"
      );

      if (radiosMarcados.length > 0) {
        exportables.push(hoja);
      }
    }

    return exportables;
  }

  async function exportarTodoWord() {

    if (!workbook) {
      mostrarOverlay({
        mensaje: "⚠️ Primero carga el Excel antes de exportar.",
        temporal: true
      });
      return;
    }

    mostrarOverlay({
      mensaje: "🔎 Analizando hojas...",
      autoCerrar: false
    });

    const hojasSeleccionadas = await obtenerHojasExportables(workbook);

    ocultarOverlay();

    if (hojasSeleccionadas.length === 0) {
      mostrarOverlay({
        mensaje: "⚠️ No hay hojas con criterios evaluados.",
        temporal: true
      });
      return;
    }

    // 🧾 Confirmación del usuario
    const mensajeHTML = `
    <strong>📄 Hojas a exportar:</strong><br><br>
    <ul style="text-align:left; padding-left:20px;">
      ${hojasSeleccionadas.map(h => `<li>${h}</li>`).join("")}
    </ul>
  `;

    const ok = await mostrarOverlay({
      mensaje: mensajeHTML,
      aceptar: true,
      cancelar: true,
      autoCerrar: false
    });

    if (!ok) {
      mostrarOverlay({
        mensaje: "❌ Exportación cancelada.",
        temporal: true
      });
      return;
    }

    // 🚀 SOLO AQUÍ empieza el trabajo pesado
    await generarWordDesdeHojas(hojasSeleccionadas);
  }

  async function generarWordDesdeHojas(hojasSeleccionadas) {

    const secciones = [];
    const menuTablas = document.getElementById("menu-tablas");

    mostrarOverlay({
      mensaje: "📄 Generando informe...",
      autoCerrar: false
    });

    for (const hoja of hojasSeleccionadas) {

      mostrarTabla(workbook, hoja);
      menuTablas.value = hoja;

      actualizarContadoresCol2();
      await new Promise(r => setTimeout(r, 800));


      const { headers, filas } = leerTablaDesdeDOM();

      if (!headers.length || !filas.length) continue;

      const section = construirSeccionWord({
        nombreTabla: hoja,
        headers,
        filas
      });

      secciones.push(section);
    }

    ocultarOverlay();

    const documento = new Document({
      styles: {
        default: {
          document: { run: { font: "Arial", size: 22 } }
        }
      },
      sections: secciones
    });

    const blob = await Packer.toBlob(documento);
    saveAs(blob, `Evaluacion_Todas_${fechaISO()}.docx`);

    mostrarOverlay({
      mensaje: "✅ Exportación completada.",
      temporal: true
    });
    cargarTextoEjemplo(RUTAS_ARCHIVOS.MensajeEjemplo);
  }

  function construirSeccionWord({ nombreTabla, headers, filas }) {

    return {
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              children: [
                new TextRun({ text: "Resultados Criterios de Evaluación", bold: true }),
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
      children: [
        new Paragraph({
          children: [new TextRun({ text: `Tabla: ${nombreTabla}`, bold: true })],
          alignment: AlignmentType.CENTER
        }),

        new Paragraph({ text: "" }),

        construirTablaResumenDocx(),
        new Paragraph({ text: "" }),

        construirTablaGraficasDocx(),
        new Paragraph({ text: "" }),

        construirTablaDocx(headers, filas),

        new Paragraph({ text: "" }),
        new Paragraph({ text: "" })
      ]
    };
  }


  //boton para exportar todo la evaluacion
  btnExportTodo.addEventListener("click", () => {
    const tipo = obtenerTipoExportacion();
    if (tipo === "pdf") {
      //alert("Exportando Todo Como PDF  en implementación");
      exportarTodoPDF();
    } else {
      //alert("Exportando Todo Como Word");
      exportarTodoWord();
    }
  });

  // === //boton para exportar la hoja ===
  btnExportWord.addEventListener("click", async () => {
    const tipo = obtenerTipoExportacion();

    if (tipo === "pdf") {
      //alert("exportar Hoja a PDF en implementación");
      await exportarHojaPDF();
    } else {
      //alert("exportar Hoja a Word");
      await exportarHojaWord();
    }
  });


  // --- BOTÓN TEST OVER ---
  /*  btntestOv.addEventListener("click", () => {
      mostrarOverlay({
        mensaje: "Overlay temporal OK",
        aceptar: true,
        cancelar: true,
        temporal: true,
        autoCerrar: true,
        tiempo: 3000,
        textoAceptar: "Aceptar",
        textoCancelar: "Cancelar"
      });
  
  
  
    }); 
    */

  function mostrarOverlay({
    mensaje,
    aceptar = false,
    cancelar = false,
    temporal = false,
    autoCerrar = true,
    tiempo = 3000,
    textoAceptar = "Aceptar",
    textoCancelar = "Cancelar"
  }) {
    const overlay = document.getElementById("overlay");
    const texto = document.getElementById("overlay-texto");
    const btnOk = document.getElementById("overlay-continuar");
    const btnCancel = document.getElementById("overlay-cancelar");

    // limpiar timer previo
    if (overlayTimer) {
      clearTimeout(overlayTimer);
      overlayTimer = null;
    }

    texto.innerHTML = mensaje;
    overlay.classList.remove("oculto");

    btnOk.style.display = aceptar ? "inline-block" : "none";
    btnCancel.style.display = cancelar ? "inline-block" : "none";

    btnOk.textContent = textoAceptar;
    btnCancel.textContent = textoCancelar;

    // ⏱️ autocierre
    if (temporal && autoCerrar) {
      overlayTimer = setTimeout(() => {
        ocultarOverlay();
      }, tiempo);
    }

    // 🔑 SI NO HAY BOTONES → NO PROMISE
    if (!aceptar && !cancelar) return;

    // 🔐 PROMISE PARA ESPERAR DECISIÓN
    return new Promise(resolve => {

      const aceptarHandler = () => {
        limpiar();
        resolve(true);
      };

      const cancelarHandler = () => {
        limpiar();
        resolve(false);
      };

      function limpiar() {
        btnOk.removeEventListener("click", aceptarHandler);
        btnCancel.removeEventListener("click", cancelarHandler);
        ocultarOverlay();
      }

      btnOk.addEventListener("click", aceptarHandler);
      btnCancel.addEventListener("click", cancelarHandler);
    });
  }



  function ocultarOverlay() {
    const overlay = document.getElementById("overlay");
    if (overlayTimer) {
      clearTimeout(overlayTimer);
      overlayTimer = null;
    }
    overlay.classList.add("oculto");
  }

  function obtenerTipoExportacion() {
    const esWord = document.getElementById("tipo-exportacion").checked;
    return esWord ? "word" : "pdf";
  }

  // funciones para exportar a

  const MARGEN_SUPERIOR = 800;
  const MARGEN_INFERIOR = 100;
  const LINE_HEIGHT = 12;

  function dividirTextoEnLineas(texto, font, size, maxWidth) {
    const palabras = texto.split(" ");
    const lineas = [];
    let linea = "";

    for (const palabra of palabras) {
      const test = linea ? linea + " " + palabra : palabra;
      const ancho = font.widthOfTextAtSize(test, size);

      if (ancho > maxWidth) {
        lineas.push(linea);
        linea = palabra;
      } else {
        linea = test;
      }
    }

    if (linea) lineas.push(linea);
    return lineas;
  }


  function descargarArchivo(bytes, nombre) {
    const blob = new Blob([bytes], { type: "application/pdf" });
    const url = URL.createObjectURL(blob);

    const link = document.createElement("a");
    link.href = url;
    link.download = nombre;
    document.body.appendChild(link);
    link.click();

    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }

  async function insertarCanvasEnPDF({
    canvasId,
    titulo = "",
    maxWidth = 480
  }) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;

    const dataUrl = canvas.toDataURL("image/png", 1.0);
    const img = await pdfDoc.embedPng(dataUrl);

    const scale = maxWidth / img.width;
    const imgHeight = img.height * scale;

    if (y - imgHeight < MARGEN_INFERIOR) {
      page = pdfDoc.addPage([595, 842]);
      y = MARGEN_SUPERIOR;
    }

    if (titulo) {
      page.drawText(titulo, {
        x: 50,
        y,
        size: 10,
        font: fontBold
      });
      y -= 14;
    }

    page.drawImage(img, {
      x: 50,
      y: y - imgHeight,
      width: maxWidth,
      height: imgHeight
    });

    y -= imgHeight + 20;
  }



  async function exportarHojaPDF() {
    const selectTablas = document.getElementById("menu-tablas");
    const nombreTabla = selectTablas?.value?.trim();

    if (!nombreTabla) {
      mostrarOverlay({
        mensaje: "⚠️ Debes seleccionar una hoja antes de exportar el PDF.",
        temporal: true,
        autoCerrar: true
      });
      return;
    }

    mostrarOverlay({
      mensaje: "Generando PDF...",
      temporal: true,
      autoCerrar: false
    });

    try {
      // ================= CONFIG =================
      const { PDFDocument, StandardFonts } = PDFLib;


      // ================= DATOS =================
      const nombreIps = document.getElementById("nombreIps")?.value?.trim() || "";
      const numeroContacto = document.getElementById("numeroContacto")?.value?.trim() || "";
      const correoElectronico = document.getElementById("correoElectronico")?.value?.trim() || "";
      const nombreTabla = selectTablas?.value || "Sin tabla";

      const { filas } = leerTablaDesdeDOM();

      // ================= PDF =================
      const pdfDoc = await PDFDocument.create();
      let page = pdfDoc.addPage([595, 842]);
      const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
      const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

      let y = MARGEN_SUPERIOR;

      // ================= ENCABEZADO =================
      page.drawText("Auditoría de Criterios de Habilitación – Resolución 3100", {
        x: 50,
        y,
        size: 14,
        font: fontBold
      });

      y -= 25;

      [
        `IPS / Profesional: ${nombreIps || "-"}`,
        `Contacto: ${numeroContacto || "-"}`,
        `Correo: ${correoElectronico || "-"}`,
        `Tabla evaluada: ${nombreTabla || "-"}`
      ].forEach(txt => {
        page.drawText(txt, { x: 50, y, size: 10, font });
        y -= 14;
      });

      y -= 20;

      // ================= COLUMNAS =================
      const colX = { criterio: 50, eval: 350, obs: 420 };
      const colWidth = { criterio: 280, eval: 60, obs: 120 };

      // ================= FUNCIONES =================
      function dibujarEncabezadoTabla() {
        page.drawText("CRITERIO", { x: colX.criterio, y, size: 9, font: fontBold });
        page.drawText("EVAL.", { x: colX.eval, y, size: 9, font: fontBold });
        page.drawText("OBSERVACIONES", { x: colX.obs, y, size: 9, font: fontBold });
        y -= 8;

        page.drawLine({
          start: { x: 50, y },
          end: { x: 545, y },
          thickness: 1
        });

        y -= 10;
      }

      function dibujarTextoConSalto({ texto, x, maxWidth, size }) {
        const lineas = dividirTextoEnLineas(texto, font, size, maxWidth);

        for (const linea of lineas) {
          if (y - LINE_HEIGHT < MARGEN_INFERIOR) {
            page = pdfDoc.addPage([595, 842]);
            y = MARGEN_SUPERIOR;
            dibujarEncabezadoTabla();
          }

          page.drawText(linea, { x, y, size, font });
          y -= LINE_HEIGHT;
        }
      }

      async function insertarCanvasEnPDF({
        canvasId,
        titulo = "",
        maxWidth = 300
      }) {
        const canvas = document.getElementById(canvasId);
        if (!canvas) return;

        const dataUrl = canvas.toDataURL("image/png", 1.0);
        const img = await pdfDoc.embedPng(dataUrl);

        const scale = maxWidth / img.width;
        const imgHeight = img.height * scale;

        if (y - imgHeight < MARGEN_INFERIOR) {
          page = pdfDoc.addPage([595, 842]);
          y = MARGEN_SUPERIOR;
        }

        if (titulo) {
          page.drawText(titulo, {
            x: 50,
            y,
            size: 10,
            font: fontBold
          });
          y -= 14;
        }

        page.drawImage(img, {
          x: 50,
          y: y - imgHeight,
          width: maxWidth,
          height: imgHeight
        });

        y -= imgHeight + 20;
      }

      function insertarResumenEnPDF() {

        const resumen = [
          { label: "Cumple", valor: "num-cumple", porcentaje: "p-cumple" },
          { label: "No cumple", valor: "num-nocumple", porcentaje: "p-nocumple" },
          { label: "No aplica", valor: "num-noaplica", porcentaje: "p-noaplica" }
        ];

        const alturaNecesaria = resumen.length * 16 + 40;

        if (y - alturaNecesaria < MARGEN_INFERIOR) {
          page = pdfDoc.addPage([595, 842]);
          y = MARGEN_SUPERIOR;
        }

        // Título
        page.drawText("Resumen", {
          x: 50,
          y,
          size: 12,
          font: fontBold
        });

        y -= 20;

        // Encabezados
        page.drawText("Estado", { x: 50, y, size: 10, font: fontBold });
        page.drawText("Cantidad", { x: 250, y, size: 10, font: fontBold });
        page.drawText("Porcentaje", { x: 350, y, size: 10, font: fontBold });

        y -= 8;

        page.drawLine({
          start: { x: 50, y },
          end: { x: 500, y },
          thickness: 1
        });

        y -= 12;

        // Filas
        resumen.forEach(r => {
          const cantidad = document.getElementById(r.valor)?.innerText || "0";
          const porcentaje = document.getElementById(r.porcentaje)?.innerText || "0";

          page.drawText(r.label, { x: 50, y, size: 10, font });
          page.drawText(cantidad, { x: 260, y, size: 10, font });
          page.drawText(porcentaje, { x: 360, y, size: 10, font });

          y -= 16;
        });

        y -= 10;
      }

      // ===== RESUMEN =====
      insertarResumenEnPDF();

      // ================= GRÁFICOS =================
      page.drawText("Gráficos de Resultados", {
        x: 50,
        y,
        size: 12,
        font: fontBold
      });

      y -= 20;

      await insertarCanvasEnPDF({
        canvasId: "graficaBarras",
        titulo: "Resultados por Criterio"
      });

      await insertarCanvasEnPDF({
        canvasId: "graficaTorta",
        titulo: "Distribución de Cumplimiento"
      });
      // ================= DETALLE =================
      page.drawText("Detalle de Criterios", {
        x: 50,
        y,
        size: 12,
        font: fontBold
      });

      y -= 20;
      dibujarEncabezadoTabla();

      for (const f of filas) {

        if (y < MARGEN_INFERIOR + 40) {
          page = pdfDoc.addPage([595, 842]);
          y = MARGEN_SUPERIOR;
          dibujarEncabezadoTabla();
        }

        if (f.tipo === "seccion") {
          page.drawText(f.titulo, { x: 50, y, size: 10, font: fontBold });
          y -= 16;
          continue;
        }

        if (f.tipo === "subtitulo") {
          page.drawText(f.subtitulo, { x: 55, y, size: 9, font: fontBold });
          y -= 14;
          continue;
        }

        const yInicioFila = y;

        dibujarTextoConSalto({
          texto: f.criterio,
          x: colX.criterio,
          maxWidth: colWidth.criterio,
          size: 9
        });

        page.drawText(f.evaluacion, {
          x: colX.eval,
          y: yInicioFila,
          size: 9,
          font
        });

        dibujarTextoConSalto({
          texto: f.observaciones || "-",
          x: colX.obs,
          maxWidth: colWidth.obs,
          size: 9
        });

        y -= 6;
      }

      // ================= GUARDAR =================
      const pdfBytes = await pdfDoc.save();

      descargarArchivo(
        pdfBytes,
        `Hoja_${sanitizeFileName(nombreTabla)}_${fechaISO()}.pdf`
      );

    } catch (error) {
      console.error("Error al generar PDF:", error);
      alert("Error al generar el PDF. Revisa la consola.");
    } finally {
      setTimeout(() => {
        ocultarOverlay();
      }, 3000);
    }
    cargarTextoEjemplo(RUTAS_ARCHIVOS.MensajeEjemplo);
  }

  //exportar todo a PDF

  const colX = {
    criterio: 50,
    eval: 360,
    obs: 420
  };

  const colWidth = {
    criterio: 290,
    eval: 50,
    obs: 115
  };


  function cambiarHoja(valor) {
    const select = document.getElementById("menu-tablas");
    select.value = valor;
    select.dispatchEvent(new Event("change"));
  }

  async function esperarTablaRenderizada(timeout = 2000) {
    const inicio = Date.now();

    while (Date.now() - inicio < timeout) {
      const contenedor = document.getElementById("contenedor-tabla");
      const radios = contenedor?.querySelectorAll('input[type="radio"]');

      if (radios && radios.length > 0) return true;

      await new Promise(r => setTimeout(r, 50));
    }

    return false;
  }

  function hojasConCriteriosEvaluados() {
    return Object.entries(respuestasPorTabla)
      .filter(([_, respuestas]) =>
        Object.values(respuestas).some(r => r.opcion)
      )
      .map(([id]) => id);
  }

  function dividirTexto(texto, font, size, maxWidth) {
    if (!texto) return ["-"];

    const palabras = texto.split(" ");
    const lineas = [];
    let linea = "";

    for (const palabra of palabras) {
      const prueba = linea ? linea + " " + palabra : palabra;
      const ancho = font.widthOfTextAtSize(prueba, size);

      if (ancho <= maxWidth) {
        linea = prueba;
      } else {
        lineas.push(linea);
        linea = palabra;
      }
    }

    if (linea) lineas.push(linea);
    return lineas;
  }

  function tipoFila(texto) {
    if (!texto) return "criterio";
    if (texto.startsWith("TT-")) return "titulo";
    if (texto.startsWith("T-")) return "subtitulo";
    return "criterio";
  }



  async function exportarTodoPDF() {

    const menu = document.getElementById("menu-tablas");

    /* ===============================
       1️⃣ Determinar hojas a exportar (DATOS)
    =============================== */
    const hojasAExportar = hojasConCriteriosEvaluados();

    if (hojasAExportar.length === 0) {
      mostrarOverlay({
        mensaje: "⚠️ Ninguna hoja tiene criterios evaluados.",
        aceptar: true
      });
      return;
    }

    /* ===============================
       2️⃣ Obtener nombres visibles (UI)
    =============================== */
    const hojasSeleccionadas = hojasAExportar.map(id => {
      const opt = [...menu.options].find(o => o.value === id);
      return opt?.text || id;
    });

    /* ===============================
       3️⃣ Overlay previo (IGUAL A WORD)
    =============================== */
    const mensajeHTML = `
    <strong>📄 Hojas a exportar:</strong><br><br>
    <ul style="text-align:left; margin:0; padding-left:20px;">
      ${hojasSeleccionadas.map(h => `<li>${h}</li>`).join("")}
    </ul>
  `;

    const ok = await mostrarOverlay({
      mensaje: mensajeHTML,
      aceptar: true,
      cancelar: true,
      autoCerrar: false,
      textoAceptar: "Aceptar",
      textoCancelar: "Cancelar"
    });

    if (!ok) return;

    mostrarOverlay({
      mensaje: "Generando PDF consolidado...",
      temporal: true,
      autoCerrar: false
    });

    try {
      const { PDFDocument, StandardFonts } = PDFLib;

      const pdfDoc = await PDFDocument.create();
      const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
      const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

      /* ===============================
         4️⃣ Recorrer hojas y exportar
      =============================== */
      for (const hojaId of hojasAExportar) {

        cambiarHoja(hojaId);
        await esperarTablaRenderizada();

        const { filas } = leerTablaDesdeDOM();

        let page = pdfDoc.addPage([595, 842]);
        let y = MARGEN_SUPERIOR;

        const nombreHoja =
          [...menu.options].find(o => o.value === hojaId)?.text || hojaId;

        page.drawText(`Hoja: ${nombreHoja}`, {
          x: 50,
          y,
          size: 14,
          font: fontBold
        });

        y -= 25;

        const colX = { criterio: 50, eval: 360, obs: 420 };

        page.drawText("CRITERIO", { x: colX.criterio, y, size: 9, font: fontBold });
        page.drawText("EVAL.", { x: colX.eval, y, size: 9, font: fontBold });
        page.drawText("OBS.", { x: colX.obs, y, size: 9, font: fontBold });

        y -= 8;

        page.drawLine({
          start: { x: 50, y },
          end: { x: 545, y },
          thickness: 1
        });

        y -= 12;

        /* ===============================
       6️⃣ Dibujar TODOS los criterios (con wrap)
    =============================== */
        for (const f of filas) {

          const lineasCriterio = dividirTexto(
            f.criterio || "-",
            font,
            9,
            colWidth.criterio
          );

          const lineasObs = dividirTexto(
            f.observaciones || "-",
            font,
            9,
            colWidth.obs
          );

          const altoFila = Math.max(
            lineasCriterio.length,
            lineasObs.length,
            1
          ) * LINE_HEIGHT;

          if (y - altoFila < MARGEN_INFERIOR + 40) {
            page = pdfDoc.addPage([595, 842]);
            y = MARGEN_SUPERIOR;
          }

          // CRITERIO (multilínea)
          lineasCriterio.forEach((linea, i) => {
            page.drawText(linea, {
              x: colX.criterio,
              y: y - i * LINE_HEIGHT,
              size: 9,
              font
            });
          });

          // EVALUACIÓN (una línea, centrada visualmente)
          page.drawText(f.evaluacion || "-", {
            x: colX.eval,
            y,
            size: 9,
            font
          });

          // OBSERVACIONES (multilínea)
          lineasObs.forEach((linea, i) => {
            page.drawText(linea, {
              x: colX.obs,
              y: y - i * LINE_HEIGHT,
              size: 9,
              font
            });
          });

          y -= altoFila;
        }


      }

      /* ===============================
         5️⃣ Guardar PDF
      =============================== */
      const pdfBytes = await pdfDoc.save();

      descargarArchivo(
        pdfBytes,
        `Auditoria_Consolidada_${fechaISO()}.pdf`
      );

    } catch (error) {
      console.error("Error al generar PDF:", error);
      mostrarOverlay({
        mensaje: "❌ Error al generar el PDF.",
        aceptar: true
      });
    } finally {
      setTimeout(() => {
        ocultarOverlay();
      }, 3000);
    }
    cargarTextoEjemplo(RUTAS_ARCHIVOS.MensajeEjemplo);
  }


  /* ==== LICENCIA DE USO ======*/
  let correosAutorizados = [];

  async function cargarCorreosAutorizados() {
    if (correosAutorizados.length > 0) return;

    const res = await fetch(RUTAS_ARCHIVOS.correosAutorizados);
    let base64 = await res.text();

    // Limpieza crítica
    base64 = base64
      .replace(/\s+/g, "")
      .trim();

    // Decodificar base64 → texto
    const texto = atob(base64);

    correosAutorizados = texto
      .split(/\r?\n/)
      .map(linea => linea.trim())
      .filter(linea => linea.length > 0)
      .map(linea => {
        // soporta TXT o CSV
        return linea.split(/[;,]/)[0]
          .toLowerCase()
          .trim();
      });
  }



  async function correoAutorizado(correo) {
    if (!correo || !correo.includes("@")) return false;

    await cargarCorreosAutorizados();
    return correosAutorizados.includes(
      correo.toLowerCase().trim()
    );
  }


  document.getElementById("tipo-exportacion")
  .addEventListener("change", async function () {

    if (this.checked) { // Word
      const correo = document.getElementById("correoElectronico")?.value || "";

      const autorizado = await correoAutorizado(correo);

      if (!autorizado) {

        const irContacto = await mostrarOverlay({
          mensaje: "La exportación a Word es una opción de pago",
          aceptar: true,
          cancelar: true,
          textoAceptar: "Contactar",
          textoCancelar: "Cerrar"
        });

        // 🔹 Si acepta → abrir página
        if (irContacto) {
          window.open(
            "https://ssh-excelencia.github.io/#contacto", // 👈 cambia la URL
            "_blank",
            "noopener,noreferrer"
          );
        }

        // 🔹 Volver a PDF
        this.checked = false;
      }
    }
  });



});
