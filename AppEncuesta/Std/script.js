let respuestas = { p1: "", p2: "", p3: "" };


function toggleMenu() {
  const menu = document.getElementById("menu");
  menu.style.display = (menu.style.display === "block") ? "none" : "block";
}

function mostrarInfo() {
  document.getElementById("infoModal").style.display = "flex";
}

function cerrarInfo() {
  document.getElementById("infoModal").style.display = "none";
}

function mostrarEncuesta() {
  document.getElementById("encuesta").style.display = "grid";
}

function seleccionar(boton, pregunta, valor) {
  respuestas[pregunta] = valor;
  // quitar selección previa
  boton.parentNode.querySelectorAll("button").forEach(b => b.classList.remove("seleccionado"));
  // marcar seleccionado
  boton.classList.add("seleccionado");
}

function limpiarRespuestas() {
  respuestas = { p1: "", p2: "", p3: "" };
  document.querySelectorAll(".opciones button").forEach(b => b.classList.remove("seleccionado"));
  document.getElementById("respuesta4").value = "";
}

function mostrarGracias() {
  const overlay = document.getElementById("graciasOverlay");
  overlay.style.display = "flex";
  setTimeout(() => {
    overlay.style.display = "none";
    limpiarRespuestas();
  }, 5000); // 5 segundos
}

function guardarCSVHorizontal() {
  const { p1, p2, p3 } = respuestas;
  const p4 = document.getElementById("respuesta4").value.trim();

  if (!p1 || !p2 || !p3) {
    alert("Por favor responde las preguntas 1, 2 y 3 antes de enviar.");
    return;
  }

  const fecha = new Date();
  const fechaTexto = fecha.toLocaleString("es-CO");
  const opcionesMes = { month: 'long' };
  const mes = fecha.toLocaleDateString("es-CO", opcionesMes);

  const fila = [fechaTexto, p1, p2, p3, p4].join(",") + "\n";
  const blob = new Blob([fila], { type: "text/csv;charset=utf-8;" });

  const nombreArchivo = `${mes}_respuestas.csv`;

  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = nombreArchivo;
  link.click();

  limpiarRespuestas();
  mostrarGracias();
}

async function consolidarArchivos() {
  try {
    const ua = navigator.userAgent.toLowerCase();
    if (!(ua.includes("chrome") && !ua.includes("edg") && !ua.includes("opr"))) {
      alert("Este proceso solo funciona en Google Chrome.");
      return;
    }

    const handles = await window.showOpenFilePicker({
      multiple: true,
      types: [{ description: "Archivos CSV", accept: { "text/csv": [".csv"] } }]
    });

    const todasFilas = [["Fecha", "Hora", "P1", "P2", "P3", "P4"]];

    for (let handle of handles) {
      const file = await handle.getFile();
      const text = await file.text();
      const lineas = text.trim().split("\n");
      lineas.forEach(l => todasFilas.push(l.split(",")));
    }

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(todasFilas);
    XLSX.utils.book_append_sheet(wb, ws, "Consolidado");

    const fecha = new Date();
    const opciones = { month: "long", year: "numeric" };
    const mesAnio = fecha.toLocaleDateString("es-CO", opciones);

    const nombreConsolidado = `Consolidado_${mesAnio}.xlsx`;
    XLSX.writeFile(wb, nombreConsolidado);

  } catch (err) {
    alert("Error al consolidar: " + err.message);
  }
}

 async function cargarNombreEmpresa() {
  try {
    //const respuesta = await fetch("empresa.txt");   // lee el txt
    const respuesta = await fetch("empresa.txt?nocache=" + new Date().getTime());
    const texto = await respuesta.text();           // convierte en string
    document.getElementById("empresa").textContent = texto.trim(); // asigna
  } catch (error) {
    console.error("Error al cargar empresa.txt:", error);
    document.getElementById("empresa").textContent = "Encuesta de Satisfacción"; // valor por defecto
  }
}
document.addEventListener("DOMContentLoaded", cargarNombreEmpresa);

window.onload = function () {
    // fuerza recarga sin usar caché
    window.location.reload(true);
  };

