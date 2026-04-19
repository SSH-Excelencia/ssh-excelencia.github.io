export function render(contenedor) {
  contenedor.innerHTML = `
   <div class="pagina-principal">

      <div class="col col-1">
        <h2>Columna 1</h2>
        <p>Panel lateral izquierdo</p>
      </div>

      <div class="col col-2">
        <h2>Pagina principal de la App</h2>
        <p>Módulo cargado dinámicamente</p>
      </div>

      <div class="col col-3">
        <h2>Columna 3</h2>
        <p>Panel lateral Derecho</p>
      </div>

    </div>
  `;
}