import { menus } from "./config/menu.js";

document.addEventListener("DOMContentLoaded", () => {

  const btnLoginTop = document.getElementById("btnLoginTop");
  const loginOverlay = document.getElementById("loginOverlay");
  const appOverlay = document.getElementById("appOverlay");

  /* ABRIR LOGIN */
  btnLoginTop.onclick = () => {
    loginOverlay.classList.remove("hidden");
    document.getElementById('usuario').focus();
   
  };

  /* LOGIN */
  document.getElementById("btnLogin").onclick = () => {
    
    const usuario = document.getElementById("usuario").value.trim();
    const password = document.getElementById("password").value.trim();
    const rol = document.getElementById("rol").value;
    
    if (!usuario) {
    alert("Ingrese usuario");
    return;
    }

    if (!password) {
    alert("Ingrese Password");
    return;
    }
   

    // 🔥 estado global
    window.AppState = {
      usuario,
      rol
    };

    localStorage.setItem("usuario", usuario);
    localStorage.setItem("rol", rol);

    // Pintar en footer
    actualizarFooter();

    // ocultar login
    loginOverlay.classList.add("hidden");

    // 🔥 mostrar app correctamente
    appOverlay.classList.remove("hidden");

    // bloquear fondo
    document.body.style.overflow = "hidden";
    limpiarLogin();
    cargarMenu(rol);
  
  };

  function actualizarFooter() {
  const col1 = document.getElementById("footer-col-1");
  const col3 = document.getElementById("footer-col-3");

  col1.innerHTML = `👤 ${AppState.usuario}`;
  col3.innerHTML = `🔐 ${AppState.rol}`;
}

  document.getElementById("togglePassword").onclick = () => {
  const input = document.getElementById("password");
  input.type = input.type === "password" ? "text" : "password";
};



  /* CANCELAR LOGIN */
  document.getElementById("btnCancelar").onclick = () => {
    loginOverlay.classList.add("hidden");
    limpiarLogin();
  };

  /* CERRAR AL HACER CLIC FUERA */
  loginOverlay.onclick = (e) => {
    if (e.target === loginOverlay) {
      loginOverlay.classList.add("hidden");
      limpiarLogin();
    }
  };

  function limpiarLogin() {
  document.getElementById("usuario").value = "";
  document.getElementById("password").value = "";
  document.getElementById("rol").value = "Usuario";
}

document.getElementById("linkRegistro").addEventListener("click", (e) => {
  e.preventDefault(); // 🔥 evita que el link recargue o suba arriba

  document.getElementById("registroOverlay").classList.remove("hidden");
});

document.getElementById("btnCerrarRegistro").onclick = () => {
  document.getElementById("NUsuario").value= "";
  document.getElementById("Tdoc").selectedIndex = 0;
  document.getElementById("NDocumento").value= "";
  document.getElementById("NDocumento").value= "";
  document.getElementById("direccion").value= "";
  document.getElementById("Ciudad").value= "";
  document.getElementById("regEmail").value= "";
  document.getElementById("regPassword").value= "";
  document.getElementById("regPassword2").value= "";
  document.getElementById("contacto").value= "";
  document.getElementById("FNacimiento").value= "";
  document.getElementById("registroOverlay").classList.add("hidden");
};

document.getElementById("regPassword2").addEventListener("blur", () => {
  const pass1 = document.getElementById("regPassword").value;
  const pass2 = document.getElementById("regPassword2").value;

  if (!pass2) return; // si está vacío no hace nada

  if (pass1 !== pass2) {
    alert("Las contraseñas no coinciden");
     document.getElementById("regPassword2").value = "";
    document.getElementById("regPassword2").focus();
  }
});

document.getElementById("btnGuardarUsuario").onclick = () => {
  
  const usuario = document.getElementById("NUsuario").value.trim();
  const email = document.getElementById("regEmail").value.trim();
  const password = document.getElementById("regPassword").value.trim();
  const rol = "Usuario";
  

  if (!usuario || !password) {
    alert("Complete los campos obligatorios");
    return;
  }

  const nuevoUsuario = {
    usuario,
    email,
    password,
    rol
  };

  console.log("Usuario creado:", nuevoUsuario);

  alert("Usuario registrado - simulado");
  

  document.getElementById("registroOverlay").classList.add("hidden");
};

  /* LOGOUT */
  document.getElementById("logout").onclick = () => {
    localStorage.removeItem("rol");

    // 🔥 ocultar app correctamente
    appOverlay.classList.add("hidden");

    // restaurar scroll
    document.body.style.overflow = "auto";
  };
  
  /* PARA EL ABATAR */
  function abrirPerfil() {
  alert("Ir a perfil");
}

function logout() {
  alert("Cerrar sesión");
  // aquí tu lógica real
}


//aun no se usa, pero esta para implemntar
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


  /* MENÚ */
function cargarMenu(rol) {

   const sidebar = document.getElementById("sidebar");
  sidebar.innerHTML = "";

  (menus[rol] || []).forEach(op => {
    const item = document.createElement("div");
    item.textContent = op.label;
    
    item.onclick = () => {
      
      cargarVista(op.module);
      sidebar.classList.remove("open");
    };

    sidebar.appendChild(item);
  });
    cargarVista("Ppal")
}

const btnMenu = document.getElementById("btnMenu");
const sidebar = document.getElementById("sidebar");

btnMenu.addEventListener("click", () => {
  sidebar.classList.toggle("open");
});

document.addEventListener("click", (e) => {
  const isClickInside = sidebar.contains(e.target);
  const isButton = e.target.id === "btnMenu";

  if (!isClickInside && !isButton) {
    sidebar.classList.remove("open");
  }
});
});

/* VISTAS */
async function cargarVista(moduleName) {
  const contenedor = document.getElementById("contenido");

  contenedor.innerHTML = "<p>Cargando...</p>";

  const mod = await import(`./modulos/${moduleName}.js`);

  mod.render(contenedor);
}

