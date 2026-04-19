// js/config/menu.js

export const menus = {
  Usuario: [
    { label: "Solicitar Cita", module: "SCita" },
    { label: "Consultar Citas", module: "CCitas" },
    { label: "Cancelar Cita", module: "Xcita" },
    { label: "Actualizar Información", module: "Actualiza" }
  ],

  Operador: [
    { label: "Gestión Pacientes", module: "Pacientes" },
    { label: "Historia Clínica", module: "Historia" },
    { label: "Asignar Citas", module: "Agenda" },
    { label: "Formulas", module: "Formulas" },
    { label: "Servicios", module: "Servicios" }
  ],

  Administrador: [
    { label: "Configuraciones", module: "Config" },
    { label: "Usuarios", module: "Usuarios" },
    { label: "Administrar Personal", module: "Personal" },
    { label: "Consultorios", module: "Consultorio" },
    { label: "Infraestructura", module: "Infraestructura" },
    { label: "Formatos Historia Clinica", module: "Formatos" }
  ]
};