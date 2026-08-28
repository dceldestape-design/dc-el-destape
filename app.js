/**
 * ==========================================================================
 * SISTEMA DE INVENTARIO Y VENTAS (LICORES & BEBIDAS) - v2.0 BI-MONEDA
 * ==========================================================================
 * Incluye catálogo de 54 licores, gestión de stock derivado, costo promedio
 * ponderado, soporte USD/CRC, exportación/importación Excel (.xlsx) y
 * sincronización bidireccional con Google Sheets.
 */

// --- Formatters ---
const fmtUSD = (n) =>
  new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" }).format(
    Number.isFinite(n) ? n : 0
  );
const fmtCRC = (n) =>
  new Intl.NumberFormat("es-CR", { style: "currency", currency: "CRC", maximumFractionDigits: 0 }).format(
    Number.isFinite(n) ? n : 0
  );
const fmtNum = (n) => new Intl.NumberFormat("es-CR").format(Number.isFinite(n) ? n : 0);
const todayStr = () => new Date().toISOString().slice(0, 10);
const uid = () => Math.random().toString(36).slice(2, 10);

// --- Catálogo Semilla de Licores y Bebidas ---
const SEED_PRODUCTS = [
  {"codigo": "WHI-001", "nombre": "Johnnie Walker Red label litro", "categoria": "WHISKY/WHISKEY", "costoRefUSD": 12.0, "costoRefCRC": 5640.0, "precioVentaUSD": 21.28, "precioVentaCRC": 10000.0, "stockInicial": 0},
  {"codigo": "WHI-002", "nombre": "Johnnie Walker Black label litro", "categoria": "WHISKY/WHISKEY", "costoRefUSD": 22.0, "costoRefCRC": 10340.0, "precioVentaUSD": 40.43, "precioVentaCRC": 19000.0, "stockInicial": 1},
  {"codigo": "WHI-003", "nombre": "Johnnie Walker Double black label litro", "categoria": "WHISKY/WHISKEY", "costoRefUSD": 30.0, "costoRefCRC": 14100.0, "precioVentaUSD": 48.94, "precioVentaCRC": 23000.0, "stockInicial": 1},
  {"codigo": "WHI-004", "nombre": "Johnnie Walker Gold Label botella", "categoria": "WHISKY/WHISKEY", "costoRefUSD": 50.0, "costoRefCRC": 23500.0, "precioVentaUSD": 85.11, "precioVentaCRC": 40000.0, "stockInicial": 0},
  {"codigo": "WHI-005", "nombre": "Johnnie Walker Green Label botella", "categoria": "WHISKY/WHISKEY", "costoRefUSD": 50.0, "costoRefCRC": 23500.0, "precioVentaUSD": 85.11, "precioVentaCRC": 40000.0, "stockInicial": 0},
  {"codigo": "WHI-006", "nombre": "Johnnie Walker Blue Label botella", "categoria": "WHISKY/WHISKEY", "costoRefUSD": 175.0, "costoRefCRC": 82250.0, "precioVentaUSD": 255.32, "precioVentaCRC": 120000.0, "stockInicial": 0},
  {"codigo": "WHI-007", "nombre": "Old Parr 12 años litro", "categoria": "WHISKY/WHISKEY", "costoRefUSD": 22.0, "costoRefCRC": 10340.0, "precioVentaUSD": 36.17, "precioVentaCRC": 17000.0, "stockInicial": 0},
  {"codigo": "WHI-008", "nombre": "Chivas 12 años litro", "categoria": "WHISKY/WHISKEY", "costoRefUSD": 22.0, "costoRefCRC": 10340.0, "precioVentaUSD": 36.17, "precioVentaCRC": 17000.0, "stockInicial": 0},
  {"codigo": "WHI-009", "nombre": "Buchanans 12 años litro", "categoria": "WHISKY/WHISKEY", "costoRefUSD": 28.0, "costoRefCRC": 13160.0, "precioVentaUSD": 46.81, "precioVentaCRC": 22000.0, "stockInicial": 0},
  {"codigo": "WHI-010", "nombre": "Buchanan's 18 años botella", "categoria": "WHISKY/WHISKEY", "costoRefUSD": 58.0, "costoRefCRC": 27260.0, "precioVentaUSD": 89.36, "precioVentaCRC": 42000.0, "stockInicial": 0},
  {"codigo": "WHI-011", "nombre": "Jack Daniel's clásico litro", "categoria": "WHISKY/WHISKEY", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 34.04, "precioVentaCRC": 16000.0, "stockInicial": 0},
  {"codigo": "WHI-012", "nombre": "Jack Daniels Apple litro", "categoria": "WHISKY/WHISKEY", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 36.17, "precioVentaCRC": 17000.0, "stockInicial": 0},
  {"codigo": "WHI-013", "nombre": "Jack Daniel's Honey Litro", "categoria": "WHISKY/WHISKEY", "costoRefUSD": 16.0, "costoRefCRC": 7520.0, "precioVentaUSD": 36.17, "precioVentaCRC": 17000.0, "stockInicial": 1},
  {"codigo": "WHI-014", "nombre": "Fireball litro", "categoria": "WHISKY/WHISKEY", "costoRefUSD": 10.0, "costoRefCRC": 4700.0, "precioVentaUSD": 21.28, "precioVentaCRC": 10000.0, "stockInicial": 0},
  {"codigo": "WHI-015", "nombre": "Ballantines litro", "categoria": "WHISKY/WHISKEY", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 21.28, "precioVentaCRC": 10000.0, "stockInicial": 0},
  {"codigo": "BOU-001", "nombre": "Jim Beam white litro", "categoria": "BOURBON", "costoRefUSD": 10.0, "costoRefCRC": 4700.0, "precioVentaUSD": 21.28, "precioVentaCRC": 10000.0, "stockInicial": 0},
  {"codigo": "BOU-002", "nombre": "Jim Beam Miel", "categoria": "BOURBON", "costoRefUSD": 10.0, "costoRefCRC": 4700.0, "precioVentaUSD": 21.28, "precioVentaCRC": 10000.0, "stockInicial": 0},
  {"codigo": "RON-001", "nombre": "Flor de caña 4 años litro", "categoria": "RON", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 14.89, "precioVentaCRC": 7000.0, "stockInicial": 0},
  {"codigo": "RON-002", "nombre": "Flor de caña 7 años litro", "categoria": "RON", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 19.15, "precioVentaCRC": 9000.0, "stockInicial": 0},
  {"codigo": "RON-003", "nombre": "Flor de caña 12 años litro", "categoria": "RON", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 38.3, "precioVentaCRC": 18000.0, "stockInicial": 0},
  {"codigo": "RON-004", "nombre": "Flor de caña 18 años botella", "categoria": "RON", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 38.3, "precioVentaCRC": 18000.0, "stockInicial": 0},
  {"codigo": "RON-005", "nombre": "Flor de caña 18 años litro", "categoria": "RON", "costoRefUSD": 24.0, "costoRefCRC": 11280.0, "precioVentaUSD": 44.68, "precioVentaCRC": 21000.0, "stockInicial": 0},
  {"codigo": "RON-006", "nombre": "Flor de caña Espresso botella", "categoria": "RON", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 19.15, "precioVentaCRC": 9000.0, "stockInicial": 0},
  {"codigo": "RON-007", "nombre": "Flor de caña Cristalino botella", "categoria": "RON", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 25.53, "precioVentaCRC": 12000.0, "stockInicial": 0},
  {"codigo": "RON-008", "nombre": "Bacardi Carta Oro Litro", "categoria": "RON", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 17.02, "precioVentaCRC": 8000.0, "stockInicial": 0},
  {"codigo": "RON-009", "nombre": "Bacardi claro litro", "categoria": "RON", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 17.02, "precioVentaCRC": 8000.0, "stockInicial": 0},
  {"codigo": "RON-010", "nombre": "Centenario 7 años litro", "categoria": "RON", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 21.28, "precioVentaCRC": 10000.0, "stockInicial": 0},
  {"codigo": "RON-011", "nombre": "Centenario 12 años botella", "categoria": "RON", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 27.66, "precioVentaCRC": 13000.0, "stockInicial": 0},
  {"codigo": "RON-012", "nombre": "Zacapa Ámbar botella", "categoria": "RON", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 42.55, "precioVentaCRC": 20000.0, "stockInicial": 0},
  {"codigo": "RON-013", "nombre": "Zacapa edición negra botella", "categoria": "RON", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 55.32, "precioVentaCRC": 26000.0, "stockInicial": 0},
  {"codigo": "RON-014", "nombre": "Zacapa 23 litro", "categoria": "RON", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 59.57, "precioVentaCRC": 28000.0, "stockInicial": 0},
  {"codigo": "RON-015", "nombre": "Malibú litro", "categoria": "RON", "costoRefUSD": 9.0, "costoRefCRC": 4230.0, "precioVentaUSD": 21.28, "precioVentaCRC": 10000.0, "stockInicial": 0},
  {"codigo": "RON-016", "nombre": "Captain Morgan Private Stock litro", "categoria": "RON", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 29.79, "precioVentaCRC": 14000.0, "stockInicial": 0},
  {"codigo": "RON-017", "nombre": "Abuelo", "categoria": "RON", "costoRefUSD": 8.0, "costoRefCRC": 3760.0, "precioVentaUSD": 17.02, "precioVentaCRC": 8000.0, "stockInicial": 2},
  {"codigo": "VOD-001", "nombre": "Absolut vodka litro", "categoria": "VODKA", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 21.28, "precioVentaCRC": 10000.0, "stockInicial": 0},
  {"codigo": "VOD-002", "nombre": "Absolut Mandrin litro", "categoria": "VODKA", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 21.28, "precioVentaCRC": 10000.0, "stockInicial": 0},
  {"codigo": "VOD-003", "nombre": "Absolut Peach litro", "categoria": "VODKA", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 21.28, "precioVentaCRC": 10000.0, "stockInicial": 0},
  {"codigo": "VOD-004", "nombre": "Hpnotiq litro", "categoria": "VODKA", "costoRefUSD": 16.0, "costoRefCRC": 7520.0, "precioVentaUSD": 29.79, "precioVentaCRC": 14000.0, "stockInicial": 0},
  {"codigo": "TEQ-001", "nombre": "Jarana claro litro", "categoria": "TEQUILA", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 14.89, "precioVentaCRC": 7000.0, "stockInicial": 0},
  {"codigo": "TEQ-002", "nombre": "1800 silver botella", "categoria": "TEQUILA", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 38.3, "precioVentaCRC": 18000.0, "stockInicial": 0},
  {"codigo": "TEQ-003", "nombre": "1800 reposado botella", "categoria": "TEQUILA", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 42.55, "precioVentaCRC": 20000.0, "stockInicial": 0},
  {"codigo": "TEQ-004", "nombre": "1800 añejo botella", "categoria": "TEQUILA", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 51.06, "precioVentaCRC": 24000.0, "stockInicial": 0},
  {"codigo": "TEQ-005", "nombre": "1800 Cristalino botella", "categoria": "TEQUILA", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 59.57, "precioVentaCRC": 28000.0, "stockInicial": 0},
  {"codigo": "TEQ-006", "nombre": "Jose Cuervo reposado litro", "categoria": "TEQUILA", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 23.4, "precioVentaCRC": 11000.0, "stockInicial": 0},
  {"codigo": "TEQ-007", "nombre": "Jose Cuervo Silver litro", "categoria": "TEQUILA", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 23.4, "precioVentaCRC": 11000.0, "stockInicial": 0},
  {"codigo": "TEQ-008", "nombre": "Don Julio Claro botella", "categoria": "TEQUILA", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 63.83, "precioVentaCRC": 30000.0, "stockInicial": 0},
  {"codigo": "TEQ-009", "nombre": "Don Julio Reposado botella", "categoria": "TEQUILA", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 74.47, "precioVentaCRC": 35000.0, "stockInicial": 0},
  {"codigo": "LIC-001", "nombre": "Anís del mono botella", "categoria": "LICORES VARIOS", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 14.89, "precioVentaCRC": 7000.0, "stockInicial": 0},
  {"codigo": "LIC-002", "nombre": "Kahlua litro", "categoria": "LICORES VARIOS", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 21.28, "precioVentaCRC": 10000.0, "stockInicial": 0},
  {"codigo": "LIC-003", "nombre": "Jager litro", "categoria": "LICORES VARIOS", "costoRefUSD": 11.0, "costoRefCRC": 5170.0, "precioVentaUSD": 21.28, "precioVentaCRC": 10000.0, "stockInicial": 0},
  {"codigo": "LIC-004", "nombre": "Jager Naranja litro", "categoria": "LICORES VARIOS", "costoRefUSD": 15.0, "costoRefCRC": 7050.0, "precioVentaUSD": 27.66, "precioVentaCRC": 13000.0, "stockInicial": 0},
  {"codigo": "LIC-005", "nombre": "Jager Manifest litro", "categoria": "LICORES VARIOS", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 42.55, "precioVentaCRC": 20000.0, "stockInicial": 0},
  {"codigo": "LIC-006", "nombre": "Aguardiente tapa roja litro", "categoria": "LICORES VARIOS", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 12.77, "precioVentaCRC": 6000.0, "stockInicial": 0},
  {"codigo": "LIC-007", "nombre": "Aguardiente tapa azul litro", "categoria": "LICORES VARIOS", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 12.77, "precioVentaCRC": 6000.0, "stockInicial": 0},
  {"codigo": "LIC-008", "nombre": "Frangelico litro", "categoria": "LICORES VARIOS", "costoRefUSD": 12.0, "costoRefCRC": 5640.0, "precioVentaUSD": 29.79, "precioVentaCRC": 14000.0, "stockInicial": 0},
  {"codigo": "LIC-009", "nombre": "Campary", "categoria": "LICORES VARIOS", "costoRefUSD": 12.0, "costoRefCRC": 5640.0, "precioVentaUSD": 21.28, "precioVentaCRC": 10000.0, "stockInicial": 0},
  {"codigo": "CRE-001", "nombre": "Tequila Rose litro", "categoria": "CREMAS", "costoRefUSD": 17.0, "costoRefCRC": 7990.0, "precioVentaUSD": 31.91, "precioVentaCRC": 15000.0, "stockInicial": 0},
  {"codigo": "CRE-002", "nombre": "Baileys litro", "categoria": "CREMAS", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 27.66, "precioVentaCRC": 13000.0, "stockInicial": 0},
  {"codigo": "CRE-003", "nombre": "Baja Rosa litro", "categoria": "CREMAS", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 21.28, "precioVentaCRC": 10000.0, "stockInicial": 0},
  {"codigo": "CRE-004", "nombre": "Sheridans", "categoria": "CREMAS", "costoRefUSD": 22.0, "costoRefCRC": 10340.0, "precioVentaUSD": 40.43, "precioVentaCRC": 19000.0, "stockInicial": 0},
  {"codigo": "CRE-005", "nombre": "Amarula litro", "categoria": "CREMAS", "costoRefUSD": 9.5, "costoRefCRC": 4465.0, "precioVentaUSD": 23.4, "precioVentaCRC": 11000.0, "stockInicial": 0},
  {"codigo": "CRE-006", "nombre": "Cerveza", "categoria": "CREMAS", "costoRefUSD": 0, "costoRefCRC": 0, "precioVentaUSD": 3.19, "precioVentaCRC": 1500.0, "stockInicial": 0},
  {"codigo": "CRE-007", "nombre": "Redbull", "categoria": "CREMAS", "costoRefUSD": 28.0, "costoRefCRC": 13160.0, "precioVentaUSD": 51.06, "precioVentaCRC": 24000.0, "stockInicial": 0},
  {"codigo": "CRE-008", "nombre": "Smirnoff", "categoria": "CREMAS", "costoRefUSD": 23.0, "costoRefCRC": 10810.0, "precioVentaUSD": 42.55, "precioVentaCRC": 20000.0, "stockInicial": 0},
  {"codigo": "CRE-009", "nombre": "Sol", "categoria": "CREMAS", "costoRefUSD": 20.0, "costoRefCRC": 9400.0, "precioVentaUSD": 36.17, "precioVentaCRC": 17000.0, "stockInicial": 0},
  {"codigo": "CRE-010", "nombre": "Cuba", "categoria": "CREMAS", "costoRefUSD": 20.0, "costoRefCRC": 9400.0, "precioVentaUSD": 36.17, "precioVentaCRC": 17000.0, "stockInicial": 0},
  {"codigo": "CRE-011", "nombre": "Corona", "categoria": "CREMAS", "costoRefUSD": 21.0, "costoRefCRC": 9870.0, "precioVentaUSD": 38.3, "precioVentaCRC": 18000.0, "stockInicial": 0},
  {"codigo": "VIN-001", "nombre": "Vino", "categoria": "VINO", "costoRefUSD": 8.0, "costoRefCRC": 3760.0, "precioVentaUSD": 13.83, "precioVentaCRC": 6500.0, "stockInicial": 0}
];

const SEED_PURCHASES = [
  { codigo: "CRE-005", cantidad: 2, costoUnitarioUSD: 9.5, tipoCambio: 470, costoUnitarioCRC: 9.5 * 470, fecha: "2026-08-10", id: "CMP-001", vendedor: "Carlos" },
  { codigo: "VIN-001", cantidad: 3, costoUnitarioUSD: 8.0, tipoCambio: 470, costoUnitarioCRC: 8.0 * 470, fecha: "2026-08-10", id: "CMP-002", vendedor: "Daniel" },
  { codigo: "LIC-004", cantidad: 2, costoUnitarioUSD: 15.0, tipoCambio: 470, costoUnitarioCRC: 15.0 * 470, fecha: "2026-08-10", id: "CMP-003", vendedor: "Carlos" },
  { codigo: "RON-005", cantidad: 1, costoUnitarioUSD: 24.0, tipoCambio: 470, costoUnitarioCRC: 24.0 * 470, fecha: "2026-08-10", id: "CMP-004", vendedor: "Daniel" },
  { codigo: "BOU-002", cantidad: 1, costoUnitarioUSD: 10.0, tipoCambio: 470, costoUnitarioCRC: 10.0 * 470, fecha: "2026-08-10", id: "CMP-005", vendedor: "Carlos" },
  { codigo: "CRE-001", cantidad: 2, costoUnitarioUSD: 17.0, tipoCambio: 470, costoUnitarioCRC: 17.0 * 470, fecha: "2026-08-10", id: "CMP-006", vendedor: "Daniel" },
  { codigo: "RON-015", cantidad: 1, costoUnitarioUSD: 11.0, tipoCambio: 470, costoUnitarioCRC: 11.0 * 470, fecha: "2026-08-10", id: "CMP-007", vendedor: "Carlos" },
  { codigo: "WHI-001", cantidad: 1, costoUnitarioUSD: 12.0, tipoCambio: 470, costoUnitarioCRC: 12.0 * 470, fecha: "2026-08-10", id: "CMP-008", vendedor: "Carlos" }
];

// --- Estado Global ---
let state = {
  productos: {}, // Mapa por código { "WHI-001": {...} }
  compras: [],
  ventas: [],
  movimientosDinero: [],
  colaSincronizacion: [], // Cola persistente para operaciones sin internet
  carrito: [],
  config: {
    sheetsUrl: "",
    tipoCambio: 520,
    nombreNegocio: "DC El Destape",
    telefonoNegocio: "+506 8992-7936"
  },
  vendedorActual: "Carlos", // "Carlos" | "Daniel"
  vistaVendedor: "Carlos",   // "Carlos" | "Daniel" | "Consolidado"
  vendedores: ["Carlos", "Daniel"],
  filtroFinanzas: "todos",  // "todos" | "Empresa" | "Carlos" | "Daniel"
  categoriaSeleccionada: "Todas",
  filtroEstadoStock: "todos", // "todos" | "constock" | "agotados"
  ordenActual: "az",
  metodoPagoSeleccionado: "Efectivo",
  escanerActivo: null,
  modoEscaner: "buscar",
  ultimaVentaCompletada: null
};

// ==========================================================================
// INICIALIZACIÓN
// ==========================================================================
document.addEventListener("DOMContentLoaded", () => {
  cargarEstadoLocal();
  comprobarLoginVendedor();
  aplicarConfiguracionUI();
  inicializarIconos();
  actualizarIndicadorOffline();
  renderizarTodo();

  // Registrar Service Worker PWA para soporte Offline-First
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js').then((reg) => {
      reg.update();
    }).catch(() => {});
  }

  // Detectores de conexión en tiempo real
  window.addEventListener('online', () => {
    actualizarIndicadorOffline();
    if (state.colaSincronizacion && state.colaSincronizacion.length > 0) {
      mostrarToast(`Conexión recuperada. Sincronizando ${state.colaSincronizacion.length} cambios pendientes... 🌐`, "info");
      procesarColaSincronizacion(false);
    } else if (state.config.sheetsUrl) {
      sincronizarConSheets(false);
    }
  });

  window.addEventListener('offline', () => {
    actualizarIndicadorOffline();
    mostrarToast("Modo Offline activo (Sin internet). Todo se guarda en tu teléfono 💾", "info");
  });

  // Procesar cola pendiente o sincronizar al abrir la app
  if (state.config.sheetsUrl && navigator.onLine) {
    if (state.colaSincronizacion && state.colaSincronizacion.length > 0) {
      procesarColaSincronizacion(false);
    } else {
      sincronizarConSheets(false);
    }
  }
});

function inicializarIconos() {
  if (window.lucide) window.lucide.createIcons();
}

// ==========================================================================
// GESTIÓN DE VENDEDORES (CARLOS Y DANIEL)
// ==========================================================================
function comprobarLoginVendedor() {
  const saved = localStorage.getItem("inv_vendedor_actual");
  if (!saved) {
    abrirModalSeleccionVendedor(true);
  } else {
    state.vendedorActual = saved;
    state.vistaVendedor = localStorage.getItem("inv_vista_vendedor") || saved;
    actualizarUIVendedor();
  }
}

function abrirModalSeleccionVendedor(forzado = false) {
  const modal = document.getElementById("modalLoginVendedor");
  const btnCerrar = document.getElementById("btnCerrarLoginVendedor");
  if (btnCerrar) {
    if (forzado) btnCerrar.classList.add("hidden");
    else btnCerrar.classList.remove("hidden");
  }
  if (modal) {
    modal.classList.remove("hidden");
    modal.classList.add("flex");
  }
  inicializarIconos();
}

function cerrarModalLoginVendedor() {
  const modal = document.getElementById("modalLoginVendedor");
  if (modal) {
    modal.classList.add("hidden");
    modal.classList.remove("flex");
  }
}

function seleccionarVendedorLogin(vendedor) {
  if (vendedor === "Consolidado") {
    state.vistaVendedor = "Consolidado";
  } else {
    state.vendedorActual = vendedor;
    state.vistaVendedor = vendedor;
    localStorage.setItem("inv_vendedor_actual", vendedor);
  }
  localStorage.setItem("inv_vista_vendedor", state.vistaVendedor);

  actualizarUIVendedor();
  cerrarModalLoginVendedor();
  renderizarTodo();
  mostrarToast(`Perfil activo: ${vendedor === "Consolidado" ? "Consolidado (Total)" : vendedor} 👤`, "success");
}

function cambiarVistaVendedor(vista) {
  state.vistaVendedor = vista;
  localStorage.setItem("inv_vista_vendedor", vista);
  actualizarUIVendedor();
  renderizarInventario();
  renderizarDashboard();
}

function actualizarUIVendedor() {
  const labelHeader = document.getElementById("headerVendedorNombre");
  if (labelHeader) labelHeader.textContent = state.vistaVendedor === "Consolidado" ? "Consolidado" : state.vendedorActual;

  const posVendedor = document.getElementById("posVendedorNombre");
  if (posVendedor) posVendedor.textContent = state.vendedorActual;

  const compraVendSelect = document.getElementById("compraVendedor");
  if (compraVendSelect) compraVendSelect.value = state.vendedorActual;

  const labelInv = document.getElementById("inventarioVendedorLabel");
  if (labelInv) labelInv.textContent = state.vistaVendedor === "Consolidado" ? "Consolidado (Total)" : `Vendedor ${state.vistaVendedor}`;

  // Tabs styling
  ["Carlos", "Daniel", "Consolidado"].forEach(v => {
    const btn = document.getElementById(`tabVendedor-${v}`);
    if (btn) {
      if (state.vistaVendedor === v) {
        btn.className = "py-2 rounded-xl bg-indigo-600 text-white shadow-md flex items-center justify-center gap-1 active:scale-95 transition-all";
      } else {
        btn.className = "py-2 rounded-xl bg-transparent text-slate-400 hover:text-white flex items-center justify-center gap-1 active:scale-95 transition-all";
      }
    }
  });
}

// ==========================================================================
// PERSISTENCIA LOCAL Y COSTOS DERIVADOS
// ==========================================================================
function cargarEstadoLocal() {
  const cfg = localStorage.getItem("inv_config_v2");
  if (cfg) state.config = { ...state.config, ...JSON.parse(cfg) };

  const prods = localStorage.getItem("inv_productos_v2");
  if (prods) {
    state.productos = JSON.parse(prods);
  } else {
    // Carga inicial semilla
    const mapa = {};
    SEED_PRODUCTS.forEach(p => { mapa[p.codigo] = p; });
    state.productos = mapa;
    state.compras = [...SEED_PURCHASES];
    guardarProductosLocal();
    guardarComprasLocal();
  }

  const comps = localStorage.getItem("inv_compras_v2");
  if (comps) state.compras = JSON.parse(comps);

  const vts = localStorage.getItem("inv_ventas_v2");
  if (vts) state.ventas = JSON.parse(vts);

  const fin = localStorage.getItem("inv_finanzas_v2");
  if (fin) state.movimientosDinero = JSON.parse(fin);

  const cola = localStorage.getItem("inv_sync_queue_v2");
  if (cola) {
    try {
      state.colaSincronizacion = JSON.parse(cola);
    } catch(e) {
      state.colaSincronizacion = [];
    }
  }
}

function guardarProductosLocal() {
  localStorage.setItem("inv_productos_v2", JSON.stringify(state.productos));
}
function guardarComprasLocal() {
  localStorage.setItem("inv_compras_v2", JSON.stringify(state.compras));
}
function guardarVentasLocal() {
  localStorage.setItem("inv_ventas_v2", JSON.stringify(state.ventas));
}
function guardarFinanzasLocal() {
  localStorage.setItem("inv_finanzas_v2", JSON.stringify(state.movimientosDinero));
}
function guardarColaLocal() {
  localStorage.setItem("inv_sync_queue_v2", JSON.stringify(state.colaSincronizacion));
  actualizarIndicadorOffline();
}
function guardarConfiguracionLocal() {
  localStorage.setItem("inv_config_v2", JSON.stringify(state.config));
}

// --- Cálculos de Stock Separado por Vendedor y Consolidado ---
function calcularStockDetalladoPorCodigo() {
  const detalle = {};
  Object.values(state.productos).forEach(p => {
    detalle[p.codigo] = {
      Carlos: 0,
      Daniel: 0,
      total: 0
    };
  });

  // Sumar compras por vendedor
  state.compras.forEach(c => {
    const cod = c.codigo;
    if (!detalle[cod]) detalle[cod] = { Carlos: 0, Daniel: 0, total: 0 };
    const cant = Number(c.cantidad || 0);
    const vend = String(c.vendedor || "Carlos").trim();
    if (vend === "Daniel") {
      detalle[cod].Daniel += cant;
    } else {
      detalle[cod].Carlos += cant;
    }
    detalle[cod].total += cant;
  });

  // Restar ventas por vendedor
  state.ventas.forEach(v => {
    const vend = String(v.vendedor || "Carlos").trim();
    const items = v.items && Array.isArray(v.items) ? v.items : (v.codigo ? [{ codigo: v.codigo, cantidad: v.cantidad }] : []);
    items.forEach(i => {
      const cod = i.codigo;
      if (!detalle[cod]) detalle[cod] = { Carlos: 0, Daniel: 0, total: 0 };
      const cant = Number(i.cantidad || 0);
      if (vend === "Daniel") {
        detalle[cod].Daniel -= cant;
      } else {
        detalle[cod].Carlos -= cant;
      }
      detalle[cod].total -= cant;
    });
  });

  return detalle;
}

function calcularStockPorCodigo(vista = state.vistaVendedor) {
  const det = calcularStockDetalladoPorCodigo();
  const mapa = {};
  Object.keys(det).forEach(cod => {
    if (vista === "Carlos") {
      mapa[cod] = det[cod].Carlos;
    } else if (vista === "Daniel") {
      mapa[cod] = det[cod].Daniel;
    } else {
      mapa[cod] = det[cod].total;
    }
  });
  return mapa;
}

function calcularCostosPorCodigo(vista = state.vistaVendedor) {
  const acc = {};
  state.compras.forEach(p => {
    const vend = String(p.vendedor || "Carlos").trim();
    if (vista === "Carlos" && vend !== "Carlos") return;
    if (vista === "Daniel" && vend !== "Daniel") return;

    const cant = Number(p.cantidad || 0);
    const cu = Number(p.costoUnitarioUSD || 0);
    const cc = Number(p.costoUnitarioCRC || (cu * (p.tipoCambio || state.config.tipoCambio)));
    if (!acc[p.codigo]) acc[p.codigo] = { cant: 0, totalUSD: 0, totalCRC: 0 };
    acc[p.codigo].cant += cant;
    acc[p.codigo].totalUSD += cant * cu;
    acc[p.codigo].totalCRC += cant * cc;
  });

  const out = {};
  Object.keys(acc).forEach(c => {
    out[c] = {
      usd: acc[c].cant > 0 ? acc[c].totalUSD / acc[c].cant : 0,
      crc: acc[c].cant > 0 ? acc[c].totalCRC / acc[c].cant : 0,
      fuente: "compras"
    };
  });

  Object.values(state.productos).forEach(p => {
    if (!out[p.codigo] && (Number(p.costoRefUSD) > 0 || Number(p.costoRefCRC) > 0)) {
      out[p.codigo] = {
        usd: Number(p.costoRefUSD) || 0,
        crc: Number(p.costoRefCRC) || 0,
        fuente: "referencia"
      };
    }
  });
  return out;
}

// ==========================================================================
// NAVEGACIÓN Y VISTAS
// ==========================================================================
function cambiarVista(vista) {
  const vistas = ["dashboard", "inventario", "ventas", "compras", "finanzas", "configuracion"];
  
  vistas.forEach(v => {
    const el = document.getElementById("view" + capitalizar(v));
    const nav = document.getElementById("navTab-" + v);
    if (el) {
      if (v === vista) {
        el.classList.remove("hidden");
        el.classList.add("active-view");
        if (nav) nav.classList.add("active");
      } else {
        el.classList.add("hidden");
        el.classList.remove("active-view");
        if (nav) nav.classList.remove("active");
      }
    }
  });

  if (vista === "dashboard") renderizarDashboard();
  if (vista === "inventario") renderizarInventario();
  if (vista === "compras") poblarSelectCompras();
  if (vista === "ventas") renderizarCarrito();
  if (vista === "finanzas") renderizarFinanzas();

  inicializarIconos();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function capitalizar(str) {
  return str.charAt(0).toUpperCase() + str.slice(1);
}

// ==========================================================================
// RENDERIZADO GENERAL
// ==========================================================================
function renderizarTodo() {
  aplicarConfiguracionUI();
  renderizarDashboard();
  renderizarInventario();
  poblarSelectCompras();
  renderizarHistorialCompras();
  renderizarCarrito();
  renderizarFinanzas();
  inicializarIconos();
}

function aplicarConfiguracionUI() {
  document.getElementById("appHeaderTitle").textContent = state.config.nombreNegocio || "Libro de Inventario";
  document.getElementById("headerExchangeRate").textContent = state.config.tipoCambio;

  const inputUrl = document.getElementById("sheetsApiUrl");
  if (inputUrl) inputUrl.value = state.config.sheetsUrl || "";

  const inputName = document.getElementById("businessNameInput");
  if (inputName) inputName.value = state.config.nombreNegocio || "Libro de Inventario";

  const inputTC = document.getElementById("exchangeRateInput");
  if (inputTC) inputTC.value = state.config.tipoCambio || 520;

  const inputPhone = document.getElementById("businessPhoneInput");
  if (inputPhone) inputPhone.value = state.config.telefonoNegocio || "";

  const inputCompraTC = document.getElementById("compraTipoCambio");
  if (inputCompraTC) inputCompraTC.value = state.config.tipoCambio || 520;

  const inputCompraFecha = document.getElementById("compraFecha");
  if (inputCompraFecha && !inputCompraFecha.value) inputCompraFecha.value = todayStr();

  actualizarBadgeConexion();
}

function actualizarBadgeConexion() {
  const badge = document.getElementById("connectionBadge");
  if (!badge) return;

  if (state.config.sheetsUrl && state.config.sheetsUrl.startsWith("https://script.google.com")) {
    badge.innerHTML = `<span class="w-1.5 h-1.5 rounded-full bg-emerald-400"></span> Conectado a Sheets`;
    badge.className = "inline-flex items-center gap-1 text-[11px] font-medium text-emerald-400";
  } else {
    badge.innerHTML = `<span class="w-1.5 h-1.5 rounded-full bg-amber-400 animate-pulse"></span> Modo Local / Demo`;
    badge.className = "inline-flex items-center gap-1 text-[11px] font-medium text-amber-400";
  }
}

// ==========================================================================
// 1. DASHBOARD
// ==========================================================================
function renderizarDashboard() {
  const stockMap = calcularStockPorCodigo();
  const costMap = calcularCostosPorCodigo();

  let costoUSD = 0, costoCRC = 0;
  let ventaUSD = 0, ventaCRC = 0;
  let totalUnidades = 0;
  let prodsConStock = 0;
  let sinExistencia = [];

  Object.values(state.productos).forEach(p => {
    const st = stockMap[p.codigo] || 0;
    if (st > 0) {
      prodsConStock++;
      totalUnidades += st;
      const c = costMap[p.codigo] || { usd: 0, crc: 0 };
      costoUSD += st * c.usd;
      costoCRC += st * c.crc;
      ventaUSD += st * Number(p.precioVentaUSD || 0);
      ventaCRC += st * Number(p.precioVentaCRC || 0);
    } else {
      sinExistencia.push(p);
    }
  });

  const margenUSD = ventaUSD - costoUSD;
  const margenCRC = ventaCRC - costoCRC;

  document.getElementById("dashCostoUSD").textContent = fmtUSD(costoUSD);
  document.getElementById("dashCostoCRC").textContent = fmtCRC(costoCRC);
  document.getElementById("dashVentaUSD").textContent = fmtUSD(ventaUSD);
  document.getElementById("dashVentaCRC").textContent = fmtCRC(ventaCRC);
  document.getElementById("dashMargenUSD").textContent = fmtUSD(margenUSD);
  document.getElementById("dashMargenCRC").textContent = fmtCRC(margenCRC);

  document.getElementById("dashTotalUnidades").textContent = fmtNum(totalUnidades);
  document.getElementById("dashProdsConStock").textContent = `${prodsConStock} de ${Object.keys(state.productos).length}`;

  // Banner sin existencias
  const alertCont = document.getElementById("stockAlertsContainer");
  const alertList = document.getElementById("stockAlertsList");
  const lowStockCount = document.getElementById("lowStockCount");

  if (sinExistencia.length > 0) {
    alertCont.classList.remove("hidden");
    lowStockCount.textContent = sinExistencia.length;
    alertList.textContent = sinExistencia.map(p => p.codigo).join(", ");
  } else {
    alertCont.classList.add("hidden");
  }

  // Últimas ventas
  const recentCont = document.getElementById("dashRecentSales");
  if (state.ventas.length === 0) {
    recentCont.innerHTML = `<div class="text-center py-5 text-slate-500 text-xs">No hay ventas registradas aún.</div>`;
  } else {
    const ultimas = state.ventas.slice(0, 6);
    recentCont.innerHTML = ultimas.map(v => {
      const fecha = v.fecha ? new Date(v.fecha).toLocaleDateString([], { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' }) : "";
      const vend = v.vendedor || "Carlos";
      const vendColor = vend === "Daniel" ? "text-violet-400 bg-violet-950/60 border-violet-500/30" : "text-blue-400 bg-blue-950/60 border-blue-500/30";

      return `
        <div class="py-2 flex items-center justify-between border-b border-slate-800/60 last:border-0">
          <div>
            <div class="flex items-center gap-1.5 mb-0.5">
              <span class="text-[9px] font-bold px-1.5 py-0.2 rounded border ${vendColor}">👤 ${vend}</span>
              <span class="text-xs font-bold text-white truncate max-w-[140px]">${v.cliente || "Venta"}</span>
            </div>
            <div class="text-[11px] text-slate-400 font-mono">${fecha} • ${v.items ? v.items.length : 1} prod(s) <span class="text-[10px] text-slate-500">(${v.metodoPago || "Efectivo"})</span></div>
          </div>
          <div class="text-right font-mono">
            <div class="text-xs font-black text-emerald-400">${fmtCRC(v.totalCRC || 0)}</div>
            <div class="text-[10px] text-slate-400">${fmtUSD(v.totalUSD || 0)}</div>
          </div>
        </div>
      `;
    }).join("");
  }
}

// ==========================================================================
// 2. INVENTARIO (PRODUCTOS)
// ==========================================================================
function filtrarEstadoStock(estado) {
  state.filtroEstadoStock = estado;
  
  const btnTodos = document.getElementById("filterStock-todos");
  const btnCon = document.getElementById("filterStock-constock");
  const btnSin = document.getElementById("filterStock-agotados");

  if (btnTodos && btnCon && btnSin) {
    btnTodos.className = estado === "todos" 
      ? "py-1.5 px-2 rounded-xl bg-indigo-600 text-white text-center border border-indigo-500/40 shadow-sm active:scale-95 transition-all"
      : "py-1.5 px-2 rounded-xl bg-slate-800 text-slate-300 text-center border border-slate-700 shadow-sm active:scale-95 transition-all";
      
    btnCon.className = estado === "constock"
      ? "py-1.5 px-2 rounded-xl bg-emerald-600 text-white text-center border border-emerald-500/40 shadow-sm active:scale-95 transition-all"
      : "py-1.5 px-2 rounded-xl bg-slate-800 text-emerald-400 text-center border border-slate-700 shadow-sm active:scale-95 transition-all";

    btnSin.className = estado === "agotados"
      ? "py-1.5 px-2 rounded-xl bg-rose-600 text-white text-center border border-rose-500/40 shadow-sm active:scale-95 transition-all"
      : "py-1.5 px-2 rounded-xl bg-slate-800 text-rose-400 text-center border border-slate-700 shadow-sm active:scale-95 transition-all";
  }

  renderizarInventario();
}

function renderizarInventario() {
  const contenedor = document.getElementById("productsList");
  const filtroTexto = (document.getElementById("searchInventory").value || "").toLowerCase().trim();
  const stockMap = calcularStockPorCodigo();
  const costMap = calcularCostosPorCodigo();

  const todosProds = Object.values(state.productos);
  const totalCount = todosProds.length;
  const conStockCount = todosProds.filter(p => (stockMap[p.codigo] || 0) > 0).length;
  const agotadosCount = totalCount - conStockCount;

  const elTodos = document.getElementById("countStockTodos");
  const elCon = document.getElementById("countStockCon");
  const elSin = document.getElementById("countStockSin");
  if (elTodos) elTodos.textContent = totalCount;
  if (elCon) elCon.textContent = conStockCount;
  if (elSin) elSin.textContent = agotadosCount;

  let lista = todosProds.filter(p => {
    const st = stockMap[p.codigo] || 0;
    const matchEstado = state.filtroEstadoStock === "todos" ||
      (state.filtroEstadoStock === "constock" && st > 0) ||
      (state.filtroEstadoStock === "agotados" && st <= 0);

    const matchCat = state.categoriaSeleccionada === "Todas" || p.categoria === state.categoriaSeleccionada;
    const matchTxt = !filtroTexto ||
      p.nombre.toLowerCase().includes(filtroTexto) ||
      p.codigo.toLowerCase().includes(filtroTexto) ||
      (p.categoria && p.categoria.toLowerCase().includes(filtroTexto));
    return matchEstado && matchCat && matchTxt;
  });

  // Ordenar
  if (state.ordenActual === "az") {
    lista.sort((a, b) => a.nombre.localeCompare(b.nombre));
  } else if (state.ordenActual === "za") {
    lista.sort((a, b) => b.nombre.localeCompare(a.nombre));
  } else if (state.ordenActual === "stock_asc") {
    lista.sort((a, b) => (stockMap[a.codigo] || 0) - (stockMap[b.codigo] || 0));
  } else if (state.ordenActual === "stock_desc") {
    lista.sort((a, b) => (stockMap[b.codigo] || 0) - (stockMap[a.codigo] || 0));
  }

  const prodCountEl = document.getElementById("prodCount");
  if (prodCountEl) prodCountEl.textContent = lista.length;

  if (lista.length === 0) {
    contenedor.innerHTML = `
      <div class="text-center py-10 text-slate-500 space-y-2">
        <i data-lucide="package-search" class="w-9 h-9 mx-auto text-slate-600"></i>
        <p class="text-xs">No se encontraron licores con ese filtro.</p>
        <button onclick="abrirModalProducto()" class="px-4 py-2 bg-indigo-600 text-white rounded-xl text-xs font-bold mt-2">
          Agregar Licor
        </button>
      </div>
    `;
    inicializarIconos();
    renderizarCategoriasPills();
    return;
  }

  const detailedMap = calcularStockDetalladoPorCodigo();

  contenedor.innerHTML = lista.map(p => {
    const det = detailedMap[p.codigo] || { Carlos: 0, Daniel: 0, total: 0 };
    const stock = stockMap[p.codigo] || 0;
    const cost = costMap[p.codigo] || { usd: 0, crc: 0 };
    const stockMinimo = Number(p.stockMinimo || 2);
    
    let badgeStock = "";
    let borderStyle = "";

    if (stock <= 0) {
      badgeStock = `
        <span class="px-2.5 py-1 rounded-xl bg-rose-950/80 border border-rose-500/50 text-rose-300 text-xs font-black font-mono flex items-center gap-1 shadow-sm">
          <span class="w-2 h-2 rounded-full bg-rose-500 animate-pulse"></span>
          ${state.vistaVendedor === "Consolidado" ? "Sin Stock (0)" : "Agotado (0)"}
        </span>
      `;
      borderStyle = "border-l-4 border-l-rose-500";
    } else if (stock <= stockMinimo) {
      badgeStock = `
        <span class="px-2.5 py-1 rounded-xl bg-amber-950/80 border border-amber-500/50 text-amber-300 text-xs font-black font-mono flex items-center gap-1 shadow-sm">
          <span class="w-2 h-2 rounded-full bg-amber-400"></span>
          Stock: ${stock} (Bajo)
        </span>
      `;
      borderStyle = "border-l-4 border-l-amber-500";
    } else {
      badgeStock = `
        <span class="px-2.5 py-1 rounded-xl bg-emerald-950/80 border border-emerald-500/50 text-emerald-300 text-xs font-black font-mono flex items-center gap-1 shadow-sm">
          <span class="w-2 h-2 rounded-full bg-emerald-400"></span>
          Stock: ${stock}
        </span>
      `;
      borderStyle = "border-l-4 border-l-emerald-500";
    }

    // Desglose por vendedor en vista consolidada o individual
    const breakdownPill = state.vistaVendedor === "Consolidado"
      ? `
        <div class="flex items-center gap-2 text-[10px] font-mono bg-slate-900/90 px-2.5 py-1 rounded-xl border border-slate-750 mt-1">
          <span class="text-blue-400 font-bold flex items-center gap-1">👤 Carlos: <b class="${det.Carlos > 0 ? 'text-emerald-400' : 'text-slate-500'} font-black">${det.Carlos}</b></span>
          <span class="text-slate-600">|</span>
          <span class="text-violet-400 font-bold flex items-center gap-1">👤 Daniel: <b class="${det.Daniel > 0 ? 'text-emerald-400' : 'text-slate-500'} font-black">${det.Daniel}</b></span>
        </div>
      `
      : "";

    const imgFormatted = formatearUrlImagen(p.imagenUrl);
    const imgHtml = imgFormatted
      ? `
        <div class="relative w-24 h-24 rounded-2xl bg-slate-900 border border-slate-700/80 flex items-center justify-center shrink-0 overflow-hidden shadow-inner group cursor-pointer">
          <img src="${imgFormatted}" alt="${p.nombre}" loading="lazy" class="w-full h-full object-cover" onerror="this.classList.add('hidden'); this.nextElementSibling.classList.remove('hidden');" onclick="abrirFotoCompleta('${imgFormatted}', '${p.nombre.replace(/'/g, "\\'")}')">
          <div class="hidden flex flex-col items-center justify-center w-full h-full text-slate-500 text-[10px]" onclick="editarProducto('${p.codigo}')">
            <i data-lucide="wine" class="w-8 h-8 text-slate-600 mb-0.5"></i>
          </div>
          <!-- Fullscreen button -->
          <button onclick="abrirFotoCompleta('${imgFormatted}', '${p.nombre.replace(/'/g, "\\'")}')" class="absolute bottom-1 right-1 w-7 h-7 bg-black/60 hover:bg-indigo-600 rounded-lg flex items-center justify-center opacity-80 transition-all active:scale-90" title="Ver en pantalla completa">
            <i data-lucide="maximize-2" class="w-4 h-4 text-white"></i>
          </button>
        </div>
      `
      : `
        <div class="w-24 h-24 rounded-2xl bg-slate-900 border border-slate-700/80 flex items-center justify-center shrink-0 text-slate-500 shadow-inner cursor-pointer" onclick="editarProducto('${p.codigo}')">
          <i data-lucide="wine" class="w-8 h-8 text-slate-500"></i>
        </div>
      `;

    return `
      <div class="p-3.5 bg-slate-800/90 hover:bg-slate-800 border border-slate-700/80 ${borderStyle} rounded-2xl shadow-lg transition-all space-y-2">
        
        <!-- Header: Code, Category & Stock Badge -->
        <div class="flex items-center justify-between gap-2">
          <div class="flex items-center gap-1.5 min-w-0">
            <span class="text-xs font-mono font-black text-amber-400 bg-slate-900 px-2 py-0.5 rounded-lg border border-slate-750">${p.codigo}</span>
            <span class="text-[11px] font-bold text-indigo-300 uppercase tracking-wide truncate bg-indigo-950/50 px-2 py-0.5 rounded-lg border border-indigo-500/20">${p.categoria || "General"}</span>
          </div>
          ${badgeStock}
        </div>

        <!-- Body: Photo + Name & Breakdown -->
        <div class="flex items-start gap-3 pt-0.5">
          ${imgHtml}
          <div class="flex-1 min-w-0 cursor-pointer" onclick="editarProducto('${p.codigo}')">
            <h4 class="text-sm font-extrabold text-white leading-snug hover:text-indigo-300 transition-colors line-clamp-2">${p.nombre}</h4>
            ${breakdownPill}
          </div>
        </div>

        <!-- Pricing & Actions Row -->
        <div class="flex items-center justify-between pt-1.5 border-t border-slate-700/50 font-mono">
          <div>
            <div class="flex items-baseline gap-2">
              <span class="text-sm font-black text-white">${fmtCRC(p.precioVentaCRC)}</span>
              <span class="text-xs font-bold text-teal-300">${fmtUSD(p.precioVentaUSD)}</span>
            </div>
            <div class="text-[10px] text-slate-400">
              Costo Ref: ${fmtUSD(cost.usd)} ${cost.fuente === 'referencia' ? '<span class="text-amber-500 font-semibold">(ref)</span>' : ''} • (${fmtCRC(cost.crc)})
            </div>
          </div>

          <div class="flex items-center gap-1.5">
            <button onclick="agregarAlCarritoPorCodigo('${p.codigo}')" title="Vender con ${state.vendedorActual}" class="px-3 py-2 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl font-bold text-xs shadow-md shadow-indigo-600/30 flex items-center gap-1 active:scale-95 transition-all">
              <i data-lucide="shopping-cart" class="w-4 h-4"></i>
              <span>Vender</span>
            </button>
            <button onclick="editarProducto('${p.codigo}')" title="Editar" class="p-2 bg-slate-900 text-slate-300 hover:text-white rounded-xl border border-slate-700 active:scale-95 transition-all">
              <i data-lucide="edit-3" class="w-4 h-4"></i>
            </button>
          </div>
        </div>

      </div>
    `;
  }).join("") + '<div class="h-16"></div>';

  renderizarCategoriasPills();
  inicializarIconos();
}

function renderizarCategoriasPills() {
  const categorias = ["Todas", ...new Set(Object.values(state.productos).map(p => p.categoria || "General"))];
  const pillsCont = document.getElementById("categoryPills");
  pillsCont.innerHTML = categorias.map(c => `
    <button onclick="filtrarCategoria('${c}')" class="cat-pill ${c === state.categoriaSeleccionada ? 'active' : ''} px-3 py-1.5 rounded-full shrink-0">
      ${c}
    </button>
  `).join("");
}

function filtrarCategoria(cat) {
  state.categoriaSeleccionada = cat;
  renderizarInventario();
}

function filtrarInventario() {
  renderizarInventario();
}

function ordenarProductos() {
  const ordenes = ["az", "za", "stock_asc", "stock_desc"];
  const labels = { az: "A-Z", za: "Z-A", stock_asc: "Menor Stock", stock_desc: "Mayor Stock" };
  let nextIdx = (ordenes.indexOf(state.ordenActual) + 1) % ordenes.length;
  state.ordenActual = ordenes[nextIdx];
  document.getElementById("sortLabel").textContent = labels[state.ordenActual];
  renderizarInventario();
}

// ==========================================================================
// FORMATEADOR DE IMÁGENES / GOOGLE DRIVE
// ==========================================================================

// --------------------------------------------------------------------------
// LIGHTBOX: Foto en pantalla completa (para mostrar a clientes)
// --------------------------------------------------------------------------
function abrirFotoCompleta(url, nombre) {
  const modal = document.getElementById("modalFotoCompleta");
  const img = document.getElementById("fotoCompletaImg");
  const nombreEl = document.getElementById("fotoCompletaNombre");

  if (!modal || !img) return;

  img.src = url;
  img.alt = nombre || "Licor";

  if (nombreEl) {
    const p = nombreEl.querySelector("p");
    if (p) p.textContent = nombre || "";
  }

  modal.classList.remove("hidden");
  modal.classList.add("flex");
  inicializarIconos();

  // Bloquear scroll de fondo
  document.body.style.overflow = "hidden";
}

function cerrarFotoCompleta() {
  const modal = document.getElementById("modalFotoCompleta");
  if (!modal) return;
  modal.classList.add("hidden");
  modal.classList.remove("flex");
  document.body.style.overflow = "";
}

// Cerrar con tecla Escape
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") cerrarFotoCompleta();
});

function formatearUrlImagen(urlOrId) {
  if (!urlOrId || typeof urlOrId !== 'string') return '';
  const trimmed = urlOrId.trim();
  if (!trimmed) return '';

  // 1. Extraer ID de enlace de Google Drive
  const matchFileD = trimmed.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (matchFileD && matchFileD[1]) {
    return `https://lh3.googleusercontent.com/d/${matchFileD[1]}`;
  }

  const matchIdParam = trimmed.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (matchIdParam && matchIdParam[1]) {
    return `https://lh3.googleusercontent.com/d/${matchIdParam[1]}`;
  }

  const matchGoogleUserContent = trimmed.match(/googleusercontent\.com\/d\/([a-zA-Z0-9_-]+)/);
  if (matchGoogleUserContent && matchGoogleUserContent[1]) {
    return `https://lh3.googleusercontent.com/d/${matchGoogleUserContent[1]}`;
  }

  // Si pegó directamente el ID alfanumérico largo de Drive
  if (/^[a-zA-Z0-9_-]{25,50}$/.test(trimmed)) {
    return `https://lh3.googleusercontent.com/d/${trimmed}`;
  }

  // 2. Si es una URL directa (http/https/data)
  if (trimmed.startsWith('http://') || trimmed.startsWith('https://') || trimmed.startsWith('data:image/')) {
    return trimmed;
  }

  return trimmed;
}

function actualizarPreviewImagenModal() {
  const input = document.getElementById("prodImagenUrl");
  const preview = document.getElementById("prodImagenPreview");
  const placeholder = document.getElementById("prodImagenPlaceholder");
  const btnLimpiar = document.getElementById("btnLimpiarImagen");

  if (!input || !preview || !placeholder) return;

  const raw = (input.value || "").trim();
  if (!raw) {
    preview.src = "";
    preview.classList.add("hidden");
    placeholder.classList.remove("hidden");
    placeholder.innerHTML = `<i data-lucide="wine" class="w-6 h-6 text-slate-600 mb-0.5"></i><span>Sin foto</span>`;
    if (btnLimpiar) btnLimpiar.classList.add("hidden");
    inicializarIconos();
    return;
  }

  const formattedUrl = formatearUrlImagen(raw);
  preview.src = formattedUrl;
  preview.classList.remove("hidden");
  placeholder.classList.add("hidden");
  if (btnLimpiar) btnLimpiar.classList.remove("hidden");
}

function onImgPreviewError() {
  const preview = document.getElementById("prodImagenPreview");
  const placeholder = document.getElementById("prodImagenPlaceholder");
  if (preview) {
    preview.classList.add("hidden");
  }
  if (placeholder) {
    placeholder.classList.remove("hidden");
    placeholder.innerHTML = `<i data-lucide="alert-circle" class="w-5 h-5 text-amber-500 mb-0.5"></i><span class="text-amber-400 text-[9px]">No carga</span>`;
    inicializarIconos();
  }
}

function limpiarImagenModal() {
  const input = document.getElementById("prodImagenUrl");
  if (input) input.value = "";
  actualizarPreviewImagenModal();
}

// ==========================================================================
// MODAL DE PRODUCTO (DUAL CURRENCY & FOTO)
// ==========================================================================
function abrirModalProducto(producto = null) {
  const modal = document.getElementById("modalProducto");
  const titulo = document.getElementById("modalProductoTitulo");
  const btnEliminar = document.getElementById("btnEliminarProducto");

  if (producto) {
    titulo.innerHTML = `<i data-lucide="edit" class="w-5 h-5 text-amber-400"></i> Editar Licor (${producto.codigo})`;
    document.getElementById("prodCodigo").value = producto.codigo;
    document.getElementById("prodCodigo").disabled = true;
    document.getElementById("prodNombre").value = producto.nombre || "";
    document.getElementById("prodCategoria").value = producto.categoria || "";
    document.getElementById("prodImagenUrl").value = producto.imagenUrl || "";
    document.getElementById("prodPrecioVentaUSD").value = producto.precioVentaUSD || 0;
    document.getElementById("prodPrecioVentaCRC").value = producto.precioVentaCRC || 0;
    document.getElementById("prodCostoRefUSD").value = producto.costoRefUSD || 0;
    document.getElementById("prodCostoRefCRC").value = producto.costoRefCRC || 0;
    document.getElementById("prodStock").value = producto.stockInicial || 0;
    document.getElementById("prodStockMinimo").value = producto.stockMinimo || 2;
    btnEliminar.classList.remove("hidden");
  } else {
    titulo.innerHTML = `<i data-lucide="wine" class="w-5 h-5 text-amber-400"></i> Nuevo Licor`;
    document.getElementById("formProducto").reset();
    document.getElementById("prodCodigo").disabled = false;
    document.getElementById("prodCodigo").value = "LIC-" + Math.floor(100 + Math.random() * 900);
    document.getElementById("prodImagenUrl").value = "";
    document.getElementById("prodStockMinimo").value = 2;
    btnEliminar.classList.add("hidden");
  }

  actualizarPreviewImagenModal();
  modal.classList.remove("hidden");
  modal.classList.add("flex");
  inicializarIconos();
}

function cerrarModalProducto() {
  const modal = document.getElementById("modalProducto");
  modal.classList.add("hidden");
  modal.classList.remove("flex");
}

function editarProducto(codigo) {
  const p = state.productos[codigo];
  if (p) abrirModalProducto(p);
}

function autoConvertirPrecio(origen) {
  const tc = Number(state.config.tipoCambio) || 520;
  if (origen === 'USD') {
    const usd = Number(document.getElementById("prodPrecioVentaUSD").value) || 0;
    document.getElementById("prodPrecioVentaCRC").value = Math.round(usd * tc);
  } else {
    const crc = Number(document.getElementById("prodPrecioVentaCRC").value) || 0;
    document.getElementById("prodPrecioVentaUSD").value = (crc / tc).toFixed(2);
  }
}

function autoConvertirCosto(origen) {
  const tc = Number(state.config.tipoCambio) || 520;
  if (origen === 'USD') {
    const usd = Number(document.getElementById("prodCostoRefUSD").value) || 0;
    document.getElementById("prodCostoRefCRC").value = Math.round(usd * tc);
  } else {
    const crc = Number(document.getElementById("prodCostoRefCRC").value) || 0;
    document.getElementById("prodCostoRefUSD").value = (crc / tc).toFixed(2);
  }
}

async function guardarProductoForm(e) {
  e.preventDefault();
  const codigo = document.getElementById("prodCodigo").value.trim().toUpperCase();
  const nombre = document.getElementById("prodNombre").value.trim();
  const categoria = document.getElementById("prodCategoria").value.trim() || "General";
  const imagenUrl = document.getElementById("prodImagenUrl").value.trim();
  const precioVentaUSD = Number(document.getElementById("prodPrecioVentaUSD").value) || 0;
  const precioVentaCRC = Number(document.getElementById("prodPrecioVentaCRC").value) || 0;
  const costoRefUSD = Number(document.getElementById("prodCostoRefUSD").value) || 0;
  const costoRefCRC = Number(document.getElementById("prodCostoRefCRC").value) || 0;
  const stockInicial = Number(document.getElementById("prodStock").value) || 0;
  const stockMinimo = Number(document.getElementById("prodStockMinimo").value) || 2;

  const prodObj = {
    codigo,
    nombre,
    categoria,
    imagenUrl,
    precioVentaUSD,
    precioVentaCRC,
    costoRefUSD,
    costoRefCRC,
    stockInicial,
    stockMinimo
  };

  const esEdicion = !!state.productos[codigo];
  state.productos[codigo] = prodObj;
  guardarProductosLocal();
  renderizarTodo();
  cerrarModalProducto();
  mostrarToast(esEdicion ? "Producto actualizado localmente" : "Producto agregado localmente", "success");

  // Encolar y sincronizar
  encolarAccionSincronizacion(esEdicion ? "actualizarProducto" : "crearProducto", { producto: prodObj });
}

async function eliminarProductoActual() {
  const codigo = document.getElementById("prodCodigo").value;
  if (!confirm(`¿Eliminar ${codigo}?`)) return;

  delete state.productos[codigo];
  guardarProductosLocal();
  renderizarTodo();
  cerrarModalProducto();
  mostrarToast("Producto eliminado localmente", "info");

  // Encolar y sincronizar
  encolarAccionSincronizacion("eliminarProducto", { codigo });
}

// ==========================================================================
// 3. COMPRAS / ENTRADAS
// ==========================================================================
function poblarSelectCompras() {
  const select = document.getElementById("compraProductoSelect");
  if (!select) return;

  const stockMap = calcularStockPorCodigo();
  const prods = Object.values(state.productos).sort((a, b) => a.nombre.localeCompare(b.nombre));

  select.innerHTML = `<option value="">-- Seleccionar de la lista --</option>` +
    prods.map(p => `
      <option value="${p.codigo}">${p.nombre} [${p.codigo}] (Stock: ${stockMap[p.codigo] || 0})</option>
    `).join("");
}

function seleccionarProductoCompra() {
  const cod = document.getElementById("compraProductoSelect").value;
  const prod = state.productos[cod];
  const tc = Number(document.getElementById("compraTipoCambio").value) || Number(state.config.tipoCambio) || 520;
  if (prod) {
    const costoUSD = Number(prod.costoRefUSD || 0);
    const costoCRC = Number(prod.costoRefCRC || (costoUSD * tc));
    document.getElementById("compraCostoUSD").value = costoUSD;
    const elCRC = document.getElementById("compraCostoCRC");
    if (elCRC) elCRC.value = costoCRC;
  }
  calcularTotalCompra();
}

function autoConvertirCompraCosto(origen) {
  const tc = Number(document.getElementById("compraTipoCambio").value) || Number(state.config.tipoCambio) || 520;
  const elUSD = document.getElementById("compraCostoUSD");
  const elCRC = document.getElementById("compraCostoCRC");
  if (origen === 'USD' && elUSD && elCRC) {
    const usd = Number(elUSD.value) || 0;
    elCRC.value = Math.round(usd * tc);
  } else if (origen === 'CRC' && elUSD && elCRC) {
    const crc = Number(elCRC.value) || 0;
    elUSD.value = (crc / tc).toFixed(2);
  }
  calcularTotalCompra();
}

function calcularTotalCompra() {
  const cant = Number(document.getElementById("compraCantidad").value) || 0;
  const costoUSD = Number(document.getElementById("compraCostoUSD").value) || 0;
  const tc = Number(document.getElementById("compraTipoCambio").value) || Number(state.config.tipoCambio) || 520;
  
  const elCRC = document.getElementById("compraCostoCRC");
  const costoCRC = elCRC && Number(elCRC.value) > 0 ? Number(elCRC.value) : (costoUSD * tc);
  const totalUSD = cant * costoUSD;
  const totalCRC = cant * costoCRC;

  const elCRCDisp = document.getElementById("compraCostoCRCDisplay");
  const elTotUSD = document.getElementById("compraTotalUSDDisplay");
  const elTotCRC = document.getElementById("compraTotalCRCDisplay");

  if (elCRCDisp) elCRCDisp.textContent = fmtCRC(costoCRC);
  if (elTotUSD) elTotUSD.textContent = fmtUSD(totalUSD);
  if (elTotCRC) elTotCRC.textContent = fmtCRC(totalCRC);
}

async function guardarCompra() {
  const codigo = document.getElementById("compraProductoSelect").value;
  const fecha = document.getElementById("compraFecha").value || todayStr();
  const vendedor = (document.getElementById("compraVendedor") ? document.getElementById("compraVendedor").value : state.vendedorActual) || "Carlos";
  const pagadoPor = (document.getElementById("compraFinanciadoPor") ? document.getElementById("compraFinanciadoPor").value : vendedor) || "Carlos";
  const cant = Number(document.getElementById("compraCantidad").value);
  const costoUSD = Number(document.getElementById("compraCostoUSD").value) || 0;
  const elCRC = document.getElementById("compraCostoCRC");
  const tc = Number(document.getElementById("compraTipoCambio").value) || Number(state.config.tipoCambio) || 520;
  const costoCRC = elCRC && Number(elCRC.value) > 0 ? Number(elCRC.value) : (costoUSD * tc);
  const proveedor = document.getElementById("compraProveedor").value.trim();
  const notas = document.getElementById("compraNotas").value.trim();

  if (!codigo || !state.productos[codigo]) {
    mostrarToast("Por favor selecciona un producto de la lista", "error");
    return;
  }
  if (cant <= 0) {
    mostrarToast("La cantidad a ingresar debe ser mayor a 0", "error");
    return;
  }

  const prod = state.productos[codigo];
  const nombre = prod ? prod.nombre : codigo;
  const idCompra = "CMP-" + Date.now().toString().slice(-6);

  const compraObj = {
    id: idCompra,
    codigo,
    nombre,
    fecha,
    vendedor,
    pagadoPor,
    cantidad: cant,
    costoUnitarioUSD: costoUSD,
    tipoCambio: tc,
    costoUnitarioCRC: costoCRC,
    proveedor: proveedor || "Proveedor General",
    notas: notas || "",
    items: [{
      codigo,
      nombre,
      vendedor,
      pagadoPor,
      cantidad: cant,
      costoUnitarioUSD: costoUSD,
      tipoCambio: tc,
      costoUnitarioCRC: costoCRC
    }]
  };

  // 1. Guardar en memoria local y actualizar cálculos
  state.compras.unshift(compraObj);
  guardarComprasLocal();

  // Reset del formulario para la próxima compra
  document.getElementById("compraCantidad").value = 1;
  document.getElementById("compraProveedor").value = "";
  document.getElementById("compraNotas").value = "";
  
  // Re-renderizar inventario, dashboard, compras y finanzas
  renderizarTodo();
  const detallePago = pagadoPor === "Empresa" ? "Caja Empresa" : `Financiada por ${pagadoPor}`;
  mostrarToast(`¡Compra asignada a ${vendedor}! (+${cant} unids • ${detallePago})`, "success");

  // Encolar y sincronizar con Google Sheets
  encolarAccionSincronizacion("registrarCompra", { compra: compraObj });
}

async function eliminarCompra(id) {
  if (!confirm("¿Deseas eliminar este registro de compra?")) return;
  state.compras = state.compras.filter(c => c.id !== id);
  guardarComprasLocal();
  renderizarTodo();
  mostrarToast("Compra eliminada localmente", "info");

  // Encolar y sincronizar
  encolarAccionSincronizacion("eliminarCompra", { id });
}

function renderizarHistorialCompras() {
  const cont = document.getElementById("comprasHistorialList");
  const countEl = document.getElementById("comprasCount");
  if (!cont) return;

  if (countEl) countEl.textContent = state.compras.length;

  if (state.compras.length === 0) {
    cont.innerHTML = `<div class="text-center py-4 text-slate-500">No hay compras registradas.</div>`;
    return;
  }

  const sorted = [...state.compras].sort((a, b) => (a.fecha < b.fecha ? 1 : -1)).slice(0, 20);
  cont.innerHTML = sorted.map(c => {
    const prod = state.productos[c.codigo];
    const nombre = prod ? prod.nombre : (c.nombre || c.codigo);
    const cant = Number(c.cantidad || 0);
    const cUSD = Number(c.costoUnitarioUSD || 0);
    const tc = Number(c.tipoCambio || state.config.tipoCambio || 520);
    const cCRC = Number(c.costoUnitarioCRC || (cUSD * tc));
    const totUSD = cant * cUSD;
    const totCRC = cant * cCRC;
    const vend = c.vendedor || "Carlos";
    const pagador = c.pagadoPor || vend;
    const vendColor = vend === "Daniel" ? "text-violet-400 bg-violet-950/60 border-violet-500/30" : "text-blue-400 bg-blue-950/60 border-blue-500/30";
    const pagoColor = pagador === "Empresa" ? "text-emerald-300 bg-emerald-950/60 border-emerald-500/30" : pagador === "Daniel" ? "text-violet-300 bg-violet-950/60 border-violet-500/30" : "text-blue-300 bg-blue-950/60 border-blue-500/30";

    return `
      <div class="p-2.5 bg-slate-900/80 rounded-xl border border-slate-800 flex justify-between items-center gap-2">
        <div class="min-w-0 flex-1">
          <div class="flex items-center gap-1.5 mb-1 flex-wrap">
            <span class="text-[10px] font-bold px-1.5 py-0.2 rounded border ${vendColor}">📦 Stock: ${vend}</span>
            <span class="text-[10px] font-bold px-1.5 py-0.2 rounded border ${pagoColor}">💳 Pagó: ${pagador}</span>
            <span class="text-[10px] text-slate-400 font-mono">${c.fecha || todayStr()}</span>
          </div>
          <div class="font-bold text-white truncate text-xs">${nombre} <span class="text-emerald-400 font-mono font-black">(+${cant})</span></div>
          <div class="text-[10px] text-slate-500 font-mono">${c.codigo} • TC: ₡${tc}</div>
        </div>
        <div class="text-right font-mono shrink-0 ml-2">
          <div class="font-black text-white text-xs">${fmtUSD(totUSD)}</div>
          <div class="text-[10px] text-slate-400">${fmtCRC(totCRC)}</div>
          <button onclick="eliminarCompra('${c.id}')" class="text-[10px] text-rose-400 hover:text-rose-300">Eliminar</button>
        </div>
      </div>
    `;
  }).join("");
}

// ==========================================================================
// 4. VENTAS (POS TÁCTIL)
// ==========================================================================
function filtrarPosProductos() {
  const txt = (document.getElementById("searchPos").value || "").toLowerCase().trim();
  const dropdown = document.getElementById("posSearchResults");
  const stockMap = calcularStockPorCodigo();

  if (!txt) {
    dropdown.classList.add("hidden");
    return;
  }

  const matches = Object.values(state.productos).filter(p =>
    p.nombre.toLowerCase().includes(txt) ||
    p.codigo.toLowerCase().includes(txt)
  ).slice(0, 6);

  if (matches.length === 0) {
    dropdown.innerHTML = `<div class="p-3 text-xs text-slate-400 text-center">No se encontró "${txt}"</div>`;
    dropdown.classList.remove("hidden");
    return;
  }

  dropdown.innerHTML = matches.map(p => {
    const st = stockMap[p.codigo] || 0;
    const imgUrl = formatearUrlImagen(p.imagenUrl);
    const imgHtml = imgUrl
      ? `<img src="${imgUrl}" alt="${p.nombre}" class="w-9 h-9 rounded-lg object-cover bg-slate-900 border border-slate-700 shrink-0" onerror="this.outerHTML='<div class=\\\'w-9 h-9 rounded-lg bg-slate-900 border border-slate-700 flex items-center justify-center text-slate-500 shrink-0\\\'>🍷</div>'">`
      : `<div class="w-9 h-9 rounded-lg bg-slate-900 border border-slate-700 flex items-center justify-center text-slate-500 shrink-0 text-xs">🍷</div>`;

    return `
      <div onclick="agregarAlCarritoPorCodigo('${p.codigo}')" class="p-2.5 hover:bg-slate-700/70 cursor-pointer flex items-center justify-between gap-2">
        <div class="flex items-center gap-2.5 min-w-0">
          ${imgHtml}
          <div class="min-w-0">
            <div class="text-xs font-bold text-white truncate">${p.nombre}</div>
            <div class="text-[10px] text-slate-400 font-mono">${p.codigo} • Stock: <b class="${st > 0 ? 'text-emerald-400' : 'text-rose-400'}">${st}</b></div>
          </div>
        </div>
        <div class="text-right font-mono shrink-0">
          <div class="text-xs font-black text-white">${fmtCRC(p.precioVentaCRC)}</div>
          <div class="text-[10px] text-teal-300">${fmtUSD(p.precioVentaUSD)}</div>
        </div>
      </div>
    `;
  }).join("");
  dropdown.classList.remove("hidden");
}

function agregarAlCarritoPorCodigo(codigo) {
  const prod = state.productos[codigo];
  if (!prod) {
    mostrarToast("Producto no encontrado", "error");
    return;
  }

  const stockMap = calcularStockPorCodigo();
  const stockDisponible = stockMap[codigo] || 0;

  const enCarrito = state.carrito.find(item => item.codigo === codigo);
  const cantActual = enCarrito ? enCarrito.cantidad : 0;

  if (cantActual + 1 > stockDisponible && stockDisponible <= 0) {
    mostrarToast(`"${prod.nombre}" no tiene stock disponible`, "error");
    return;
  }

  if (enCarrito) {
    enCarrito.cantidad += 1;
  } else {
    state.carrito.push({
      codigo: prod.codigo,
      nombre: prod.nombre,
      imagenUrl: prod.imagenUrl || "",
      precioVentaCRC: Number(prod.precioVentaCRC || 0),
      precioVentaUSD: Number(prod.precioVentaUSD || 0),
      costoRefUSD: Number(prod.costoRefUSD || 0),
      costoRefCRC: Number(prod.costoRefCRC || 0),
      cantidad: 1,
      stockMaximo: stockDisponible
    });
  }

  reproducirBeep();
  document.getElementById("searchPos").value = "";
  document.getElementById("posSearchResults").classList.add("hidden");
  renderizarCarrito();
  mostrarToast(`Agregado: ${prod.nombre}`, "success");
}

function modificarCantidadCarrito(codigo, delta) {
  const item = state.carrito.find(i => i.codigo === codigo);
  if (!item) return;

  const nuevo = item.cantidad + delta;
  if (nuevo <= 0) {
    eliminarDelCarrito(codigo);
    return;
  }
  item.cantidad = nuevo;
  renderizarCarrito();
}

function eliminarDelCarrito(codigo) {
  state.carrito = state.carrito.filter(i => i.codigo !== codigo);
  renderizarCarrito();
}

function vaciarCarrito() {
  state.carrito = [];
  renderizarCarrito();
}

function renderizarCarrito() {
  const cont = document.getElementById("cartItemsList");
  const countEl = document.getElementById("cartCount");
  const totalCRCEl = document.getElementById("cartTotalCRC");
  const totalUSDEl = document.getElementById("cartTotalUSD");

  if (!cont) return;

  let totalCRC = 0;
  let totalUSD = 0;
  let totalItems = 0;

  state.carrito.forEach(i => {
    totalCRC += (i.cantidad * i.precioVentaCRC);
    totalUSD += (i.cantidad * i.precioVentaUSD);
    totalItems += i.cantidad;
  });

  countEl.textContent = totalItems;
  totalCRCEl.textContent = fmtCRC(totalCRC);
  totalUSDEl.textContent = `${fmtUSD(totalUSD)} USD`;

  if (state.carrito.length === 0) {
    cont.innerHTML = `
      <div class="flex flex-col items-center justify-center py-6 text-slate-500 text-xs">
        <i data-lucide="shopping-cart" class="w-8 h-8 stroke-1 mb-1 text-slate-600"></i>
        <span>Carrito vacío. Agrega licores para vender.</span>
      </div>
    `;
  } else {
    cont.innerHTML = state.carrito.map(item => {
      const imgUrl = formatearUrlImagen(item.imagenUrl);
      const imgHtml = imgUrl
        ? `<img src="${imgUrl}" alt="${item.nombre}" class="w-8 h-8 rounded-lg object-cover bg-slate-900 border border-slate-700 shrink-0" onerror="this.outerHTML='<div class=\\\'w-8 h-8 rounded-lg bg-slate-900 border border-slate-700 flex items-center justify-center text-slate-500 shrink-0 text-xs\\\'>🍷</div>'">`
        : `<div class="w-8 h-8 rounded-lg bg-slate-900 border border-slate-700 flex items-center justify-center text-slate-500 shrink-0 text-xs">🍷</div>`;

      return `
        <div class="p-2.5 bg-slate-900/90 rounded-xl border border-slate-800 flex items-center justify-between gap-2">
          <div class="flex items-center gap-2 min-w-0 flex-1">
            ${imgHtml}
            <div class="min-w-0 flex-1">
              <h5 class="text-xs font-bold text-white truncate">${item.nombre}</h5>
              <div class="text-[11px] text-slate-400 font-mono">${fmtCRC(item.precioVentaCRC)} (${fmtUSD(item.precioVentaUSD)})</div>
            </div>
          </div>

          <div class="flex items-center gap-1.5 bg-slate-800 rounded-lg p-1">
            <button onclick="modificarCantidadCarrito('${item.codigo}', -1)" class="w-6 h-6 rounded bg-slate-700 text-white font-bold text-xs flex items-center justify-center active:scale-95">-</button>
            <span class="text-xs font-bold text-white w-5 text-center font-mono">${item.cantidad}</span>
            <button onclick="modificarCantidadCarrito('${item.codigo}', 1)" class="w-6 h-6 rounded bg-slate-700 text-white font-bold text-xs flex items-center justify-center active:scale-95">+</button>
          </div>

          <div class="text-right min-w-[70px] font-mono">
            <div class="text-xs font-black text-emerald-400">${fmtCRC(item.cantidad * item.precioVentaCRC)}</div>
            <button onclick="eliminarDelCarrito('${item.codigo}')" class="text-[10px] text-rose-400 hover:text-rose-300">Quitar</button>
          </div>
        </div>
      `;
    }).join("");
  }

  calcularCambio();
  inicializarIconos();
}

function setPaymentMethod(metodo) {
  state.metodoPagoSeleccionado = metodo;
  document.querySelectorAll(".pay-btn").forEach(btn => {
    if (btn.textContent.includes(metodo.split(" ")[0])) {
      btn.classList.add("active");
    } else {
      btn.classList.remove("active");
    }
  });

  const cashHelper = document.getElementById("cashHelper");
  if (metodo === "Efectivo") {
    cashHelper.classList.remove("hidden");
  } else {
    cashHelper.classList.add("hidden");
  }
}

function calcularCambio() {
  const recibido = Number(document.getElementById("cashReceived").value) || 0;
  let totalCRC = 0;
  state.carrito.forEach(i => totalCRC += (i.cantidad * i.precioVentaCRC));

  const cambio = recibido - totalCRC;
  const cambioEl = document.getElementById("cashChange");
  if (recibido > 0) {
    cambioEl.textContent = fmtCRC(Math.max(0, cambio));
  } else {
    cambioEl.textContent = fmtCRC(0);
  }
}

async function completarVenta() {
  if (state.carrito.length === 0) {
    mostrarToast("El carrito está vacío", "error");
    return;
  }

  const vendedor = state.vendedorActual || "Carlos";
  const stockMap = calcularStockPorCodigo(vendedor);

  for (let item of state.carrito) {
    const disp = stockMap[item.codigo] || 0;
    if (item.cantidad > disp) {
      mostrarToast(`Stock insuficiente de ${item.nombre} para ${vendedor} (Disponible: ${disp})`, "error");
      return;
    }
  }

  let totalCRC = 0, totalUSD = 0;
  let gananciaCRC = 0, gananciaUSD = 0;

  state.carrito.forEach(item => {
    const subCRC = item.cantidad * item.precioVentaCRC;
    const subUSD = item.cantidad * item.precioVentaUSD;
    const cCRC = item.cantidad * (item.costoRefCRC || 0);
    const cUSD = item.cantidad * (item.costoRefUSD || 0);

    totalCRC += subCRC;
    totalUSD += subUSD;
    gananciaCRC += (subCRC - cCRC);
    gananciaUSD += (subUSD - cUSD);
  });

  const cliente = document.getElementById("customerName").value.trim() || "Cliente General";
  const idVenta = "VTA-" + Date.now().toString().slice(-6);

  const ventaObj = {
    id: idVenta,
    fecha: new Date().toISOString(),
    vendedor: vendedor,
    items: [...state.carrito],
    totalCRC,
    totalUSD,
    gananciaCRC,
    gananciaUSD,
    cliente,
    metodoPago: state.metodoPagoSeleccionado
  };

  state.ventas.unshift(ventaObj);
  guardarVentasLocal();

  state.ultimaVentaCompletada = ventaObj;

  if (window.confetti) {
    window.confetti({ particleCount: 80, spread: 60, origin: { y: 0.8 } });
  }

  state.carrito = [];
  document.getElementById("cashReceived").value = "";
  document.getElementById("customerName").value = "";
  renderizarTodo();

  abrirModalRecibo(ventaObj);

  // Encolar y sincronizar con Google Sheets
  encolarAccionSincronizacion("registrarVenta", { venta: ventaObj });
}

// ==========================================================================
// MODAL RECIBO Y WHATSAPP
// ==========================================================================
function abrirModalRecibo(venta) {
  const modal = document.getElementById("modalRecibo");
  document.getElementById("reciboNegocio").textContent = state.config.nombreNegocio || "Libro de Inventario";
  document.getElementById("reciboId").textContent = `${venta.id} (👤 ${venta.vendedor || state.vendedorActual})`;
  document.getElementById("reciboFecha").textContent = new Date(venta.fecha).toLocaleString([], { dateStyle: 'short', timeStyle: 'short' });
  document.getElementById("reciboMetodo").textContent = venta.metodoPago;
  document.getElementById("reciboTotalCRC").textContent = fmtCRC(venta.totalCRC);
  document.getElementById("reciboTotalUSD").textContent = `(${fmtUSD(venta.totalUSD)} USD)`;

  const itemsCont = document.getElementById("reciboItems");
  itemsCont.innerHTML = venta.items.map(i => `
    <div class="flex justify-between py-1 font-mono">
      <div>
        <span class="font-bold">${i.cantidad}x</span> ${i.nombre}
      </div>
      <span class="font-bold text-slate-800">${fmtCRC(i.cantidad * i.precioVentaCRC)}</span>
    </div>
  `).join("");

  modal.classList.remove("hidden");
  modal.classList.add("flex");
  inicializarIconos();
}

function cerrarModalRecibo() {
  const modal = document.getElementById("modalRecibo");
  modal.classList.add("hidden");
  modal.classList.remove("flex");
}

function compartirReciboWhatsApp() {
  if (!state.ultimaVentaCompletada) return;
  const v = state.ultimaVentaCompletada;

  const negocio = state.config.nombreNegocio || "DC EL DESTAPE LICORES";
  const telefono = state.config.telefonoNegocio || "+506 8992-7936";
  const fecha = new Date(v.fecha).toLocaleString([], { dateStyle: 'short', timeStyle: 'short' });
  const vendedor = v.vendedor || state.vendedorActual || "Carlos";

  let texto = `🥃 *${negocio.toUpperCase()}* 🥃\n`;
  texto += `📞 *Tel:* ${telefono}\n`;
  texto += `--------------------------------\n`;
  texto += `🧾 *COMPROBANTE DE COMPRA*\n`;
  texto += `📅 Fecha: ${fecha}\n`;
  texto += `🔢 N° Ticket: ${v.id}\n`;
  texto += `👤 Atendido por: ${vendedor}\n`;
  texto += `👥 Cliente: ${v.cliente || "General"}\n`;
  texto += `--------------------------------\n`;
  
  v.items.forEach(i => {
    texto += `• ${i.cantidad}x ${i.nombre} = ${fmtCRC(i.cantidad * i.precioVentaCRC)} (${fmtUSD(i.cantidad * i.precioVentaUSD)})\n`;
  });

  texto += `--------------------------------\n`;
  texto += `💳 *Método de Pago:* ${v.metodoPago || "Efectivo"}\n`;
  texto += `💰 *TOTAL CRC:* ${fmtCRC(v.totalCRC)}\n`;
  texto += `💵 *TOTAL USD:* ${fmtUSD(v.totalUSD)}\n\n`;
  texto += `¡Muchas gracias por su preferencia! 🙏🥂`;

  const encoded = encodeURIComponent(texto);
  const phoneClean = telefono ? telefono.replace(/[^0-9]/g, "") : "";
  const url = phoneClean ? `https://wa.me/${phoneClean}?text=${encoded}` : `https://wa.me/?text=${encoded}`;
  window.open(url, "_blank");
}

// ==========================================================================
// 5. MÓDULO DE FINANZAS Y MANEJO DE 3 CUENTAS (EMPRESA, CARLOS, DANIEL)
// ==========================================================================
function calcularSaldosFinancieros() {
  const tcActual = Number(state.config.tipoCambio) || 520;

  // 1. Ingresos a la Caja de Empresa por Ventas
  let totalVentasCRC = 0;
  let totalVentasUSD = 0;
  state.ventas.forEach(v => {
    totalVentasCRC += Number(v.totalCRC) || 0;
    totalVentasUSD += Number(v.totalUSD) || ((Number(v.totalCRC) || 0) / tcActual);
  });

  // 2. Compras según quién las pagó / financió
  let carlosFinanciaComprasCRC = 0;
  let carlosFinanciaComprasUSD = 0;
  let danielFinanciaComprasCRC = 0;
  let danielFinanciaComprasUSD = 0;
  let empresaPagaComprasCRC = 0;
  let empresaPagaComprasUSD = 0;

  state.compras.forEach(c => {
    const cant = Number(c.cantidad) || 1;
    const tc = Number(c.tipoCambio) || tcActual;
    const cUSD = cant * (Number(c.costoUnitarioUSD) || 0);
    const cCRC = cant * (Number(c.costoUnitarioCRC) || (cUSD * tc));
    const pagador = c.pagadoPor || c.vendedor || "Carlos";

    if (pagador === "Carlos") {
      carlosFinanciaComprasCRC += cCRC;
      carlosFinanciaComprasUSD += cUSD;
    } else if (pagador === "Daniel") {
      danielFinanciaComprasCRC += cCRC;
      danielFinanciaComprasUSD += cUSD;
    } else { // "Empresa"
      empresaPagaComprasCRC += cCRC;
      empresaPagaComprasUSD += cUSD;
    }
  });

  // 3. Movimientos directos de dinero (Aportes, Pagos de Deuda, Gastos)
  let carlosAportesDirectosCRC = 0;
  let carlosAportesDirectosUSD = 0;
  let carlosReembolsosCRC = 0;
  let carlosReembolsosUSD = 0;

  let danielAportesDirectosCRC = 0;
  let danielAportesDirectosUSD = 0;
  let danielReembolsosCRC = 0;
  let danielReembolsosUSD = 0;

  let empresaGastosCRC = 0;
  let empresaGastosUSD = 0;
  let empresaCapitalPropioCRC = 0;  // Saldo inicial o capital propio de la empresa
  let empresaCapitalPropioUSD = 0;

  state.movimientosDinero.forEach(m => {
    const tc = Number(m.tipoCambio) || tcActual;
    const mCRC = Number(m.montoCRC) || ((Number(m.montoUSD) || 0) * tc);
    const mUSD = Number(m.montoUSD) || (mCRC / tc);

    if (m.tipo === "aporte_capital") {
      if (m.socio === "Carlos") {
        carlosAportesDirectosCRC += mCRC;
        carlosAportesDirectosUSD += mUSD;
      } else if (m.socio === "Daniel") {
        danielAportesDirectosCRC += mCRC;
        danielAportesDirectosUSD += mUSD;
      } else {
        // Socio vacío o "Empresa" = capital propio / saldo inicial de la empresa
        empresaCapitalPropioCRC += mCRC;
        empresaCapitalPropioUSD += mUSD;
      }
    } else if (m.tipo === "pago_socio") {
      if (m.socio === "Carlos") {
        carlosReembolsosCRC += mCRC;
        carlosReembolsosUSD += mUSD;
      } else if (m.socio === "Daniel") {
        danielReembolsosCRC += mCRC;
        danielReembolsosUSD += mUSD;
      }
    } else if (m.tipo === "gasto_operativo") {
      empresaGastosCRC += mCRC;
      empresaGastosUSD += mUSD;
    }
  });

  // Totales aportados y deudas de socios
  const carlosTotalAportadoCRC = carlosFinanciaComprasCRC + carlosAportesDirectosCRC;
  const carlosTotalAportadoUSD = carlosFinanciaComprasUSD + carlosAportesDirectosUSD;
  const carlosDeudaCRC = carlosTotalAportadoCRC - carlosReembolsosCRC;
  const carlosDeudaUSD = carlosTotalAportadoUSD - carlosReembolsosUSD;

  const danielTotalAportadoCRC = danielFinanciaComprasCRC + danielAportesDirectosCRC;
  const danielTotalAportadoUSD = danielFinanciaComprasUSD + danielAportesDirectosUSD;
  const danielDeudaCRC = danielTotalAportadoCRC - danielReembolsosCRC;
  const danielDeudaUSD = danielTotalAportadoUSD - danielReembolsosUSD;

  // Saldo real en caja de la Empresa
  const totalAportesSociosCRC = carlosAportesDirectosCRC + danielAportesDirectosCRC;
  const totalAportesSociosUSD = carlosAportesDirectosUSD + danielAportesDirectosUSD;
  const totalReembolsosSociosCRC = carlosReembolsosCRC + danielReembolsosCRC;
  const totalReembolsosSociosUSD = carlosReembolsosUSD + danielReembolsosUSD;

  // Saldo = Ventas + Capital Propio Empresa + Aportes Socios - Compras Empresa - Reembolsos - Gastos
  const saldoEmpresaCRC = totalVentasCRC + empresaCapitalPropioCRC + totalAportesSociosCRC - empresaPagaComprasCRC - totalReembolsosSociosCRC - empresaGastosCRC;
  const saldoEmpresaUSD = totalVentasUSD + empresaCapitalPropioUSD + totalAportesSociosUSD - empresaPagaComprasUSD - totalReembolsosSociosUSD - empresaGastosUSD;

  return {
    tcActual,
    totalVentasCRC, totalVentasUSD,
    empresa: {
      saldoCRC: saldoEmpresaCRC,
      saldoUSD: saldoEmpresaUSD,
      ventasCRC: totalVentasCRC,
      gastosCRC: empresaGastosCRC + empresaPagaComprasCRC
    },
    carlos: {
      deudaCRC: carlosDeudaCRC,
      deudaUSD: carlosDeudaUSD,
      totalAportadoCRC: carlosTotalAportadoCRC,
      totalAportadoUSD: carlosTotalAportadoUSD,
      reembolsadoCRC: carlosReembolsosCRC
    },
    daniel: {
      deudaCRC: danielDeudaCRC,
      deudaUSD: danielDeudaUSD,
      totalAportadoCRC: danielTotalAportadoCRC,
      totalAportadoUSD: danielTotalAportadoUSD,
      reembolsadoCRC: danielReembolsosCRC
    }
  };
}

function renderizarFinanzas() {
  const fin = calcularSaldosFinancieros();

  // Actualizar Caja Empresa
  const elEmpresaCRC = document.getElementById("finSaldoEmpresaCRC");
  const elEmpresaUSD = document.getElementById("finSaldoEmpresaUSD");
  if (elEmpresaCRC) elEmpresaCRC.textContent = fmtCRC(fin.empresa.saldoCRC);
  if (elEmpresaUSD) elEmpresaUSD.textContent = fmtUSD(fin.empresa.saldoUSD);

  // Actualizar Carlos
  const elCarlosCRC = document.getElementById("finDeudaCarlosCRC");
  const elCarlosUSD = document.getElementById("finDeudaCarlosUSD");
  if (elCarlosCRC) elCarlosCRC.textContent = fmtCRC(fin.carlos.deudaCRC);
  if (elCarlosUSD) elCarlosUSD.textContent = `(${fmtUSD(fin.carlos.deudaUSD)} USD)`;

  // Actualizar Daniel
  const elDanielCRC = document.getElementById("finDeudaDanielCRC");
  const elDanielUSD = document.getElementById("finDeudaDanielUSD");
  if (elDanielCRC) elDanielCRC.textContent = fmtCRC(fin.daniel.deudaCRC);
  if (elDanielUSD) elDanielUSD.textContent = `(${fmtUSD(fin.daniel.deudaUSD)} USD)`;

  renderizarHistorialFinanzas();
}

function obtenerListaVentasConsolidadas() {
  const map = new Map();
  
  (state.ventas || []).forEach(v => {
    const id = v.id || ("VTA-" + (v.fecha || ""));
    if (!map.has(id)) {
      map.set(id, {
        id: id,
        fecha: v.fecha,
        vendedor: v.vendedor || "Carlos",
        cliente: v.cliente || "Cliente General",
        metodoPago: v.metodoPago || "Efectivo",
        totalCRC: 0,
        totalUSD: 0,
        itemsSummary: []
      });
    }

    const sale = map.get(id);
    if (v.items && Array.isArray(v.items)) {
      sale.totalCRC = Number(v.totalCRC) || 0;
      sale.totalUSD = Number(v.totalUSD) || 0;
      sale.itemsSummary = v.items.map(i => `${i.cantidad}x ${i.nombre || i.codigo}`);
    } else {
      const cant = Number(v.cantidad) || 1;
      const totCRC = Number(v.totalCRC) || (cant * (Number(v.precioCRC) || 0));
      const totUSD = Number(v.totalUSD) || (cant * (Number(v.precioUSD) || 0));
      sale.totalCRC += totCRC;
      sale.totalUSD += totUSD;
      sale.itemsSummary.push(`${cant}x ${v.nombre || v.codigo}`);
    }
  });

  return Array.from(map.values());
}

function renderizarHistorialFinanzas() {
  const cont = document.getElementById("finMovimientosList");
  const countEl = document.getElementById("finMovimientosCount");
  if (!cont) return;

  const filtro = state.filtroFinanzas || "todos";

  // 1. Movimientos directos (Aportes, Pagos a socios, Gastos)
  const movsDirectos = (state.movimientosDinero || []).map(m => ({
    origen: 'movimiento',
    id: m.id,
    fecha: m.fecha,
    tipo: m.tipo,
    socio: m.socio || "",
    cuentaOrigen: m.cuentaOrigen || "",
    cuentaDestino: m.cuentaDestino || "",
    montoCRC: Number(m.montoCRC) || 0,
    montoUSD: Number(m.montoUSD) || 0,
    metodoPago: m.metodoPago || "SINPE Móvil",
    notas: m.notas || "",
    registradoPor: m.registradoPor || m.socio || "Carlos"
  }));

  // 2. Inyecciones de Dinero por Ventas del POS
  const ventasConsolidadas = obtenerListaVentasConsolidadas().map(v => ({
    origen: 'venta',
    id: v.id,
    fecha: v.fecha,
    tipo: 'venta_pos',
    socio: v.vendedor || "Carlos",
    cuentaOrigen: `Venta POS (${v.vendedor || "Carlos"})`,
    cuentaDestino: "Caja Empresa",
    montoCRC: v.totalCRC,
    montoUSD: v.totalUSD,
    metodoPago: v.metodoPago || "Efectivo",
    notas: `${v.itemsSummary.join(", ")}${v.cliente && v.cliente !== 'Cliente General' ? ` • Cl: ${v.cliente}` : ''}`,
    registradoPor: v.vendedor || "Carlos"
  }));

  // 3. Egresos por Compras pagadas directamente con fondos de la Empresa
  const comprasEmpresa = (state.compras || [])
    .filter(c => (c.pagadoPor === "Empresa" || c.financiadoPor === "Empresa"))
    .map(c => {
      const cant = Number(c.cantidad) || 1;
      const tc = Number(c.tipoCambio) || (state.config.tipoCambio || 520);
      const cUSD = cant * (Number(c.costoUnitarioUSD) || 0);
      const cCRC = cant * (Number(c.costoUnitarioCRC) || (cUSD * tc));
      return {
        origen: 'compra_empresa',
        id: c.id,
        fecha: c.fecha,
        tipo: 'compra_empresa',
        socio: c.vendedor || "Carlos",
        cuentaOrigen: "Caja Empresa",
        cuentaDestino: c.proveedor || "Proveedor",
        montoCRC: cCRC,
        montoUSD: cUSD,
        metodoPago: "Caja Empresa",
        notas: `Compra ${cant}x ${c.nombre || c.codigo} (${c.proveedor || 'Proveedor'})`,
        registradoPor: c.vendedor || "Carlos"
      };
    });

  // Combinar todos los movimientos
  let todosLosMovimientos = [...movsDirectos, ...ventasConsolidadas, ...comprasEmpresa];

  // Ordenar cronológicamente (más recientes primero)
  todosLosMovimientos.sort((a, b) => {
    const tA = a.fecha ? new Date(a.fecha).getTime() : 0;
    const tB = b.fecha ? new Date(b.fecha).getTime() : 0;
    return tB - tA;
  });

  // Filtrar según la pestaña seleccionada
  let lista = todosLosMovimientos;
  if (filtro === "Empresa") {
    lista = lista.filter(m => 
      m.tipo === "venta_pos" ||
      m.tipo === "compra_empresa" ||
      m.cuentaOrigen === "Empresa" || 
      m.cuentaDestino === "Empresa" || 
      m.tipo === "gasto_operativo" ||
      m.tipo === "aporte_capital" ||
      m.tipo === "pago_socio"
    );
  } else if (filtro === "Carlos") {
    lista = lista.filter(m => 
      m.socio === "Carlos" || 
      m.registradoPor === "Carlos" || 
      m.cuentaOrigen === "Carlos" || 
      m.cuentaDestino === "Carlos" ||
      (m.tipo === "venta_pos" && m.socio === "Carlos")
    );
  } else if (filtro === "Daniel") {
    lista = lista.filter(m => 
      m.socio === "Daniel" || 
      m.registradoPor === "Daniel" || 
      m.cuentaOrigen === "Daniel" || 
      m.cuentaDestino === "Daniel" ||
      (m.tipo === "venta_pos" && m.socio === "Daniel")
    );
  }

  if (countEl) countEl.textContent = lista.length;

  if (lista.length === 0) {
    cont.innerHTML = `
      <div class="text-center py-8 text-slate-500 space-y-2">
        <i data-lucide="wallet" class="w-8 h-8 mx-auto text-slate-600"></i>
        <p class="text-xs">No hay movimientos financieros registrados aún.</p>
        <p class="text-[10px] text-slate-600">Registra ventas, aportes o pagos para ver el flujo de caja en vivo.</p>
      </div>
    `;
    inicializarIconos();
    return;
  }

  cont.innerHTML = lista.map(m => {
    const isVenta = m.tipo === "venta_pos";
    const isAporte = m.tipo === "aporte_capital";
    const isPagoSocio = m.tipo === "pago_socio";
    const isGasto = m.tipo === "gasto_operativo";
    const isCompraEmpresa = m.tipo === "compra_empresa";

    let colorBadge = "bg-blue-950/60 text-blue-400 border-blue-500/30";
    let iconName = "plus-circle";
    let tituloTipo = "Aporte de Capital";
    let subtitulo = `${m.socio || "Socio"} ➔ Caja Empresa`;
    let signoMonto = "+";
    let colorMonto = "text-emerald-400";

    if (isVenta) {
      const vendBadge = m.socio === "Daniel" ? "text-violet-300" : "text-blue-300";
      colorBadge = "bg-emerald-950/80 text-emerald-300 border-emerald-500/40 shadow-sm shadow-emerald-500/20";
      iconName = "shopping-cart";
      tituloTipo = `Ingreso x Venta (${m.socio || "POS"})`;
      subtitulo = `Venta POS <span class="${vendBadge} font-bold">(${m.socio})</span> ➔ Caja Empresa`;
      signoMonto = "+";
      colorMonto = "text-emerald-400 font-black";
    } else if (isPagoSocio) {
      colorBadge = "bg-emerald-950/60 text-emerald-400 border-emerald-500/30";
      iconName = "arrow-up-right";
      tituloTipo = "Pago / Abono a Socio";
      subtitulo = `Caja Empresa ➔ ${m.socio || "Socio"}`;
      signoMonto = "-";
      colorMonto = "text-blue-400 font-black";
    } else if (isGasto) {
      colorBadge = "bg-rose-950/60 text-rose-400 border-rose-500/30";
      iconName = "receipt";
      tituloTipo = "Gasto Operativo";
      subtitulo = `Caja Empresa ➔ Gastos`;
      signoMonto = "-";
      colorMonto = "text-rose-400 font-black";
    } else if (isCompraEmpresa) {
      colorBadge = "bg-amber-950/60 text-amber-300 border-amber-500/30";
      iconName = "truck";
      tituloTipo = "Compra Pagada x Empresa";
      subtitulo = `Caja Empresa ➔ Proveedor`;
      signoMonto = "-";
      colorMonto = "text-amber-400 font-black";
    }

    const fechaStr = m.fecha ? new Date(m.fecha).toLocaleString([], { dateStyle: 'short', timeStyle: 'short' }) : todayStr();
    const montoCRC = Number(m.montoCRC) || 0;
    const montoUSD = Number(m.montoUSD) || 0;

    const actionRight = (m.origen === 'movimiento') ? `
      <button onclick="eliminarMovimientoDinero('${m.id}')" title="Eliminar transacción" class="text-slate-500 hover:text-rose-400 p-1 transition-all">
        <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
      </button>
    ` : `
      <span class="text-[9px] font-mono px-1.5 py-0.5 rounded bg-slate-800 text-slate-400 border border-slate-700">
        ${isVenta ? 'POS ⚡' : 'COMPRA 📦'}
      </span>
    `;

    return `
      <div class="p-3 bg-slate-900/90 border border-slate-800 rounded-xl space-y-1.5 hover:border-slate-700 transition-all">
        <div class="flex items-center justify-between">
          <div class="flex items-center gap-2">
            <span class="px-2 py-0.5 rounded-lg text-[10px] font-bold border ${colorBadge} flex items-center gap-1">
              <i data-lucide="${iconName}" class="w-3 h-3"></i>
              ${tituloTipo}
            </span>
            <span class="text-[10px] text-slate-400 font-mono">${fechaStr}</span>
          </div>
          ${actionRight}
        </div>

        <div class="flex items-center justify-between gap-2">
          <div class="min-w-0 flex-1">
            <div class="font-bold text-white text-xs truncate">${subtitulo}</div>
            <div class="text-[10px] text-slate-400 flex items-center gap-1 mt-0.5 truncate">
              <span>💳 ${m.metodoPago || 'Efectivo'}</span>
              ${m.notas ? `<span class="truncate">• 📝 ${m.notas}</span>` : ''}
            </div>
          </div>
          <div class="text-right font-mono shrink-0">
            <div class="text-sm font-black ${colorMonto}">${signoMonto}${fmtCRC(montoCRC)}</div>
            <div class="text-[10px] text-slate-400">(${fmtUSD(montoUSD)})</div>
          </div>
        </div>
      </div>
    `;
  }).join("");

  inicializarIconos();
}

function filtrarHistorialFinanzas(filtro) {
  state.filtroFinanzas = filtro;
  ["todos", "empresa", "carlos", "daniel"].forEach(f => {
    const btn = document.getElementById(`finFilter-${f}`);
    if (btn) {
      if (f.toLowerCase() === filtro.toLowerCase()) {
        btn.className = "py-1 rounded-lg bg-indigo-600 text-white text-center font-bold";
      } else {
        btn.className = "py-1 rounded-lg bg-slate-900 text-slate-400 hover:text-white text-center font-bold";
      }
    }
  });
  renderizarHistorialFinanzas();
}

function abrirModalMovimientoDinero(tipoPredefinido = "aporte_capital", socioPredefinido = "") {
  const modal = document.getElementById("modalMovimientoDinero");
  if (!modal) return;

  const tipoSelect = document.getElementById("finTipoMovimiento");
  if (tipoSelect) tipoSelect.value = tipoPredefinido;

  const socioSelect = document.getElementById("finSocioSelect");
  if (socioSelect) {
    if (socioPredefinido) socioSelect.value = socioPredefinido;
    else socioSelect.value = state.vendedorActual || "Carlos";
  }

  const inputFecha = document.getElementById("finFecha");
  if (inputFecha) inputFecha.value = todayStr();

  const inputTC = document.getElementById("finTipoCambio");
  if (inputTC) inputTC.value = state.config.tipoCambio || 520;

  const inputCRC = document.getElementById("finMontoCRC");
  if (inputCRC) inputCRC.value = "";

  const inputUSD = document.getElementById("finMontoUSD");
  if (inputUSD) inputUSD.value = "";

  const inputNotas = document.getElementById("finNotas");
  if (inputNotas) inputNotas.value = "";

  actualizarCamposModalFinanzas();

  modal.classList.remove("hidden");
  modal.classList.add("flex");
  inicializarIconos();
}

function cerrarModalMovimientoDinero() {
  const modal = document.getElementById("modalMovimientoDinero");
  if (modal) {
    modal.classList.add("hidden");
    modal.classList.remove("flex");
  }
}

function actualizarCamposModalFinanzas() {
  const tipo = document.getElementById("finTipoMovimiento").value;
  const grupoSocio = document.getElementById("finGrupoSocio");
  const titulo = document.getElementById("modalFinanzasTitulo");

  if (tipo === "gasto_operativo") {
    if (grupoSocio) grupoSocio.classList.add("hidden");
    if (titulo) titulo.innerHTML = `<i data-lucide="receipt" class="w-4 h-4 text-rose-400"></i><span>Gasto Operativo de Empresa</span>`;
  } else if (tipo === "pago_socio") {
    if (grupoSocio) grupoSocio.classList.remove("hidden");
    if (titulo) titulo.innerHTML = `<i data-lucide="arrow-up-right" class="w-4 h-4 text-emerald-400"></i><span>Pagar / Abono a Socio</span>`;
  } else {
    if (grupoSocio) grupoSocio.classList.remove("hidden");
    if (titulo) titulo.innerHTML = `<i data-lucide="plus-circle" class="w-4 h-4 text-blue-400"></i><span>Aporte de Capital a Empresa</span>`;
  }
  inicializarIconos();
}

function autoConvertirFinanzas(origen) {
  const tc = Number(document.getElementById("finTipoCambio").value) || Number(state.config.tipoCambio) || 520;
  const elCRC = document.getElementById("finMontoCRC");
  const elUSD = document.getElementById("finMontoUSD");

  if (origen === 'CRC' && elCRC && elUSD) {
    const crc = Number(elCRC.value) || 0;
    elUSD.value = crc > 0 ? (crc / tc).toFixed(2) : "";
  } else if (origen === 'USD' && elCRC && elUSD) {
    const usd = Number(elUSD.value) || 0;
    elCRC.value = usd > 0 ? Math.round(usd * tc) : "";
  }
}

async function guardarMovimientoDinero(e) {
  if (e) e.preventDefault();

  const tipo = document.getElementById("finTipoMovimiento").value;
  const socio = tipo !== "gasto_operativo" ? document.getElementById("finSocioSelect").value : "";
  const fecha = document.getElementById("finFecha").value || todayStr();
  const montoCRC = Number(document.getElementById("finMontoCRC").value) || 0;
  const tc = Number(document.getElementById("finTipoCambio").value) || Number(state.config.tipoCambio) || 520;
  const montoUSD = Number(document.getElementById("finMontoUSD").value) || (montoCRC / tc);
  const metodoPago = document.getElementById("finMetodoPago").value;
  const notas = document.getElementById("finNotas").value.trim();

  if (montoCRC <= 0) {
    mostrarToast("El monto debe ser mayor a 0", "error");
    return;
  }

  let cuentaOrigen = "Empresa";
  let cuentaDestino = "Empresa";

  if (tipo === "aporte_capital") {
    cuentaOrigen = socio;
    cuentaDestino = "Empresa";
  } else if (tipo === "pago_socio") {
    cuentaOrigen = "Empresa";
    cuentaDestino = socio;
  } else if (tipo === "gasto_operativo") {
    cuentaOrigen = "Empresa";
    cuentaDestino = "Gasto Operativo";
  }

  const movObj = {
    id: "MOV-" + Date.now().toString().slice(-6),
    fecha,
    tipo,
    cuentaOrigen,
    cuentaDestino,
    socio,
    montoCRC,
    montoUSD,
    tipoCambio: tc,
    metodoPago,
    notas,
    registradoPor: state.vendedorActual || "Carlos"
  };

  state.movimientosDinero.unshift(movObj);
  guardarFinanzasLocal();
  renderizarFinanzas();
  cerrarModalMovimientoDinero();

  const msg = tipo === "aporte_capital" 
    ? `¡Aporte de ${socio} registrado! (+${fmtCRC(montoCRC)})` 
    : tipo === "pago_socio"
    ? `¡Pago a ${socio} registrado! (-${fmtCRC(montoCRC)})`
    : `Gasto registrado (-${fmtCRC(montoCRC)})`;

  mostrarToast(msg, "success");

  // Encolar y sincronizar con Google Sheets
  encolarAccionSincronizacion("registrarMovimiento", { movimiento: movObj });
}

async function eliminarMovimientoDinero(id) {
  if (!confirm("¿Deseas eliminar este movimiento financiero?")) return;

  state.movimientosDinero = state.movimientosDinero.filter(m => m.id !== id);
  guardarFinanzasLocal();
  renderizarFinanzas();
  mostrarToast("Movimiento eliminado localmente", "info");

  // Encolar y sincronizar con Google Sheets
  encolarAccionSincronizacion("eliminarMovimiento", { id });
}

// ==========================================================================
// 6. IMPORTACIÓN Y EXPORTACIÓN EXCEL (.XLSX)
// ==========================================================================
async function exportarLibroExcel() {
  if (!window.XLSX) {
    mostrarToast("Cargando motor de Excel...", "info");
    return;
  }
  const XLSX = window.XLSX;
  const wb = XLSX.utils.book_new();

  // Hoja Productos
  const prodRows = Object.values(state.productos).map(p => ({
    codigo: p.codigo,
    nombre: p.nombre,
    categoria: p.categoria || "",
    imagenUrl: p.imagenUrl || "",
    precioVentaUSD: p.precioVentaUSD || 0,
    precioVentaCRC: p.precioVentaCRC || 0,
    stockInicial: p.stockInicial || 0,
    costoRefUSD: p.costoRefUSD || 0,
    costoRefCRC: p.costoRefCRC || 0
  }));

  // Hoja Compras
  const compRows = state.compras.map(p => ({
    id: p.id,
    fecha: p.fecha,
    vendedor: p.vendedor || "Carlos",
    pagadoPor: p.pagadoPor || p.vendedor || "Carlos",
    codigo: p.codigo,
    cantidad: p.cantidad,
    costoUnitarioUSD: p.costoUnitarioUSD,
    tipoCambio: p.tipoCambio,
    costoUnitarioCRC: p.costoUnitarioCRC,
    totalUSD: Number(p.cantidad || 0) * Number(p.costoUnitarioUSD || 0),
    totalCRC: Number(p.cantidad || 0) * Number(p.costoUnitarioCRC || 0),
    proveedor: p.proveedor || "",
    notas: p.notas || ""
  }));

  // Hoja Ventas
  const ventRows = [];
  state.ventas.forEach(v => {
    if (v.items) {
      v.items.forEach(i => {
        ventRows.push({
          id: v.id,
          fecha: v.fecha ? v.fecha.slice(0, 16).replace("T", " ") : todayStr(),
          vendedor: v.vendedor || "Carlos",
          codigo: i.codigo,
          nombre: i.nombre || "",
          cantidad: i.cantidad,
          precioUnitarioUSD: i.precioVentaUSD,
          precioUnitarioCRC: i.precioVentaCRC,
          totalUSD: (i.cantidad || 1) * (i.precioVentaUSD || 0),
          totalCRC: (i.cantidad || 1) * (i.precioVentaCRC || 0),
          cliente: v.cliente || "General",
          metodoPago: v.metodoPago || "Efectivo"
        });
      });
    }
  });

  // Hoja Finanzas
  const finRows = state.movimientosDinero.map(m => ({
    id: m.id,
    fecha: m.fecha,
    tipo: m.tipo,
    cuentaOrigen: m.cuentaOrigen,
    cuentaDestino: m.cuentaDestino,
    socio: m.socio || "",
    montoCRC: m.montoCRC,
    montoUSD: m.montoUSD,
    tipoCambio: m.tipoCambio,
    metodoPago: m.metodoPago || "SINPE Móvil",
    notas: m.notas || "",
    registradoPor: m.registradoPor || "Carlos"
  }));

  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(prodRows), "Productos");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(compRows), "Compras");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ventRows), "Ventas");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(finRows), "Finanzas");

  const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([out], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `inventario_licores_${todayStr()}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  mostrarToast("Archivo Excel descargado con éxito 📊", "success");
}

async function importarArchivoExcel(event) {
  const file = event.target.files[0];
  if (!file) return;

  mostrarToast("Leyendo archivo Excel...", "info");
  try {
    const XLSX = window.XLSX;
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
        
        // Hojas
        const prodKey = wb.SheetNames.find(n => ["productos", "catalogo", "products"].includes(n.trim().toLowerCase()));
        const compKey = wb.SheetNames.find(n => ["compras", "purchases"].includes(n.trim().toLowerCase()));
        const ventKey = wb.SheetNames.find(n => ["ventas", "sales"].includes(n.trim().toLowerCase()));

        let importadosCount = 0;

        if (prodKey) {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets[prodKey], { defval: "" });
          rows.forEach(r => {
            const codigo = String(r.codigo || r.Codigo || r.SKU || "").trim();
            if (codigo) {
              state.productos[codigo] = {
                codigo,
                nombre: String(r.nombre || r.Nombre || codigo).trim(),
                categoria: String(r.categoria || r.Categoria || "General").trim(),
                imagenUrl: String(r.imagenUrl || r.Imagen_URL || r.imagen || r.foto || "").trim(),
                precioVentaUSD: Number(r.precioVentaUSD || r.PrecioVentaUSD || 0),
                precioVentaCRC: Number(r.precioVentaCRC || r.PrecioVentaCRC || 0),
                stockInicial: Number(r.stockInicial || r.StockInicial || 0),
                costoRefUSD: Number(r.costoRefUSD || r.CostoRefUSD || 0),
                costoRefCRC: Number(r.costoRefCRC || r.CostoRefCRC || 0)
              };
              importadosCount++;
            }
          });
          guardarProductosLocal();
        }

        if (compKey) {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets[compKey], { defval: "" });
          const comprasNuevas = rows.map(r => ({
            id: String(r.id || uid()),
            fecha: r.fecha ? new Date(r.fecha).toISOString().slice(0, 10) : todayStr(),
            vendedor: String(r.vendedor || r.Vendedor || "Carlos").trim(),
            codigo: String(r.codigo || r.Codigo).trim(),
            cantidad: Number(r.cantidad || 1),
            costoUnitarioUSD: Number(r.costoUnitarioUSD || 0),
            tipoCambio: Number(r.tipoCambio || state.config.tipoCambio || 520),
            costoUnitarioCRC: Number(r.costoUnitarioCRC || 0),
            proveedor: String(r.proveedor || r.Proveedor || ""),
            notas: String(r.notas || r.Notas || "")
          })).filter(c => c.codigo);
          state.compras = [...comprasNuevas, ...state.compras];
          guardarComprasLocal();
        }

        renderizarTodo();
        mostrarToast(`¡Importación exitosa! (${importadosCount} productos)`, "success");
      } catch (err) {
        mostrarToast("Error al procesar las hojas del Excel", "error");
      }
    };
    reader.readAsArrayBuffer(file);
  } catch (e) {
    mostrarToast("No se pudo leer el archivo", "error");
  }
}

// ==========================================================================
// 6. ESCÁNER DE CÓDIGO DE BARRAS / QR
// ==========================================================================
function abrirEscaner(modo = "buscar") {
  state.modoEscaner = modo;
  const modal = document.getElementById("modalEscaner");
  modal.classList.remove("hidden");
  modal.classList.add("flex");
  inicializarIconos();

  setTimeout(() => {
    iniciarCamara();
  }, 200);
}

function iniciarCamara() {
  if (state.escanerActivo) {
    state.escanerActivo.stop().catch(() => {});
  }

  const html5QrCode = new Html5Qrcode("reader");
  state.escanerActivo = html5QrCode;

  html5QrCode.start(
    { facingMode: "environment" },
    { fps: 15, qrbox: { width: 250, height: 250 }, aspectRatio: 1.0 },
    (decodedText) => onCodigoEscaneado(decodedText),
    () => {}
  ).catch(err => {
    mostrarToast("No se pudo acceder a la cámara", "error");
    cerrarEscaner();
  });
}

function cerrarEscaner() {
  if (state.escanerActivo) {
    state.escanerActivo.stop().then(() => state.escanerActivo = null).catch(() => state.escanerActivo = null);
  }
  const modal = document.getElementById("modalEscaner");
  modal.classList.add("hidden");
  modal.classList.remove("flex");
}

function onCodigoEscaneado(codigo) {
  reproducirBeep();
  cerrarEscaner();

  if (state.modoEscaner === "nuevo_producto") {
    document.getElementById("prodCodigo").value = codigo;
    mostrarToast(`Código: ${codigo}`, "success");
  } else if (state.modoEscaner === "buscar") {
    document.getElementById("searchInventory").value = codigo;
    cambiarVista("inventario");
    filtrarInventario();
  } else if (state.modoEscaner === "compra") {
    const select = document.getElementById("compraProductoSelect");
    select.value = codigo;
    seleccionarProductoCompra();
  } else if (state.modoEscaner === "venta") {
    agregarAlCarritoPorCodigo(codigo);
  }
}

function reproducirBeep() {
  try {
    const ctx = new (window.AudioContext || window.webkitAudioContext)();
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.type = "sine";
    osc.frequency.setValueAtTime(880, ctx.currentTime);
    gain.gain.setValueAtTime(0.2, ctx.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.00001, ctx.currentTime + 0.15);
    osc.connect(gain);
    gain.connect(ctx.destination);
    osc.start();
    osc.stop(ctx.currentTime + 0.15);
  } catch(e) {}
}

// ==========================================================================
// 7. CONEXIÓN API CON GOOGLE APPS SCRIPT Y MOTOR OFFLINE-FIRST
// ==========================================================================
let sincronizandoCola = false;

function encolarAccionSincronizacion(accion, datos) {
  const item = {
    id: "SYNC-" + Date.now() + "-" + Math.random().toString(36).slice(2, 6),
    accion: accion,
    datos: datos,
    fecha: new Date().toISOString()
  };

  if (!state.colaSincronizacion) state.colaSincronizacion = [];
  state.colaSincronizacion.push(item);
  guardarColaLocal();
  actualizarIndicadorOffline();

  // Intentar sincronizar en segundo plano si hay conexión
  if (navigator.onLine && state.config.sheetsUrl) {
    procesarColaSincronizacion(false);
  }
}

async function procesarColaSincronizacion(mostrarFeedback = false) {
  if (sincronizandoCola) return;

  actualizarIndicadorOffline();

  if (!navigator.onLine) {
    if (mostrarFeedback) {
      mostrarToast(`Sin conexión. ${state.colaSincronizacion.length} cambios guardados localmente 💾`, "info");
    }
    return;
  }

  if (!state.config.sheetsUrl) {
    if (mostrarFeedback) {
      mostrarToast("Configura la URL de Google Sheets en Ajustes ⚙️", "error");
    }
    return;
  }

  if (!state.colaSincronizacion || state.colaSincronizacion.length === 0) {
    if (mostrarFeedback) {
      await sincronizarConSheets(true);
    }
    actualizarIndicadorOffline();
    return;
  }

  sincronizandoCola = true;
  const icon = document.getElementById("syncIcon");
  if (icon) icon.classList.add("animate-spin");

  if (mostrarFeedback) {
    mostrarToast(`Subiendo ${state.colaSincronizacion.length} cambios pendientes a Google Sheets... ☁️`, "info");
  }

  try {
    while (state.colaSincronizacion.length > 0) {
      const item = state.colaSincronizacion[0];
      try {
        await enviarPeticionSheets(item.accion, item.datos);
        // Si no arrojó excepción de red, se procesó
        state.colaSincronizacion.shift();
        guardarColaLocal();
        actualizarIndicadorOffline();
      } catch (err) {
        console.warn("Fallo temporal de red al procesar item de sincronización:", item, err);
        break; // Detener bucle y mantener los ítems restantes en cola
      }
    }

    if (state.colaSincronizacion.length === 0) {
      mostrarToast("¡Todos los cambios sin internet se sincronizaron con Google Sheets! 🚀", "success");
      // Descargar datos consolidados
      await sincronizarConSheets(false);
    } else {
      mostrarToast(`Quedan ${state.colaSincronizacion.length} cambios pendientes por sincronizar`, "info");
    }
  } catch (globalErr) {
    console.error("Error al procesar cola de sincronización:", globalErr);
  } finally {
    sincronizandoCola = false;
    if (icon) icon.classList.remove("animate-spin");
    actualizarIndicadorOffline();
  }
}

function actualizarIndicadorOffline() {
  const isOnline = navigator.onLine;
  const pendingCount = (state.colaSincronizacion || []).length;
  const badge = document.getElementById("syncPendingBadge");
  const banner = document.getElementById("offlineSyncBanner");
  const bannerText = document.getElementById("offlineSyncText");

  if (badge) {
    if (pendingCount > 0) {
      badge.textContent = pendingCount;
      badge.classList.remove("hidden");
      badge.classList.add("flex");
    } else {
      badge.classList.add("hidden");
      badge.classList.remove("flex");
    }
  }

  if (banner && bannerText) {
    if (!isOnline) {
      bannerText.textContent = pendingCount > 0 
        ? `📡 Sin Internet (${pendingCount} cambios guardados localmente)`
        : `📡 Modo Offline: Trabajando 100% en memoria local`;
      banner.className = "max-w-md mx-auto px-3.5 py-1.5 mt-2 bg-amber-950/90 border border-amber-500/50 rounded-xl text-amber-200 text-xs font-bold flex items-center justify-between shadow-lg animate-pulse";
      banner.classList.remove("hidden");
    } else if (pendingCount > 0) {
      bannerText.textContent = `🟡 ${pendingCount} cambio(s) pendiente(s) por subir a Sheets`;
      banner.className = "max-w-md mx-auto px-3.5 py-1.5 mt-2 bg-indigo-950/90 border border-indigo-500/50 rounded-xl text-indigo-200 text-xs font-bold flex items-center justify-between shadow-lg";
      banner.classList.remove("hidden");
    } else {
      banner.classList.add("hidden");
    }
  }

  actualizarBadgeConexion();
}

function actualizarBadgeConexion() {
  const badge = document.getElementById("sheetsConnectionBadge");
  if (!badge) return;

  const isOnline = navigator.onLine;
  const hasUrl = !!state.config.sheetsUrl;
  const pendingCount = (state.colaSincronizacion || []).length;

  if (!isOnline) {
    badge.className = "px-2 py-0.5 rounded-full text-[10px] font-bold bg-amber-950 text-amber-400 border border-amber-500/40 flex items-center gap-1";
    badge.innerHTML = `<span class="w-1.5 h-1.5 rounded-full bg-amber-400"></span> Modo Offline`;
  } else if (pendingCount > 0) {
    badge.className = "px-2 py-0.5 rounded-full text-[10px] font-bold bg-indigo-950 text-indigo-300 border border-indigo-500/40 flex items-center gap-1";
    badge.innerHTML = `<span class="w-1.5 h-1.5 rounded-full bg-amber-400 animate-ping"></span> ${pendingCount} pendientes`;
  } else if (hasUrl) {
    badge.className = "px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-950 text-emerald-400 border border-emerald-500/40 flex items-center gap-1";
    badge.innerHTML = `<span class="w-1.5 h-1.5 rounded-full bg-emerald-400"></span> Conectado a Sheets`;
  } else {
    badge.className = "px-2 py-0.5 rounded-full text-[10px] font-bold bg-slate-800 text-slate-400 border border-slate-700 flex items-center gap-1";
    badge.innerHTML = `<span class="w-1.5 h-1.5 rounded-full bg-slate-500"></span> Sin URL Sheets`;
  }
}

async function enviarPeticionSheets(accion, datos = {}) {
  if (!state.config.sheetsUrl) throw new Error("No hay URL de Sheets configurada");
  const payload = { action: accion, ...datos };
  
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), 12000);

  try {
    const res = await fetch(state.config.sheetsUrl, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
      signal: controller.signal
    });
    clearTimeout(timeoutId);
    return res;
  } catch (err) {
    clearTimeout(timeoutId);
    throw err;
  }
}

async function sincronizarConSheets(mostrarMensaje = true) {
  // Si hay cambios pendientes en la cola, procesarlos primero
  if (state.colaSincronizacion && state.colaSincronizacion.length > 0) {
    return await procesarColaSincronizacion(mostrarMensaje);
  }

  if (!state.config.sheetsUrl) {
    if (mostrarMensaje) mostrarToast("Configura la URL de Google Sheets en Ajustes", "error");
    return;
  }

  if (!navigator.onLine) {
    if (mostrarMensaje) mostrarToast("Sin conexión a internet. Los datos locales están seguros.", "info");
    return;
  }

  const icon = document.getElementById("syncIcon");
  if (icon) icon.classList.add("animate-spin");

  try {
    const url = `${state.config.sheetsUrl}?action=getTodo`;
    const resp = await fetch(url);
    const json = await resp.json();

    if (json.success && json.data) {
      if (json.data.productos && json.data.productos.length > 0) {
        const mapa = {};
        json.data.productos.forEach(p => { mapa[p.codigo] = p; });
        state.productos = mapa;
        guardarProductosLocal();
      }
      if (json.data.ultimasCompras) {
        state.compras = json.data.ultimasCompras;
        guardarComprasLocal();
      }
      if (json.data.ultimasVentas) {
        state.ventas = json.data.ultimasVentas;
        guardarVentasLocal();
      }
      if (json.data.finanzas) {
        state.movimientosDinero = json.data.finanzas;
        guardarFinanzasLocal();
      }

      renderizarTodo();
      actualizarBadgeConexion();
      if (mostrarMensaje) mostrarToast("¡Sincronizado con Google Sheets! 📊", "success");
    }
  } catch (err) {
    if (mostrarMensaje) mostrarToast("Error al conectar con Google Sheets", "error");
  } finally {
    if (icon) icon.classList.remove("animate-spin");
  }
}

function guardarConfiguracionSheets() {
  const url = document.getElementById("sheetsApiUrl").value.trim();
  state.config.sheetsUrl = url;
  guardarConfiguracionLocal();
  actualizarBadgeConexion();
  mostrarToast("URL guardada", "success");
  if (url) sincronizarConSheets(true);
}

async function probarConexionSheets() {
  const url = document.getElementById("sheetsApiUrl").value.trim();
  if (!url) {
    mostrarToast("Ingresa una URL primero", "error");
    return;
  }
  mostrarToast("Probando conexión...", "info");
  try {
    const resp = await fetch(`${url}?action=ping`);
    const json = await resp.json();
    if (json.success) mostrarToast("¡Conexión Exitosa con Google Sheets! 🎉", "success");
  } catch(e) {
    mostrarToast("Verifica que la Web App tenga acceso público", "error");
  }
}

function guardarPreferenciasNegocio() {
  state.config.nombreNegocio = document.getElementById("businessNameInput").value.trim() || "Libro de Inventario";
  state.config.tipoCambio = Number(document.getElementById("exchangeRateInput").value) || 520;
  state.config.telefonoNegocio = document.getElementById("businessPhoneInput").value.trim();
  guardarConfiguracionLocal();
  renderizarTodo();
  mostrarToast("Ajustes actualizados", "success");
}

function recargarCatalogoSemilla() {
  if (confirm("¿Restaurar el catálogo con los 54 licores originales?")) {
    const mapa = {};
    SEED_PRODUCTS.forEach(p => { mapa[p.codigo] = p; });
    state.productos = mapa;
    state.compras = [...SEED_PURCHASES];
    state.ventas = [];
    state.movimientosDinero = [];
    state.colaSincronizacion = [];
    guardarProductosLocal();
    guardarComprasLocal();
    guardarVentasLocal();
    guardarFinanzasLocal();
    guardarColaLocal();
    renderizarTodo();
    mostrarToast("Catálogo de 54 licores restaurado", "info");
  }
}

function limpiarCacheLocal() {
  if (confirm("¿Borrar todos los datos locales?")) {
    localStorage.clear();
    location.reload();
  }
}

// ==========================================================================
// TOAST NOTIFICACIONES
// ==========================================================================
function mostrarToast(mensaje, tipo = "info") {
  const toast = document.getElementById("toast");
  const box = document.getElementById("toastBox");
  const msgEl = document.getElementById("toastMsg");
  const iconEl = document.getElementById("toastIcon");

  msgEl.textContent = mensaje;

  if (tipo === "success") {
    box.className = "px-4 py-3 rounded-2xl shadow-2xl flex items-center gap-3 text-sm font-bold bg-emerald-600 text-white border-emerald-400";
    iconEl.setAttribute("data-lucide", "check-circle-2");
  } else if (tipo === "error") {
    box.className = "px-4 py-3 rounded-2xl shadow-2xl flex items-center gap-3 text-sm font-bold bg-rose-600 text-white border-rose-400";
    iconEl.setAttribute("data-lucide", "alert-circle");
  } else {
    box.className = "px-4 py-3 rounded-2xl shadow-2xl flex items-center gap-3 text-sm font-bold bg-indigo-600 text-white border-indigo-400";
    iconEl.setAttribute("data-lucide", "info");
  }

  inicializarIconos();
  toast.classList.remove("-translate-y-20", "opacity-0", "pointer-events-none");
  toast.classList.add("translate-y-0", "opacity-100");

  setTimeout(() => {
    toast.classList.add("-translate-y-20", "opacity-0", "pointer-events-none");
    toast.classList.remove("translate-y-0", "opacity-100");
  }, 3000);
}
