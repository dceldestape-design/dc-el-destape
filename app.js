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

// --- Catálogo y Compras Iniciales (Vacíos por defecto - Carga 100% desde Google Sheets) ---
const SEED_PRODUCTS = [];
const SEED_PURCHASES = [];

// --- Estado Global ---
let state = {
  productos: {}, // Mapa por código { "WHI-001": {...} }
  compras: [],
  ventas: [],
  movimientosDinero: [],
  colaSincronizacion: [], // Cola persistente para operaciones sin internet
  carrito: [],
  listaCompraActual: [], // Carrito de compras multi-producto
  // --- Maestro de Clientes, Pedidos y Cuentas ---
  clientes: {}, // Mapa por ID { "CLI-xxx": { id, nombre, telefono, puntos, fechaRegistro, ultimaVenta } }
  clienteSeleccionado: null, // cliente activo en la venta actual
  descuentoPuntosAplicado: 0, // descuento en CRC aplicado de puntos en la venta actual
  pedidos: [], // Lista de pedidos / encargos de clientes
  cuentas: [], // Lista de cuentas pendientes (Por Cobrar / Por Pagar)
  filtroCuentas: "",
  filtroTipoCuenta: "Por Cobrar", // "Por Cobrar" | "Por Pagar" | "todos"
  modoPOS: "venta", // "venta" | "pedido"
  config: {
    sheetsUrl: "",
    tipoCambio: 520,
    nombreNegocio: "DC El Destape",
    telefonoNegocio: "+506 8992-7936",
    // Configuración del sistema de puntos de fidelización (Predefinidos)
    puntosRazonCRC: 20,        // cada ₡20 = 1 punto
    puntosValorCRC: 1,         // 1 punto = ₡1 de descuento
    puntosMinimosCanje: 4000   // mínimo 4000 puntos para poder canjear
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
  ultimaVentaCompletada: null,
  ultimoPedidoCompletado: null,
  // Filtro para la vista de clientes
  filtroClientes: ""
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
    mostrarToast("Modo Offline activo (sin internet). Todo se guarda en tu teléfono 💾", "info");
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
  try {
    if (window.lucide && typeof window.lucide.createIcons === "function") {
      window.lucide.createIcons();
    } else {
      setTimeout(() => {
        if (window.lucide && typeof window.lucide.createIcons === "function") {
          window.lucide.createIcons();
        }
      }, 100);
    }
  } catch (e) {
    console.warn("Lucide icons init:", e);
  }
}

// ==========================================================================
// GESTIÓN DE VENDEDORES (CARLOS Y DANIEL)
// ==========================================================================
function comprobarLoginVendedor() {
  // Siempre solicitar selección obligatoria de vendedor al ingresar a la app
  abrirModalSeleccionVendedor(true);
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
  }
  localStorage.setItem("inv_vista_vendedor", state.vistaVendedor);

  actualizarUIVendedor();
  cerrarModalLoginVendedor();
  renderizarTodo();
  mostrarToast(`¡Bienvenido! Perfil activo: ${vendedor === "Consolidado" ? "Consolidado (Total)" : vendedor} 👤`, "success");
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
    try { state.productos = JSON.parse(prods); } catch(e) { state.productos = {}; }
  } else {
    state.productos = {};
  }

  const comps = localStorage.getItem("inv_compras_v2");
  if (comps) {
    try { state.compras = JSON.parse(comps); } catch(e) { state.compras = []; }
  } else {
    state.compras = [];
  }

  const vts = localStorage.getItem("inv_ventas_v2");
  if (vts) state.ventas = JSON.parse(vts);

  const peds = localStorage.getItem("inv_pedidos_v2");
  if (peds) {
    try { state.pedidos = JSON.parse(peds); } catch(e) { state.pedidos = []; }
  }

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

  const cli = localStorage.getItem("inv_clientes_v2");
  if (cli) {
    try { state.clientes = JSON.parse(cli); } catch(e) { state.clientes = {}; }
  }

  const ctas = localStorage.getItem("inv_cuentas_v2");
  if (ctas) {
    try { state.cuentas = JSON.parse(ctas); } catch(e) { state.cuentas = []; }
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
function guardarPedidosLocal() {
  localStorage.setItem("inv_pedidos_v2", JSON.stringify(state.pedidos));
}
function guardarCuentasLocal() {
  localStorage.setItem("inv_cuentas_v2", JSON.stringify(state.cuentas));
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
function guardarClientesLocal() {
  localStorage.setItem("inv_clientes_v2", JSON.stringify(state.clientes));
}

// --- Cálculos de Stock Separado por Vendedor y Consolidado ---
function calcularStockDetalladoPorCodigo() {
  const detalle = {};
  
  // 1. Inicializar mapa con todos los productos del catálogo
  Object.values(state.productos).forEach(p => {
    const cod = String(p.codigo || "").trim().toUpperCase();
    if (!cod) return;
    const init = Number(p.stockInicial || 0);
    detalle[cod] = {
      Carlos: init,
      Daniel: 0,
      total: init
    };
  });

  // 2. Sumar compras por vendedor
  state.compras.forEach(c => {
    const cod = String(c.codigo || "").trim().toUpperCase();
    if (!cod) return;
    if (!detalle[cod]) detalle[cod] = { Carlos: 0, Daniel: 0, total: 0 };
    
    // Si la compra tiene items[] anidados
    if (c.items && Array.isArray(c.items) && c.items.length > 0) {
      c.items.forEach(ci => {
        const ciCod = String(ci.codigo || cod).trim().toUpperCase();
        if (!detalle[ciCod]) detalle[ciCod] = { Carlos: 0, Daniel: 0, total: 0 };
        const cant = Number(ci.cantidad || 0);
        const vend = String(ci.vendedor || c.vendedor || "Carlos").trim();
        if (vend === "Daniel") {
          detalle[ciCod].Daniel += cant;
        } else {
          detalle[ciCod].Carlos += cant;
        }
        detalle[ciCod].total += cant;
      });
    } else {
      const cant = Number(c.cantidad || 0);
      const vend = String(c.vendedor || "Carlos").trim();
      if (vend === "Daniel") {
        detalle[cod].Daniel += cant;
      } else {
        detalle[cod].Carlos += cant;
      }
      detalle[cod].total += cant;
    }
  });

  // 3. Restar ventas por vendedor
  state.ventas.forEach(v => {
    const vend = String(v.vendedor || "Carlos").trim();
    let items = [];
    if (v.items && Array.isArray(v.items)) {
      items = v.items;
    } else if (typeof v.items === "string") {
      try { items = JSON.parse(v.items); } catch(e) { items = []; }
    } else if (v.codigo) {
      items = [{ codigo: v.codigo, cantidad: v.cantidad }];
    }

    items.forEach(i => {
      const cod = String(i.codigo || "").trim().toUpperCase();
      if (!cod) return;
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
  const vistas = ["dashboard", "inventario", "ventas", "compras", "finanzas", "configuracion", "clientes", "cuentas"];
  
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
  if (vista === "clientes") renderizarClientes();
  if (vista === "cuentas") renderizarCuentas();
  if (vista === "configuracion") {
    cargarConfigPuntosUI();
    const surl = document.getElementById("sheetsApiUrl");
    if (surl && state.config.sheetsApiUrl) surl.value = state.config.sheetsApiUrl;
    const bname = document.getElementById("businessNameInput");
    if (bname && state.config.nombreNegocio) bname.value = state.config.nombreNegocio;
    const tc = document.getElementById("exchangeRateInput");
    if (tc && state.config.tipoCambio) tc.value = state.config.tipoCambio;
    const ph = document.getElementById("businessPhoneInput");
    if (ph && state.config.telefonoNegocio) ph.value = state.config.telefonoNegocio;
  }

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
  renderizarClientes();
  renderizarCuentas();
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

  // --- Resumen de Cuentas Pendientes en Dashboard (Por Cobrar y Por Pagar) ---
  let totCobrarCRC = 0;
  let totPagarCRC = 0;
  (state.cuentas || []).forEach(cta => {
    const saldo = Number(cta.saldoPendienteCRC || 0);
    if ((cta.estado || "Pendiente") !== "Pagado" && saldo > 0) {
      if (cta.tipo === "Por Cobrar") {
        totCobrarCRC += saldo;
      } else {
        totPagarCRC += saldo;
      }
    }
  });

  const dashCobrarEl = document.getElementById("dashCobrarCRC");
  const dashPagarEl = document.getElementById("dashPagarCRC");
  if (dashCobrarEl) dashCobrarEl.textContent = fmtCRC(totCobrarCRC);
  if (dashPagarEl) dashPagarEl.textContent = fmtCRC(totPagarCRC);

  // --- Pedidos Pendientes de Clientes (Consolidado para Proveedor) ---
  renderizarConsolidadoPedidosDashboard();

  // Últimas ventas
  const recentCont = document.getElementById("dashRecentSales");
  if (state.ventas.length === 0) {
    recentCont.innerHTML = `<div class="text-center py-5 text-slate-500 text-xs">No hay ventas registradas aún.</div>`;
  } else {
    const ultimas = state.ventas.slice(0, 10);
    recentCont.innerHTML = ultimas.map((v, idx) => {
      const fecha = v.fecha ? new Date(v.fecha).toLocaleDateString([], { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' }) : "";
      const vend = v.vendedor || "Carlos";
      const vendColor = vend === "Daniel" ? "text-violet-400 bg-violet-950/60 border-violet-500/30" : "text-blue-400 bg-blue-950/60 border-blue-500/30";
      const totCRC = Number(v.totalCRC || 0);
      const totUSD = Number(v.totalUSD || 0);
      const esPagoLuego = (v.metodoPago || "").toLowerCase().includes("luego") || (v.metodoPago || "").toLowerCase().includes("crédito");
      const vUid = v.id ? `${v.id}_${v.codigo || ''}_${idx}` : `VTA_ROW_${idx}`;

      // Buscar si existe una cuenta asociada a esta venta
      const cuentaAsociada = (state.cuentas || []).find(cta => 
        cta.tipo === "Por Cobrar" && (
          cta.referenciaId === v.id ||
          cta.referenciaId === vUid ||
          (v.id && cta.referenciaId && cta.referenciaId.startsWith(v.id)) ||
          (cta.entidad && v.cliente && cta.entidad.toLowerCase() === v.cliente.toLowerCase() && Math.abs(Number(cta.montoTotalCRC) - totCRC) < 1)
        )
      );

      // Determinar estado real de la cuenta
      let estadoBadgeHtml = '';
      if (cuentaAsociada) {
        const est = cuentaAsociada.estado || "Pendiente";
        if (est === "Pagado") {
          estadoBadgeHtml = `<span class="text-[9px] font-bold px-1.5 py-0.2 rounded bg-emerald-950/80 border border-emerald-500/40 text-emerald-300">✅ Liquidada</span>`;
        } else if (est === "Parcial") {
          estadoBadgeHtml = `<span class="text-[9px] font-bold px-1.5 py-0.2 rounded bg-amber-950/80 border border-amber-500/40 text-amber-300">⏳ Abono Parcial</span>`;
        } else {
          estadoBadgeHtml = `<span class="text-[9px] font-bold px-1.5 py-0.2 rounded bg-rose-950/80 border border-rose-500/40 text-rose-300">🕒 Por Cobrar</span>`;
        }
      } else if (esPagoLuego) {
        estadoBadgeHtml = `<span class="text-[9px] font-bold px-1.5 py-0.2 rounded bg-amber-950/60 border border-amber-500/40 text-amber-300">🕒 Pago Luego</span>`;
      }

      return `
        <div class="py-2.5 flex items-center justify-between border-b border-slate-800/60 last:border-0 gap-2">
          <div class="min-w-0 flex-1">
            <div class="flex items-center gap-1.5 mb-0.5 flex-wrap">
              <span class="text-[9px] font-bold px-1.5 py-0.2 rounded border ${vendColor}">👤 ${vend}</span>
              <span class="text-xs font-bold text-white truncate max-w-[150px]">${v.cliente || "Venta General"}</span>
              ${estadoBadgeHtml}
            </div>
            <div class="text-[11px] text-slate-400 font-mono">
              ${fecha} • ${v.nombre ? `${v.nombre} (${v.cantidad || 1}x)` : `${v.items ? v.items.length : 1} prod(s)`} <span class="text-[10px] text-slate-500">(${v.metodoPago || "Efectivo"})</span>
            </div>
          </div>
          <div class="text-right font-mono shrink-0 space-y-0.5">
            <div class="text-xs font-black text-emerald-400">${fmtCRC(totCRC)}</div>
            <div class="text-[10px] text-slate-400">${fmtUSD(totUSD)}</div>
            <div class="pt-0.5">
              ${!cuentaAsociada && !esPagoLuego ? `
                <button onclick="pasarVentaIndividualACuentasPorCobrar(${idx})" title="Pasar a Cuentas por Cobrar" class="text-[9.5px] font-bold px-2 py-0.5 rounded bg-amber-950/60 hover:bg-amber-900 border border-amber-500/40 text-amber-300 active:scale-95 transition-all">
                  + Cta Cobrar
                </button>
              ` : (cuentaAsociada && cuentaAsociada.estado === "Pagado" ? `
                <span class="text-[9px] font-bold px-1.5 py-0.2 rounded bg-emerald-950/60 border border-emerald-500/30 text-emerald-300">Cobrado</span>
              ` : `
                <button onclick="cambiarVista('cuentas')" title="Ver en Cuentas por Cobrar" class="text-[9px] font-bold px-1.5 py-0.2 rounded bg-indigo-950/60 hover:bg-indigo-900 border border-indigo-500/40 text-indigo-300 active:scale-95 transition-all">
                  Ver Cuenta
                </button>
              `)}
            </div>
          </div>
        </div>
      `;
    }).join("");
  }
}

// ==========================================================================
// RENDERIZAR CONSOLIDADO DE PEDIDOS EN EL DASHBOARD
// ==========================================================================
function renderizarConsolidadoPedidosDashboard() {
  const container = document.getElementById("dashPedidosContainer");
  const badge = document.getElementById("dashPedidosBadge");
  const consolidadoLista = document.getElementById("dashConsolidadoLista");
  const pedidosList = document.getElementById("dashPedidosList");
  if (!container || !consolidadoLista || !pedidosList) return;

  const pedidosPendientes = (state.pedidos || []).filter(p => p.estado === "pendiente" || !p.estado);

  if (pedidosPendientes.length === 0) {
    badge.textContent = "0 pendientes";
    badge.className = "text-[11px] font-bold font-mono px-2 py-0.5 rounded-full bg-slate-800 text-slate-400 border border-slate-700";
    consolidadoLista.innerHTML = `
      <div class="py-3 text-center text-xs text-slate-400 font-sans">
        ✨ No hay pedidos pendientes de clientes. Todo al día.
      </div>
    `;
    pedidosList.innerHTML = `
      <div class="py-2 text-center text-[11px] text-slate-500 font-sans">
        Usa el modo <b>"Encargo / Pedido"</b> en el TPV para registrar solicitudes.
      </div>
    `;
    return;
  }

  badge.textContent = `${pedidosPendientes.length} pendiente(s)`;
  badge.className = "text-[11px] font-bold font-mono px-2 py-0.5 rounded-full bg-amber-500/20 text-amber-300 border border-amber-500/30";

  // 1. Agrupar productos solicitados (Consolidado de botellas)
  const mapaConsolidado = {};
  pedidosPendientes.forEach(ped => {
    (ped.items || []).forEach(it => {
      const cod = it.codigo || it.nombre;
      if (!mapaConsolidado[cod]) {
        mapaConsolidado[cod] = {
          codigo: it.codigo,
          nombre: it.nombre,
          cantidad: 0
        };
      }
      mapaConsolidado[cod].cantidad += Number(it.cantidad || 1);
    });
  });

  const consolidadoArray = Object.values(mapaConsolidado);
  consolidadoLista.innerHTML = consolidadoArray.map(item => `
    <div class="py-1.5 flex items-center justify-between">
      <div class="flex items-center gap-2 min-w-0">
        <span class="font-black text-amber-400 bg-amber-950/80 px-2 py-0.5 rounded-lg border border-amber-500/40 text-xs shrink-0">${item.cantidad}x</span>
        <span class="text-xs font-bold text-white truncate">${item.nombre}</span>
      </div>
      <span class="text-[10px] text-slate-400 font-mono shrink-0">${item.codigo}</span>
    </div>
  `).join("");

  // 2. Renderizar lista detallada de pedidos por cliente
  pedidosList.innerHTML = pedidosPendientes.map(ped => {
    const fecha = ped.fecha ? new Date(ped.fecha).toLocaleDateString([], { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' }) : "";
    const itemsResumen = (ped.items || []).map(i => `${i.cantidad}x ${i.nombre}`).join(", ");
    let totalBotellasPed = 0;
    (ped.items || []).forEach(i => totalBotellasPed += Number(i.cantidad || 1));

    return `
      <div class="p-2.5 bg-slate-950/70 border border-slate-800 rounded-xl flex items-center justify-between gap-2.5">
        <div class="min-w-0 flex-1">
          <div class="flex items-center gap-1.5 mb-0.5">
            <span class="text-xs font-bold text-white truncate">${ped.cliente || 'Cliente General'}</span>
            ${ped.clienteTelefono ? `<span class="text-[10px] text-amber-400/90 font-mono">(${ped.clienteTelefono})</span>` : ''}
            <span class="text-[9px] text-slate-400 bg-slate-900 px-1.5 py-0.2 rounded border border-slate-800 font-mono">${fecha}</span>
          </div>
          <p class="text-[11px] text-slate-300 font-mono truncate leading-tight">${itemsResumen}</p>
          <div class="text-[10px] text-amber-400 font-mono mt-0.5">
            Total botellas: <b class="text-white">${totalBotellasPed} unids</b> • Anotó: ${ped.vendedor || 'Carlos'}
          </div>
        </div>

        <div class="flex items-center gap-1 shrink-0">
          <button type="button" onclick="marcarPedidoComprado('${ped.id}')" title="Marcar como ya comprado al proveedor" class="px-2.5 py-1.5 bg-emerald-600/20 hover:bg-emerald-600/30 border border-emerald-500/40 text-emerald-300 font-bold text-[11px] rounded-lg active:scale-95 transition-all flex items-center gap-1">
            <i data-lucide="check" class="w-3.5 h-3.5"></i>
            <span>Comprado</span>
          </button>
          <button type="button" onclick="cancelarPedido('${ped.id}')" title="Cancelar pedido" class="p-1.5 text-slate-500 hover:text-rose-400 rounded-lg active:scale-95 transition-all">
            <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
          </button>
        </div>
      </div>
    `;
  }).join("");

  inicializarIconos();
}

// ============================================================
// COPIAR LISTA DE PEDIDO AL PROVEEDOR
// ============================================================
function copiarPedidoProveedorTexto() {
  const pedidosPendientes = (state.pedidos || []).filter(p => p.estado === "pendiente" || !p.estado);
  if (pedidosPendientes.length === 0) {
    mostrarToast("No hay pedidos pendientes para solicitar", "info");
    return;
  }

  const mapaConsolidado = {};
  pedidosPendientes.forEach(ped => {
    (ped.items || []).forEach(it => {
      const cod = it.codigo || it.nombre;
      if (!mapaConsolidado[cod]) {
        mapaConsolidado[cod] = {
          codigo: it.codigo,
          nombre: it.nombre,
          cantidad: 0
        };
      }
      mapaConsolidado[cod].cantidad += Number(it.cantidad || 1);
    });
  });

  const negocio = state.config.nombreNegocio || "DC EL DESTAPE";
  const fecha = new Date().toLocaleDateString([], { dateStyle: 'medium' });

  let texto = `🍷 *PEDIDO CONSOLIDADO PARA PROVEEDOR* 🍷\n`;
  texto += `🏢 *${negocio.toUpperCase()}*\n`;
  texto += `📅 *Fecha:* ${fecha}\n`;
  texto += `----------------------------------------\n`;
  texto += `📦 *DETALLE DE BOTELLAS SOLICITADAS:*\n`;

  let totalBotellas = 0;
  Object.values(mapaConsolidado).forEach(item => {
    texto += `• *${item.cantidad}x* ${item.nombre} (Cod: ${item.codigo})\n`;
    totalBotellas += item.cantidad;
  });

  texto += `----------------------------------------\n`;
  texto += `📊 *TOTAL A ENCARGAR:* ${totalBotellas} unidades\n`;
  texto += `📌 *Pedidos de clientes atendidos:* ${pedidosPendientes.length}`;

  navigator.clipboard.writeText(texto).then(() => {
    mostrarToast(`📋 Lista de ${totalBotellas} botellas copiada al portapapeles`, "success");
  }).catch(() => {
    // Fallback prompt
    prompt("Copia el texto del pedido para enviar al proveedor:", texto);
  });
}

function marcarPedidoComprado(idPedido) {
  const ped = (state.pedidos || []).find(p => p.id === idPedido);
  if (!ped) return;

  ped.estado = "comprado";
  ped.fechaComprado = new Date().toISOString();
  guardarPedidosLocal();
  renderizarDashboard();
  mostrarToast(`Pedido de ${ped.cliente || 'cliente'} marcado como comprado ✅`, "success");

  // Encolar y sincronizar con Google Sheets
  encolarAccionSincronizacion("marcarPedidoComprado", { id: idPedido });
}

function cancelarPedido(idPedido) {
  if (!confirm("¿Deseas eliminar este encargo?")) return;
  state.pedidos = (state.pedidos || []).filter(p => p.id !== idPedido);
  guardarPedidosLocal();
  renderizarDashboard();
  mostrarToast("Pedido eliminado", "info");

  // Encolar y sincronizar con Google Sheets
  encolarAccionSincronizacion("eliminarPedido", { id: idPedido });
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
    const existeEnOtrasCategorias = filtroTexto ? todosProds.filter(p => 
      p.nombre.toLowerCase().includes(filtroTexto) || 
      p.codigo.toLowerCase().includes(filtroTexto)
    ) : [];

    let sugerenciaHtml = "";
    if (existeEnOtrasCategorias.length > 0 && state.categoriaSeleccionada !== "Todas") {
      const cats = Array.from(new Set(existeEnOtrasCategorias.map(x => x.categoria || "Otras"))).join(", ");
      sugerenciaHtml = `
        <div class="p-3 bg-amber-950/40 border border-amber-500/40 rounded-2xl space-y-2 mt-2">
          <p class="text-xs text-amber-200">🔍 Se encontraron <b>${existeEnOtrasCategorias.length}</b> resultado(s) en <b>${cats}</b>:</p>
          <button type="button" onclick="filtrarCategoria('Todas')" class="px-3 py-1.5 bg-amber-500 hover:bg-amber-400 text-slate-950 font-bold text-xs rounded-xl active:scale-95 transition-all shadow-md">
            Ver resultados en "Todas las categorías"
          </button>
        </div>
      `;
    }

    contenedor.innerHTML = `
      <div class="text-center py-10 text-slate-500 space-y-3">
        <i data-lucide="package-search" class="w-10 h-10 mx-auto text-slate-600"></i>
        <div>
          <p class="text-xs font-semibold text-slate-300">No se encontraron licores ${filtroTexto ? `para "<b>${filtroTexto}</b>"` : ''} en la categoría <span class="text-indigo-400 font-bold">${state.categoriaSeleccionada}</span>.</p>
        </div>
        ${sugerenciaHtml}
        <div class="pt-2 flex items-center justify-center gap-2">
          <button type="button" onclick="limpiarFiltrosInventario()" class="px-3.5 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-bold rounded-xl border border-slate-700 active:scale-95 transition-all">
            Limpiar Filtros
          </button>
          <button type="button" onclick="abrirModalProducto()" class="px-4 py-2 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl text-xs font-bold shadow-lg shadow-indigo-600/30 active:scale-95 transition-all">
            + Agregar Licor
          </button>
        </div>
      </div>
    `;
    inicializarIconos();
    renderizarCategoriasPills();
    return;
  }

  const detailedMap = calcularStockDetalladoPorCodigo();

  contenedor.innerHTML = lista.map(p => {
    const det = detailedMap[p.codigo] || { Carlos: 0, Daniel: 0, total: 0 };
    const stockVisual = state.vistaVendedor === "Carlos" 
      ? det.Carlos 
      : (state.vistaVendedor === "Daniel" ? det.Daniel : det.total);

    const isLow = stockVisual <= (p.stockMinimo || 2) && stockVisual > 0;
    const isOut = stockVisual <= 0;

    let stockBadgeClass = "bg-emerald-950/80 text-emerald-300 border-emerald-500/40";
    let stockStatusText = `${stockVisual} unids`;
    if (isOut) {
      stockBadgeClass = "bg-rose-950/80 text-rose-300 border-rose-500/40";
      stockStatusText = "Agotado (0)";
    } else if (isLow) {
      stockBadgeClass = "bg-amber-950/80 text-amber-300 border-amber-500/40";
      stockStatusText = `Bajo: ${stockVisual}`;
    }

    const imgFormatted = formatearUrlImagen(p.imagenUrl);
    const hasImg = !!imgFormatted;

    return `
      <div class="bg-gradient-to-br from-slate-900 via-slate-900/95 to-slate-950 border border-slate-800 hover:border-slate-700/80 rounded-2xl p-3 shadow-lg space-y-2.5 transition-all">
        <!-- Header: Code, Category & Stock Badge -->
        <div class="flex items-center justify-between gap-2">
          <div class="flex items-center gap-1.5 min-w-0">
            <span class="font-mono text-[10px] font-bold text-slate-400 bg-slate-800 px-1.5 py-0.5 rounded border border-slate-700/60 shrink-0">${p.codigo}</span>
            <span class="text-[10px] font-semibold text-slate-400 truncate">${p.categoria || 'Licor'}</span>
          </div>
          <span class="text-[10px] font-bold px-2 py-0.5 rounded-full border ${stockBadgeClass} shrink-0">
            ${stockStatusText}
          </span>
        </div>

        <!-- Body: Photo + Name & Breakdown -->
        <div class="flex gap-3 items-center">
          <div class="relative w-20 h-20 rounded-2xl bg-slate-800 border border-slate-700/80 flex items-center justify-center shrink-0 overflow-hidden shadow-md cursor-pointer group foto-producto-btn" data-url="${imgFormatted || ''}" data-nombre="${p.nombre.replace(/"/g, '&quot;')}" title="Toca para ver foto completa">
            ${hasImg ? `
              <img src="${imgFormatted}" alt="${p.nombre}" loading="lazy" class="w-full h-full object-cover group-hover:scale-105 transition-transform" 
                onerror="this.onerror=null; this.style.display='none'; this.nextElementSibling.classList.remove('hidden'); this.nextElementSibling.classList.add('flex');">
              <div class="hidden flex-col items-center justify-center text-slate-500 text-[10px] w-full h-full">
                <i data-lucide="wine" class="w-7 h-7 text-slate-600"></i>
              </div>
            ` : `
              <div class="flex flex-col items-center justify-center text-slate-500 text-[10px]">
                <i data-lucide="wine" class="w-7 h-7 text-slate-600"></i>
                <span class="text-[8px] text-slate-500 mt-0.5">Sin foto</span>
              </div>
            `}
            <div class="absolute bottom-1 right-1 bg-black/60 backdrop-blur-sm rounded-md p-0.5 text-white/80 opacity-70 group-hover:opacity-100 transition-opacity">
              <i data-lucide="zoom-in" class="w-3 h-3"></i>
            </div>
          </div>

          <div class="min-w-0 flex-1">
            <h4 class="text-xs font-black text-white leading-snug cursor-pointer hover:text-indigo-300 transition-colors line-clamp-2" onclick="abrirModalProducto('${p.codigo}')">
              ${p.nombre}
            </h4>
            <div class="flex items-center gap-2 mt-1.5 text-[10px] font-mono text-slate-400">
              <span class="text-blue-300">C: <b>${det.Carlos}</b></span>
              <span class="text-slate-600">|</span>
              <span class="text-violet-300">D: <b>${det.Daniel}</b></span>
              <span class="text-slate-600">|</span>
              <span class="text-amber-300">Tot: <b>${det.total}</b></span>
            </div>
          </div>
        </div>

        <!-- Pricing & Actions Row -->
        <div class="flex items-center justify-between pt-2 border-t border-slate-800/80 text-xs font-mono">
          <div>
            <div class="text-[10px] text-slate-400 font-sans">Precio Venta</div>
            <div class="font-black text-emerald-400 text-sm">${fmtCRC(p.precioVentaCRC)}</div>
            <div class="text-[10px] text-slate-400 font-normal">${fmtUSD(p.precioVentaUSD)} USD</div>
          </div>

          <div class="flex items-center gap-1">
            <button onclick="abrirModalProducto('${p.codigo}')" class="p-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl active:scale-95 transition-all" title="Editar Producto">
              <i data-lucide="edit-2" class="w-3.5 h-3.5"></i>
            </button>
            <button onclick="agregarAlCarritoPorCodigo('${p.codigo}')" class="py-1.5 px-3 bg-indigo-600 hover:bg-indigo-500 text-white font-sans font-bold text-[11px] rounded-xl flex items-center gap-1 active:scale-95 shadow-md shadow-indigo-600/20 transition-all">
              <i data-lucide="plus" class="w-3.5 h-3.5"></i>
              Vender
            </button>
          </div>
        </div>
      </div>
    `;
  }).join("");

  inicializarIconos();
  renderizarCategoriasPills();
}

function limpiarFiltrosInventario() {
  const s = document.getElementById("searchInventory");
  if (s) s.value = "";
  state.categoriaSeleccionada = "Todas";
  state.filtroEstadoStock = "todos";
  renderizarTodo();
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
// MÓDULO: MAESTRO DE CLIENTES Y PUNTOS DE FIDELIZACIÓN
// ==========================================================================

// --- CRUD Clientes ---
function guardarCliente(obj) {
  // obj: { nombre, telefono, puntos?, fechaRegistro? }
  const tel = (obj.telefono || "").trim().replace(/\s+/g, "");
  if (!obj.nombre || !tel) { mostrarToast("Nombre y teléfono son requeridos.", "error"); return null; }

  // Buscar si ya existe por teléfono
  const existente = Object.values(state.clientes).find(c => c.telefono === tel);
  const id = existente ? existente.id : "CLI-" + Date.now().toString().slice(-8);
  const ahora = new Date().toISOString();

  const clienteObj = {
    id,
    nombre: obj.nombre.trim(),
    telefono: tel,
    puntos: existente ? (obj.puntos !== undefined ? obj.puntos : existente.puntos) : (obj.puntos || 0),
    fechaRegistro: existente ? existente.fechaRegistro : (obj.fechaRegistro || ahora),
    ultimaVenta: existente ? existente.ultimaVenta : null
  };

  state.clientes[id] = clienteObj;
  guardarClientesLocal();
  encolarAccionSincronizacion("guardarCliente", { cliente: clienteObj });
  return clienteObj;
}

function buscarClientePorTelefono(tel) {
  const t = (tel || "").trim().replace(/\s+/g, "");
  return Object.values(state.clientes).find(c => c.telefono === t) || null;
}

function buscarClientesPorQuery(q) {
  if (!q) return Object.values(state.clientes);
  const ql = q.toLowerCase();
  return Object.values(state.clientes).filter(c =>
    c.nombre.toLowerCase().includes(ql) || c.telefono.includes(ql)
  );
}

function actualizarPuntosCliente(id, delta) {
  if (!state.clientes[id]) return;
  state.clientes[id].puntos = Math.max(0, (state.clientes[id].puntos || 0) + delta);
  guardarClientesLocal();
  encolarAccionSincronizacion("actualizarPuntos", { telefono: state.clientes[id].telefono, puntos: state.clientes[id].puntos });
}

// --- Selección de cliente en el POS ---
function seleccionarCliente(id) {
  state.clienteSeleccionado = state.clientes[id] || null;
  state.descuentoPuntosAplicado = 0;

  const input = document.getElementById("posClienteInput");
  if (input && state.clienteSeleccionado) {
    input.value = state.clienteSeleccionado.nombre;
  }

  const dropdown = document.getElementById("clienteDropdown");
  if (dropdown) {
    dropdown.classList.add("hidden");
    dropdown.innerHTML = "";
  }

  renderizarPanelCliente();
  renderizarCarrito();
}
function deseleccionarCliente() {
  state.clienteSeleccionado = null;
  state.descuentoPuntosAplicado = 0;
  const input = document.getElementById("posClienteInput");
  if (input) input.value = "";
  const dropdown = document.getElementById("clienteDropdown");
  if (dropdown) dropdown.classList.add("hidden");
  renderizarPanelCliente();
  renderizarCarrito();
}

function renderizarPanelCliente() {
  const panel = document.getElementById("panelPuntosCliente");
  if (!panel) return;

  const cli = state.clienteSeleccionado;
  if (!cli) {
    panel.classList.add("hidden");
    panel.innerHTML = "";
    return;
  }

  const config = state.config;
  const puntos = cli.puntos || 0;
  const valorCanje = Math.floor(puntos * (config.puntosValorCRC || 1));
  const puedesCanjear = puntos >= (config.puntosMinimosCanje || 100);
  const yaCanjeado = state.descuentoPuntosAplicado > 0;

  panel.classList.remove("hidden");
  panel.innerHTML = `
    <div class="p-3 bg-amber-950/40 border border-amber-500/40 rounded-2xl space-y-2">
      <div class="flex items-center justify-between">
        <div class="flex items-center gap-2">
          <span class="w-6 h-6 rounded-full bg-amber-500/20 border border-amber-500/40 flex items-center justify-center text-xs">👤</span>
          <div>
            <div class="font-bold text-white text-xs">${cli.nombre}</div>
            <div class="text-[10px] text-slate-400 font-mono">${cli.telefono}</div>
          </div>
        </div>
        <button onclick="deseleccionarCliente()" class="text-slate-500 hover:text-rose-400 text-xs">✕</button>
      </div>
      <div class="flex items-center justify-between text-xs">
        <span class="text-slate-300">Puntos acumulados:</span>
        <span class="font-black text-amber-300 font-mono">${puntos.toLocaleString()} pts <span class="text-amber-500/70 font-normal">(vale ${fmtCRC(valorCanje)})</span></span>
      </div>
      ${!yaCanjeado && puedesCanjear ? `
      <button onclick="abrirModalCanjeoPuntos()" class="w-full py-2 bg-amber-500/20 hover:bg-amber-500/30 border border-amber-500/40 text-amber-300 font-bold text-xs rounded-xl active:scale-95 transition-all flex items-center justify-center gap-1.5">
        🎁 Canjear puntos como descuento
      </button>
      ` : yaCanjeado ? `
      <div class="flex items-center justify-between py-1.5 px-2 bg-emerald-950/40 border border-emerald-500/30 rounded-xl">
        <span class="text-xs text-emerald-300 font-bold">✓ Descuento aplicado:</span>
        <span class="text-xs font-black text-emerald-400 font-mono">-${fmtCRC(state.descuentoPuntosAplicado)}</span>
        <button onclick="quitarCanjeoPuntos()" class="text-[10px] text-rose-400 hover:text-rose-300 ml-2">Quitar</button>
      </div>
      ` : `
      <div class="text-[11px] text-slate-500 text-center">Mínimo ${config.puntosMinimosCanje} puntos para canjear.</div>
      `}
    </div>
  `;
  inicializarIconos();
}
function abrirModalCanjeoPuntos() {
  const cli = state.clienteSeleccionado;
  if (!cli) return;
  const config = state.config;
  const puntos = cli.puntos || 0;

  // Calcular total actual del carrito (sin descuento)
  let totalCarritoCRC = 0;
  state.carrito.forEach(i => totalCarritoCRC += i.cantidad * i.precioVentaCRC);

  const maxDescuento = Math.min(puntos * (config.puntosValorCRC || 1), totalCarritoCRC);

  const modal = document.getElementById("modalCanjeoPuntos");
  if (!modal) return;
  document.getElementById("canjePuntosDisponibles").textContent = puntos.toLocaleString();
  document.getElementById("canjePuntosValor").textContent = fmtCRC(puntos * (config.puntosValorCRC || 1));
  document.getElementById("canjePuntosMax").textContent = fmtCRC(maxDescuento);

  const slider = document.getElementById("canjePuntosSlider");
  const maxPuntos = Math.floor(maxDescuento / (config.puntosValorCRC || 1));
  slider.max = maxPuntos;
  slider.value = maxPuntos;
  actualizarSliderCanje();

  modal.classList.remove("hidden");
  modal.classList.add("flex");
}

function actualizarSliderCanje() {
  const slider = document.getElementById("canjePuntosSlider");
  const config = state.config;
  const puntosAUsar = Number(slider.value) || 0;
  const descuento = puntosAUsar * (config.puntosValorCRC || 1);
  const el = document.getElementById("canjePuntosResumen");
  if (el) el.textContent = `${puntosAUsar.toLocaleString()} puntos = -${fmtCRC(descuento)} de descuento`;
}

function confirmarCanjeoPuntos() {
  const slider = document.getElementById("canjePuntosSlider");
  const puntosAUsar = Number(slider.value) || 0;
  if (puntosAUsar <= 0) { mostrarToast("Seleccioná al menos 1 punto.", "error"); return; }

  state.descuentoPuntosAplicado = puntosAUsar * (state.config.puntosValorCRC || 1);
  cerrarModalCanjeoPuntos();
  renderizarPanelCliente();
  renderizarCarrito();
  mostrarToast(`Descuento de ${fmtCRC(state.descuentoPuntosAplicado)} aplicado 🎁`, "success");
}

function quitarCanjeoPuntos() {
  state.descuentoPuntosAplicado = 0;
  renderizarPanelCliente();
  renderizarCarrito();
  mostrarToast("Descuento de puntos eliminado.", "info");
}

function cerrarModalCanjeoPuntos() {
  const modal = document.getElementById("modalCanjeoPuntos");
  if (modal) { modal.classList.add("hidden"); modal.classList.remove("flex"); }
}

// --- Búsqueda de cliente en el POS ---
function onClienteInputChange() {
  const q = (document.getElementById("posClienteInput").value || "").trim();
  const dropdown = document.getElementById("clienteDropdown");

  if (!q || q.length < 1) {
    dropdown.classList.add("hidden");
    return;
  }

  const resultados = buscarClientesPorQuery(q).slice(0, 6);
  if (resultados.length === 0) {
    dropdown.innerHTML = `
      <div class="p-2 text-xs text-slate-400">No encontrado.
        <button onclick="abrirModalNuevoCliente()" class="text-indigo-400 font-bold hover:underline ml-1">➕ Crear cliente</button>
      </div>`;
    dropdown.classList.remove("hidden");
    return;
  }

  dropdown.innerHTML = resultados.map(c => `
    <div onclick="seleccionarCliente('${c.id}')" class="px-3 py-2.5 hover:bg-slate-700 cursor-pointer flex items-center justify-between gap-2">
      <div>
        <div class="text-xs font-bold text-white">${c.nombre}</div>
        <div class="text-[11px] text-slate-400 font-mono">${c.telefono}</div>
      </div>
      <span class="text-[11px] font-bold text-amber-400 font-mono bg-amber-950/40 px-2 py-0.5 rounded-lg">🏅 ${(c.puntos||0).toLocaleString()} pts</span>
    </div>
  `).join("") + `
    <div onclick="abrirModalNuevoCliente()" class="px-3 py-2 hover:bg-slate-700 cursor-pointer text-indigo-400 font-bold text-xs flex items-center gap-1.5">
      <span>➕ Nuevo cliente</span>
    </div>
  `;
  dropdown.classList.remove("hidden");
}

function cerrarDropdownCliente(e) {
  if (!e.target.closest("#clienteDropdown") && !e.target.closest("#posClienteInput")) {
    const dd = document.getElementById("clienteDropdown");
    if (dd) dd.classList.add("hidden");
  }
}

function abrirModalNuevoCliente(prefillNombre = "") {
  _editandoClienteId = null;
  const dd = document.getElementById("clienteDropdown");
  if (dd) dd.classList.add("hidden");

  // Solo pre-llenar si se pasa explícitamente como argumento, de lo contrario siempre en blanco
  document.getElementById("modalClienteNombre").value = typeof prefillNombre === "string" ? prefillNombre : "";
  document.getElementById("modalClienteTelefono").value = "";
  document.getElementById("modalClientePuntos").value = "0";
  document.getElementById("modalClienteId").textContent = "Nuevo cliente";
  document.getElementById("btnEliminarCliente").classList.add("hidden");

  const modal = document.getElementById("modalCliente");
  if (modal) { modal.classList.remove("hidden"); modal.classList.add("flex"); }
  
  setTimeout(() => {
    document.getElementById("modalClienteNombre")?.focus();
  }, 100);

  inicializarIconos();
}

function abrirModalEditarCliente(id) {
  _editandoClienteId = id;
  const cli = state.clientes[id];
  if (!cli) return;

  document.getElementById("modalClienteNombre").value = cli.nombre;
  document.getElementById("modalClienteTelefono").value = cli.telefono;
  document.getElementById("modalClientePuntos").value = cli.puntos || 0;
  document.getElementById("modalClienteId").textContent = cli.id;
  document.getElementById("btnEliminarCliente").classList.remove("hidden");

  const modal = document.getElementById("modalCliente");
  if (modal) { modal.classList.remove("hidden"); modal.classList.add("flex"); }
  inicializarIconos();
}

function guardarClienteForm() {
  const nombreInput = document.getElementById("modalClienteNombre");
  const telInput = document.getElementById("modalClienteTelefono");
  const puntosInput = document.getElementById("modalClientePuntos");

  const nombre = (nombreInput?.value || "").trim();
  const telefono = (telInput?.value || "").trim();
  const puntos = parseInt(puntosInput?.value) || 0;

  if (!nombre || !telefono) { mostrarToast("Nombre y teléfono son requeridos.", "error"); return; }

  const clienteObj = _editandoClienteId
    ? { ...(state.clientes[_editandoClienteId] || {}), nombre, telefono, puntos }
    : { nombre, telefono, puntos };

  const guardado = guardarCliente(clienteObj);
  cerrarModalCliente();

  if (guardado) {
    seleccionarCliente(guardado.id);
    renderizarClientes();
    mostrarToast(`Cliente ${nombre} guardado 👤`, "success");
  }
}

function cerrarModalCliente() {
  _editandoClienteId = null;
  const nombreInput = document.getElementById("modalClienteNombre");
  const telInput = document.getElementById("modalClienteTelefono");
  const puntosInput = document.getElementById("modalClientePuntos");
  if (nombreInput) nombreInput.value = "";
  if (telInput) telInput.value = "";
  if (puntosInput) puntosInput.value = "0";

  const modal = document.getElementById("modalCliente");
  if (modal) {
    modal.classList.add("hidden");
    modal.classList.remove("flex");
  }
}

function eliminarClienteActual() {
  if (!_editandoClienteId) return;
  const cli = state.clientes[_editandoClienteId];
  if (!cli) return;
  if (!confirm(`¿Eliminar a ${cli.nombre}? Esta acción no se puede deshacer.`)) return;

  const idAEliminar = _editandoClienteId;
  delete state.clientes[idAEliminar];
  guardarClientesLocal();
  cerrarModalCliente();
  if (state.clienteSeleccionado && state.clienteSeleccionado.id === idAEliminar) {
    deseleccionarCliente();
  }
  renderizarClientes();
  mostrarToast("Cliente eliminado.", "info");
}

// --- Vista Maestro de Clientes ---
function renderizarClientes() {
  const cont = document.getElementById("listaClientesMaestro");
  if (!cont) return;

  const q = state.filtroClientes || "";
  const lista = buscarClientesPorQuery(q).sort((a, b) => (b.puntos || 0) - (a.puntos || 0));
  const totalClientes = document.getElementById("contadorClientes");
  if (totalClientes) totalClientes.textContent = lista.length;

  if (lista.length === 0) {
    cont.innerHTML = `
      <div class="flex flex-col items-center justify-center py-10 text-slate-500 text-xs space-y-2">
        <i data-lucide="users" class="w-10 h-10 stroke-1 text-slate-600"></i>
        <span>${q ? "No hay clientes que coincidan." : "Aún no hay clientes registrados."}</span>
        <button onclick="abrirModalNuevoCliente()" class="px-4 py-2 bg-indigo-600 text-white rounded-xl font-bold text-xs active:scale-95 mt-2">➕ Agregar primer cliente</button>
      </div>
    `;
    inicializarIconos();
    return;
  }

  cont.innerHTML = lista.map(c => {
    const ultimaVenta = c.ultimaVenta ? new Date(c.ultimaVenta).toLocaleDateString() : "—";
    const puntos = c.puntos || 0;
    const valorPuntos = puntos * (state.config.puntosValorCRC || 1);
    return `
      <div onclick="abrirModalEditarCliente('${c.id}')" class="p-3 bg-slate-800/90 border border-slate-700/80 rounded-2xl flex items-center gap-3 cursor-pointer hover:bg-slate-800 active:scale-[0.99] transition-all">
        <div class="w-10 h-10 rounded-full bg-indigo-900/60 border border-indigo-500/30 flex items-center justify-center shrink-0 text-lg font-black text-indigo-300">
          ${c.nombre.charAt(0).toUpperCase()}
        </div>
        <div class="flex-1 min-w-0">
          <div class="text-sm font-bold text-white truncate">${c.nombre}</div>
          <div class="text-[11px] text-slate-400 font-mono">${c.telefono} • Última: ${ultimaVenta}</div>
        </div>
        <div class="text-right shrink-0">
          <div class="text-sm font-black text-amber-400 font-mono">🏅 ${puntos.toLocaleString()}</div>
          <div class="text-[10px] text-slate-500">${fmtCRC(valorPuntos)}</div>
        </div>
      </div>
    `;
  }).join("") + '<div class="h-4"></div>';
  inicializarIconos();
}

// ==========================================================================
// MÓDULO: CUENTAS PENDIENTES (POR COBRAR Y POR PAGAR)
// ==========================================================================

function filtrarTipoCuenta(tipo) {
  state.filtroTipoCuenta = tipo;
  ["cobrar", "pagar", "todos"].forEach(t => {
    const btn = document.getElementById("tabCuentas-" + t);
    if (!btn) return;
    if ((t === "cobrar" && tipo === "Por Cobrar") || (t === "pagar" && tipo === "Por Pagar") || (t === "todos" && tipo === "todos")) {
      btn.className = "py-2 rounded-lg bg-emerald-600 text-white shadow-md text-center transition-all";
    } else {
      btn.className = "py-2 rounded-lg bg-transparent text-slate-400 hover:text-white text-center transition-all";
    }
  });
  renderizarCuentas();
}

function renderizarCuentas() {
  const cont = document.getElementById("listaCuentasPendientes");
  if (!cont) return;

  const q = (state.filtroCuentas || "").toLowerCase().trim();
  const filtroTipo = state.filtroTipoCuenta || "Por Cobrar";

  // Calcular métricas generales
  let totCobrarCRC = 0, totCobrarUSD = 0, countCobrar = 0;
  let totPagarCRC = 0, totPagarUSD = 0, countPagar = 0;

  state.cuentas.forEach(cta => {
    const saldoCRC = Number(cta.saldoPendienteCRC || 0);
    const saldoUSD = Number(cta.saldoPendienteUSD || 0);
    const estado = cta.estado || "Pendiente";

    if (estado !== "Pagado" && saldoCRC > 0) {
      if (cta.tipo === "Por Cobrar") {
        totCobrarCRC += saldoCRC;
        totCobrarUSD += saldoUSD;
        countCobrar++;
      } else {
        totPagarCRC += saldoCRC;
        totPagarUSD += saldoUSD;
        countPagar++;
      }
    }
  });

  // Actualizar Cards de resumen
  const elCobrarCRC = document.getElementById("cuentasTotalCobrarCRC");
  const elCobrarUSD = document.getElementById("cuentasTotalCobrarUSD");
  const elBadgeCobrar = document.getElementById("badgeCuentasCobrar");
  if (elCobrarCRC) elCobrarCRC.textContent = fmtCRC(totCobrarCRC);
  if (elCobrarUSD) elCobrarUSD.textContent = fmtUSD(totCobrarUSD);
  if (elBadgeCobrar) elBadgeCobrar.textContent = countCobrar;

  const elPagarCRC = document.getElementById("cuentasTotalPagarCRC");
  const elPagarUSD = document.getElementById("cuentasTotalPagarUSD");
  const elBadgePagar = document.getElementById("badgeCuentasPagar");
  if (elPagarCRC) elPagarCRC.textContent = fmtCRC(totPagarCRC);
  if (elPagarUSD) elPagarUSD.textContent = fmtUSD(totPagarUSD);
  if (elBadgePagar) elBadgePagar.textContent = countPagar;

  const elCountTotal = document.getElementById("cuentasCountTotal");
  if (elCountTotal) elCountTotal.textContent = state.cuentas.length;

  // Actualizar también widget en Dashboard
  const dashCobrar = document.getElementById("dashCobrarCRC");
  const dashPagar = document.getElementById("dashPagarCRC");
  if (dashCobrar) dashCobrar.textContent = fmtCRC(totCobrarCRC);
  if (dashPagar) dashPagar.textContent = fmtCRC(totPagarCRC);

  // Filtrar lista para mostrar
  let lista = state.cuentas.filter(cta => {
    if (filtroTipo !== "todos" && cta.tipo !== filtroTipo) return false;
    if (q) {
      const ent = (cta.entidad || "").toLowerCase();
      const tel = (cta.telefono || "").toLowerCase();
      const ref = (cta.referenciaId || "").toLowerCase();
      const not = (cta.notas || "").toLowerCase();
      if (!ent.includes(q) && !tel.includes(q) && !ref.includes(q) && !not.includes(q)) return false;
    }
    return true;
  });

  // Ordenar: primero las pendientes con mayor saldo, luego las pagadas
  lista.sort((a, b) => {
    const aPag = a.estado === "Pagado" ? 1 : 0;
    const bPag = b.estado === "Pagado" ? 1 : 0;
    if (aPag !== bPag) return aPag - bPag;
    return (Number(b.saldoPendienteCRC || 0)) - (Number(a.saldoPendienteCRC || 0));
  });

  if (lista.length === 0) {
    cont.innerHTML = `
      <div class="flex flex-col items-center justify-center py-12 text-slate-500 text-xs space-y-2 bg-slate-900/60 rounded-2xl border border-slate-800">
        <i data-lucide="badge-check" class="w-10 h-10 stroke-1 text-slate-600"></i>
        <span>${q ? "No hay cuentas que coincidan con la búsqueda." : "No hay cuentas registradas en esta sección."}</span>
        <span class="text-[11px] text-slate-500 max-w-xs text-center mt-1">
          Las cuentas por cobrar se generan automáticamente al vender con <b>'Pago Luego'</b> seleccionando un cliente registrado.
        </span>
      </div>
    `;
    inicializarIconos();
    return;
  }

  cont.innerHTML = lista.map(cta => {
    const esCobrar = cta.tipo === "Por Cobrar";
    const esPagado = cta.estado === "Pagado" || (Number(cta.saldoPendienteCRC || 0) <= 0);
    const badgeColor = esPagado
      ? "bg-slate-800 text-slate-400 border-slate-700"
      : (esCobrar ? "bg-emerald-950/80 text-emerald-300 border-emerald-500/40" : "bg-rose-950/80 text-rose-300 border-rose-500/40");
    const estadoTexto = esPagado ? "✅ Liquidado" : (cta.estado === "Parcial" ? "⏳ Parcial" : "🔴 Pendiente");
    const iconTipo = esCobrar ? "arrow-down-left" : "arrow-up-right";
    const iconColor = esCobrar ? "text-emerald-400 bg-emerald-950/60" : "text-rose-400 bg-rose-950/60";

    const saldoCRC = Number(cta.saldoPendienteCRC !== undefined ? cta.saldoPendienteCRC : cta.montoTotalCRC);
    const saldoUSD = Number(cta.saldoPendienteUSD !== undefined ? cta.saldoPendienteUSD : cta.montoTotalUSD);
    const totalCRC = Number(cta.montoTotalCRC || 0);

    const fechaStr = cta.fecha ? new Date(cta.fecha).toLocaleDateString() : "";

    return `
      <div class="bg-gradient-to-br from-slate-900 via-slate-900/95 to-slate-950 border ${esPagado ? 'border-slate-800 opacity-70' : 'border-slate-700/90'} rounded-2xl p-3.5 shadow-lg space-y-2.5 transition-all">
        <!-- Top row: Type badge, status, date -->
        <div class="flex items-center justify-between gap-2">
          <div class="flex items-center gap-1.5 min-w-0">
            <span class="p-1 rounded-lg ${iconColor} flex items-center justify-center">
              <i data-lucide="${iconTipo}" class="w-3.5 h-3.5"></i>
            </span>
            <span class="text-[10px] font-black uppercase px-2 py-0.5 rounded-full border ${badgeColor}">
              ${cta.tipo}
            </span>
            <span class="text-[10px] font-bold px-1.5 py-0.5 rounded bg-slate-800 text-slate-300 border border-slate-700 font-mono">${estadoTexto}</span>
          </div>
          <span class="text-[10px] text-slate-500 font-mono shrink-0">${fechaStr}</span>
        </div>

        <!-- Middle: Entity Name, Phone, Ref, Notes -->
        <div class="flex justify-between items-start gap-2">
          <div class="min-w-0 flex-1">
            <div class="text-xs font-black text-white truncate">${cta.entidad || 'Sin nombre'}</div>
            <div class="text-[10px] text-slate-400 font-mono mt-0.5 flex items-center gap-2 flex-wrap">
              ${cta.telefono ? `<span>📞 ${cta.telefono}</span>` : ''}
              ${cta.referenciaId ? `<span>Ref: <b class="text-indigo-300">${cta.referenciaId}</b></span>` : ''}
              ${cta.vendedor ? `<span>Vend: <b>${cta.vendedor}</b></span>` : ''}
            </div>
            ${cta.notas ? `<div class="text-[10px] text-slate-400 italic mt-1 bg-slate-950/60 p-1.5 rounded-lg border border-slate-800/80">${cta.notas}</div>` : ''}
          </div>

          <!-- Balance Amounts -->
          <div class="text-right font-mono shrink-0">
            <div class="text-[10px] text-slate-400">Saldo Pendiente:</div>
            <div class="text-sm font-black ${esPagado ? 'text-slate-400 line-through' : (esCobrar ? 'text-emerald-400' : 'text-rose-400')}">${fmtCRC(saldoCRC)}</div>
            <div class="text-[10px] text-slate-400">${fmtUSD(saldoUSD)}</div>
            ${totalCRC > saldoCRC ? `<div class="text-[9px] text-slate-500">Total orig: ${fmtCRC(totalCRC)}</div>` : ''}
          </div>
        </div>

        <!-- Action Buttons -->
        <div class="flex items-center justify-between pt-2 border-t border-slate-800/80 gap-2">
          <div class="flex items-center gap-1.5">
            ${cta.telefono ? `
              <a href="https://wa.me/506${cta.telefono.replace(/[^0-9]/g, '')}?text=Hola%20${encodeURIComponent(cta.entidad)},%20te%20saludamos%20de%20DC%20El%20Destape.%20Te%20recordamos%20el%20saldo%20pendiente%20de%20${encodeURIComponent(fmtCRC(saldoCRC))}.%20¡Pura%20vida!" target="_blank"
                class="px-2 py-1 bg-emerald-950/60 hover:bg-emerald-900 border border-emerald-500/30 text-emerald-400 rounded-lg text-[10px] font-bold flex items-center gap-1 active:scale-95">
                <i data-lucide="message-circle" class="w-3 h-3"></i>
                <span>WhatsApp</span>
              </a>
            ` : ''}
            <button onclick="eliminarCuenta('${cta.id}')" class="px-2 py-1 bg-slate-800 hover:bg-rose-950/60 text-slate-400 hover:text-rose-300 rounded-lg text-[10px] active:scale-95 transition-all">
              Eliminar
            </button>
          </div>

          ${!esPagado ? `
            <div class="flex items-center gap-1.5">
              <button onclick="abrirModalAbonoCuenta('${cta.id}')" class="px-3 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl text-xs font-bold flex items-center gap-1 active:scale-95 shadow-md shadow-indigo-600/30">
                <i data-lucide="plus" class="w-3.5 h-3.5"></i>
                <span>Abonar</span>
              </button>
              <button onclick="liquidarCuentaDirecto('${cta.id}')" class="px-2.5 py-1.5 bg-emerald-600/20 hover:bg-emerald-600/40 border border-emerald-500/40 text-emerald-300 rounded-xl text-xs font-bold flex items-center gap-1 active:scale-95">
                <i data-lucide="check" class="w-3.5 h-3.5"></i>
                <span>Liquidar</span>
              </button>
            </div>
          ` : `
            <span class="text-[10px] text-emerald-400 font-bold font-mono flex items-center gap-1">
              <i data-lucide="check-circle" class="w-3.5 h-3.5"></i> Pagado totalmente
            </span>
          `}
        </div>
      </div>
    `;
  }).join("") + '<div class="h-6"></div>';

  inicializarIconos();
}

// --- Abonos a Cuentas ---
function abrirModalAbonoCuenta(idCuenta) {
  const cta = state.cuentas.find(c => c.id === idCuenta);
  if (!cta) return;

  document.getElementById("abonoCuentaId").value = idCuenta;
  document.getElementById("abonoEntidadNombre").textContent = `${cta.tipo}: ${cta.entidad}`;
  
  const saldoCRC = Number(cta.saldoPendienteCRC !== undefined ? cta.saldoPendienteCRC : cta.montoTotalCRC);
  const saldoUSD = Number(cta.saldoPendienteUSD !== undefined ? cta.saldoPendienteUSD : cta.montoTotalUSD);

  document.getElementById("abonoSaldoActualCRC").textContent = fmtCRC(saldoCRC);
  document.getElementById("abonoSaldoActualUSD").textContent = fmtUSD(saldoUSD);

  document.getElementById("abonoMontoCRC").value = "";
  document.getElementById("abonoMontoUSD").value = "";
  document.getElementById("abonoNota").value = "";

  const modal = document.getElementById("modalAbonoCuenta");
  if (modal) {
    modal.classList.remove("hidden");
    modal.classList.add("flex");
  }
  inicializarIconos();
  setTimeout(() => document.getElementById("abonoMontoCRC").focus(), 150);
}

function cerrarModalAbonoCuenta() {
  const modal = document.getElementById("modalAbonoCuenta");
  if (modal) {
    modal.classList.add("hidden");
    modal.classList.remove("flex");
  }
}

function llenarAbonoTotal() {
  const idCuenta = document.getElementById("abonoCuentaId").value;
  const cta = state.cuentas.find(c => c.id === idCuenta);
  if (!cta) return;

  const saldoCRC = Number(cta.saldoPendienteCRC !== undefined ? cta.saldoPendienteCRC : cta.montoTotalCRC);
  const tc = Number(state.config.tipoCambio || 520);
  const saldoUSD = parseFloat((saldoCRC / tc).toFixed(2));

  document.getElementById("abonoMontoCRC").value = saldoCRC;
  document.getElementById("abonoMontoUSD").value = saldoUSD;
}

function autoConvertirAbono(fuente) {
  const tc = Number(state.config.tipoCambio || 520);
  const inCRC = document.getElementById("abonoMontoCRC");
  const inUSD = document.getElementById("abonoMontoUSD");

  if (fuente === "CRC") {
    const valCRC = parseFloat(inCRC.value) || 0;
    inUSD.value = valCRC > 0 ? (valCRC / tc).toFixed(2) : "";
  } else {
    const valUSD = parseFloat(inUSD.value) || 0;
    inCRC.value = valUSD > 0 ? Math.round(valUSD * tc) : "";
  }
}

function guardarAbonoCuenta() {
  const idCuenta = document.getElementById("abonoCuentaId").value;
  const cta = state.cuentas.find(c => c.id === idCuenta);
  if (!cta) return;

  const abonoCRC = parseFloat(document.getElementById("abonoMontoCRC").value) || 0;
  const abonoUSD = parseFloat(document.getElementById("abonoMontoUSD").value) || 0;
  const metodo = document.getElementById("abonoMetodoPago").value;
  const nota = document.getElementById("abonoNota").value.trim();

  if (abonoCRC <= 0 && abonoUSD <= 0) {
    mostrarToast("Ingresa un monto válido para el abono.", "error");
    return;
  }

  const saldoAnteriorCRC = Number(cta.saldoPendienteCRC !== undefined ? cta.saldoPendienteCRC : cta.montoTotalCRC);
  const saldoAnteriorUSD = Number(cta.saldoPendienteUSD !== undefined ? cta.saldoPendienteUSD : cta.montoTotalUSD);

  const nuevoSaldoCRC = Math.max(0, saldoAnteriorCRC - abonoCRC);
  const nuevoSaldoUSD = Math.max(0, saldoAnteriorUSD - abonoUSD);
  const nuevoEstado = nuevoSaldoCRC <= 0 ? "Pagado" : "Parcial";

  cta.saldoPendienteCRC = nuevoSaldoCRC;
  cta.saldoPendienteUSD = nuevoSaldoUSD;
  cta.estado = nuevoEstado;
  const notaExtra = ` [Abono ₡${abonoCRC} (${metodo}) el ${new Date().toLocaleDateString()}${nota ? ': ' + nota : ''}]`;
  cta.notas = (cta.notas || "") + notaExtra;

  guardarCuentasLocal();

  // Encolar acción para sincronizar con Google Sheets
  encolarAccionSincronizacion("abonarCuenta", {
    id: idCuenta,
    abonoCRC,
    abonoUSD,
    notas: `${metodo} - ${nota}`
  });

  cerrarModalAbonoCuenta();
  renderizarCuentas();
  renderizarFinanzas();
  renderizarDashboard();

  if (nuevoEstado === "Pagado") {
    mostrarToast(`¡Cuenta de ${cta.entidad} liquidada completamente! Dinero ingresado a Caja 💵`, "success");
    if (window.confetti) window.confetti({ particleCount: 70, spread: 60, origin: { y: 0.8 } });
  } else {
    mostrarToast(`Abono de ${fmtCRC(abonoCRC)} registrado. Saldo restante: ${fmtCRC(nuevoSaldoCRC)} 💵`, "success");
  }
}

function liquidarCuentaDirecto(idCuenta) {
  const cta = state.cuentas.find(c => c.id === idCuenta);
  if (!cta) return;

  const saldoCRC = Number(cta.saldoPendienteCRC !== undefined ? cta.saldoPendienteCRC : cta.montoTotalCRC);
  const saldoUSD = Number(cta.saldoPendienteUSD !== undefined ? cta.saldoPendienteUSD : cta.montoTotalUSD);

  if (!confirm(`¿Confirmas liquidar el saldo total de ${fmtCRC(saldoCRC)} de ${cta.entidad}?`)) return;

  cta.saldoPendienteCRC = 0;
  cta.saldoPendienteUSD = 0;
  cta.estado = "Pagado";
  cta.notas = (cta.notas || "") + ` [Liquidado total ₡${saldoCRC} el ${new Date().toLocaleDateString()}]`;

  guardarCuentasLocal();
  encolarAccionSincronizacion("abonarCuenta", {
    id: idCuenta,
    abonoCRC: saldoCRC,
    abonoUSD: saldoUSD,
    notas: "Liquidación completa"
  });

  renderizarCuentas();
  renderizarFinanzas();
  renderizarDashboard();
  mostrarToast(`¡Cuenta de ${cta.entidad} liquidada totalmente! Dinero sumado a Caja 💵`, "success");
  if (window.confetti) window.confetti({ particleCount: 70, spread: 60, origin: { y: 0.8 } });
}

function eliminarCuenta(idCuenta) {
  if (!confirm("¿Deseas eliminar este registro de cuenta pendiente?")) return;
  state.cuentas = state.cuentas.filter(c => c.id !== idCuenta);
  guardarCuentasLocal();
  renderizarCuentas();
  renderizarFinanzas();
  renderizarDashboard();
  mostrarToast("Cuenta eliminada.", "info");
  encolarAccionSincronizacion("eliminarCuenta", { id: idCuenta });
}

// --- Crear Cuenta Manual ---
function setTipoNuevaCuenta(tipo) {
  document.getElementById("nuevaCtaTipo").value = tipo;
  const btnCobrar = document.getElementById("btnTipoCtaCobrar");
  const btnPagar = document.getElementById("btnTipoCtaPagar");
  const lbl = document.getElementById("lblNuevaCtaEntidad");

  if (tipo === "Por Cobrar") {
    btnCobrar.className = "py-2.5 rounded-xl bg-emerald-600 text-white text-center border border-transparent";
    btnPagar.className = "py-2.5 rounded-xl bg-slate-800 text-slate-400 text-center border border-slate-700";
    if (lbl) lbl.textContent = "Nombre del Cliente *";
  } else {
    btnPagar.className = "py-2.5 rounded-xl bg-rose-600 text-white text-center border border-transparent";
    btnCobrar.className = "py-2.5 rounded-xl bg-slate-800 text-slate-400 text-center border border-slate-700";
    if (lbl) lbl.textContent = "Nombre del Proveedor *";
  }
}

function abrirModalNuevaCuenta() {
  document.getElementById("nuevaCtaEntidad").value = "";
  document.getElementById("nuevaCtaTelefono").value = "";
  document.getElementById("nuevaCtaMontoCRC").value = "";
  document.getElementById("nuevaCtaMontoUSD").value = "";
  document.getElementById("nuevaCtaVencimiento").value = "";
  document.getElementById("nuevaCtaNotas").value = "";
  setTipoNuevaCuenta("Por Cobrar");

  const modal = document.getElementById("modalNuevaCuenta");
  if (modal) {
    modal.classList.remove("hidden");
    modal.classList.add("flex");
  }
  inicializarIconos();
}

function cerrarModalNuevaCuenta() {
  const modal = document.getElementById("modalNuevaCuenta");
  if (modal) {
    modal.classList.add("hidden");
    modal.classList.remove("flex");
  }
}

function autoConvertirNuevaCta(fuente) {
  const tc = Number(state.config.tipoCambio || 520);
  const inCRC = document.getElementById("nuevaCtaMontoCRC");
  const inUSD = document.getElementById("nuevaCtaMontoUSD");

  if (fuente === "CRC") {
    const valCRC = parseFloat(inCRC.value) || 0;
    inUSD.value = valCRC > 0 ? (valCRC / tc).toFixed(2) : "";
  } else {
    const valUSD = parseFloat(inUSD.value) || 0;
    inCRC.value = valUSD > 0 ? Math.round(valUSD * tc) : "";
  }
}

function guardarNuevaCuentaManual() {
  const tipo = document.getElementById("nuevaCtaTipo").value || "Por Cobrar";
  const entidad = document.getElementById("nuevaCtaEntidad").value.trim();
  const telefono = document.getElementById("nuevaCtaTelefono").value.trim();
  const montoCRC = parseFloat(document.getElementById("nuevaCtaMontoCRC").value) || 0;
  const montoUSD = parseFloat(document.getElementById("nuevaCtaMontoUSD").value) || 0;
  const vencimiento = document.getElementById("nuevaCtaVencimiento").value;
  const vendedor = document.getElementById("nuevaCtaVendedor").value || "Carlos";
  const notas = document.getElementById("nuevaCtaNotas").value.trim();

  if (!entidad) {
    mostrarToast("Ingresa el nombre del cliente o proveedor.", "error");
    return;
  }
  if (montoCRC <= 0 && montoUSD <= 0) {
    mostrarToast("Ingresa un monto válido.", "error");
    return;
  }

  const id = "CTA-" + Date.now().toString().slice(-6);
  const cuentaObj = {
    id,
    fecha: new Date().toISOString(),
    tipo,
    entidad,
    telefono,
    referenciaId: "MANUAL",
    montoTotalCRC: montoCRC,
    montoTotalUSD: montoUSD,
    saldoPendienteCRC: montoCRC,
    saldoPendienteUSD: montoUSD,
    estado: "Pendiente",
    fechaVencimiento: vencimiento,
    vendedor,
    notas
  };

  state.cuentas.unshift(cuentaObj);
  guardarCuentasLocal();
  encolarAccionSincronizacion("registrarCuenta", { cuenta: cuentaObj });

  cerrarModalNuevaCuenta();
  renderizarCuentas();
  mostrarToast(`Cuenta pendiente para ${entidad} guardada con éxito 📋`, "success");
}

// --- Editar precio de ítem en el carrito ---
function editarPrecioCarrito(codigo) {
  const item = state.carrito.find(i => i.codigo === codigo);
  if (!item) return;

  document.getElementById("editPrecioProductoNombre").textContent = item.nombre;
  document.getElementById("editPrecioOriginal").textContent = fmtCRC(item.precioVentaCRC);
  document.getElementById("editPrecioInput").value = item.precioVentaCRC;
  document.getElementById("editPrecioCodigo").value = codigo;

  const modal = document.getElementById("modalEditarPrecio");
  if (modal) { modal.classList.remove("hidden"); modal.classList.add("flex"); }
  setTimeout(() => document.getElementById("editPrecioInput").focus(), 100);
}

function aplicarNuevoPrecioCarrito() {
  const codigo = document.getElementById("editPrecioCodigo").value;
  const nuevoPrecio = parseFloat(document.getElementById("editPrecioInput").value) || 0;
  if (nuevoPrecio < 0) { mostrarToast("El precio no puede ser negativo.", "error"); return; }

  const item = state.carrito.find(i => i.codigo === codigo);
  if (item) {
    const tc = state.config.tipoCambio || 520;
    item.precioVentaCRC = nuevoPrecio;
    item.precioVentaUSD = parseFloat((nuevoPrecio / tc).toFixed(2));
    item._precioEditado = true;
  }

  cerrarModalEditarPrecio();
  renderizarCarrito();
  mostrarToast("Precio actualizado correctamente.", "success");
}

function cerrarModalEditarPrecio() {
  const modal = document.getElementById("modalEditarPrecio");
  if (modal) { modal.classList.add("hidden"); modal.classList.remove("flex"); }
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

  if (!url) {
    mostrarToast("Este licor no tiene foto asignada", "info");
    return;
  }

  // Intentar cargar la URL principal; si falla intentar thumbnail más grande
  img.onerror = function() {
    this.onerror = null;
    // Extraer driveId si está en el src
    const idMatch = this.src.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (idMatch) {
      this.src = `https://drive.google.com/thumbnail?id=${idMatch[1]}&sz=w1200`;
    } else {
      // Sin fallback posible — ocultar imagen rota
      this.style.display = 'none';
    }
  };

  // Asegurarse de que el src siempre se actualice correctamente
  img.src = "";
  img.style.display = "";
  img.src = formatearUrlImagen(url);
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

// Delegación de eventos para fotos de productos en inventario (compatible móvil)
document.addEventListener("click", (e) => {
  const btn = e.target.closest(".foto-producto-btn");
  if (btn) {
    e.preventDefault();
    e.stopPropagation();
    const url = btn.dataset.url || "";
    const nombre = btn.dataset.nombre || "";
    abrirFotoCompleta(url, nombre);
  }
}, { passive: false });



function formatearUrlImagen(urlOrId) {
  if (!urlOrId || typeof urlOrId !== 'string') return '';
  const trimmed = urlOrId.trim();
  if (!trimmed) return '';


  // 1. Data URLs directas (Base64)
  if (trimmed.startsWith('data:image/')) {
    return trimmed;
  }

  // 2. Extraer ID de Google Drive (varios formatos conocidos)
  let driveId = null;

  // Formato /file/d/ID/view o /file/d/ID
  const matchFileD = trimmed.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (matchFileD && matchFileD[1]) driveId = matchFileD[1];

  // Formato id=ID o ?id=ID
  if (!driveId) {
    const matchIdParam = trimmed.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (matchIdParam && matchIdParam[1]) driveId = matchIdParam[1];
  }

  // Formato lh3.googleusercontent.com/d/ID
  if (!driveId) {
    const matchGoogleUserContent = trimmed.match(/googleusercontent\.com\/d\/([a-zA-Z0-9_-]+)/);
    if (matchGoogleUserContent && matchGoogleUserContent[1]) driveId = matchGoogleUserContent[1];
  }

  // Formato drive.google.com/open?id=ID o /uc?id=ID
  if (!driveId) {
    const matchUc = trimmed.match(/drive\.google\.com\/(?:uc|open)\?.*id=([a-zA-Z0-9_-]+)/);
    if (matchUc && matchUc[1]) driveId = matchUc[1];
  }

  // Si pegó directamente el ID alfanumérico de Drive (25 a 50 caracteres)
  if (!driveId && /^[a-zA-Z0-9_-]{25,50}$/.test(trimmed)) {
    driveId = trimmed;
  }

  if (driveId) {
    // drive.google.com/thumbnail?id=ID&sz=w400 funciona en móvil y escritorio sin restricciones CORS
    // lh3.googleusercontent.com/d/ID puede fallar en Android/iOS sin autenticación
    return `https://drive.google.com/thumbnail?id=${driveId}&sz=w400`;
  }

  // 3. URLs web directas (http/https)
  if (trimmed.startsWith('http://') || trimmed.startsWith('https://')) {
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
  
  // Intentar fallback si es de Google Drive
  if (preview && preview.src && preview.src.includes("lh3.googleusercontent.com/d/")) {
    const id = preview.src.split("/d/")[1];
    preview.src = `https://drive.google.com/thumbnail?id=${id}&sz=w500`;
    return;
  }

  if (preview) {
    preview.classList.add("hidden");
  }
  if (placeholder) {
    placeholder.classList.remove("hidden");
    placeholder.innerHTML = `<i data-lucide="alert-circle" class="w-6 h-6 text-amber-500 mb-0.5"></i><span class="text-amber-400 text-[10px]">No cargó imagen</span>`;
    inicializarIconos();
  }
}

function manejarSubidaImagenProducto(e) {
  const file = e.target.files && e.target.files[0];
  if (!file) return;

  if (!file.type.startsWith("image/")) {
    mostrarToast("El archivo seleccionado no es una imagen válida", "error");
    return;
  }

  mostrarToast("Procesando y optimizando imagen... ⏳", "info");

  const reader = new FileReader();
  reader.onload = (event) => {
    const img = new Image();
    img.onload = () => {
      // Redimensionar / comprimir imagen a máx 600px para no saturar memoria/almacenamiento
      const maxDim = 600;
      let width = img.width;
      let height = img.height;

      if (width > maxDim || height > maxDim) {
        if (width > height) {
          height = Math.round((height * maxDim) / width);
          width = maxDim;
        } else {
          width = Math.round((width * maxDim) / height);
          height = maxDim;
        }
      }

      const canvas = document.createElement("canvas");
      canvas.width = width;
      canvas.height = height;
      const ctx = canvas.getContext("2d");
      ctx.drawImage(img, 0, 0, width, height);

      const base64Optimizada = canvas.toDataURL("image/jpeg", 0.82);
      const input = document.getElementById("prodImagenUrl");
      if (input) {
        input.value = base64Optimizada;
        actualizarPreviewImagenModal();
        mostrarToast("Foto cargada con éxito 📸", "success");
      }
    };
    img.src = event.target.result;
  };
  reader.readAsDataURL(file);
}

function limpiarImagenModal() {
  const input = document.getElementById("prodImagenUrl");
  if (input) input.value = "";
  actualizarPreviewImagenModal();
}

// ==========================================================================
// MODAL DE PRODUCTO (DUAL CURRENCY & FOTO)
// ==========================================================================
function abrirModalProducto(productoOcodigo = null) {
  const modal = document.getElementById("modalProducto");
  const titulo = document.getElementById("modalProductoTitulo");
  const btnEliminar = document.getElementById("btnEliminarProducto");

  let producto = null;
  if (typeof productoOcodigo === "string") {
    producto = state.productos[productoOcodigo] || null;
  } else if (productoOcodigo && typeof productoOcodigo === "object") {
    producto = productoOcodigo;
  }

  if (producto) {
    titulo.innerHTML = `<i data-lucide="edit" class="w-5 h-5 text-amber-400"></i> Editar Licor (${producto.codigo})`;
    document.getElementById("prodCodigo").value = producto.codigo || "";
    document.getElementById("prodCodigo").disabled = true;
    document.getElementById("prodNombre").value = producto.nombre || "";
    document.getElementById("prodCategoria").value = producto.categoria || "";
    document.getElementById("prodImagenUrl").value = producto.imagenUrl || "";
    const cUSD = Number(producto.costoRefUSD || 0);
    const pUSD = Number(producto.precioVentaUSD || 0);
    document.getElementById("prodCostoRefUSD").value = cUSD;
    document.getElementById("prodCostoRefCRC").value = Number(producto.costoRefCRC || 0);
    
    // Si ya tiene costo y precio, calcular el margen que tiene actualmente
    let margenActual = 70;
    if (cUSD > 0 && pUSD >= cUSD) {
      margenActual = Number((((pUSD - cUSD) / cUSD) * 100).toFixed(1));
    }
    document.getElementById("prodMargenPorcentaje").value = margenActual;
    
    document.getElementById("prodPrecioVentaUSD").value = pUSD;
    document.getElementById("prodPrecioVentaCRC").value = Number(producto.precioVentaCRC || 0);
    document.getElementById("prodStock").value = Number(producto.stockInicial || 0);
    document.getElementById("prodStockMinimo").value = Number(producto.stockMinimo || 2);
    
    calcularMargenYPrecioRecomendado('costoUSD', false);
    if (btnEliminar) btnEliminar.classList.remove("hidden");
  } else {
    titulo.innerHTML = `<i data-lucide="wine" class="w-5 h-5 text-amber-400"></i> Nuevo Licor`;
    document.getElementById("formProducto").reset();
    document.getElementById("prodCodigo").disabled = false;
    document.getElementById("prodCodigo").value = "LIC-" + Math.floor(100 + Math.random() * 900);
    document.getElementById("prodImagenUrl").value = "";
    document.getElementById("prodMargenPorcentaje").value = 70;
    document.getElementById("prodStockMinimo").value = 2;
    document.getElementById("prodCostoRefUSD").value = "";
    document.getElementById("prodCostoRefCRC").value = "0";
    document.getElementById("prodPrecioRecomendadoUSD").value = "0";
    document.getElementById("prodPrecioVentaUSD").value = "";
    document.getElementById("prodPrecioVentaCRC").value = "";
    if (btnEliminar) btnEliminar.classList.add("hidden");
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

// --------------------------------------------------------------------------
// CÁLCULO DE MARGEN (%) Y PRECIO RECOMENDADO SEGÚN FÓRMULA DEL USUARIO:
// Precio Recomendado = Precio Costo + (Precio Costo / 100 * Margen%)
// --------------------------------------------------------------------------
function calcularMargenYPrecioRecomendado(origen = 'costoUSD', autoAplicarVenta = true) {
  const tc = Number(state.config.tipoCambio) || 520;
  const costoUSD = parseFloat(document.getElementById("prodCostoRefUSD").value) || 0;
  const margenPct = parseFloat(document.getElementById("prodMargenPorcentaje").value) || 0;

  // 1. Costo Ref. en CRC automático: Costo USD * Tipo de Cambio
  const costoCRC = Math.round(costoUSD * tc);
  const inCostoCRC = document.getElementById("prodCostoRefCRC");
  if (inCostoCRC) inCostoCRC.value = costoCRC;

  // 2. Fórmula exacta: Costo + (Costo / 100 * Margen%)
  const precioRecomendadoUSD = costoUSD + (costoUSD * (margenPct / 100));
  const inRecUSD = document.getElementById("prodPrecioRecomendadoUSD");
  if (inRecUSD) inRecUSD.value = precioRecomendadoUSD > 0 ? precioRecomendadoUSD.toFixed(2) : "0";

  // 3. Si es un producto nuevo o se cambia el costo/margen, sugerir en los campos de venta si están vacíos o si autoAplicarVenta es true
  const inVentaUSD = document.getElementById("prodPrecioVentaUSD");
  const inVentaCRC = document.getElementById("prodPrecioVentaCRC");
  
  if (autoAplicarVenta && inVentaUSD) {
    inVentaUSD.value = precioRecomendadoUSD > 0 ? precioRecomendadoUSD.toFixed(2) : "";
    if (inVentaCRC) {
      inVentaCRC.value = precioRecomendadoUSD > 0 ? Math.round(precioRecomendadoUSD * tc) : "";
    }
  }
}

function aplicarPrecioRecomendado() {
  const recUSD = parseFloat(document.getElementById("prodPrecioRecomendadoUSD").value) || 0;
  if (recUSD <= 0) {
    mostrarToast("Ingresa primero el precio de costo en USD.", "info");
    return;
  }
  const tc = Number(state.config.tipoCambio) || 520;
  document.getElementById("prodPrecioVentaUSD").value = recUSD.toFixed(2);
  document.getElementById("prodPrecioVentaCRC").value = Math.round(recUSD * tc);
  mostrarToast("Precio recomendado aplicado a la venta 💵", "success");
}

function autoConvertirPrecio(origen) {
  const tc = Number(state.config.tipoCambio) || 520;
  if (origen === 'USD') {
    const usd = Number(document.getElementById("prodPrecioVentaUSD").value) || 0;
    document.getElementById("prodPrecioVentaCRC").value = usd > 0 ? Math.round(usd * tc) : "";
  } else {
    const crc = Number(document.getElementById("prodPrecioVentaCRC").value) || 0;
    document.getElementById("prodPrecioVentaUSD").value = crc > 0 ? (crc / tc).toFixed(2) : "";
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
  
  // Si la categoría actual filtrada no coincide con la del nuevo producto, poner en "Todas"
  if (state.categoriaSeleccionada !== "Todas" && state.categoriaSeleccionada !== categoria) {
    state.categoriaSeleccionada = "Todas";
  }
  
  // Limpiar texto de búsqueda para mostrar la lista completa
  const inSearch = document.getElementById("searchInventory");
  if (inSearch && inSearch.value) inSearch.value = "";

  renderizarTodo();
  cerrarModalProducto();
  mostrarToast(esEdicion ? "Producto actualizado correctamente." : "Producto agregado correctamente.", "success");

  // Encolar y sincronizar
  encolarAccionSincronizacion(esEdicion ? "actualizarProducto" : "crearProducto", { producto: prodObj });
}

async function eliminarProductoActual() {
  const codigo = document.getElementById("prodCodigo").value;
  if (!confirm(`¿Eliminar definitivamente el producto ${codigo}?`)) return;

  delete state.productos[codigo];
  guardarProductosLocal();

  const inSearch = document.getElementById("searchInventory");
  if (inSearch && inSearch.value) inSearch.value = "";

  renderizarTodo();
  cerrarModalProducto();
  mostrarToast("Producto eliminado correctamente.", "info");

  // Encolar y sincronizar
  encolarAccionSincronizacion("eliminarProducto", { codigo });
}

// ==========================================================================
// 3. COMPRAS / ENTRADAS (Búsqueda predictiva y selección)
// ==========================================================================
function poblarSelectCompras() {
  const cod = document.getElementById("compraProductoCodigo")?.value;
  if (!cod) {
    limpiarSeleccionProductoCompra(false);
  }
}

function filtrarProductosCompra(q) {
  const dropdown = document.getElementById("compraProductoSugerencias");
  if (!dropdown) return;

  const ql = (q || "").trim().toLowerCase();
  const prods = Object.values(state.productos).sort((a, b) => a.nombre.localeCompare(b.nombre));
  const stockMap = calcularStockPorCodigo();

  const filtrados = ql.length === 0
    ? prods.slice(0, 8)
    : prods.filter(p => 
        p.nombre.toLowerCase().includes(ql) || 
        p.codigo.toLowerCase().includes(ql) || 
        (p.categoria && p.categoria.toLowerCase().includes(ql))
      ).slice(0, 10);

  if (filtrados.length === 0) {
    dropdown.innerHTML = `
      <div class="p-3 text-xs text-slate-400 text-center">
        No se encontró ningún producto con "<strong>${q}</strong>".
      </div>
    `;
    dropdown.classList.remove("hidden");
    return;
  }

  dropdown.innerHTML = filtrados.map(p => {
    const imgUrl = formatearUrlImagen(p.imagenUrl);
    const imgHtml = imgUrl
      ? `<img src="${imgUrl}" alt="${p.nombre}" class="w-8 h-8 rounded-lg object-cover bg-slate-900 border border-slate-700 shrink-0" onerror="this.outerHTML='<div class=\\'w-8 h-8 rounded-lg bg-slate-900 border border-slate-700 flex items-center justify-center text-xs shrink-0\\'>🍷</div>'">`
      : `<div class="w-8 h-8 rounded-lg bg-slate-900 border border-slate-700 flex items-center justify-center text-xs shrink-0">🍷</div>`;

    const stock = stockMap[p.codigo] || 0;
    return `
      <div onclick="seleccionarProductoCompraPorCodigo('${p.codigo}')" class="px-3 py-2.5 hover:bg-slate-700/80 cursor-pointer flex items-center justify-between gap-2.5 transition-colors">
        <div class="flex items-center gap-2.5 min-w-0 flex-1">
          ${imgHtml}
          <div class="min-w-0 flex-1">
            <div class="text-xs font-bold text-white truncate">${p.nombre}</div>
            <div class="text-[11px] text-slate-400 font-mono">${p.codigo} • ${p.categoria || 'Licor'}</div>
          </div>
        </div>
        <div class="text-right shrink-0">
          <div class="text-[11px] font-mono font-bold ${stock > 0 ? 'text-emerald-400' : 'text-slate-400'}">Stock: ${stock}</div>
          <div class="text-[10px] text-slate-500 font-mono">$${(p.costoRefUSD||0).toFixed(2)}</div>
        </div>
      </div>
    `;
  }).join("");

  dropdown.classList.remove("hidden");
  inicializarIconos();
}

function seleccionarProductoCompraPorCodigo(codigo) {
  const prod = state.productos[codigo];
  if (!prod) return;

  const tc = Number(document.getElementById("compraTipoCambio").value) || Number(state.config.tipoCambio) || 520;

  // Llenar campos del editor rápido de item
  document.getElementById("compraItemEditorCodigo").value = codigo;
  document.getElementById("compraItemEditorNombre").textContent = `${prod.nombre} (${prod.codigo})`;
  
  const imgUrl = formatearUrlImagen(prod.imagenUrl);
  const container = document.getElementById("compraItemEditorImg");
  if (container) {
    container.innerHTML = imgUrl
      ? `<img src="${imgUrl}" alt="${prod.nombre}" class="w-7 h-7 rounded-lg object-cover" onerror="this.outerHTML='🍷'">`
      : `🍷`;
  }

  // Pre-cargar costos de referencia
  const costoUSD = Number(prod.costoRefUSD || 0);
  const costoCRC = Number(prod.costoRefCRC || (costoUSD * tc));
  document.getElementById("compraItemEditorCantidad").value = 1;
  document.getElementById("compraItemEditorCostoUSD").value = costoUSD;
  document.getElementById("compraItemEditorCostoCRC").value = costoCRC;

  // Mostrar editor de item y ocultar dropdown
  document.getElementById("compraItemEditor")?.classList.remove("hidden");
  document.getElementById("compraProductoSugerencias")?.classList.add("hidden");
  document.getElementById("compraProductoBusqueda").value = "";

  document.getElementById("compraItemEditorCantidad")?.focus();
  inicializarIconos();
}

function cancelarItemCompra() {
  document.getElementById("compraItemEditor")?.classList.add("hidden");
  document.getElementById("compraProductoBusqueda").value = "";
}

function autoConvertirItemCompraCosto(origen) {
  const tc = Number(document.getElementById("compraTipoCambio").value) || Number(state.config.tipoCambio) || 520;
  const elUSD = document.getElementById("compraItemEditorCostoUSD");
  const elCRC = document.getElementById("compraItemEditorCostoCRC");
  if (origen === 'USD' && elUSD && elCRC) {
    const usd = Number(elUSD.value) || 0;
    elCRC.value = Math.round(usd * tc);
  } else if (origen === 'CRC' && elUSD && elCRC) {
    const crc = Number(elCRC.value) || 0;
    elUSD.value = (crc / tc).toFixed(2);
  }
}

function calcularSubtotalItemCompra() {
  // Función auxiliar para reactividad
}

function agregarItemAListaCompra() {
  const codigo = document.getElementById("compraItemEditorCodigo")?.value;
  const cant = Number(document.getElementById("compraItemEditorCantidad")?.value) || 0;
  const costoUSD = Number(document.getElementById("compraItemEditorCostoUSD")?.value) || 0;
  const costoCRC = Number(document.getElementById("compraItemEditorCostoCRC")?.value) || 0;

  if (!codigo || !state.productos[codigo]) {
    mostrarToast("Producto no válido", "error");
    return;
  }
  if (cant <= 0) {
    mostrarToast("La cantidad debe ser mayor a 0", "error");
    return;
  }

  const prod = state.productos[codigo];
  if (!state.listaCompraActual) state.listaCompraActual = [];

  const existente = state.listaCompraActual.find(i => i.codigo === codigo);
  if (existente) {
    existente.cantidad += cant;
    existente.costoUnitarioUSD = costoUSD;
    existente.costoUnitarioCRC = costoCRC;
  } else {
    state.listaCompraActual.push({
      codigo: prod.codigo,
      nombre: prod.nombre,
      imagenUrl: prod.imagenUrl || "",
      cantidad: cant,
      costoUnitarioUSD: costoUSD,
      costoUnitarioCRC: costoCRC
    });
  }

  document.getElementById("compraItemEditor")?.classList.add("hidden");
  document.getElementById("compraProductoBusqueda").value = "";
  renderizarListaCompraActual();
  mostrarToast(`+${cant} ${prod.nombre} añadido a la compra`, "success");
}

function modificarCantidadItemCompra(codigo, delta) {
  const item = state.listaCompraActual.find(i => i.codigo === codigo);
  if (!item) return;

  const nuevo = item.cantidad + delta;
  if (nuevo <= 0) {
    quitarItemDeListaCompra(codigo);
    return;
  }
  item.cantidad = nuevo;
  renderizarListaCompraActual();
}

function quitarItemDeListaCompra(codigo) {
  state.listaCompraActual = (state.listaCompraActual || []).filter(i => i.codigo !== codigo);
  renderizarListaCompraActual();
}

function vaciarListaCompra() {
  state.listaCompraActual = [];
  renderizarListaCompraActual();
}

function recalcularTotalesListaCompra() {
  renderizarListaCompraActual();
}

function renderizarListaCompraActual() {
  const cont = document.getElementById("compraListaItems");
  const countEl = document.getElementById("compraListaCount");
  const totalCRCEl = document.getElementById("compraTotalCRCDisplay");
  const totalUSDEl = document.getElementById("compraTotalUSDDisplay");
  if (!cont) return;

  const items = state.listaCompraActual || [];
  if (countEl) countEl.textContent = items.reduce((acc, i) => acc + i.cantidad, 0);

  let totalCRC = 0;
  let totalUSD = 0;

  if (items.length === 0) {
    cont.innerHTML = `
      <div class="py-4 text-center text-slate-500 text-xs">
        No has agregado productos a esta compra todavía.
      </div>
    `;
  } else {
    cont.innerHTML = items.map(item => {
      const subCRC = item.cantidad * item.costoUnitarioCRC;
      const subUSD = item.cantidad * item.costoUnitarioUSD;
      totalCRC += subCRC;
      totalUSD += subUSD;

      const imgUrl = formatearUrlImagen(item.imagenUrl);
      const imgHtml = imgUrl
        ? `<img src="${imgUrl}" alt="${item.nombre}" class="w-8 h-8 rounded-lg object-cover bg-slate-900 border border-slate-700 shrink-0" onerror="this.outerHTML='🍷'">`
        : `🍷`;

      return `
        <div class="p-2 bg-slate-900/90 rounded-xl border border-slate-800 flex items-center justify-between gap-2">
          <div class="flex items-center gap-2 min-w-0 flex-1">
            <div class="w-8 h-8 rounded-lg bg-slate-800 border border-slate-700 flex items-center justify-center shrink-0 text-xs">${imgHtml}</div>
            <div class="min-w-0 flex-1">
              <h5 class="text-xs font-bold text-white truncate">${item.nombre}</h5>
              <div class="text-[10px] text-slate-400 font-mono">Costo: ₡${item.costoUnitarioCRC.toLocaleString()} ($${item.costoUnitarioUSD.toFixed(2)})</div>
            </div>
          </div>

          <div class="flex items-center gap-1.5 bg-slate-800 rounded-lg p-1">
            <button type="button" onclick="modificarCantidadItemCompra('${item.codigo}', -1)" class="w-5 h-5 rounded bg-slate-700 text-white font-bold text-xs flex items-center justify-center active:scale-95">-</button>
            <span class="text-xs font-bold text-white w-4 text-center font-mono">${item.cantidad}</span>
            <button type="button" onclick="modificarCantidadItemCompra('${item.codigo}', 1)" class="w-5 h-5 rounded bg-slate-700 text-white font-bold text-xs flex items-center justify-center active:scale-95">+</button>
          </div>

          <div class="text-right min-w-[65px] font-mono">
            <div class="text-xs font-bold text-emerald-400">${fmtCRC(subCRC)}</div>
            <button type="button" onclick="quitarItemDeListaCompra('${item.codigo}')" class="text-[10px] text-rose-400 hover:text-rose-300">Quitar</button>
          </div>
        </div>
      `;
    }).join("");
  }

  if (totalCRCEl) totalCRCEl.textContent = fmtCRC(totalCRC);
  if (totalUSDEl) totalUSDEl.textContent = `${fmtUSD(totalUSD)} USD`;
  inicializarIconos();
}

function limpiarSeleccionProductoCompra(hacerFocus = true) {
  const busqueda = document.getElementById("compraProductoBusqueda");
  if (busqueda) busqueda.value = "";
  document.getElementById("compraItemEditor")?.classList.add("hidden");
  document.getElementById("compraProductoSugerencias")?.classList.add("hidden");
  if (hacerFocus && busqueda) busqueda.focus();
}

function cerrarDropdownProductosCompra(e) {
  if (!e.target.closest("#compraProductoSugerencias") && !e.target.closest("#compraProductoBusqueda")) {
    const dd = document.getElementById("compraProductoSugerencias");
    if (dd) dd.classList.add("hidden");
  }
}

async function guardarCompra() {
  const items = state.listaCompraActual || [];
  if (items.length === 0) {
    mostrarToast("Agrega al menos un producto a la lista de compra", "error");
    return;
  }

  const fecha = document.getElementById("compraFecha").value || todayStr();
  const vendedor = (document.getElementById("compraVendedor") ? document.getElementById("compraVendedor").value : state.vendedorActual) || "Carlos";
  const pagadoPor = (document.getElementById("compraFinanciadoPor") ? document.getElementById("compraFinanciadoPor").value : vendedor) || "Carlos";
  const tc = Number(document.getElementById("compraTipoCambio").value) || Number(state.config.tipoCambio) || 520;
  const proveedor = document.getElementById("compraProveedor").value.trim();
  const notas = document.getElementById("compraNotas").value.trim();

  let totalCant = 0;
  let totalUSD = 0;
  let totalCRC = 0;

  const itemsNormalizados = items.map(it => {
    totalCant += it.cantidad;
    const subUSD = it.cantidad * it.costoUnitarioUSD;
    const subCRC = it.cantidad * it.costoUnitarioCRC;
    totalUSD += subUSD;
    totalCRC += subCRC;
    return {
      codigo: it.codigo,
      nombre: it.nombre,
      vendedor,
      pagadoPor,
      cantidad: it.cantidad,
      costoUnitarioUSD: it.costoUnitarioUSD,
      tipoCambio: tc,
      costoUnitarioCRC: it.costoUnitarioCRC
    };
  });

  const idCompra = "CMP-" + Date.now().toString().slice(-6);

  const esCredito = document.getElementById("compraEsCredito") ? document.getElementById("compraEsCredito").checked : false;

  const compraObj = {
    id: idCompra,
    codigo: items[0].codigo,
    nombre: items.length === 1 ? items[0].nombre : `${items[0].nombre} +${items.length - 1} licores`,
    fecha,
    vendedor,
    pagadoPor: esCredito ? "Pendiente (Crédito)" : pagadoPor,
    esCredito,
    cantidad: totalCant,
    costoUnitarioUSD: totalUSD / (totalCant || 1),
    tipoCambio: tc,
    costoUnitarioCRC: totalCRC / (totalCant || 1),
    totalUSD,
    totalCRC,
    proveedor: proveedor || "Proveedor General",
    notas: notas || "",
    items: itemsNormalizados
  };

  // 1. Guardar en compras
  state.compras.unshift(compraObj);
  guardarComprasLocal();

  // 2. Si es a crédito, registrar automáticamente en Cuentas por Pagar
  if (esCredito) {
    const cuentaObj = {
      id: "CTA-" + Date.now().toString().slice(-6),
      fecha,
      tipo: "Por Pagar",
      entidad: proveedor || "Proveedor General",
      telefono: "",
      referenciaId: idCompra,
      montoTotalCRC: totalCRC,
      montoTotalUSD: totalUSD,
      saldoPendienteCRC: totalCRC,
      saldoPendienteUSD: totalUSD,
      estado: "Pendiente",
      fechaVencimiento: "",
      vendedor,
      notas: `Compra ${idCompra} (${compraObj.nombre}) a crédito a proveedor`
    };
    state.cuentas.unshift(cuentaObj);
    guardarCuentasLocal();
    encolarAccionSincronizacion("registrarCuenta", { cuenta: cuentaObj });
  }

  // 3. Limpiar lista de compra actual y campos
  state.listaCompraActual = [];
  document.getElementById("compraProveedor").value = "";
  document.getElementById("compraNotas").value = "";
  const chkCredito = document.getElementById("compraEsCredito");
  if (chkCredito) chkCredito.checked = false;
  limpiarSeleccionProductoCompra(false);
  renderizarListaCompraActual();

  // 4. Re-renderizar todo
  renderizarTodo();
  
  const detallePago = esCredito ? "A Crédito (Registrado en Cuentas por Pagar)" : (pagadoPor === "Empresa" ? "Caja Empresa" : `Financiada por ${pagadoPor}`);
  mostrarToast(`¡Compra de ${totalCant} botellas guardada! (${detallePago}) 📦`, "success");

  if (window.confetti) {
    window.confetti({ particleCount: 70, spread: 60, origin: { y: 0.8 } });
  }

  // 5. Encolar y sincronizar con Google Sheets
  encolarAccionSincronizacion("registrarCompra", { compra: compraObj });
}

async function eliminarCompra(id) {
  if (!confirm("¿Deseas eliminar este registro de compra?")) return;
  state.compras = state.compras.filter(c => c.id !== id);
  guardarComprasLocal();
  renderizarTodo();
  mostrarToast("Compra eliminada localmente.", "info");

  // Encolar y sincronizar
  encolarAccionSincronizacion("eliminarCompra", { id });
}

function pasarCompraACuentasPorPagar(idCompra) {
  const c = state.compras.find(x => x.id === idCompra);
  if (!c) return;

  // Verificar si ya existe en cuentas
  const yaExiste = state.cuentas.find(cta => cta.referenciaId === idCompra && cta.tipo === "Por Pagar");
  if (yaExiste) {
    mostrarToast("Esta compra ya está registrada en Cuentas por Pagar.", "info");
    cambiarVista("cuentas");
    return;
  }

  const cant = Number(c.cantidad || 0);
  const cUSD = Number(c.costoUnitarioUSD || 0);
  const tc = Number(c.tipoCambio || state.config.tipoCambio || 520);
  const cCRC = Number(c.costoUnitarioCRC || (cUSD * tc));
  const totUSD = Number(c.totalUSD || (cant * cUSD));
  const totCRC = Number(c.totalCRC || (cant * cCRC));

  const cuentaObj = {
    id: "CTA-" + Date.now().toString().slice(-6),
    fecha: c.fecha || todayStr(),
    tipo: "Por Pagar",
    entidad: c.proveedor || "Proveedor General",
    telefono: "",
    referenciaId: c.id,
    montoTotalCRC: totCRC,
    montoTotalUSD: totUSD,
    saldoPendienteCRC: totCRC,
    saldoPendienteUSD: totUSD,
    estado: "Pendiente",
    fechaVencimiento: "",
    vendedor: c.vendedor || "Carlos",
    notas: `Compra ${c.id} (${c.nombre || c.codigo}) agregada a cuentas por pagar`
  };

  state.cuentas.unshift(cuentaObj);
  guardarCuentasLocal();
  encolarAccionSincronizacion("registrarCuenta", { cuenta: cuentaObj });
  renderizarTodo();
  mostrarToast(`Compra ${c.id} agregada a Cuentas por Pagar 📋`, "success");
  cambiarVista("cuentas");
}

function pasarVentaIndividualACuentasPorCobrar(idxVenta) {
  const v = state.ventas[idxVenta];
  if (!v) {
    mostrarToast("Venta no encontrada.", "error");
    return;
  }

  const vUid = v.id ? `${v.id}_${v.codigo || ''}_${idxVenta}` : `VTA_ROW_${idxVenta}`;

  // Verificar si ESTA venta exacta ya existe en cuentas
  const yaExiste = (state.cuentas || []).find(cta => 
    cta.referenciaId === vUid || 
    (cta.referenciaId === v.id && (!cta.notas || cta.notas.includes(v.codigo || '')))
  );
  
  if (yaExiste) {
    mostrarToast("Esta venta ya está registrada en Cuentas por Cobrar.", "info");
    cambiarVista("cuentas");
    return;
  }

  const totCRC = Number(v.totalCRC || 0);
  const totUSD = Number(v.totalUSD || 0);
  const cliNombre = v.cliente || "Cliente General";
  
  if (!cliNombre || cliNombre.toLowerCase() === "cliente general") {
    mostrarToast("⚠️ No se puede pasar a CXC: la venta no tiene cliente específico asignado.", "error");
    return;
  }

  const nombreProd = v.nombre || (v.codigo ? `Licor (${v.codigo})` : "Venta");
  const cantProd = Number(v.cantidad || 1);
  
  // Buscar teléfono si está en el maestro de clientes
  let tel = "";
  if (v.clienteTelefono) {
    tel = v.clienteTelefono;
  } else if (state.clientes) {
    const matchCli = Object.values(state.clientes).find(c => c.nombre && c.nombre.toLowerCase() === cliNombre.toLowerCase());
    if (matchCli) tel = matchCli.telefono || "";
  }

  const cuentaObj = {
    id: "CTA-" + Date.now().toString().slice(-6),
    fecha: v.fecha || new Date().toISOString(),
    tipo: "Por Cobrar",
    entidad: cliNombre,
    telefono: tel,
    referenciaId: vUid, // ID único por fila para no colisionar con otras ventas del mismo cliente
    montoTotalCRC: totCRC,
    montoTotalUSD: totUSD,
    saldoPendienteCRC: totCRC,
    saldoPendienteUSD: totUSD,
    estado: "Pendiente",
    fechaVencimiento: "",
    vendedor: v.vendedor || "Carlos",
    notas: `Venta: ${cantProd}x ${nombreProd} (${v.metodoPago || 'Efectivo'})`
  };

  if (!state.cuentas) state.cuentas = [];
  state.cuentas.unshift(cuentaObj);
  guardarCuentasLocal();
  encolarAccionSincronizacion("registrarCuenta", { cuenta: cuentaObj });
  renderizarTodo();
  mostrarToast(`Venta (${cantProd}x ${nombreProd}) de ${cliNombre} agregada a Cuentas por Cobrar 📋`, "success");
  cambiarVista("cuentas");
}

function pasarVentaACuentasPorCobrar(idVenta) {
  const idx = state.ventas.findIndex(x => x.id === idVenta);
  if (idx !== -1) {
    pasarVentaIndividualACuentasPorCobrar(idx);
  }
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
    const totUSD = Number(c.totalUSD || (cant * cUSD));
    const totCRC = Number(c.totalCRC || (cant * cCRC));
    const vend = c.vendedor || "Carlos";
    const pagador = c.pagadoPor || vend;
    const vendColor = vend === "Daniel" ? "text-violet-400 bg-violet-950/60 border-violet-500/30" : "text-blue-400 bg-blue-950/60 border-blue-500/30";
    const esPend = pagador.includes("Crédito") || pagador.includes("Pendiente") || c.esCredito;
    const pagoColor = esPend ? "text-amber-300 bg-amber-950/60 border-amber-500/40" : (pagador === "Empresa" ? "text-emerald-300 bg-emerald-950/60 border-emerald-500/30" : pagador === "Daniel" ? "text-violet-300 bg-violet-950/60 border-violet-500/30" : "text-blue-300 bg-blue-950/60 border-blue-500/30");

    const yaEnCuentas = state.cuentas.some(cta => cta.referenciaId === c.id && cta.tipo === "Por Pagar");

    return `
      <div class="p-2.5 bg-slate-900/80 rounded-xl border border-slate-800 flex justify-between items-center gap-2">
        <div class="min-w-0 flex-1">
          <div class="flex items-center gap-1.5 mb-1 flex-wrap">
            <span class="text-[10px] font-bold px-1.5 py-0.2 rounded border ${vendColor}">📦 Stock: ${vend}</span>
            <span class="text-[10px] font-bold px-1.5 py-0.2 rounded border ${pagoColor}">💳 Pagó: ${pagador}</span>
            <span class="text-[10px] text-slate-400 font-mono">${c.fecha || todayStr()}</span>
          </div>
          <div class="font-bold text-white truncate text-xs">${nombre} <span class="text-emerald-400 font-mono font-black">(+${cant})</span></div>
          <div class="text-[10px] text-slate-500 font-mono">${c.codigo} • Proveedor: <b>${c.proveedor || "General"}</b></div>
        </div>
        <div class="text-right font-mono shrink-0 ml-2 space-y-1">
          <div class="font-black text-white text-xs">${fmtUSD(totUSD)}</div>
          <div class="text-[10px] text-slate-400">${fmtCRC(totCRC)}</div>
          <div class="flex items-center justify-end gap-1.5 pt-0.5">
            ${!yaEnCuentas ? `
              <button onclick="pasarCompraACuentasPorPagar('${c.id}')" title="Agregar a Cuentas por Pagar" class="text-[10px] text-amber-400 hover:text-amber-300 font-bold bg-amber-950/40 px-1.5 py-0.5 rounded border border-amber-500/30">
                + Cta Pagar
              </button>
            ` : `
              <span class="text-[9px] text-emerald-400 font-bold font-sans">En Cuentas</span>
            `}
            <button onclick="eliminarCompra('${c.id}')" class="text-[10px] text-rose-400 hover:text-rose-300">Eliminar</button>
          </div>
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
    dropdown.innerHTML = `<div class="p-3 text-xs text-slate-400 text-center">No se encontró "${txt}".</div>`;
    dropdown.classList.remove("hidden");
    return;
  }

  dropdown.innerHTML = matches.map(p => {
    const st = stockMap[p.codigo] || 0;
    const imgUrl = formatearUrlImagen(p.imagenUrl);
    const imgHtml = imgUrl
      ? `<img src="${imgUrl}" alt="${p.nombre}" class="w-9 h-9 rounded-lg object-cover bg-slate-900 border border-slate-700 shrink-0" onerror="this.outerHTML='<div class=\\'w-9 h-9 rounded-lg bg-slate-900 border border-slate-700 flex items-center justify-center text-slate-500 shrink-0\\'>🍷</div>'">`
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
    mostrarToast("Producto no encontrado.", "error");
    return;
  }

  const stockMap = calcularStockPorCodigo();
  const stockDisponible = stockMap[codigo] || 0;

  const enCarrito = state.carrito.find(item => item.codigo === codigo);

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

  if (stockDisponible <= 0) {
    mostrarToast(`Agregado: ${prod.nombre} (⚠️ Sin stock registrado).`, "info");
  } else {
    mostrarToast(`Agregado: ${prod.nombre}.`, "success");
  }

  reproducirBeep();
  const searchInput = document.getElementById("searchPos");
  if (searchInput) searchInput.value = "";
  const searchResults = document.getElementById("posSearchResults");
  if (searchResults) {
    searchResults.classList.add("hidden");
    searchResults.innerHTML = "";
  }
  renderizarCarrito();
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

function cambiarModoPOS(modo) {
  state.modoPOS = modo;
  const btnVenta = document.getElementById("btnModoVenta");
  const btnPedido = document.getElementById("btnModoPedido");
  const btnCheckout = document.getElementById("btnCheckout");
  const cartIcon = document.getElementById("cartHeaderIcon");
  const cartTitle = document.getElementById("cartHeaderTitle");
  const panelPuntos = document.getElementById("panelPuntosCliente");
  const totalesYPagos = document.getElementById("posTotalesYPagosContainer");
  const bannerPedido = document.getElementById("posBannerModoPedido");

  if (modo === "pedido") {
    if (btnVenta) {
      btnVenta.className = "py-2 rounded-lg bg-transparent text-slate-400 hover:text-white flex items-center justify-center gap-1.5 active:scale-95 transition-all";
    }
    if (btnPedido) {
      btnPedido.className = "py-2 rounded-lg bg-amber-600 text-white shadow-md flex items-center justify-center gap-1.5 active:scale-95 transition-all";
    }
    if (btnCheckout) {
      btnCheckout.className = "w-full py-3.5 bg-gradient-to-r from-amber-500 to-orange-600 text-white text-sm font-extrabold rounded-xl shadow-lg shadow-amber-500/25 active:scale-95 transition-all flex items-center justify-center gap-2";
      btnCheckout.innerHTML = `<i data-lucide="clipboard-check" class="w-5 h-5"></i><span>GUARDAR ENCARGO DE BOTELLAS</span>`;
    }
    if (cartIcon) cartIcon.className = "w-4 h-4 text-amber-400";
    if (cartTitle) cartTitle.innerHTML = `Lista de Encargo (<span id="cartCount">${state.carrito.length}</span>)`;
    
    // Ocultar montos, métodos de pago y puntos
    if (totalesYPagos) totalesYPagos.classList.add("hidden");
    if (panelPuntos) panelPuntos.classList.add("hidden");
    if (bannerPedido) bannerPedido.classList.remove("hidden");

    mostrarToast("Modo 'Encargo / Pedido' (solo cantidades) 📋", "info");
  } else {
    if (btnVenta) {
      btnVenta.className = "py-2 rounded-lg bg-emerald-600 text-white shadow-md flex items-center justify-center gap-1.5 active:scale-95 transition-all";
    }
    if (btnPedido) {
      btnPedido.className = "py-2 rounded-lg bg-transparent text-slate-400 hover:text-white flex items-center justify-center gap-1.5 active:scale-95 transition-all";
    }
    if (btnCheckout) {
      btnCheckout.className = "w-full py-3.5 bg-gradient-to-r from-emerald-500 to-teal-600 text-white text-sm font-extrabold rounded-xl shadow-lg shadow-emerald-500/25 active:scale-95 transition-all flex items-center justify-center gap-2";
      btnCheckout.innerHTML = `<i data-lucide="check" class="w-5 h-5"></i><span>COMPLETAR VENTA</span>`;
    }
    if (cartIcon) cartIcon.className = "w-4 h-4 text-emerald-400";
    if (cartTitle) cartTitle.innerHTML = `Carrito de Venta (<span id="cartCount">${state.carrito.length}</span>)`;
    
    // Mostrar montos, métodos de pago y puntos
    if (totalesYPagos) totalesYPagos.classList.remove("hidden");
    if (bannerPedido) bannerPedido.classList.add("hidden");
    if (state.clienteSeleccionado && panelPuntos) panelPuntos.classList.remove("hidden");

    mostrarToast("Modo 'Venta Directa' activado 🛍️", "info");
  }

  inicializarIconos();
  renderizarCarrito();
}

function renderizarCarrito() {
  const cont = document.getElementById("cartItemsList");
  const countEl = document.getElementById("cartCount");
  const totalCRCEl = document.getElementById("cartTotalCRC");
  const totalUSDEl = document.getElementById("cartTotalUSD");
  const pedidoTotalUnidades = document.getElementById("pedidoTotalUnidades");
  if (!cont) return;

  const esModoPedido = state.modoPOS === "pedido";

  let totalBrutoCRC = 0;
  let totalUSD = 0;
  let totalItems = 0;

  state.carrito.forEach(i => {
    totalBrutoCRC += (i.cantidad * i.precioVentaCRC);
    totalUSD += (i.cantidad * i.precioVentaUSD);
    totalItems += i.cantidad;
  });

  const descuento = state.descuentoPuntosAplicado || 0;
  const totalFinalCRC = Math.max(0, totalBrutoCRC - descuento);
  const tc = state.config.tipoCambio || 520;
  const totalFinalUSD = totalFinalCRC / tc;

  if (countEl) countEl.textContent = totalItems;
  if (pedidoTotalUnidades) pedidoTotalUnidades.textContent = `${totalItems} unids`;

  if (totalCRCEl && totalUSDEl) {
    if (descuento > 0) {
      totalCRCEl.innerHTML = `
        <span class="line-through text-slate-500 text-base font-bold">${fmtCRC(totalBrutoCRC)}</span>
        <span class="text-emerald-400 text-2xl font-black">${fmtCRC(totalFinalCRC)}</span>
        <span class="block text-[10px] text-amber-400 font-normal">-${fmtCRC(descuento)} descuento de puntos 🎁</span>
      `;
      totalUSDEl.textContent = `${fmtUSD(totalFinalUSD)} USD`;
    } else {
      totalCRCEl.textContent = fmtCRC(totalFinalCRC);
      totalUSDEl.textContent = `${fmtUSD(totalUSD)} USD`;
    }
  }

  if (state.carrito.length === 0) {
    cont.innerHTML = `
      <div class="flex flex-col items-center justify-center py-6 text-slate-500 text-xs">
        <i data-lucide="${esModoPedido ? 'clipboard-list' : 'shopping-cart'}" class="w-8 h-8 stroke-1 mb-1 text-slate-600"></i>
        <span>${esModoPedido ? 'Lista de encargo vacía. Agrega los licores pedidos.' : 'Carrito vacío. Agrega licores para vender.'}</span>
      </div>
    `;
  } else {
    cont.innerHTML = state.carrito.map(item => {
      const imgUrl = formatearUrlImagen(item.imagenUrl);
      const imgHtml = imgUrl
        ? `<img src="${imgUrl}" alt="${item.nombre}" class="w-8 h-8 rounded-lg object-cover bg-slate-900 border border-slate-700 shrink-0" onerror="this.outerHTML='<div class=\\'w-8 h-8 rounded-lg bg-slate-900 border border-slate-700 flex items-center justify-center text-slate-500 shrink-0 text-xs\\'>🍷</div>'">`
        : `<div class="w-8 h-8 rounded-lg bg-slate-900 border border-slate-700 flex items-center justify-center text-slate-500 shrink-0 text-xs">🍷</div>`;

      const editado = item._precioEditado 
        ? 'text-amber-300 font-bold bg-amber-500/20 border-amber-500/40 hover:bg-amber-500/30' 
        : 'text-slate-400 bg-slate-800/80 border-slate-700/60 hover:text-amber-300 hover:bg-slate-750';

      return `
        <div class="p-2.5 bg-slate-900/90 rounded-xl border border-slate-800 flex items-center justify-between gap-2">
          <div class="flex items-center gap-2.5 min-w-0 flex-1">
            ${imgHtml}
            <div class="min-w-0 flex-1">
              <h5 class="text-xs font-bold text-white truncate">${item.nombre}</h5>
              ${esModoPedido ? `
                <div class="text-[10px] text-amber-400 font-mono">Encargo: <b class="text-white">${item.cantidad} botella(s)</b></div>
              ` : `
                <button type="button" onclick="editarPrecioCarrito('${item.codigo}')" class="text-[11px] font-mono flex items-center gap-1.5 px-2 py-0.5 rounded-lg border transition-all active:scale-95 mt-0.5 ${editado}" title="Clic para editar precio">
                  <span>${fmtCRC(item.precioVentaCRC)}</span>
                  <span class="text-[10px] text-amber-400 flex items-center gap-0.5">✏️ ${item._precioEditado ? '<span class="text-[9px] font-sans font-bold text-amber-300 uppercase">Editado</span>' : ''}</span>
                </button>
              `}
            </div>
          </div>

          <div class="flex items-center gap-1.5 bg-slate-800 rounded-lg p-1">
            <button onclick="modificarCantidadCarrito('${item.codigo}', -1)" class="w-6 h-6 rounded bg-slate-700 text-white font-bold text-xs flex items-center justify-center active:scale-95">-</button>
            <span class="text-xs font-bold text-white w-5 text-center font-mono">${item.cantidad}</span>
            <button onclick="modificarCantidadCarrito('${item.codigo}', 1)" class="w-6 h-6 rounded bg-slate-700 text-white font-bold text-xs flex items-center justify-center active:scale-95">+</button>
          </div>

          <div class="text-right min-w-[60px] font-mono">
            ${esModoPedido ? `
              <span class="text-xs font-black text-amber-400">${item.cantidad}x</span>
            ` : `
              <div class="text-xs font-black text-emerald-400">${fmtCRC(item.cantidad * item.precioVentaCRC)}</div>
            `}
            <button onclick="eliminarDelCarrito('${item.codigo}')" class="text-[10px] text-rose-400 hover:text-rose-300 block ml-auto">Quitar</button>
          </div>
        </div>
      `;
    }).join("");
  }

  // Actualizar panel de fidelización si hay cliente seleccionado y no es pedido
  if (state.clienteSeleccionado && !esModoPedido) {
    renderizarPanelCliente();
  }

  calcularCambio();
  inicializarIconos();
}
function setPaymentMethod(metodo) {
  state.metodoPagoSeleccionado = metodo;
  document.querySelectorAll(".pay-btn").forEach(btn => {
    const txt = btn.textContent.trim().toLowerCase();
    const target = metodo.toLowerCase();
    if (txt.includes(target) || (target === "sinpe / transf." && txt.includes("sinpe")) || (target === "pago luego" && txt.includes("luego"))) {
      btn.classList.add("active");
      btn.classList.remove("bg-slate-900", "text-slate-300");
      if (metodo === "Pago Luego") {
        btn.classList.add("bg-amber-600", "text-white");
      } else {
        btn.classList.add("bg-indigo-600", "text-white");
      }
    } else {
      btn.classList.remove("active", "bg-indigo-600", "bg-amber-600", "text-white");
      btn.classList.add("bg-slate-900", "text-slate-300");
    }
  });

  const cashHelper = document.getElementById("cashHelper");
  const pagoLuegoHelper = document.getElementById("pagoLuegoHelper");

  if (cashHelper) {
    if (metodo === "Efectivo") cashHelper.classList.remove("hidden");
    else cashHelper.classList.add("hidden");
  }

  if (pagoLuegoHelper) {
    if (metodo === "Pago Luego") pagoLuegoHelper.classList.remove("hidden");
    else pagoLuegoHelper.classList.add("hidden");
  }
}

function calcularCambio() {
  const recibido = Number(document.getElementById("cashReceived").value) || 0;
  let totalBrutoCRC = 0;
  state.carrito.forEach(i => totalBrutoCRC += (i.cantidad * i.precioVentaCRC));
  const totalFinal = Math.max(0, totalBrutoCRC - (state.descuentoPuntosAplicado || 0));

  const cambio = recibido - totalFinal;
  const cambioEl = document.getElementById("cashChange");
  if (recibido > 0) {
    cambioEl.textContent = fmtCRC(Math.max(0, cambio));
  } else {
    cambioEl.textContent = fmtCRC(0);
  }
}

function cambiarModoPOS(modo) {
  state.modoPOS = modo;
  const btnVenta = document.getElementById("btnModoVenta");
  const btnPedido = document.getElementById("btnModoPedido");
  const btnCheckout = document.getElementById("btnCheckout");
  const cartIcon = document.getElementById("cartHeaderIcon");
  const cartTitle = document.getElementById("cartHeaderTitle");
  const cashHelper = document.getElementById("cashHelper");

  if (modo === "pedido") {
    if (btnVenta) {
      btnVenta.className = "py-2 rounded-lg bg-transparent text-slate-400 hover:text-white flex items-center justify-center gap-1.5 active:scale-95 transition-all";
    }
    if (btnPedido) {
      btnPedido.className = "py-2 rounded-lg bg-amber-600 text-white shadow-md flex items-center justify-center gap-1.5 active:scale-95 transition-all";
    }
    if (btnCheckout) {
      btnCheckout.className = "w-full py-3.5 bg-gradient-to-r from-amber-500 to-orange-600 text-white text-sm font-extrabold rounded-xl shadow-lg shadow-amber-500/25 active:scale-95 transition-all flex items-center justify-center gap-2";
      btnCheckout.innerHTML = `<i data-lucide="clipboard-check" class="w-5 h-5"></i><span>GUARDAR PEDIDO / ENCARGO</span>`;
    }
    if (cartIcon) cartIcon.className = "w-4 h-4 text-amber-400";
    if (cartTitle) cartTitle.innerHTML = `Lista de Encargo (<span id="cartCount">${state.carrito.length}</span>)`;
    if (cashHelper) cashHelper.classList.add("hidden");
    mostrarToast("Modo 'Encargo / Pedido' activado 📋", "info");
  } else {
    if (btnVenta) {
      btnVenta.className = "py-2 rounded-lg bg-emerald-600 text-white shadow-md flex items-center justify-center gap-1.5 active:scale-95 transition-all";
    }
    if (btnPedido) {
      btnPedido.className = "py-2 rounded-lg bg-transparent text-slate-400 hover:text-white flex items-center justify-center gap-1.5 active:scale-95 transition-all";
    }
    if (btnCheckout) {
      btnCheckout.className = "w-full py-3.5 bg-gradient-to-r from-emerald-500 to-teal-600 text-white text-sm font-extrabold rounded-xl shadow-lg shadow-emerald-500/25 active:scale-95 transition-all flex items-center justify-center gap-2";
      btnCheckout.innerHTML = `<i data-lucide="check" class="w-5 h-5"></i><span>COMPLETAR VENTA</span>`;
    }
    if (cartIcon) cartIcon.className = "w-4 h-4 text-emerald-400";
    if (cartTitle) cartTitle.innerHTML = `Carrito de Venta (<span id="cartCount">${state.carrito.length}</span>)`;
    if (state.metodoPagoSeleccionado === "Efectivo" && cashHelper) cashHelper.classList.remove("hidden");
    mostrarToast("Modo 'Venta Directa' activado 🛍️", "info");
  }

  inicializarIconos();
  renderizarCarrito();
}

async function completarVenta() {
  if (state.modoPOS === "pedido") {
    return guardarPedidoCliente();
  }

  if (state.carrito.length === 0) {
    mostrarToast("El carrito está vacío.", "error");
    return;
  }

  const vendedor = state.vendedorActual || "Carlos";
  let totalBrutoCRC = 0, totalUSD = 0;
  let gananciaCRC = 0, gananciaUSD = 0;

  state.carrito.forEach(item => {
    const subCRC = item.cantidad * item.precioVentaCRC;
    const subUSD = item.cantidad * item.precioVentaUSD;
    const cCRC = item.cantidad * (item.costoRefCRC || 0);
    const cUSD = item.cantidad * (item.costoRefUSD || 0);

    totalBrutoCRC += subCRC;
    totalUSD += subUSD;
    gananciaCRC += (subCRC - cCRC);
    gananciaUSD += (subUSD - cUSD);
  });

  // --- Descuento y puntos ---
  const descuentoPuntos = state.descuentoPuntosAplicado || 0;
  const puntosCanjados = Math.floor(descuentoPuntos / (state.config.puntosValorCRC || 1));
  const totalFinalCRC = Math.max(0, totalBrutoCRC - descuentoPuntos);
  const puntosGanados = Math.floor(totalFinalCRC / (state.config.puntosRazonCRC || 100));

  // --- Nombre/ID de cliente ---
  const cli = state.clienteSeleccionado;
  const clienteInputVal = document.getElementById("posClienteInput")?.value?.trim();
  const clienteNombre = cli ? cli.nombre : (clienteInputVal || "Cliente General");
  const clienteId = cli ? cli.id : null;
  const clienteTelefono = cli ? cli.telefono : null;

  // Validación estricta: No se puede vender con 'Pago Luego' si no hay un cliente seleccionado/asignado
  if (state.metodoPagoSeleccionado === "Pago Luego") {
    if (!clienteNombre || clienteNombre.toLowerCase() === "cliente general") {
      mostrarToast("⚠️ Para vender con 'Pago Luego' debes asignar un cliente registrado.", "error");
      const inCli = document.getElementById("posClienteInput");
      if (inCli) {
        inCli.focus();
        inCli.classList.add("border-amber-500", "animate-pulse");
        setTimeout(() => inCli.classList.remove("border-amber-500", "animate-pulse"), 2500);
      }
      return;
    }
  }

  const idVenta = "VTA-" + Date.now().toString().slice(-6);

  const ventaObj = {
    id: idVenta,
    fecha: new Date().toISOString(),
    vendedor,
    items: [...state.carrito],
    totalCRC: totalBrutoCRC,
    totalFinalCRC,
    totalUSD,
    gananciaCRC: gananciaCRC - descuentoPuntos,
    gananciaUSD,
    cliente: clienteNombre,
    clienteId,
    clienteTelefono,
    metodoPago: state.metodoPagoSeleccionado,
    descuentoPuntos,
    puntosGanados,
    puntosCanjados
  };

  // --- Actualizar puntos del cliente ---
  if (cli) {
    state.clientes[cli.id].ultimaVenta = ventaObj.fecha;
    actualizarPuntosCliente(cli.id, puntosGanados - puntosCanjados);
  }

  state.ventas.unshift(ventaObj);
  guardarVentasLocal();
  state.ultimaVentaCompletada = ventaObj;

  // --- Si es "Pago Luego", crear automáticamente registro en Cuentas por Cobrar ---
  if (state.metodoPagoSeleccionado === "Pago Luego") {
    const cuentaObj = {
      id: "CTA-" + Date.now().toString().slice(-6),
      fecha: ventaObj.fecha,
      tipo: "Por Cobrar",
      entidad: clienteNombre,
      telefono: clienteTelefono || "",
      referenciaId: idVenta,
      montoTotalCRC: totalFinalCRC,
      montoTotalUSD: totalUSD,
      saldoPendienteCRC: totalFinalCRC,
      saldoPendienteUSD: totalUSD,
      estado: "Pendiente",
      fechaVencimiento: "",
      vendedor,
      notas: `Venta POS ${idVenta} a crédito / pago posterior`
    };
    state.cuentas.unshift(cuentaObj);
    guardarCuentasLocal();
    encolarAccionSincronizacion("registrarCuenta", { cuenta: cuentaObj });
    mostrarToast(`Venta guardada y añadida a Cuentas por Cobrar (${clienteNombre}) 📋`, "info");
  }

  if (window.confetti) {
    window.confetti({ particleCount: 80, spread: 60, origin: { y: 0.8 } });
  }

  // --- Limpiar estado post-venta ---
  state.carrito = [];
  state.clienteSeleccionado = null;
  state.descuentoPuntosAplicado = 0;
  document.getElementById("cashReceived").value = "";
  const clienteInput = document.getElementById("posClienteInput");
  if (clienteInput) clienteInput.value = "";

  renderizarTodo();
  renderizarPanelCliente();

  abrirModalRecibo(ventaObj);

  // Encolar y sincronizar con Google Sheets
  encolarAccionSincronizacion("registrarVenta", { venta: ventaObj });
}

// ==========================================================================
// GUARDAR PEDIDO DE CLIENTE (ENCARGO)
// ==========================================================================
function guardarPedidoCliente() {
  if (state.carrito.length === 0) {
    mostrarToast("Agrega licores al pedido antes de guardar", "error");
    return;
  }

  const vendedor = state.vendedorActual || "Carlos";
  const cli = state.clienteSeleccionado;
  const clienteInputVal = document.getElementById("posClienteInput")?.value?.trim();
  const clienteNombre = cli ? cli.nombre : (clienteInputVal || "Cliente General");
  const clienteTelefono = cli ? cli.telefono : "";

  let totalCRC = 0;
  let totalUSD = 0;
  state.carrito.forEach(i => {
    totalCRC += (i.cantidad * i.precioVentaCRC);
    totalUSD += (i.cantidad * i.precioVentaUSD);
  });

  const idPedido = "PED-" + Date.now().toString().slice(-6);
  const pedidoObj = {
    id: idPedido,
    fecha: new Date().toISOString(),
    vendedor,
    cliente: clienteNombre,
    clienteId: cli ? cli.id : null,
    clienteTelefono: clienteTelefono,
    items: [...state.carrito],
    totalCRC,
    totalUSD,
    estado: "pendiente" // "pendiente" | "comprado"
  };

  if (!state.pedidos) state.pedidos = [];
  state.pedidos.unshift(pedidoObj);
  guardarPedidosLocal();
  state.ultimoPedidoCompletado = pedidoObj;

  if (window.confetti) {
    window.confetti({ particleCount: 60, spread: 50, origin: { y: 0.8 } });
  }

  // Limpiar carrito y campos
  state.carrito = [];
  state.clienteSeleccionado = null;
  const clienteInput = document.getElementById("posClienteInput");
  if (clienteInput) clienteInput.value = "";

  renderizarTodo();
  renderizarPanelCliente();

  mostrarToast(`📋 Pedido ${idPedido} guardado con éxito. Se agregó al consolidado del proveedor.`, "success");
  abrirModalRecibo(pedidoObj, true);

  // Encolar y sincronizar con Google Sheets
  encolarAccionSincronizacion("registrarPedido", { pedido: pedidoObj });
}

// ==========================================================================
// MODAL RECIBO Y WHATSAPP
// ==========================================================================
function abrirModalRecibo(venta, esPedido = false) {
  const modal = document.getElementById("modalRecibo");
  document.getElementById("reciboNegocio").textContent = state.config.nombreNegocio || "DC EL DESTAPE LICORES";
  document.getElementById("reciboId").textContent = `${venta.id} (👤 ${venta.vendedor || state.vendedorActual})`;
  document.getElementById("reciboFecha").textContent = new Date(venta.fecha).toLocaleString([], { dateStyle: 'short', timeStyle: 'short' });
  document.getElementById("reciboMetodo").textContent = esPedido ? "Encargo de Botellas" : (venta.metodoPago || "Efectivo");
  
  const totalCRCEl = document.getElementById("reciboTotalCRC");
  const totalUSDEl = document.getElementById("reciboTotalUSD");

  if (esPedido) {
    let totalBotellas = 0;
    (venta.items || []).forEach(i => totalBotellas += Number(i.cantidad || 1));
    totalCRCEl.textContent = `${totalBotellas} botella(s)`;
    totalUSDEl.textContent = "(Encargo al proveedor)";
  } else {
    totalCRCEl.textContent = fmtCRC(venta.totalFinalCRC || venta.totalCRC);
    totalUSDEl.textContent = `(${fmtUSD(venta.totalUSD)} USD)`;
  }

  const puntosRow = document.getElementById("reciboPuntosRow");
  const puntosGanadosEl = document.getElementById("reciboPuntosGanados");
  if (puntosRow && puntosGanadosEl) {
    if (!esPedido && venta.puntosGanados && venta.puntosGanados > 0) {
      puntosRow.classList.remove("hidden");
      puntosRow.classList.add("flex");
      puntosGanadosEl.textContent = `+${venta.puntosGanados.toLocaleString()} pts (${venta.cliente || 'Cliente'})`;
    } else {
      puntosRow.classList.add("hidden");
      puntosRow.classList.remove("flex");
    }
  }

  const itemsCont = document.getElementById("reciboItems");
  if (itemsCont) {
    itemsCont.innerHTML = venta.items.map(i => `
      <div class="flex justify-between py-1 font-mono">
        <div>
          <span class="font-bold text-amber-600 font-sans">${i.cantidad}x</span> ${i.nombre}
        </div>
        <span class="font-bold ${esPedido ? 'text-amber-600' : 'text-slate-800'}">${esPedido ? `${i.cantidad} unids` : fmtCRC(i.cantidad * i.precioVentaCRC)}</span>
      </div>
    `).join("");
  }

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
  const v = state.modoPOS === "pedido" ? (state.ultimoPedidoCompletado || state.ultimaVentaCompletada) : (state.ultimaVentaCompletada || state.ultimoPedidoCompletado);
  if (!v) return;

  const esPedido = v.id && v.id.startsWith("PED-");
  const negocio = state.config.nombreNegocio || "DC EL DESTAPE LICORES";
  const telefono = state.config.telefonoNegocio || "+506 8992-7936";
  const fecha = new Date(v.fecha).toLocaleString([], { dateStyle: 'short', timeStyle: 'short' });
  const vendedor = v.vendedor || state.vendedorActual || "Carlos";

  let texto = `🍷 *${negocio.toUpperCase()}* 🍷\n`;
  texto += `📱 *Tel:* ${telefono}\n`;
  texto += `--------------------------------\n`;
  texto += esPedido ? `📋 *COMPROBANTE DE ENCARGO*\n` : `🧾 *COMPROBANTE DE COMPRA*\n`;
  texto += `📅 Fecha: ${fecha}\n`;
  texto += `🎫 N°: ${v.id}\n`;
  texto += `👤 Atendido por: ${vendedor}\n`;
  texto += `👤 Cliente: ${v.cliente || "General"}\n`;
  texto += `--------------------------------\n`;
  
  let totalBotellas = 0;
  v.items.forEach(i => {
    totalBotellas += Number(i.cantidad || 1);
    if (esPedido) {
      texto += `• *${i.cantidad}x* ${i.nombre}\n`;
    } else {
      texto += `• ${i.cantidad}x ${i.nombre} = ${fmtCRC(i.cantidad * i.precioVentaCRC)} (${fmtUSD(i.cantidad * i.precioVentaUSD)})\n`;
    }
  });

  if (!esPedido && v.descuentoPuntos && v.descuentoPuntos > 0) {
    texto += `🎁 *Descuento Puntos:* -${fmtCRC(v.descuentoPuntos)}\n`;
  }
  if (!esPedido && v.puntosGanados && v.puntosGanados > 0) {
    texto += `✨ *Puntos Ganados:* +${v.puntosGanados.toLocaleString()} pts\n`;
  }

  texto += `--------------------------------\n`;
  if (esPedido) {
    texto += `📦 *TOTAL BOTELLAS ENCARGADAS:* ${totalBotellas} unids\n`;
    texto += `📌 *Estado:* Pedido registrado (en gestión con proveedor)\n\n`;
    texto += `¡Hemos anotado tu pedido de licores! Te contactaremos tan pronto las tengamos disponibles. 🍷`;
  } else {
    texto += `💳 *Método de Pago:* ${v.metodoPago || "Efectivo"}\n`;
    texto += `💵 *TOTAL CRC:* ${fmtCRC(v.totalFinalCRC || v.totalCRC)}\n`;
    texto += `💵 *TOTAL USD:* ${fmtUSD(v.totalUSD)}\n\n`;
    texto += `¡Muchas gracias por su preferencia! 🍷`;
  }

  const encoded = encodeURIComponent(texto);
  const phoneClient = v.clienteTelefono ? v.clienteTelefono.replace(/[^0-9]/g, "") : "";
  const url = phoneClient ? `https://wa.me/506${phoneClient}?text=${encoded}` : `https://wa.me/?text=${encoded}`;
  window.open(url, "_blank");
}
function calcularSaldosFinancieros() {
  const tcActual = Number(state.config.tipoCambio) || 520;

  // 1. Total Ventas Facturadas
  let totalVentasCRC = 0;
  let totalVentasUSD = 0;

  state.ventas.forEach(v => {
    totalVentasCRC += Number(v.totalCRC) || 0;
    totalVentasUSD += Number(v.totalUSD) || ((Number(v.totalCRC) || 0) / tcActual);
  });

  // 1.1 Cuentas por Cobrar Pendientes (Dinero que aún no ha ingresado físicamente a caja)
  let totalCxcPendienteCRC = 0;
  let totalCxcPendienteUSD = 0;

  (state.cuentas || []).forEach(cta => {
    if (cta.tipo === "Por Cobrar" && (cta.estado || "Pendiente") !== "Pagado") {
      const sCRC = Number(cta.saldoPendienteCRC || 0);
      const sUSD = Number(cta.saldoPendienteUSD || 0) || (tcActual > 0 ? (sCRC / tcActual) : 0);
      if (sCRC > 0) {
        totalCxcPendienteCRC += sCRC;
        totalCxcPendienteUSD += sUSD;
      }
    }
  });

  // Ventas efectivamente cobradas en caja = Total Ventas - Saldo Pendiente de Cobro
  const ventasEfectivamenteCobradasCRC = Math.max(0, totalVentasCRC - totalCxcPendienteCRC);
  const ventasEfectivamenteCobradasUSD = tcActual > 0 ? (ventasEfectivamenteCobradasCRC / tcActual) : 0;

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

  // Saldo real en caja de la Empresa (Solo dinero cobrado efectivamente)
  const totalAportesSociosCRC = carlosAportesDirectosCRC + danielAportesDirectosCRC;
  const totalAportesSociosUSD = carlosAportesDirectosUSD + danielAportesDirectosUSD;
  const totalReembolsosSociosCRC = carlosReembolsosCRC + danielReembolsosCRC;
  const totalReembolsosSociosUSD = carlosReembolsosUSD + danielReembolsosUSD;

  // Saldo Real = Ventas Cobradas en Mano + Capital Propio Empresa + Aportes Socios - Compras Empresa - Reembolsos - Gastos
  const saldoEmpresaCRC = ventasEfectivamenteCobradasCRC + empresaCapitalPropioCRC + totalAportesSociosCRC - empresaPagaComprasCRC - totalReembolsosSociosCRC - empresaGastosCRC;
  const saldoEmpresaUSD = tcActual > 0 ? (saldoEmpresaCRC / tcActual) : 0;

  return {
    tcActual,
    totalVentasCRC,
    totalVentasUSD: tcActual > 0 ? (totalVentasCRC / tcActual) : 0,
    cxcPendienteCRC: totalCxcPendienteCRC,
    cxcPendienteUSD: totalCxcPendienteUSD,
    ventasCobradasCRC: ventasEfectivamenteCobradasCRC,
    ventasCobradasUSD: ventasEfectivamenteCobradasUSD,
    empresa: {
      saldoCRC: saldoEmpresaCRC,
      saldoUSD: saldoEmpresaUSD,
      ventasCRC: totalVentasCRC,
      cxcPendienteCRC: totalCxcPendienteCRC,
      cxcPendienteUSD: totalCxcPendienteUSD,
      gastosCRC: empresaGastosCRC + empresaPagaComprasCRC
    },
    carlos: {
      deudaCRC: carlosDeudaCRC,
      deudaUSD: tcActual > 0 ? (carlosDeudaCRC / tcActual) : 0,
      totalAportadoCRC: carlosTotalAportadoCRC,
      totalAportadoUSD: tcActual > 0 ? (carlosTotalAportadoCRC / tcActual) : 0,
      reembolsadoCRC: carlosReembolsosCRC
    },
    daniel: {
      deudaCRC: danielDeudaCRC,
      deudaUSD: tcActual > 0 ? (danielDeudaCRC / tcActual) : 0,
      totalAportadoCRC: danielTotalAportadoCRC,
      totalAportadoUSD: tcActual > 0 ? (danielTotalAportadoCRC / tcActual) : 0,
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

  // Rubro Extra: Cuentas por Cobrar Pendientes (Dinero que falta por ingresar)
  const elCxcCRC = document.getElementById("finCxcPendienteCRC");
  const elCxcUSD = document.getElementById("finCxcPendienteUSD");
  const elCxcBadge = document.getElementById("finCxcBadge");
  if (elCxcCRC) elCxcCRC.textContent = fmtCRC(fin.empresa.cxcPendienteCRC);
  if (elCxcUSD) elCxcUSD.textContent = `(${fmtUSD(fin.empresa.cxcPendienteUSD)} USD)`;
  if (elCxcBadge) {
    if (fin.empresa.cxcPendienteCRC > 0) {
      elCxcBadge.classList.remove("hidden");
    } else {
      elCxcBadge.classList.add("hidden");
    }
  }

  // Desglose de Ventas Cobradas
  const elVentasCobradas = document.getElementById("finVentasCobradasCRC");
  if (elVentasCobradas) elVentasCobradas.textContent = fmtCRC(fin.ventasCobradasCRC);

  const elVentasFacturadas = document.getElementById("finVentasFacturadasCRC");
  if (elVentasFacturadas) elVentasFacturadas.textContent = fmtCRC(fin.totalVentasCRC);

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
    let subtitulo = `${m.socio || "Socio"} -> Caja Empresa`;
    let signoMonto = "+";
    let colorMonto = "text-emerald-400";

    if (isVenta) {
      const vendBadge = m.socio === "Daniel" ? "text-violet-300" : "text-blue-300";
      colorBadge = "bg-emerald-950/80 text-emerald-300 border-emerald-500/40 shadow-sm shadow-emerald-500/20";
      iconName = "shopping-cart";
      tituloTipo = `Ingreso x Venta (${m.socio || "POS"})`;
      subtitulo = `Venta POS <span class="${vendBadge} font-bold">(${m.socio})</span> -> Caja Empresa`;
      signoMonto = "+";
      colorMonto = "text-emerald-400 font-black";
    } else if (isPagoSocio) {
      colorBadge = "bg-emerald-950/60 text-emerald-400 border-emerald-500/30";
      iconName = "arrow-up-right";
      tituloTipo = "Pago / Abono a Socio";
      subtitulo = `Caja Empresa -> ${m.socio || "Socio"}`;
      signoMonto = "-";
      colorMonto = "text-blue-400 font-black";
    } else if (isGasto) {
      colorBadge = "bg-rose-950/60 text-rose-400 border-rose-500/30";
      iconName = "receipt";
      tituloTipo = "Gasto Operativo";
      subtitulo = `Caja Empresa -> Gastos`;
      signoMonto = "-";
      colorMonto = "text-rose-400 font-black";
    } else if (isCompraEmpresa) {
      colorBadge = "bg-amber-950/60 text-amber-300 border-amber-500/30";
      iconName = "truck";
      tituloTipo = "Compra Pagada x Empresa";
      subtitulo = `Caja Empresa -> Proveedor`;
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
        ${isVenta ? 'POS 🛒' : 'COMPRA 📦'}
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
    mostrarToast("El monto debe ser mayor a 0.", "error");
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
    : `Gasto registrado (-${fmtCRC(montoCRC)}).`;

  mostrarToast(msg, "success");

  // Encolar y sincronizar con Google Sheets
  encolarAccionSincronizacion("registrarMovimiento", { movimiento: movObj });
}

async function eliminarMovimientoDinero(id) {
  if (!confirm("¿Deseas eliminar este movimiento financiero?")) return;

  state.movimientosDinero = state.movimientosDinero.filter(m => m.id !== id);
  guardarFinanzasLocal();
  renderizarFinanzas();
  mostrarToast("Movimiento eliminado localmente.", "info");

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
        mostrarToast(`¡Importación exitosa! (${importadosCount} productos).`, "success");
      } catch (err) {
        mostrarToast("Error al procesar las hojas del Excel.", "error");
      }
    };
    reader.readAsArrayBuffer(file);
  } catch (e) {
    mostrarToast("No se pudo leer el archivo.", "error");
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
    mostrarToast("No se pudo acceder a la cámara.", "error");
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
    seleccionarProductoCompraPorCodigo(codigo);
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
      mostrarToast("Configura la URL de Google Sheets en Ajustes ⚙️.", "error");
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
      if (mostrarFeedback) {
        mostrarToast("¡Todos los cambios sin internet se sincronizaron con Google Sheets! 🚀", "success");
      }
      // Descargar datos frescos sin volver a procesar cola (ya está vacía)
      sincronizandoCola = false; // liberar flag antes del GET
      await _descargarDatosSheets(false);
      return;
    } else {
      mostrarToast(`Quedan ${state.colaSincronizacion.length} cambios pendientes por sincronizar.`, "info");
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
      bannerText.textContent = `🟡 ${pendingCount} cambio(s) pendiente(s) por subir a Sheets.`;
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
  if (!state.config.sheetsUrl) throw new Error("No hay URL de Sheets configurada.");
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

async function _descargarDatosSheets(mostrarMensaje = false) {
  if (!state.config.sheetsUrl) return;
  if (!navigator.onLine) return;

  const icon = document.getElementById("syncIcon");
  if (icon) icon.classList.add("animate-spin");

  try {
    const url = `${state.config.sheetsUrl}?action=getTodo&t=${Date.now()}`;
    const resp = await fetch(url, { cache: "no-store" });
    const json = await resp.json();

    console.log("[SYNC] Respuesta de Sheets:", JSON.stringify({ 
      success: json.success, 
      pedidosCount: json.data ? (json.data.pedidos || []).length : "N/A",
      productosCount: json.data ? (json.data.productos || []).length : "N/A",
      pedidosPrimeros: json.data && json.data.pedidos ? json.data.pedidos.slice(0,2) : []
    }));

    if (json.success && json.data) {
      if (json.data.productos && json.data.productos.length > 0) {
        const mapa = { ...state.productos };
        json.data.productos.forEach(p => { 
          if (p && p.codigo) mapa[p.codigo] = p; 
        });
        state.productos = mapa;
        guardarProductosLocal();
      }
      if (json.data.ultimasCompras && json.data.ultimasCompras.length > 0) {
        state.compras = json.data.ultimasCompras;
        guardarComprasLocal();
      }
      if (json.data.ultimasVentas && json.data.ultimasVentas.length > 0) {
        state.ventas = json.data.ultimasVentas;
        guardarVentasLocal();
      }
      if (json.data.finanzas && json.data.finanzas.length > 0) {
        state.movimientosDinero = json.data.finanzas;
        guardarFinanzasLocal();
      }
      if (json.data.clientes) {
        const mapaCli = {};
        if (Array.isArray(json.data.clientes)) {
          json.data.clientes.forEach(c => { if (c.id) mapaCli[c.id] = c; });
        } else if (typeof json.data.clientes === "object") {
          Object.assign(mapaCli, json.data.clientes);
        }
        if (Object.keys(mapaCli).length > 0) {
          state.clientes = mapaCli;
          guardarClientesLocal();
        }
      }
      // Pedidos: fusionar Sheets + pendientes locales en cola
      if (json.data.pedidos && Array.isArray(json.data.pedidos)) {
        const idsPedidosSheets = new Set(json.data.pedidos.map(p => p.id));
        const pedidosSoloLocales = (state.pedidos || []).filter(p =>
          !idsPedidosSheets.has(p.id) &&
          (state.colaSincronizacion || []).some(q => q.datos && q.datos.pedido && q.datos.pedido.id === p.id)
        );
        state.pedidos = [...json.data.pedidos, ...pedidosSoloLocales];
        guardarPedidosLocal();
        console.log("[SYNC] Pedidos cargados:", state.pedidos.length, "| Locales pendientes cola:", pedidosSoloLocales.length);
      }
      if (json.data.cuentas && Array.isArray(json.data.cuentas)) {
        const idsCuentasSheets = new Set(json.data.cuentas.map(c => c.id || c.referenciaId));
        const cuentasSoloLocales = (state.cuentas || []).filter(c =>
          !idsCuentasSheets.has(c.id) && !idsCuentasSheets.has(c.referenciaId) &&
          (state.colaSincronizacion || []).some(q => q.datos && q.datos.cuenta && (q.datos.cuenta.id === c.id || q.datos.cuenta.referenciaId === c.referenciaId))
        );
        state.cuentas = [...json.data.cuentas, ...cuentasSoloLocales];
        guardarCuentasLocal();
      }

      renderizarTodo();
      actualizarBadgeConexion();
      if (mostrarMensaje) {
        const nPed = (state.pedidos || []).filter(p => p.estado === "pendiente" || !p.estado).length;
        mostrarToast(`📊 Sincronizado — ${nPed} encargo(s) pendiente(s)`, "success");
      }
    } else {
      console.warn("[SYNC] Respuesta inesperada:", json);
      if (mostrarMensaje) mostrarToast("Error: Sheets no devolvió datos válidos.", "error");
    }
  } catch (err) {
    console.error("[SYNC] Error al conectar:", err);
    if (mostrarMensaje) mostrarToast("Error al conectar con Google Sheets: " + err.message, "error");
  } finally {
    if (icon) icon.classList.remove("animate-spin");
  }
}

async function sincronizarConSheets(mostrarMensaje = true) {
  if (!state.config.sheetsUrl) {
    if (mostrarMensaje) mostrarToast("Configura la URL de Google Sheets en Ajustes.", "error");
    return;
  }
  if (!navigator.onLine) {
    if (mostrarMensaje) mostrarToast("Sin conexión a internet. Los datos locales están seguros.", "info");
    return;
  }

  // Subir cola en segundo plano (sin bloquear el GET de descarga)
  if (state.colaSincronizacion && state.colaSincronizacion.length > 0) {
    procesarColaSincronizacion(false);
  }

  // Siempre descargar datos frescos (pedidos, productos, clientes, etc.)
  await _descargarDatosSheets(mostrarMensaje);
}

function guardarConfiguracionSheets() {
  const url = document.getElementById("sheetsApiUrl").value.trim();
  state.config.sheetsUrl = url;
  guardarConfiguracionLocal();
  actualizarBadgeConexion();
  mostrarToast("URL guardada.", "success");
  if (url) sincronizarConSheets(true);
}

async function probarConexionSheets() {
  const url = document.getElementById("sheetsApiUrl").value.trim();
  if (!url) {
    mostrarToast("Ingresa una URL primero.", "error");
    return;
  }
  mostrarToast("Probando conexión...", "info");
  try {
    const resp = await fetch(`${url}?action=ping`);
    const json = await resp.json();
    if (json.success) mostrarToast("¡Conexión Exitosa con Google Sheets! 🎉", "success");
  } catch(e) {
    mostrarToast("Verifica que la Web App tenga acceso público.", "error");
  }
}

async function diagnosticarPedidos() {
  if (!state.config.sheetsUrl) {
    mostrarToast("Ingresa la URL de Sheets en Ajustes.", "error");
    return;
  }
  mostrarToast("Consultando encargos en Sheets...", "info");
  try {
    const resp = await fetch(`${state.config.sheetsUrl}?action=getPedidos&t=${Date.now()}`, { cache: "no-store" });
    const json = await resp.json();
    console.log("[DIAG] getPedidos completo:", json);

    if (json && json.success) {
      let listaPedidos = [];
      if (Array.isArray(json.data)) {
        listaPedidos = json.data;
      } else if (json.data && Array.isArray(json.data.pedidos)) {
        listaPedidos = json.data.pedidos;
      }

      const total = listaPedidos.length;
      const pendientes = listaPedidos.filter(p => p.estado === "pendiente" || !p.estado).length;
      
      state.pedidos = listaPedidos;
      guardarPedidosLocal();
      renderizarTodo();

      if (total === 0) {
        mostrarToast("⚠️ La hoja Pedidos está creada en Sheets pero no tiene filas con encargos.", "info");
      } else {
        mostrarToast(`✅ ${total} encargo(s) sincronizado(s) (${pendientes} pendiente(s))`, "success");
      }
    } else {
      mostrarToast("❌ Error en respuesta: " + JSON.stringify(json).slice(0, 120), "error");
    }
  } catch(e) {
    mostrarToast("❌ Error de red: " + e.message, "error");
    console.error("[DIAG] Error:", e);
  }
}

function guardarPreferenciasNegocio() {
  state.config.nombreNegocio = document.getElementById("businessNameInput").value.trim() || "Libro de Inventario";
  state.config.tipoCambio = Number(document.getElementById("exchangeRateInput").value) || 520;
  state.config.telefonoNegocio = document.getElementById("businessPhoneInput").value.trim();
  guardarConfiguracionLocal();
  renderizarTodo();
  mostrarToast("Ajustes actualizados.", "success");
}

function guardarConfigPuntos() {
  const razon = Number(document.getElementById("puntosRazonCRCInput").value);
  const valor = Number(document.getElementById("puntosValorCRCInput").value);
  const minimo = Number(document.getElementById("puntosMinimosCajeInput").value);

  if (!razon || razon < 1) {
    mostrarToast("Ingresa una razón válida (ej: 100).", "error");
    return;
  }
  if (!valor || valor < 1) {
    mostrarToast("Ingresa un valor de punto válido (ej: 5).", "error");
    return;
  }
  if (!minimo || minimo < 1) {
    mostrarToast("Ingresa un mínimo de puntos válido (ej: 100).", "error");
    return;
  }

  state.config.puntosRazonCRC = razon;
  state.config.puntosValorCRC = valor;
  state.config.puntosMinimosCanje = minimo;
  guardarConfiguracionLocal();

  // Mostrar resumen
  const box = document.getElementById("puntosSummaryBox");
  const txt = document.getElementById("puntosSummaryText");
  if (box && txt) {
    txt.innerHTML = `
      • Por cada ₡${razon.toLocaleString()} gastados → <b>1 punto</b><br>
      • 1 punto equivale a <b>₡${valor.toLocaleString()}</b> de descuento<br>
      • Mínimo <b>${minimo} puntos</b> para poder canjear<br>
      • Ej: con 500 puntos → descuento de <b>${fmtCRC(500 * valor)}</b>
    `;
    box.classList.remove("hidden");
  }

  mostrarToast("✅ Configuración de puntos guardada.", "success");
}

function cargarConfigPuntosUI() {
  const c = state.config;
  const razonEl = document.getElementById("puntosRazonCRCInput");
  const valorEl = document.getElementById("puntosValorCRCInput");
  const minimoEl = document.getElementById("puntosMinimosCajeInput");
  if (razonEl) razonEl.value = c.puntosRazonCRC || 20;
  if (valorEl) valorEl.value = c.puntosValorCRC !== undefined ? c.puntosValorCRC : 1;
  if (minimoEl) minimoEl.value = c.puntosMinimosCanje || 4000;
}

async function recargarCatalogoSemilla() {
  if (confirm("¿Descargar y actualizar todos los datos limpios directamente desde Google Sheets?")) {
    if (!state.config.sheetsUrl) {
      mostrarToast("Debes ingresar la URL de Google Sheets primero en Ajustes.", "error");
      return;
    }
    await sincronizarConSheets(true);
  }
}

async function forzarActualizacionApp() {
  if (!confirm("¿Deseas vaciar el caché del navegador y forzar la descarga de la versión más reciente?")) return;

  mostrarToast("Vaciando caché y buscando última versión... ⏳", "info");

  try {
    // 1. Limpiar todos los cachés de CacheStorage (Service Worker)
    if ('caches' in window) {
      const cacheKeys = await caches.keys();
      await Promise.all(cacheKeys.map(key => caches.delete(key)));
    }

    // 2. Desregistrar Service Workers activos para obligar instalación limpia
    if ('serviceWorker' in navigator) {
      const registrations = await navigator.serviceWorker.getRegistrations();
      for (const reg of registrations) {
        await reg.unregister();
      }
    }

    mostrarToast("¡Caché eliminado! Recargando aplicación... 🚀", "success");

    // 3. Forzar recarga con bypass de caché
    setTimeout(() => {
      window.location.href = window.location.origin + window.location.pathname + '?v=' + Date.now();
    }, 800);
  } catch (err) {
    console.error("Error al limpiar caché:", err);
    window.location.reload(true);
  }
}

function limpiarCacheLocal() {
  if (confirm("¿Borrar todos los datos locales de prueba y reiniciar la aplicación?")) {
    const sheetsUrl = state.config.sheetsUrl;
    localStorage.clear();
    if (sheetsUrl) {
      localStorage.setItem("inv_config_v2", JSON.stringify({ sheetsUrl }));
    }
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