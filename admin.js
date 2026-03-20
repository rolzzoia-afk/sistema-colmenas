// ═══════════════════════════════════════════════════════════════════
// admin.js — Panel de Administrador "Ojo de Dios"
// Solo lectura del inventario + control de versión del sistema
// ═══════════════════════════════════════════════════════════════════

// ── Lista de emails con acceso de administrador ──
// Agregar aquí los emails autorizados para usar el panel admin.
// Idealmente esto debería leerse de Firestore (colección "configuracion/admins"),
// pero como respaldo mínimo se valida localmente.
const ADMINS_AUTORIZADOS = [
    // Agregar emails de administradores aquí, ej:
    // "admin@empresa.com",
    // "jefe@empresa.com"
];

let unsubVersion = null;
let unsubColmena = null;
let unsubSeriales = null;
let unsubHistorial = null;
let versionActualFirebase = null;    // Valor actual en Firebase
let usuarioSeleccionado = null;      // Email del usuario actualmente seleccionado
let historialOperaciones = [];       // Cache local del historial
let operacionSeleccionada = null;    // Operación activa en el modal

// ─── ESPERAR A QUE FIREBASE ESTÉ LISTO ──────────────────────────────────────

document.addEventListener("DOMContentLoaded", () => {
    const esperar = setInterval(() => {
        if (window.fbAuth) {
            clearInterval(esperar);

            window.fbOnAuth(window.fbAuth, async (user) => {
                if (!user) {
                    mostrarLogin();
                } else {
                    // ── Validación de rol admin ──
                    const esAdmin = await verificarRolAdmin(user.email);
                    if (!esAdmin) {
                        alert("⛔ Acceso denegado. Tu cuenta no tiene permisos de administrador.");
                        window.fbSignOut(window.fbAuth).then(() => location.reload());
                        return;
                    }
                    document.getElementById("app-container").style.display = "block";
                    document.getElementById("adminEmail").textContent = user.email;
                    inicializarPanel();
                }
            });
        }
    }, 100);
});

// Verificar si el email tiene rol de admin (primero Firestore, luego lista local)
async function verificarRolAdmin(email) {
    // Intento 1: verificar en Firestore (colección configuracion/admins)
    try {
        const docRef = window.fbDoc(window.fbDb, "configuracion", "admins");
        const snap = await window.fbGetDoc(docRef);
        if (snap.exists()) {
            const data = snap.data();
            const listaAdmins = data.emails || [];
            if (listaAdmins.includes(email)) return true;
            if (listaAdmins.length > 0) return false; // Lista existe pero el email no está
        }
    } catch (e) {
        console.warn("No se pudo verificar admins en Firestore:", e.message);
    }

    // Intento 2: lista local hardcodeada (fallback)
    if (ADMINS_AUTORIZADOS.length > 0) {
        return ADMINS_AUTORIZADOS.includes(email);
    }

    // Si no hay lista configurada en ningún lado, permitir acceso (compatibilidad)
    console.warn("⚠️ No hay lista de admins configurada. Cualquier usuario autenticado tiene acceso.");
    return true;
}

// ─── LOGIN ──────────────────────────────────────────────────────────────────

function mostrarLogin() {
    document.getElementById("app-container").style.display = "none";
    document.body.insertAdjacentHTML("afterbegin", `
        <div class="login-container" id="loginBox">
            <h2>Admin — Ojo de Dios</h2>
            <input type="email" id="email" placeholder="Correo del administrador" />
            <input type="password" id="password" placeholder="Contraseña" />
            <button id="btnLogin">Ingresar</button>
            <p class="error" id="loginError"></p>
        </div>
    `);
    document.getElementById("btnLogin").addEventListener("click", () => {
        const email = document.getElementById("email").value;
        const pass = document.getElementById("password").value;
        window.fbSignIn(window.fbAuth, email, pass)
            .then(() => {
                const box = document.getElementById("loginBox");
                if (box) box.remove();
            })
            .catch(err => {
                document.getElementById("loginError").textContent = err.message;
            });
    });
}

// ─── INICIALIZAR PANEL ──────────────────────────────────────────────────────

function inicializarPanel() {
    // Logout
    document.getElementById("btnLogout").addEventListener("click", () => {
        limpiarListeners();
        window.fbSignOut(window.fbAuth).then(() => location.reload());
    });

    // Monitor de versión en tiempo real
    iniciarMonitorVersion();

    // Cargar lista de usuarios
    cargarUsuarios();

    // Botón de forzar actualización
    document.getElementById("btnForzarUpdate").addEventListener("click", forzarActualizacion);

    // Selector de usuario
    document.getElementById("selectUsuario").addEventListener("change", (e) => {
        const email = e.target.value;
        if (email) {
            usuarioSeleccionado = email;
            suscribirInventarioUsuario(email);
            suscribirHistorialUsuario(email);
        } else {
            usuarioSeleccionado = null;
            limpiarTablaInventario();
            limpiarTablaHistorial();
        }
    });

    // Modal: botones de cerrar
    document.getElementById("btnCerrarModal").addEventListener("click", cerrarModal);
    document.getElementById("btnCerrarModal2").addEventListener("click", cerrarModal);
    document.getElementById("btnRollback").addEventListener("click", ejecutarRollback);
}

// ─── MONITOR DE VERSIÓN (onSnapshot) ────────────────────────────────────────

function iniciarMonitorVersion() {
    const docRef = window.fbDoc(window.fbDb, "configuracion", "version_minima");

    unsubVersion = window.fbOnSnapshot(docRef, (snap) => {
        if (snap.exists()) {
            versionActualFirebase = snap.data().version;
            document.getElementById("versionActual").textContent = `v${versionActualFirebase}`;
            document.getElementById("btnForzarUpdate").disabled = false;
        } else {
            document.getElementById("versionActual").textContent = "Sin configurar";
            versionActualFirebase = null;
            document.getElementById("btnForzarUpdate").disabled = false;
        }
    }, (error) => {
        console.error("Error monitor versión:", error);
        document.getElementById("versionActual").textContent = "Error";
    });
}

// ─── FORZAR ACTUALIZACIÓN (subir versión) ───────────────────────────────────

async function forzarActualizacion() {
    const btn = document.getElementById("btnForzarUpdate");

    let nuevaVersion;
    if (versionActualFirebase) {
        const num = parseFloat(versionActualFirebase);
        nuevaVersion = (Math.round((num + 0.1) * 10) / 10).toFixed(1);
    } else {
        nuevaVersion = "1.0";
    }

    const confirmar = confirm(
        `¿Estás seguro de subir la versión de "${versionActualFirebase || 'N/A'}" a "${nuevaVersion}"?\n\n` +
        `Esto forzará una recarga inmediata en TODOS los navegadores del taller.`
    );
    if (!confirmar) return;

    btn.disabled = true;
    btn.textContent = "Actualizando...";

    try {
        const docRef = window.fbDoc(window.fbDb, "configuracion", "version_minima");
        await window.fbSetDoc(docRef, {
            version: nuevaVersion,
            actualizadoPor: window.fbAuth.currentUser.email,
            fechaActualizacion: new Date().toISOString()
        });
        // El onSnapshot actualizará la UI automáticamente
    } catch (error) {
        console.error("Error actualizando versión:", error);
        alert("Error al actualizar la versión: " + error.message);
    } finally {
        btn.disabled = false;
        btn.textContent = "Forzar Actualización en Taller (Subir Versión)";
    }
}

// ─── CARGAR USUARIOS ────────────────────────────────────────────────────────

// Lista conocida de emails de taller (fallback cuando Firestore no lista documentos padre)
const USUARIOS_CONOCIDOS = [];

async function cargarUsuarios() {
    const select = document.getElementById("selectUsuario");
    let usuarios = [];

    try {
        // Intento 1: leer documentos directos de la colección "usuarios"
        console.log("🔍 Consultando colección 'usuarios'...");
        const colRef = window.fbCollection(window.fbDb, "usuarios");
        const snapshot = await window.fbGetDocs(colRef);
        console.log("📊 Documentos encontrados en 'usuarios':", snapshot.size);

        snapshot.forEach(docSnap => {
            console.log("  → doc:", docSnap.id);
            usuarios.push(docSnap.id);
        });
    } catch (error) {
        console.error("❌ Error leyendo colección 'usuarios':", error.code, error.message);
    }

    // Intento 2: si no encontró nada, usar el email del admin logueado como primer candidato
    if (usuarios.length === 0) {
        console.log("⚠️ 0 documentos en 'usuarios'. Activando detección alternativa...");

        // Agregar usuarios conocidos configurados manualmente
        usuarios = [...USUARIOS_CONOCIDOS];

        // Agregar el propio email del admin como candidato
        const adminEmail = window.fbAuth.currentUser?.email;
        if (adminEmail && !usuarios.includes(adminEmail)) {
            usuarios.push(adminEmail);
        }

        // Intentar descubrir usuarios probando rutas de inventario conocidas
        if (usuarios.length > 0) {
            const usuariosVerificados = [];
            for (const email of usuarios) {
                try {
                    const invRef = window.fbDoc(window.fbDb, "usuarios", email, "inventario", "datos");
                    const invSnap = await window.fbGetDoc(invRef);
                    if (invSnap.exists()) {
                        usuariosVerificados.push(email);
                        console.log("  ✅ Verificado (tiene inventario):", email);
                    } else {
                        // Probar colmena_final
                        const colRef = window.fbDoc(window.fbDb, "usuarios", email, "colmena_final", "datos");
                        const colSnap = await window.fbGetDoc(colRef);
                        if (colSnap.exists()) {
                            usuariosVerificados.push(email);
                            console.log("  ✅ Verificado (tiene colmena_final):", email);
                        } else {
                            console.log("  ⚪ Sin datos:", email);
                            // Aún así incluirlo para que pueda intentar
                            usuariosVerificados.push(email);
                        }
                    }
                } catch (e) {
                    console.warn("  ⚠️ Error verificando", email, ":", e.message);
                    usuariosVerificados.push(email); // incluir de todos modos
                }
            }
            usuarios = usuariosVerificados;
        }

        console.log("📋 Usuarios detectados (fallback):", usuarios);
    }

    // Poblar el selector
    select.innerHTML = '<option value="">-- Seleccionar usuario --</option>';
    usuarios.sort();
    usuarios.forEach(email => {
        const opt = document.createElement("option");
        opt.value = email;
        opt.textContent = email;
        select.appendChild(opt);
    });

    // Si solo hay 1 usuario, seleccionarlo automáticamente y cargar inventario
    if (usuarios.length === 1) {
        select.value = usuarios[0];
        usuarioSeleccionado = usuarios[0];
        console.log("👁️ Auto-seleccionando único usuario:", usuarios[0]);
        suscribirInventarioUsuario(usuarios[0]);
        suscribirHistorialUsuario(usuarios[0]);
    } else if (usuarios.length === 0) {
        console.warn("⚠️ No se detectaron usuarios. Cargando inventario del admin como fallback global...");
        select.innerHTML = '<option value="">Sin usuarios detectados (mostrando global)</option>';
        // Fallback global: usar el email del admin logueado
        const adminEmail = window.fbAuth.currentUser?.email;
        if (adminEmail) {
            usuarioSeleccionado = adminEmail;
            suscribirInventarioUsuario(adminEmail);
            suscribirHistorialUsuario(adminEmail);
        }
    }
}

// ─── SUSCRIPCIÓN AL INVENTARIO DE UN USUARIO (onSnapshot) ───────────────────

function suscribirInventarioUsuario(email) {
    // Limpiar listeners anteriores
    if (unsubColmena) { unsubColmena(); unsubColmena = null; }
    if (unsubSeriales) { unsubSeriales(); unsubSeriales = null; }

    const tbody = document.getElementById("tbodyInventario");
    tbody.innerHTML = '<tr><td colspan="7" class="loading-msg">Cargando inventario...</td></tr>';

    let colmenas = [];
    let seriales = [];
    let colmenaCargada = false;
    let serialesCargados = false;

    // Listener de colmena_final
    const colmenaRef = window.fbDoc(window.fbDb, "usuarios", email, "colmena_final", "datos");
    unsubColmena = window.fbOnSnapshot(colmenaRef, (snap) => {
        if (snap.exists() && snap.data().data) {
            colmenas = JSON.parse(snap.data().data) || [];
        } else {
            colmenas = [];
        }
        colmenaCargada = true;
        if (serialesCargados) renderizarInventario(colmenas, seriales);
    }, (error) => {
        console.error("Error listener colmena:", error);
        colmenaCargada = true;
        if (serialesCargados) renderizarInventario(colmenas, seriales);
    });

    // Listener de seriales
    const serialesRef = window.fbCollection(window.fbDb, "usuarios", email, "maestro_seriales");
    unsubSeriales = window.fbOnSnapshot(serialesRef, (querySnap) => {
        seriales = [];
        querySnap.forEach(docSnap => {
            seriales.push(docSnap.data());
        });
        serialesCargados = true;
        if (colmenaCargada) renderizarInventario(colmenas, seriales);
    }, (error) => {
        console.error("Error listener seriales:", error);
        serialesCargados = true;
        if (colmenaCargada) renderizarInventario(colmenas, seriales);
    });
}

// ─── RENDERIZAR TABLA DE INVENTARIO ─────────────────────────────────────────

function renderizarInventario(colmenas, seriales) {
    const tbody = document.getElementById("tbodyInventario");

    // Construir mapa de seriales por código para cruzar con colmenas
    const serialesPorCodigo = {};
    seriales.forEach(s => {
        const cod = (s.codigo || "").toUpperCase().trim();
        if (!serialesPorCodigo[cod]) serialesPorCodigo[cod] = [];
        serialesPorCodigo[cod].push(s);
    });

    // Construir filas combinando colmenas + seriales
    const filas = [];

    // Filas de colmenas con su serial asociado si existe
    colmenas.forEach(c => {
        const cod = (c.cod || "").toUpperCase().trim();
        const serial = c.serial || null;
        filas.push({
            codigo: c.cod || "-",
            medida_cm: c.medida_cm || (c.medida_mm ? (c.medida_mm / 10).toFixed(1) : "-"),
            colmena: c.n_colmena || "-",
            lote: serial ? serial.lote : "-",
            serial: serial ? serial.serial : "-",
            fecha: serial ? serial.fecha : "-",
            estado: serial ? serial.estado : "disponible"
        });
    });

    // Si no hay colmenas, mostrar seriales directamente
    if (colmenas.length === 0 && seriales.length > 0) {
        seriales.forEach(s => {
            filas.push({
                codigo: s.codigo || "-",
                medida_cm: "-",
                colmena: "-",
                lote: s.lote || "-",
                serial: s.serial || "-",
                fecha: s.fecha || "-",
                estado: s.estado || "disponible"
            });
        });
    }

    // Stats
    const totalDisponibles = filas.filter(f => f.estado === "disponible").length;
    const totalOcupados = filas.filter(f => f.estado === "ocupado").length;
    const codigosUnicos = new Set(filas.map(f => f.codigo)).size;

    document.getElementById("statTotal").textContent = filas.length;
    document.getElementById("statDisponibles").textContent = totalDisponibles;
    document.getElementById("statOcupados").textContent = totalOcupados;
    document.getElementById("statCodigos").textContent = codigosUnicos;

    if (filas.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="loading-msg">Este usuario no tiene inventario</td></tr>';
        return;
    }

    tbody.innerHTML = filas.map(f => {
        const estadoBadge = f.estado === "ocupado"
            ? '<span class="badge-ocupado">OCUPADO</span>'
            : '<span class="badge-disponible">DISPONIBLE</span>';
        return `<tr>
            <td>${f.codigo}</td>
            <td>${f.medida_cm}</td>
            <td>${f.colmena}</td>
            <td>${f.lote}</td>
            <td>${f.serial}</td>
            <td>${formatearFecha(f.fecha)}</td>
            <td>${estadoBadge}</td>
        </tr>`;
    }).join("");
}

// ─── UTILIDADES ─────────────────────────────────────────────────────────────

function formatearFecha(fecha) {
    if (!fecha || fecha === "-" || fecha === "undefined" || fecha === null) return "-";
    if (typeof fecha === "number" && fecha > 40000) {
        const d = new Date((fecha - 25569) * 86400 * 1000);
        return `${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${d.getFullYear()}`;
    }
    const d = new Date(fecha);
    if (isNaN(d.getTime())) return String(fecha).includes("/") ? fecha : "-";
    return `${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${d.getFullYear()}`;
}

function limpiarTablaInventario() {
    document.getElementById("tbodyInventario").innerHTML =
        '<tr><td colspan="7" class="loading-msg">Seleccione un usuario para ver su inventario</td></tr>';
    document.getElementById("statTotal").textContent = "0";
    document.getElementById("statDisponibles").textContent = "0";
    document.getElementById("statOcupados").textContent = "0";
    document.getElementById("statCodigos").textContent = "0";
}

function limpiarTablaHistorial() {
    if (unsubHistorial) { unsubHistorial(); unsubHistorial = null; }
    historialOperaciones = [];
    document.getElementById("tbodyHistorial").innerHTML =
        '<tr><td colspan="4" class="loading-msg">Seleccione un usuario para ver su historial</td></tr>';
}

// ─── HISTORIAL DE OPERACIONES (onSnapshot) ──────────────────────────────────

function suscribirHistorialUsuario(email) {
    if (unsubHistorial) { unsubHistorial(); unsubHistorial = null; }

    const tbody = document.getElementById("tbodyHistorial");
    tbody.innerHTML = '<tr><td colspan="4" class="loading-msg">Cargando historial...</td></tr>';

    const colRef = window.fbCollection(window.fbDb, "usuarios", email, "historial_operaciones");
    const q = window.fbQuery(colRef, window.fbOrderBy("fecha", "desc"));

    unsubHistorial = window.fbOnSnapshot(q, (querySnap) => {
        historialOperaciones = [];
        querySnap.forEach(docSnap => {
            historialOperaciones.push({ id: docSnap.id, ...docSnap.data() });
        });
        renderizarHistorial();
    }, (error) => {
        console.error("Error listener historial:", error);
        tbody.innerHTML = '<tr><td colspan="4" class="loading-msg">Error cargando historial</td></tr>';
    });
}

function renderizarHistorial() {
    const tbody = document.getElementById("tbodyHistorial");

    if (historialOperaciones.length === 0) {
        tbody.innerHTML = '<tr><td colspan="4" class="loading-msg">Sin operaciones registradas</td></tr>';
        return;
    }

    tbody.innerHTML = historialOperaciones.map((op, idx) => {
        const fecha = new Date(op.fecha);
        const fechaStr = `${String(fecha.getDate()).padStart(2,"0")}/${String(fecha.getMonth()+1).padStart(2,"0")}/${fecha.getFullYear()} ${String(fecha.getHours()).padStart(2,"0")}:${String(fecha.getMinutes()).padStart(2,"0")}`;
        const esRevertido = op.estado === "REVERTIDO";
        const esCascada = op.estado === "ANULADO_POR_CASCADA";
        const esAnulado = esRevertido || esCascada;
        const rowStyle = esAnulado ? 'style="background:rgba(231,76,60,0.15);"' : '';
        let badgeHtml = '';
        if (esRevertido) {
            badgeHtml = `<span style="background:#e74c3c;color:white;padding:2px 8px;border-radius:3px;font-size:10px;font-weight:bold;margin-left:8px;">❌ REVERTIDO</span>`;
        } else if (esCascada) {
            badgeHtml = `<span style="background:#e67e22;color:white;padding:2px 8px;border-radius:3px;font-size:10px;font-weight:bold;margin-left:8px;">⚠️ ANULADO (CASCADA)</span>`;
        }
        const motivoHtml = esAnulado && op.motivo_error
            ? `<div style="font-size:10px;color:#e57373;margin-top:3px;font-style:italic;">Motivo: ${op.motivo_error}</div>`
            : '';
        // Diagnóstico: mostrar cuántas colmenas tiene el snapshot guardado
        let snapshotInfo = '-';
        try {
            const snapColmenas = JSON.parse(op.snapshot_inventario || '[]');
            const snapSeriales = JSON.parse(op.snapshot_seriales || '[]');
            const disponibles = snapSeriales.filter(s => s.estado === 'disponible').length;
            snapshotInfo = `<span style="font-size:10px;color:#00d2ff;">${snapColmenas.length} colmenas / ${disponibles} tubos</span>`;
        } catch(e) { snapshotInfo = '<span style="color:#e74c3c;font-size:10px;">⚠️ Sin snapshot</span>'; }

        return `<tr ${rowStyle}>
            <td>${fechaStr}${badgeHtml}</td>
            <td>${op.nombre_excel || '-'}${motivoHtml}</td>
            <td>${op.total_cortes || '-'}</td>
            <td>${snapshotInfo}</td>
            <td>
                <button onclick="abrirModalHistorial(${idx})" style="background:#3498db;color:white;border:none;padding:5px 12px;border-radius:4px;cursor:pointer;font-size:11px;">Ver Detalle</button>
            </td>
        </tr>`;
    }).join("");
}

// ─── MODAL DE DETALLE + ROLLBACK ────────────────────────────────────────────

function abrirModalHistorial(idx) {
    operacionSeleccionada = historialOperaciones[idx];
    if (!operacionSeleccionada) return;

    const fecha = new Date(operacionSeleccionada.fecha);
    const fechaStr = `${String(fecha.getDate()).padStart(2,"0")}/${String(fecha.getMonth()+1).padStart(2,"0")}/${fecha.getFullYear()} ${String(fecha.getHours()).padStart(2,"0")}:${String(fecha.getMinutes()).padStart(2,"0")}:${String(fecha.getSeconds()).padStart(2,"0")}`;

    document.getElementById("modalTitulo").textContent = `Operación: ${operacionSeleccionada.nombre_excel || 'Sin nombre'}`;
    document.getElementById("modalSubtitulo").textContent = `Fecha: ${fechaStr} | Cortes: ${operacionSeleccionada.total_cortes || 0} | Usuario: ${operacionSeleccionada.usuario || '-'}`;

    // Parsear resultados y renderizar tabla
    const tbodyPlan = document.getElementById("tbodyModalPlan");
    try {
        const resultados = JSON.parse(operacionSeleccionada.resultados || '[]');
        tbodyPlan.innerHTML = resultados.map(r => {
            const codigoReal = r.codigo || r.codigo_original || '';
            const codigoDisplay = (r.codigo_original && r.codigo && r.codigo_original !== r.codigo)
                ? `${r.codigo_original} &rarr; ${r.codigo}`
                : codigoReal;

            let accion = (r.fuente || '').toUpperCase();
            let badgeStyle = 'background:#9b59b6;color:white;padding:2px 8px;border-radius:3px;';
            if (r.fuente === 'tubo_nuevo') badgeStyle = 'background:#f39c12;color:white;padding:2px 8px;border-radius:3px;';
            else if (r.fuente === 'reemplazo') badgeStyle = 'background:#3498db;color:white;padding:2px 8px;border-radius:3px;';

            let filas = `<tr>
                <td>${r.colmena || r.nombreMaterialNuevo || '-'}</td>
                <td>${codigoDisplay}</td>
                <td>${r.color || '-'}</td>
                <td>${r.medida_cm || '-'}</td>
                <td>${r.medida_origen || '-'}</td>
                <td>${r.sobrante_cm || '-'}</td>
                <td><span style="${badgeStyle}">${accion}</span></td>
            </tr>`;

            if (r.sobrante_cm > 0) {
                let accionSobrante, estilo;
                if (r.es_intermedio) {
                    accionSobrante = 'RESERVAR EN MESA';
                    estilo = 'background:#fff3e0;color:#e65100;';
                } else if (r.es_desecho) {
                    accionSobrante = 'DESECHAR MERMA';
                    estilo = 'background:#ffebee;color:#c62828;';
                } else {
                    accionSobrante = 'GUARDAR SOBRANTE';
                    estilo = 'background:#e8f5e9;color:#2e7d32;';
                }
                filas += `<tr style="${estilo}">
                    <td>${r.colmena_sobrante || '-'}</td>
                    <td>${codigoDisplay}</td>
                    <td></td>
                    <td colspan="2">${r.sobrante_cm} cm</td>
                    <td></td>
                    <td><strong>${accionSobrante}</strong></td>
                </tr>`;
            }
            return filas;
        }).join('') || '<tr><td colspan="7">Sin datos de resultados</td></tr>';
    } catch (e) {
        tbodyPlan.innerHTML = '<tr><td colspan="7">Error parseando resultados</td></tr>';
    }

    document.getElementById("modalHistorial").style.display = "block";
}

function cerrarModal() {
    document.getElementById("modalHistorial").style.display = "none";
    operacionSeleccionada = null;
}

async function ejecutarRollback() {
    if (!operacionSeleccionada || !usuarioSeleccionado) {
        alert("No hay operación seleccionada o usuario activo.");
        return;
    }

    const fecha = new Date(operacionSeleccionada.fecha);
    const fechaStr = `${String(fecha.getDate()).padStart(2,"0")}/${String(fecha.getMonth()+1).padStart(2,"0")}/${fecha.getFullYear()} ${String(fecha.getHours()).padStart(2,"0")}:${String(fecha.getMinutes()).padStart(2,"0")}`;

    // Pedir motivo del error (dato para futura IA)
    const motivoError = prompt(
        `⚠️ ROLLBACK — ${fechaStr}\n\n` +
        `¿Cuál fue el error?\n` +
        `(Este dato entrenará a la IA en el futuro)\n\n` +
        `Describe brevemente qué salió mal:`
    );
    if (motivoError === null || motivoError.trim() === "") {
        alert("Rollback cancelado. Debes ingresar un motivo para continuar.");
        return;
    }

    const confirmar = confirm(
        `⚠️ OPERACIÓN DESTRUCTIVA ⚠️\n\n` +
        `Vas a REVERTIR el inventario de "${usuarioSeleccionado}" al estado previo a la operación del ${fechaStr}.\n\n` +
        `Motivo: "${motivoError.trim()}"\n\n` +
        `Esto SOBREESCRIBIRÁ:\n` +
        `• Colmenas (colmena_final)\n` +
        `• Maestro de Seriales (maestro_seriales)\n\n` +
        `¿Estás absolutamente seguro?`
    );
    if (!confirmar) return;

    const btnRollback = document.getElementById("btnRollback");
    btnRollback.disabled = true;
    btnRollback.textContent = "Revirtiendo...";

    try {
        const db = window.fbDb;

        // 1. Restaurar colmena_final
        const snapshotColmenas = JSON.parse(operacionSeleccionada.snapshot_inventario || '[]');
        await window.fbSetDoc(
            window.fbDoc(db, "usuarios", usuarioSeleccionado, "colmena_final", "datos"),
            {
                data: JSON.stringify(snapshotColmenas),
                fechaActualizacion: new Date().toISOString(),
                restauradoPor: window.fbAuth.currentUser.email,
                esRollback: true
            }
        );
        console.log(`✅ colmena_final restaurada: ${snapshotColmenas.length} registros`);

        // 2. Restaurar maestro_seriales (borrar todos + re-crear)
        const snapshotSeriales = JSON.parse(operacionSeleccionada.snapshot_seriales || '[]');

        // 2a. Borrar todos los documentos actuales del maestro_seriales
        const serialesRef = window.fbCollection(db, "usuarios", usuarioSeleccionado, "maestro_seriales");
        const serialesActuales = await window.fbGetDocs(serialesRef);

        // Usar batches de 500 (límite de Firestore)
        let batch = window.fbWriteBatch(db);
        let batchCount = 0;

        for (const docSnap of serialesActuales.docs) {
            batch.delete(docSnap.ref);
            batchCount++;
            if (batchCount === 499) {
                await batch.commit();
                batch = window.fbWriteBatch(db);
                batchCount = 0;
            }
        }
        if (batchCount > 0) await batch.commit();
        console.log(`🗑️ ${serialesActuales.size} seriales eliminados`);

        // 2b. Re-crear seriales del snapshot
        let batchInsert = window.fbWriteBatch(db);
        let insertCount = 0;

        for (let i = 0; i < snapshotSeriales.length; i++) {
            const s = snapshotSeriales[i];
            const docId = `${s.codigo}_${s.lote}_${s.paquete}_${s.serial}`.replace(/[\/\s]/g, '_');
            const docRef = window.fbDoc(db, "usuarios", usuarioSeleccionado, "maestro_seriales", docId);
            batchInsert.set(docRef, s);
            insertCount++;
            if (insertCount === 499) {
                await batchInsert.commit();
                batchInsert = window.fbWriteBatch(db);
                insertCount = 0;
            }
        }
        if (insertCount > 0) await batchInsert.commit();
        console.log(`✅ ${snapshotSeriales.length} seriales restaurados`);

        // 3. Anulación en cascada: etiquetar esta operación + todas las posteriores
        const fechaReversion = new Date().toISOString();
        const adminEmail = window.fbAuth.currentUser.email;
        const fechaOperacion = operacionSeleccionada.fecha;

        // Consultar todas las operaciones con fecha >= a la seleccionada
        const histColRef = window.fbCollection(db, "usuarios", usuarioSeleccionado, "historial_operaciones");
        const cascadaQuery = window.fbQuery(histColRef, window.fbWhere("fecha", ">=", fechaOperacion));
        const cascadaSnap = await window.fbGetDocs(cascadaQuery);

        let contadorRevertido = 0;
        let contadorCascada = 0;

        for (const docSnap of cascadaSnap.docs) {
            const docRef = window.fbDoc(db, "usuarios", usuarioSeleccionado, "historial_operaciones", docSnap.id);
            if (docSnap.id === operacionSeleccionada.id) {
                // El documento clickeado: REVERTIDO
                await window.fbUpdateDoc(docRef, {
                    estado: "REVERTIDO",
                    motivo_error: motivoError.trim(),
                    fecha_reversion: fechaReversion,
                    revertido_por: adminEmail
                });
                contadorRevertido++;
            } else {
                // Documentos posteriores: ANULADO_POR_CASCADA
                await window.fbUpdateDoc(docRef, {
                    estado: "ANULADO_POR_CASCADA",
                    motivo_error: "Anulado automáticamente porque se revirtió una operación anterior en la línea de tiempo.",
                    fecha_reversion: fechaReversion,
                    revertido_por: "SISTEMA"
                });
                contadorCascada++;
            }
        }
        console.log(`🏷️ Historial etiquetado — REVERTIDO: ${contadorRevertido}, CASCADA: ${contadorCascada}`);

        const cascadaMsg = contadorCascada > 0
            ? `\n• ${contadorCascada} operación(es) posterior(es) anulada(s) por cascada`
            : '';
        alert(`Rollback completado exitosamente.\n\nMotivo: "${motivoError.trim()}"\n\n• ${snapshotColmenas.length} colmenas restauradas\n• ${snapshotSeriales.length} seriales restaurados${cascadaMsg}`);
        cerrarModal();

    } catch (error) {
        console.error("Error en rollback:", error);
        alert("Error durante el rollback: " + error.message);
    } finally {
        btnRollback.disabled = false;
        btnRollback.textContent = "❌ Anular esta operación (y restaurar inventario previo)";
    }
}

function limpiarListeners() {
    if (unsubVersion) { unsubVersion(); unsubVersion = null; }
    if (unsubColmena) { unsubColmena(); unsubColmena = null; }
    if (unsubSeriales) { unsubSeriales(); unsubSeriales = null; }
    if (unsubHistorial) { unsubHistorial(); unsubHistorial = null; }
}