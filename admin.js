// ═══════════════════════════════════════════════════════════════════
// admin.js — Panel de Administrador "Ojo de Dios"
// Solo lectura del inventario + control de versión del sistema
// ═══════════════════════════════════════════════════════════════════

let unsubVersion = null;
let unsubColmena = null;
let unsubSeriales = null;
let versionActualFirebase = null;    // Valor actual en Firebase

// ─── ESPERAR A QUE FIREBASE ESTÉ LISTO ──────────────────────────────────────

document.addEventListener("DOMContentLoaded", () => {
    const esperar = setInterval(() => {
        if (window.fbAuth) {
            clearInterval(esperar);

            window.fbOnAuth(window.fbAuth, (user) => {
                if (!user) {
                    mostrarLogin();
                } else {
                    document.getElementById("app-container").style.display = "block";
                    document.getElementById("adminEmail").textContent = user.email;
                    inicializarPanel();
                }
            });
        }
    }, 100);
});

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
            suscribirInventarioUsuario(email);
        } else {
            limpiarTablaInventario();
        }
    });
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

async function cargarUsuarios() {
    const select = document.getElementById("selectUsuario");
    try {
        const colRef = window.fbCollection(window.fbDb, "usuarios");
        const snapshot = await window.fbGetDocs(colRef);

        select.innerHTML = '<option value="">-- Seleccionar usuario --</option>';

        const usuarios = [];
        snapshot.forEach(docSnap => {
            usuarios.push(docSnap.id);
        });
        usuarios.sort();

        usuarios.forEach(email => {
            const opt = document.createElement("option");
            opt.value = email;
            opt.textContent = email;
            select.appendChild(opt);
        });

        if (usuarios.length === 0) {
            select.innerHTML = '<option value="">No se encontraron usuarios</option>';
        }
    } catch (error) {
        console.error("Error cargando usuarios:", error);
        select.innerHTML = '<option value="">Error cargando usuarios</option>';
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

function limpiarListeners() {
    if (unsubVersion) { unsubVersion(); unsubVersion = null; }
    if (unsubColmena) { unsubColmena(); unsubColmena = null; }
    if (unsubSeriales) { unsubSeriales(); unsubSeriales = null; }
}
