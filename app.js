// Función para limpiar números y manejar tanto comas como puntos decimales
function limpiarNumero(valor) {
    if (valor === null || valor === undefined) return 0;
    // Convertir a string, reemplazar coma por punto y quitar espacios
    let str = String(valor).replace(',', '.').trim();
    let num = parseFloat(str);
    return isNaN(num) ? 0 : num;
}

const VERSION_ACTUAL = "4.0";

const MM_TUBO_ORIGINAL = 5780;
const MM_KERF = 3;
const STOCK_MINIMO = 10; // Umbral de alerta: códigos con ≤ este número de tubos enteros disponibles

const SistemaInventario = {
    ordenes: [],
    colmenas: [],
    catalogoReemplazos: {},
    catalogoColores: {},   // { "E01": "Aluminio", "E02": "Bronce", ... }
    catalogoMedidas: {},   // { "E01": 578, "E02": 600, ... } medida real del tubo en cm
    logs: [],
    colmenasDisponibles: [],
    datosCrudosOrdenes: [],
    resultadosOptimizacion: [],
    colmenasHistorico: [],
    mermas: [],
    seriales: [],
    historialMermas: [],
    colmenaCruda: [],          // Copia inmutable de colmenas al cargar Excel
    ordenesCrudas: [],         // Copia inmutable de órdenes antes de optimizar
    serialesCrudos: [],        // Copia inmutable de seriales antes de optimizar
    overridesNuevos: {},       // Cola de medidas reales rectificadas: { 'E02': [579.2, 579.5] }
    catalogoAccesorios: {}     // Diccionario DESCRIPCIÓN|COLOR → código: { 'CENEFA OVALADA|ALUMINIO': 'E15' }
};

// Variables globales para manejo de colmena desde Firebase
let colmenaActual = null;        // Colmena descargada de Firebase al inicio de sesión
let usandoColmenaManual = false; // true si el usuario subió un archivo manualmente

// Flags de carga inicial para onSnapshot (escudo anti-inventario-vacío)
let inventarioCargado = false;
let serialesCargados = false;
let procesandoExcel = false;     // true mientras el usuario procesa un archivo Excel

// Desuscriptores de onSnapshot para limpiar al cerrar sesión
let unsubInventario = null;
let unsubSeriales = null;

function formatearFecha(fecha) {
    if (!fecha || fecha === '-' || fecha === 'undefined' || fecha === null) return '-';
    // Si es número serie de Excel (46083 -> 04/03/2026)
    if (typeof fecha === 'number' && fecha > 40000) {
        const dExcel = new Date((fecha - 25569) * 86400 * 1000);
        return `${String(dExcel.getDate()).padStart(2,'0')}/${String(dExcel.getMonth()+1).padStart(2,'0')}/${dExcel.getFullYear()}`;
    }
    // Si es string o Date objeto
    const d = new Date(fecha);
    if (isNaN(d.getTime())) return String(fecha).includes('/') ? fecha : '-';
    return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
}
// Verificar versión mínima contra Firebase
async function verificarVersionMinima() {
    try {
        const db = window.firebaseDb;
        const versionDoc = await window.firebaseGetDoc(window.firebaseDoc(db, "configuracion", "version_minima"));
        if (versionDoc.exists()) {
            const versionMinima = versionDoc.data().version;
            if (versionMinima && VERSION_ACTUAL < versionMinima) {
                console.warn(`Versión ${VERSION_ACTUAL} obsoleta. Mínima requerida: ${versionMinima}`);
                localStorage.clear();
                alert("Hay una nueva versión del sistema. Se recargará la página.");
                window.location.reload(true);
            }
        }
    } catch (error) {
        console.error("Error verificando versión mínima:", error);
    }
}

// Verificar sesión activa
document.addEventListener("DOMContentLoaded", () => {

  const esperarFirebase = setInterval(() => {
    if (window.firebaseAuth) {
      clearInterval(esperarFirebase);

      // Verificar versión mínima antes de continuar
      verificarVersionMinima();

      window.firebaseOnAuth(window.firebaseAuth, (user) => {
        if (!user) {
          // Limpiar listeners al cerrar sesión
          if (unsubInventario) { unsubInventario(); unsubInventario = null; }
          if (unsubSeriales) { unsubSeriales(); unsubSeriales = null; }
          inventarioCargado = false;
          serialesCargados = false;
          mostrarLogin();
        } else {
          console.log("Usuario logueado:", user.email);

          // Suscribir listeners en tiempo real (onSnapshot)
          cargarDesdeFirestore();
          cargarSerialesDesdeFirestore();

          // 🔐 Logout
          const btn = document.getElementById("btnLogout");
          if (btn) {
            btn.addEventListener("click", () => {
              window.firebaseSignOut(window.firebaseAuth)
              .then(() => {
                location.reload();
              });
            });
          }
        }
      });
    }
  }, 100);
});

// Función para guardar el sistema en localStorage
function guardarSistema() {
    localStorage.setItem('sistemaInventario', JSON.stringify(SistemaInventario));
}

// Función para guardar en Firestore
async function guardarEnFirestore() {
    // Mostrar el overlay de carga
    document.getElementById('loading-overlay').style.display = 'flex';
    
    try {
        const user = window.firebaseAuth.currentUser;
        if (!user) {
            log("No hay usuario logueado para guardar en Firestore", "error");
            return;
        }

        const db = window.firebaseDb;

        log("Guardando inventario en Firestore...", "info");
        
        // Convertir todo el objeto SistemaInventario a string para evitar problemas con arrays anidados
        const sistemaInventarioString = JSON.stringify(SistemaInventario);
        
        const datosAhorro = {
            data: sistemaInventarioString,
            fechaActualizacion: new Date().toISOString()
        };
        
        await window.firebaseSetDoc(
            window.firebaseDoc(db, "usuarios", user.email, "inventario", "datos"),
            datosAhorro
        );

        // Asegurar que el documento padre del usuario exista para que el admin lo detecte
        await window.firebaseSetDoc(
            window.firebaseDoc(db, "usuarios", user.email),
            { ultimaActividad: new Date().toISOString() },
            { merge: true }
        );

        console.log("✅ Inventario guardado en Firestore correctamente");
        log("✅ Inventario guardado exitosamente en Firestore", "success");
    } catch (error) {
        console.error("Error guardando en Firestore:", error);
        log("❌ Error guardando en Firestore: " + error.message, "error");
    } finally {
        // Ocultar el overlay de carga independientemente del resultado
        document.getElementById('loading-overlay').style.display = 'none';
    }
}

// Controla el escudo de carga inicial: deshabilita inputs hasta que los datos estén listos
function actualizarEscudoCarga() {
    const listo = inventarioCargado && serialesCargados;
    // Habilitar/deshabilitar inputs de archivos Excel
    ['fileOrdenes', 'fileColmenas', 'fileCatalogo', 'fileEstructura'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.disabled = !listo;
    });
    // Ocultar overlay cuando todo esté listo
    if (listo) {
        document.getElementById('loading-overlay').style.display = 'none';
        verificarListo();
        console.log("✅ Datos iniciales cargados — interfaz habilitada");
    }
}

// Función para cargar desde Firestore con suscripción en tiempo real (onSnapshot)
function cargarDesdeFirestore() {
    document.getElementById('loading-overlay').style.display = 'flex';

    const user = window.firebaseAuth.currentUser;
    if (!user || !user.email) {
        console.log("No hay usuario logueado para cargar desde Firestore");
        document.getElementById('loading-overlay').style.display = 'none';
        return;
    }

    const db = window.firebaseDb;
    const docRef = window.firebaseDoc(db, "usuarios", user.email, "inventario", "datos");

    // Desuscribir listener anterior si existe (ej. cambio de sesión)
    if (unsubInventario) unsubInventario();

    // Suscripción en tiempo real al inventario principal
    unsubInventario = window.firebaseOnSnapshot(
        docRef,
        { includeMetadataChanges: false },
        (snap) => {
            // Ignorar escrituras locales pendientes para evitar doble procesamiento
            if (snap.metadata.hasPendingWrites) return;

            if (snap.exists()) {
                const docData = snap.data();
                if (docData && docData.data) {
                    const datos = JSON.parse(docData.data);
                    // Preservar colmenas si ya están gestionadas por el listener de colmena_final
                    const colmenasActuales = SistemaInventario.colmenas;
                    Object.assign(SistemaInventario, datos);
                    if (colmenaActual && colmenaActual.length > 0) {
                        SistemaInventario.colmenas = colmenasActuales;
                    }

                    // Actualizar UI solo si no se está procesando un Excel
                    if (!procesandoExcel) {
                        actualizarTablaOrdenes();
                        actualizarTablaCatalogo();
                        verificarListo();
                    }

                    if (!inventarioCargado) {
                        inventarioCargado = true;
                        console.log("✅ Inventario cargado (primer snapshot)");
                        actualizarEscudoCarga();
                    } else {
                        console.log("🔄 Inventario actualizado en tiempo real");
                    }
                }
            } else {
                console.log("No hay inventario guardado aún en Firestore");
                if (!inventarioCargado) {
                    inventarioCargado = true;
                    actualizarEscudoCarga();
                }
            }
        },
        (error) => {
            console.error("Error en listener onSnapshot de inventario:", error.code, error.message);
            // Si falla el listener, marcar como cargado para no bloquear la UI indefinidamente
            if (!inventarioCargado) {
                inventarioCargado = true;
                actualizarEscudoCarga();
            }
        }
    );

    // Cargar colmena_final y registrar listener real-time
    cargarColmenaFinalDesdeFirestore();
}

// Función para guardar resultados de optimización en Firestore
async function guardarResultadoOptimizacion(resultados) {
    // Mostrar el overlay de carga
    document.getElementById('loading-overlay').style.display = 'flex';
    
    try {
        const user = window.firebaseAuth.currentUser;
        if (!user) {
            console.error("No hay usuario logueado");
            return Promise.reject("No hay usuario logueado");
        }

        const db = window.firebaseDb;

        // Añadir timestamp a cada resultado
        const timestamp = new Date().toISOString();
        
        // Convertir los resultados a string para evitar problemas con arrays anidados
        const resultadosString = JSON.stringify(resultados.map(r => ({
            ...r,
            fechaGuardado: timestamp
        })));
        
        const datosConTimestamp = {
            usuario: user.email,
            timestamp: timestamp,
            resultadosString: resultadosString
        };

        // Guardar en la colección 'historial_optimizaciones' usando una referencia de documento
        const docRef = window.firebaseDoc(db, "historial_optimizaciones", timestamp + "_" + user.email);
        await window.firebaseSetDoc(docRef, datosConTimestamp);

        console.log("✅ Resultados de optimización guardados en historial:", timestamp);
        log("✅ Resultados de optimización guardados correctamente", "success");
        return Promise.resolve(timestamp);
    } catch (error) {
        console.error("Error guardando resultados de optimización:", error);
        log("❌ Error guardando resultados: " + error.message, "error");
        return Promise.reject(error);
    } finally {
        // Ocultar el overlay de carga independientemente del resultado
        document.getElementById('loading-overlay').style.display = 'none';
    }
}

// ─── FUNCIONES DE COLMENA FINAL (Firebase) ───────────────────────────────────

// Actualiza el indicador visual de fuente de colmena en la UI
function actualizarIndicadorFuente(esManual) {
    const indicador = document.getElementById('indicadorFuenteColmena');
    if (!indicador) return;
    if (esManual) {
        indicador.textContent = '📂 Usando inventario manual cargado';
        indicador.style.backgroundColor = '#27ae60';
        indicador.style.display = 'inline-block';
    } else {
        indicador.textContent = '☁️ Usando inventario sincronizado de Firebase';
        indicador.style.backgroundColor = '#3498db';
        indicador.style.display = 'inline-block';
    }
}

// Carga la colmena_final guardada en Firebase y la asigna a colmenaActual
async function cargarColmenaFinalDesdeFirestore() {
    const user = window.firebaseAuth.currentUser;
    if (!user || !user.email) return;

    const db = window.firebaseDb;
    const docRef = window.firebaseDoc(db, "usuarios", user.email, "colmena_final", "datos");

    // Registrar el listener SIEMPRE (independiente de si el doc existe o no)
    // onSnapshot dispara inmediatamente con el estado actual y luego en cada cambio
    window.firebaseOnSnapshot(
        docRef,
        { includeMetadataChanges: false },
        (snap) => {
            if (snap.exists() && !snap.metadata.hasPendingWrites) {
                const data = snap.data();
                if (data && data.data) {
                    colmenaActual = JSON.parse(data.data) || [];
                    SistemaInventario.colmenas = colmenaActual;
                    actualizarTablaColmenas();
                    actualizarIndicadorFuente(false);
                    verificarListo();
                    console.log("🔄 Colmena final actualizada desde Firebase:", colmenaActual.length, "registros");
                }
            } else if (!snap.exists()) {
                // Sin colmena_final guardada: usar SistemaInventario.colmenas como fallback
                if (SistemaInventario.colmenas && SistemaInventario.colmenas.length > 0) {
                    colmenaActual = SistemaInventario.colmenas;
                    actualizarIndicadorFuente(false);
                    verificarListo();
                    console.log(`ℹ️ Sin colmena_final en Firebase. Usando colmenas del inventario (${colmenaActual.length})`);
                } else {
                    console.log("ℹ️ No hay colmena_final ni colmenas en inventario aún.");
                }
            }
        },
        (error) => {
            // Error callback: evita "Message channel closed" por promesas sin handler
            console.error("Error en listener onSnapshot de colmena_final:", error.code, error.message);
        }
    );

    // Forzar lectura desde servidor para la sesión inicial (sin caché)
    try {
        const docSnap = await window.firebaseGetDocFromServer(docRef);
        if (docSnap && docSnap.exists()) {
            const docData = docSnap.data();
            if (docData && docData.data) {
                colmenaActual = JSON.parse(docData.data);
                SistemaInventario.colmenas = colmenaActual;
                actualizarTablaColmenas();
                actualizarIndicadorFuente(false);
                verificarListo();
                console.log(`✅ Colmena final cargada desde servidor: ${colmenaActual.length} registros`);
            }
        }
    } catch (error) {
        console.warn("getDocFromServer falló, el listener onSnapshot seguirá activo:", error.message);
    }
}

// Guarda las colmenas disponibles post-optimización en Firebase como colmena_final
async function guardarColmenaFinalEnFirestore() {
    document.getElementById('loading-overlay').style.display = 'flex';
    try {
        const user = window.firebaseAuth.currentUser;
        if (!user) {
            log("No hay usuario logueado para guardar colmena final", "error");
            return;
        }

        const db = window.firebaseDb;

        // Extraer colmenas disponibles del histórico y mapear al formato simple
        // IMPORTANTE: incluir serial para preservar la trazabilidad en optimizaciones sucesivas
        const colmenasFinal = SistemaInventario.colmenasHistorico
            .filter(c => c.estado === 'disponible' && c.medida_mm > 0 && c.cod)
            .map(c => ({
                n_colmena: c.n_colmena,
                medida_mm: c.medida_mm,
                medida_cm: c.medida_cm,
                cod: c.cod,
                serial: c.serial || null
            }));

        const datos = {
            data: JSON.stringify(colmenasFinal),
            fechaActualizacion: new Date().toISOString()
        };

        // Asegurar que el documento padre del usuario exista
        await window.firebaseSetDoc(
            window.firebaseDoc(db, "usuarios", user.email),
            { ultimaActividad: new Date().toISOString() },
            { merge: true }
        );

        await window.firebaseSetDoc(
            window.firebaseDoc(db, "usuarios", user.email, "colmena_final", "datos"),
            datos
        );

        // Actualizar colmenaActual con el resultado para la sesión actual
        colmenaActual = colmenasFinal;

        // Actualizar SistemaInventario.colmenas y la tabla de la UI para que
        // coincida con el Excel exportado (colmenas disponibles post-optimización)
        SistemaInventario.colmenas = colmenasFinal;
        actualizarTablaColmenas();
        const estadoEl = document.getElementById('estadoColmenas');
        if (estadoEl) {
            estadoEl.textContent = `✓ ${colmenasFinal.length} colmenas (sincronizadas)`;
            estadoEl.className = 'estado-archivo estado-ok';
        }

        console.log(`✅ Colmena final guardada en Firebase: ${colmenasFinal.length} registros`);
        log(`✅ Colmena final guardada en Firebase (${colmenasFinal.length} colmenas disponibles)`, "success");
    } catch (error) {
        console.error("Error guardando colmena_final en Firestore:", error);
        log("❌ Error guardando colmena final en Firebase: " + error.message, "error");
    } finally {
        document.getElementById('loading-overlay').style.display = 'none';
    }
}

// ─────────────────────────────────────────────────────────────────────────────

// Función para cargar el sistema desde localStorage
function cargarSistema() {
    const datosGuardados = localStorage.getItem('sistemaInventario');
    if (datosGuardados) {
        const datos = JSON.parse(datosGuardados);
        SistemaInventario.ordenes = datos.ordenes || [];
        SistemaInventario.colmenas = datos.colmenas || [];
        SistemaInventario.catalogoReemplazos = datos.catalogoReemplazos || {};
        SistemaInventario.catalogoColores = datos.catalogoColores || {};
        SistemaInventario.catalogoMedidas = datos.catalogoMedidas || {};
        SistemaInventario.logs = datos.logs || [];
        SistemaInventario.colmenasDisponibles = datos.colmenasDisponibles || [];
        SistemaInventario.datosCrudosOrdenes = datos.datosCrudosOrdenes || [];
        SistemaInventario.resultadosOptimizacion = datos.resultadosOptimizacion || [];
        SistemaInventario.colmenasHistorico = datos.colmenasHistorico || [];
        SistemaInventario.mermas = datos.mermas || [];
        SistemaInventario.catalogoAccesorios = datos.catalogoAccesorios || {};
    }
}

// Cargar automáticamente al inicio
cargarSistema();

// ─── FUNCIONES DE SERIALES (Estructura de Inventario) ───────────────────────────

// Función para cargar el archivo de estructura de inventario
async function cargarEstructuraInventario(event) {
    const file = event.target.files[0];
    if (!file) return;
    procesandoExcel = true;
    document.getElementById('loading-overlay').style.display = 'flex';

    try {
        const datos = await leerExcelCompleto(file);
        const filaEncabezado = detectarFilaEncabezado(datos);
        const encabezados = datos[filaEncabezado];
        
        // Detectar columnas requeridas: FECHA, CODIGO, LOTE, PAQUETE, SERIAL
        let colFecha = -1, colCodigo = -1, colLote = -1, colPaquete = -1, colSerial = -1;
        
        for (let i = 0; i < encabezados.length; i++) {
            const enc = String(encabezados[i] || '').trim().toUpperCase();
            if (enc.includes('FECHA')) colFecha = i;
            if (enc.includes('CODIGO') || enc === 'COD') colCodigo = i;
            if (enc.includes('LOTE')) colLote = i;
            if (enc.includes('PAQUETE')) colPaquete = i;
            if (enc.includes('SERIAL')) colSerial = i;
        }
        
        // Verificar que se encontraron todas las columnas requeridas
        if (colFecha === -1 || colCodigo === -1 || colLote === -1 || colPaquete === -1 || colSerial === -1) {
            alert('No se encontraron todas las columnas requeridas: FECHA, CODIGO, LOTE, PAQUETE, SERIAL');
            document.getElementById('loading-overlay').style.display = 'none';
            return;
        }
        
        // Procesar los datos
        SistemaInventario.seriales = [];
        
        for (let i = filaEncabezado + 1; i < datos.length; i++) {
            const fila = datos[i];
            if (!fila) continue;
            
            // Asegurarse de que los campos lote, paquete y serial se capturen siempre como strings limpios
            const fechaLimpia = formatearFecha(fila[colFecha]);
            const codigoStr = String(fila[colCodigo] || '').trim().toUpperCase();
            const loteStr = String(fila[colLote] || '').trim() || '-';
            const paqueteStr = String(fila[colPaquete] || '').trim() || '-';
            const serialStr = String(fila[colSerial] || '').trim() || '-';
            
            if (codigoStr) {
                // "Serial" en el archivo es la CANTIDAD de tubos en ese paquete
                const cantidadStr = String(fila[colSerial] || '1').trim();
                const cantidadTubos = parseInt(cantidadStr, 10) || 1;
                for (let t = 0; t < cantidadTubos; t++) {
                    SistemaInventario.seriales.push({
                        fecha: fechaLimpia,
                        codigo: codigoStr,
                        lote: loteStr,
                        paquete: paqueteStr,
                        serial: (t + 1).toString(), // Numeración secuencial: tubo 1, 2, ... N dentro del paquete
                        estado: 'disponible'
                    });
                }
            }
        }
        
        // Ordenar los seriales por FECHA (más antigua primero), LOTE, PAQUETE, SERIAL
        SistemaInventario.seriales.sort((a, b) => {
            // Primero por fecha (más antigua primero)
            const fechaA = new Date(a.fecha);
            const fechaB = new Date(b.fecha);
            if (fechaA < fechaB) return -1;
            if (fechaA > fechaB) return 1;
            
            // Luego por lote (numérico si es posible)
            const loteANum = parseInt(a.lote);
            const loteBNum = parseInt(b.lote);
            if (!isNaN(loteANum) && !isNaN(loteBNum)) {
                if (loteANum < loteBNum) return -1;
                if (loteANum > loteBNum) return 1;
            } else {
                const cmpLote = a.lote.localeCompare(b.lote);
                if (cmpLote !== 0) return cmpLote;
            }
            
            // Luego por paquete (numérico si es posible)
            const paqueteANum = parseInt(a.paquete);
            const paqueteBNum = parseInt(b.paquete);
            if (!isNaN(paqueteANum) && !isNaN(paqueteBNum)) {
                if (paqueteANum < paqueteBNum) return -1;
                if (paqueteANum > paqueteBNum) return 1;
            } else {
                const cmpPaquete = a.paquete.localeCompare(b.paquete);
                if (cmpPaquete !== 0) return cmpPaquete;
            }
            
            // Finalmente por serial (numérico si es posible)
            const serialANum = parseInt(a.serial);
            const serialBNum = parseInt(b.serial);
            if (!isNaN(serialANum) && !isNaN(serialBNum)) {
                return serialANum - serialBNum;
            }
            return a.serial.localeCompare(b.serial);
        });
        
        // Actualizar la tabla de seriales
        actualizarTablaSeriales();
        
        // Verificar alertas de stock tras cargar el inventario
        verificarAlertasStock();
        
        // Actualizar el estado en la UI
        document.getElementById('estadoEstructura').textContent = `✓ ${SistemaInventario.seriales.length} seriales`;
        document.getElementById('estadoEstructura').className = 'estado-archivo estado-ok';
        
        log(`🏷️ Estructura de inventario cargada: ${SistemaInventario.seriales.length} seriales`, 'success');
        
        // Guardar en Firebase
        await guardarSerialesEnFirestore();
        
    } catch (e) {
        alert('Error: ' + e.message);
        console.error(e);
    } finally {
        procesandoExcel = false;
        document.getElementById('loading-overlay').style.display = 'none';
    }
}

// Función para actualizar la tabla de seriales en la UI
function actualizarTablaSeriales() {
    const tbody = document.getElementById('tbodySeriales');
    if (!tbody) return;
    
    if (SistemaInventario.seriales.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6">Sin datos</td></tr>';
        return;
    }
    
    // Mostrar solo los primeros 100 seriales para no sobrecargar la UI
    const serialesAMostrar = SistemaInventario.seriales.slice(0, 100);
    
    tbody.innerHTML = serialesAMostrar.map(s => {
        let estadoBadge = '';
        switch(s.estado) {
            case 'disponible': 
                estadoBadge = '<span style="background-color: #27ae60; color: white; padding: 2px 6px; border-radius: 3px; font-size: 10px;">DISPONIBLE</span>';
                break;
            case 'ocupado': 
                estadoBadge = '<span style="background-color: #e74c3c; color: white; padding: 2px 6px; border-radius: 3px; font-size: 10px;">OCUPADO</span>';
                break;
            default: 
                estadoBadge = s.estado || '';
        }
        
        return `<tr>
            <td>${s.codigo}</td>
            <td>${formatearFecha(s.fecha)}</td>
            <td>${s.lote}</td>
            <td>${s.paquete}</td>
            <td>${s.serial}</td>
            <td>${estadoBadge}</td>
        </tr>`;
    }).join('');
    
    if (SistemaInventario.seriales.length > 100) {
        tbody.innerHTML += `<tr><td colspan="6">... y ${SistemaInventario.seriales.length - 100} más</td></tr>`;
    }
}

// Función para guardar los seriales en Firestore usando writeBatch
async function guardarSerialesEnFirestore() {
    document.getElementById('loading-overlay').style.display = 'flex';
    
    try {
        const user = window.firebaseAuth.currentUser;
        if (!user) {
            log("No hay usuario logueado para guardar seriales", "error");
            document.getElementById('loading-overlay').style.display = 'none';
            return;
        }

        const db = window.firebaseDb;
        if (!db) throw new Error("Firestore no está inicializado");
        
        log("Guardando seriales en Firestore...", "info");
        
        const coleccionRef = window.firebaseCollection(db, "usuarios", user.email, "maestro_seriales");
        
        // ── FASE 1: Eliminar todos los documentos existentes (en batches de 499) ──
        const docsExistentes = await window.firebaseGetDocs(coleccionRef);
        let batch = window.firebaseWriteBatch(db);
        let opCount = 0;

        for (const docSnap of docsExistentes.docs) {
            batch.delete(docSnap.ref);
            opCount++;
            if (opCount >= 499) {
                await batch.commit();
                batch = window.firebaseWriteBatch(db);
                opCount = 0;
            }
        }
        if (opCount > 0) {
            await batch.commit();
        }
        
        // ── FASE 2: Insertar los nuevos documentos (en batches de 499) ──
        batch = window.firebaseWriteBatch(db);
        opCount = 0;

        SistemaInventario.seriales.forEach(serial => {
            const docId = `${serial.codigo}_${serial.lote}_${serial.paquete}_${serial.serial}`.replace(/[\/\s]/g, '_');
            const docRef = window.firebaseDoc(coleccionRef, docId);
            batch.set(docRef, {
                ...serial,
                usuario: user.email,
                fechaActualizacion: new Date().toISOString()
            });
            opCount++;
            if (opCount >= 499) {
                // Crear promesa y resetear (no await dentro de forEach)
                batch.commit();
                batch = window.firebaseWriteBatch(db);
                opCount = 0;
            }
        });
        if (opCount > 0) {
            await batch.commit();
        }
        
        console.log(`✅ ${SistemaInventario.seriales.length} seriales guardados en Firestore`);
        log(`✅ ${SistemaInventario.seriales.length} seriales guardados en Firestore`, "success");
        
    } catch (error) {
        console.error("Error guardando seriales en Firestore:", error);
        if (error.code === 'permission-denied' || error.message.includes('permission')) {
            log("❌ Error de permisos: No tienes acceso a Firestore. Verifica las reglas de Firebase.", "error");
        } else {
            log("❌ Error guardando seriales en Firestore: " + error.message, "error");
        }
    } finally {
        document.getElementById('loading-overlay').style.display = 'none';
    }
}


// Lógica de ordenamiento de seriales (extraída para reutilizar)
function ordenarSeriales(seriales) {
    seriales.sort((a, b) => {
        // Por fecha (más antigua primero)
        const fechaA = new Date(a.fecha);
        const fechaB = new Date(b.fecha);
        if (fechaA < fechaB) return -1;
        if (fechaA > fechaB) return 1;

        // Por lote
        const loteANum = parseInt(a.lote);
        const loteBNum = parseInt(b.lote);
        if (!isNaN(loteANum) && !isNaN(loteBNum)) {
            if (loteANum < loteBNum) return -1;
            if (loteANum > loteBNum) return 1;
        } else {
            const cmpLote = a.lote.localeCompare(b.lote);
            if (cmpLote !== 0) return cmpLote;
        }

        // Por paquete
        const paqueteANum = parseInt(a.paquete);
        const paqueteBNum = parseInt(b.paquete);
        if (!isNaN(paqueteANum) && !isNaN(paqueteBNum)) {
            if (paqueteANum < paqueteBNum) return -1;
            if (paqueteANum > paqueteBNum) return 1;
        } else {
            const cmpPaquete = a.paquete.localeCompare(b.paquete);
            if (cmpPaquete !== 0) return cmpPaquete;
        }

        // Por serial
        const serialANum = parseInt(a.serial);
        const serialBNum = parseInt(b.serial);
        if (!isNaN(serialANum) && !isNaN(serialBNum)) {
            return serialANum - serialBNum;
        }
        return a.serial.localeCompare(b.serial);
    });
}

// Función para cargar los seriales desde Firestore con suscripción en tiempo real
function cargarSerialesDesdeFirestore() {
    const user = window.firebaseAuth.currentUser;
    if (!user) return;

    const db = window.firebaseDb;

    // Desuscribir listener anterior si existe
    if (unsubSeriales) unsubSeriales();

    const coleccionRef = window.firebaseCollection(db, "usuarios", user.email, "maestro_seriales");

    unsubSeriales = window.firebaseOnSnapshot(
        coleccionRef,
        { includeMetadataChanges: false },
        (querySnapshot) => {
            SistemaInventario.seriales = [];

            querySnapshot.forEach(doc => {
                const data = doc.data();
                SistemaInventario.seriales.push({
                    fecha: data.fecha,
                    codigo: data.codigo,
                    lote: data.lote,
                    paquete: data.paquete,
                    serial: data.serial,
                    estado: data.estado || 'disponible'
                });
            });

            ordenarSeriales(SistemaInventario.seriales);

            // Actualizar UI solo si no se está procesando un Excel
            if (!procesandoExcel) {
                actualizarTablaSeriales();
                verificarAlertasStock();
            }

            const estadoEl = document.getElementById('estadoEstructura');
            if (estadoEl) {
                estadoEl.textContent = `✓ ${SistemaInventario.seriales.length} seriales (sincronizados)`;
                estadoEl.className = 'estado-archivo estado-ok';
            }

            if (!serialesCargados) {
                serialesCargados = true;
                console.log(`✅ ${SistemaInventario.seriales.length} seriales cargados (primer snapshot)`);
                actualizarEscudoCarga();
            } else {
                console.log(`🔄 Seriales actualizados en tiempo real: ${SistemaInventario.seriales.length}`);
            }
        },
        (error) => {
            console.error("Error en listener onSnapshot de seriales:", error.code, error.message);
            if (!serialesCargados) {
                serialesCargados = true;
                actualizarEscudoCarga();
            }
        }
    );
}

// Función para buscar un serial disponible para un código específico
function buscarSerialDisponible(codigo) {
    // Normalizar el código
    const codigoNormalizado = normalizarCodigo(codigo);

    // Buscar directamente el índice en el array fuente (FIFO)
    const idx = SistemaInventario.seriales.findIndex(s =>
        normalizarCodigo(s.codigo) === codigoNormalizado && s.estado === 'disponible'
    );

    if (idx === -1) return null;

    // CRÍTICO: Marcar como ocupado INMEDIATAMENTE para evitar que el mismo
    // serial sea seleccionado en la siguiente iteración del bucle de optimización
    SistemaInventario.seriales[idx].estado = 'ocupado';

    return SistemaInventario.seriales[idx];
}

// Función para marcar un serial como ocupado (complementa buscarSerialDisponible)
function marcarSerialComoOcupado(serial, ordenId) {
    // Buscar el serial en la lista (usar normalizarCodigo para consistencia)
    const idx = SistemaInventario.seriales.findIndex(s =>
        normalizarCodigo(s.codigo) === normalizarCodigo(serial.codigo) &&
        s.lote === serial.lote &&
        s.paquete === serial.paquete &&
        s.serial === serial.serial
    );

    if (idx !== -1) {
        // Actualizar estado y metadata de trazabilidad
        SistemaInventario.seriales[idx].estado = 'ocupado';
        SistemaInventario.seriales[idx].ordenId = ordenId;
        SistemaInventario.seriales[idx].fechaUso = new Date().toISOString();

        // Actualizar la tabla
        actualizarTablaSeriales();

        return true;
    }

    return false;
}

const COLUMNAS_ESPECIFICAS = [
    { key: 'con_tira', titulo: 'CON TIRA', buscar: ['con tira'] },
    { key: 'peso_u', titulo: 'PESO U', buscar: ['peso u'] },
    { key: 'peso_interno', titulo: 'PESO INTERNO', buscar: ['peso interno'] },
    { key: 'pletina', titulo: 'PLETINA', buscar: ['pletina'] },
    { key: 'perfil_cortina', titulo: 'PERFIL [CORTINA VERTICAL]', buscar: ['perfil [cortina vertical]'] },
    { key: 'varilla_cortina', titulo: 'VARILLA [CORTINA VERTICAL]', buscar: ['varilla [cortina vertical]'] },
    { key: 'perfil_izq_int', titulo: 'PERFIL (IZQ) INT', buscar: ['perfil (izq) int'] },
    { key: 'color_perfil', titulo: 'COLOR PERFIL', buscar: ['color perfil'] },
    { key: 'separador_superior', titulo: 'SEPARADOR SUPERIOR', buscar: ['separador superior'] },
    { key: 'separador_lateral', titulo: 'SEPARADOR LATERAL', buscar: ['separador lateral'] },
    { key: 'perfil_der_int', titulo: 'PERFIL (DER) INT', buscar: ['perfil (der) int'] },
    { key: 'color_peso_inf', titulo: 'COLOR PESO INF. SOFT LIGHT', buscar: ['color peso inf. soft light'] },
    { key: 'peso_soft_light', titulo: 'PESO SOFT LIGHT', buscar: ['peso soft light'] },
    { key: 'cenefa_delantera', titulo: 'CENEFA DELANTERA', buscar: ['cenefa_delantera', 'cenefa delantea'] },
    { key: 'cenefa_trasera', titulo: 'CENEFA TRASERA', buscar: ['cenefa_trasera', 'cenefa trasera'] },
    { key: 'cenefa_ovalada', titulo: 'CENEFA OVALADA', buscar: ['cenefa ovalada', 'cenefa_ovalada'] },
    { key: 'perfil_base', titulo: 'PERFIL BASE', buscar: ['perfil base'] }
];

const COLUMNAS_ESPECIALES = {
    'medida': { key: 'medida_cm', titulo: 'Medida(cm)', esNumero: true, buscar: ['ancho real', 'alto real', 'medida', 'corte', 'longitud'] },
    'codsec': { key: 'codSec', titulo: 'COD SEC', esNumero: false, buscar: ['cod sec'] },
    'tuberia': { key: 'tuberia', titulo: 'TUBERIA', esNumero: false, buscar: ['tuberia'] },
    'codigo': { key: 'codigoExtraido', titulo: 'Código', esNumero: false, buscar: [], extraerDe: 'tuberia' },
    'reemplazo': { key: 'reemplazo', titulo: 'Reemplazo', esNumero: false, buscar: [], deriving: true },
    'ubic': { key: 'ubic', titulo: 'UBIC.', esNumero: false, buscar: ['ubic.'] },
    'color': { key: 'color', titulo: 'COLOR', esNumero: false, buscar: ['color accesorios'] },
    'tubo': { key: 'tubo', titulo: 'TUBO', esNumero: true, buscar: ['tubo'] },
    'peso': { key: 'peso', titulo: 'PESO', esNumero: true, buscar: ['peso'] },
    'ot': { key: 'otAsignada', titulo: 'OT', esNumero: false, buscar: ['ot'] }
};

function ordenarColmena(a, b) {
    const extraerNumero = (str) => {
        const match = String(str).match(/([A-Za-z]*)(\d+)/);
        if (match) {
            const letra = match[1].toUpperCase();
            const numero = parseInt(match[2], 10);
            return { letra: letra, numero: numero };
        }
        return { letra: String(str).toUpperCase(), numero: 0 };
    };
    
    const aParsed = extraerNumero(a);
    const bParsed = extraerNumero(b);
    
    if (aParsed.letra < bParsed.letra) return -1;
    if (aParsed.letra > bParsed.letra) return 1;
    if (aParsed.numero < bParsed.numero) return -1;
    if (aParsed.numero > bParsed.numero) return 1;
    return 0;
}

function detectarFilaEncabezado(datos) {
    // Palabras que identifican inequívocamente una fila de encabezado (match exacto)
    const EXACTOS = new Set(['ot', 'tuberia', 'tubo', 'medida', 'corte', 'longitud']);
    // Palabras que se buscan como substrings (son suficientemente específicas)
    const PARCIALES = ['ancho real', 'alto real', 'cod sec'];

    for (let i = 0; i < Math.min(datos.length, 30); i++) {
        if (!datos[i]) continue;
        const celdas = datos[i]
            .filter(c => c !== null && c !== undefined && String(c).trim() !== '')
            .map(c => String(c).trim().toLowerCase());
        if (celdas.length === 0) continue;

        let hits = 0;
        for (const c of celdas) {
            if (EXACTOS.has(c) || PARCIALES.some(p => c.includes(p))) hits++;
        }
        // ≥2 indicadores, o cualquiera muy específico como 'ot' o 'tuberia'
        if (hits >= 2 || celdas.some(c => c === 'ot' || c === 'tuberia')) return i;
    }

    // Fallback: primera fila con cualquier texto
    for (let i = 0; i < Math.min(datos.length, 30); i++) {
        if (!datos[i]) continue;
        if (datos[i].some(cell => typeof cell === 'string' && cell.trim().length > 0)) return i;
    }
    return 0;
}

function leerExcelCompleto(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellNF: true });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const range = XLSX.utils.decode_range(firstSheet['!ref'] || 'A1');
            const datos = [];
            for (let R = range.s.r; R <= range.e.r; ++R) {
                const fila = [];
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const addr = XLSX.utils.encode_cell({ r: R, c: C });
                    const cell = firstSheet[addr];
                    if (!cell) fila.push(null);
                    else if (cell.f && cell.f.includes('[')) fila.push(cell.v !== undefined ? cell.v : null);
                    else fila.push(cell.v);
                }
                datos.push(fila);
            }
            resolve(datos);
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// Función para formatear fecha ISO a formato DD/MM/YYYY
// Maneja tanto strings ISO como objetos Date de JavaScript



function formatearNumero(valor) {
    if (valor === null || valor === undefined || valor === '') return null;
    // Usar limpiarNumero para manejar tanto comas como puntos
    const num = limpiarNumero(valor);
    return num === 0 ? null : Math.round(num * 100) / 100;
}

function parsearNumero(valor) {
    if (valor === null || valor === undefined || valor === '') return null;
    // Usar limpiarNumero para manejar tanto comas como puntos
    const num = limpiarNumero(valor);
    return num === 0 ? null : num;
}

  // Función para validar y limpiar datos de órdenes o colmenas
  function validarYLimpiarDatos(datos, tipo) {
      if (!Array.isArray(datos) || datos.length === 0) {
          console.warn("No hay datos para validar");
          return [];
      }

      const datosLimpios = [];

      if (tipo === 'orden') {
          // Validar órdenes: debe tener 'Pedido' y 'Medida'
          for (let i = 0; i < datos.length; i++) {
              const item = datos[i];
              
              // Verificar que tenga las propiedades requeridas
              if (!item.hasOwnProperty('Pedido') || !item.hasOwnProperty('Medida')) {
                  console.warn(`Registro ${i + 1} eliminado: falta 'Pedido' o 'Medida'`);
                  continue;
              }

              // Verificar que los datos no estén vacíos
              const pedido = item.Pedido;
              const medidaRaw = item.Medida;

              if (pedido === null || pedido === undefined || String(pedido).trim() === '') {
                  console.warn(`Registro ${i + 1} eliminado: Pedido vacío`);
                  continue;
              }

              // Convertir Medida a número (float) usando limpiarNumero
              let medidaNum = null;
              if (medidaRaw !== null && medidaRaw !== undefined && medidaRaw !== '') {
                  medidaNum = limpiarNumero(medidaRaw);
              }

              if (medidaNum === null || isNaN(medidaNum) || medidaNum <= 0) {
                  console.warn(`Registro ${i + 1} eliminado: Medida inválida`);
                  continue;
              }

              // Normalizar textos (quitar espacios en blanco)
              const itemLimpio = {
                  Pedido: String(pedido).trim(),
                  Medida: medidaNum
              };

              // Copiar otras propiedades existentes y normalizar códigos
              for (const key of Object.keys(item)) {
                  if (key !== 'Pedido' && key !== 'Medida') {
                      let valor = item[key];
                      if (typeof valor === 'string') {
                          valor = valor.trim();
                      }
                      itemLimpio[key] = valor;
                  }
              }

              datosLimpios.push(itemLimpio);
          }
      } else if (tipo === 'colmena') {
          // Validar colmenas: debe tener 'N° Colmena' y 'Medida (cm)'
          for (let i = 0; i < datos.length; i++) {
              const item = datos[i];
              
              // Verificar que tenga las propiedades requeridas
              if (!item.hasOwnProperty('N° Colmena') || !item.hasOwnProperty('Medida (cm)')) {
                  console.warn(`Registro ${i + 1} eliminado: falta 'N° Colmena' o 'Medida (cm)'`);
                  continue;
              }

              // Verificar que los datos no estén vacíos
              const nColmena = item['N° Colmena'];
              const medidaRaw = item['Medida (cm)'];

              if (nColmena === null || nColmena === undefined || String(nColmena).trim() === '') {
                  console.warn(`Registro ${i + 1} eliminado: N° Colmena vacío`);
                  continue;
              }

              // Convertir Medida (cm) a número usando limpiarNumero
              let medidaNum = null;
              if (medidaRaw !== null && medidaRaw !== undefined && medidaRaw !== '') {
                  medidaNum = limpiarNumero(medidaRaw);
              }

              if (medidaNum === null || isNaN(medidaNum) || medidaNum <= 0) {
                  console.warn(`Registro ${i + 1} eliminado: Medida (cm) inválida`);
                  continue;
              }

              // Normalizar textos (quitar espacios en blanco)
              const itemLimpio = {
                  'N° Colmena': String(nColmena).trim(),
                  'Medida (cm)': medidaNum
              };

              // Copiar otras propiedades existentes y normalizar códigos
              for (const key of Object.keys(item)) {
                  if (key !== 'N° Colmena' && key !== 'Medida (cm)') {
                      let valor = item[key];
                      if (typeof valor === 'string') {
                          valor = valor.trim();
                      }
                      itemLimpio[key] = valor;
                  }
              }

              datosLimpios.push(itemLimpio);
          }
      } else {
          console.error("Tipo no válido. Use 'orden' o 'colmena'");
          return [];
      }

      console.log(`✅ Validación completada: ${datosLimpios.length} registros válidos de ${datos.length}`);
      return datosLimpios;
  }

  function extraerCodigoDesdeTuberia(tuberia) {
    if (!tuberia) return null;
    const str = String(tuberia).trim();
    const match = str.match(/_([A-Za-z0-9]+)$/);
    if (match) return match[1].toUpperCase();
    return str.toUpperCase();
}

// Términos cortos o ambiguos que requieren match exacto para no colisionar
// ej: 'tubo' ⊂ 'tuberia', 'ot' ⊂ 'nota'/'rotacion'
const TERMINOS_EXACTOS = new Set(['tubo', 'ot', 'peso', 'color']);

function _coincide(enc, busqueda) {
    return TERMINOS_EXACTOS.has(busqueda) ? enc === busqueda : enc.includes(busqueda);
}

function detectarColumnasExcel(encabezados) {
    const mapeo = {};
    for (let i = 0; i < encabezados.length; i++) {
        const enc = String(encabezados[i] || '').trim().toLowerCase();
        for (const [nombre, config] of Object.entries(COLUMNAS_ESPECIALES)) {
            if (mapeo[nombre] !== undefined) continue;
            if (config.buscar && config.buscar.length > 0) {
                for (const busqueda of config.buscar) {
                    if (_coincide(enc, busqueda)) { mapeo[nombre] = i; break; }
                }
            }
        }
    }
    for (let i = 0; i < encabezados.length; i++) {
        const enc = String(encabezados[i] || '').trim().toLowerCase();
        for (const colEsp of COLUMNAS_ESPECIFICAS) {
            const key = colEsp.key;
            if (mapeo[key] !== undefined) continue;
            if (colEsp.buscar && colEsp.buscar.length > 0) {
                for (const busqueda of colEsp.buscar) {
                    if (_coincide(enc, busqueda)) { mapeo[key] = i; break; }
                }
            }
        }
    }
    return mapeo;
}

function detectarColumnasConDatos(ordenes) {
    const columnasConDatos = [];
    columnasConDatos.push({ key: 'id', titulo: 'ID', esNumero: false });
    for (const [nombre, config] of Object.entries(COLUMNAS_ESPECIALES)) {
        const tieneDatos = ordenes.some(ord => { const valor = ord[nombre]; return valor !== null && valor !== undefined; });
        if (tieneDatos) columnasConDatos.push({ key: nombre, titulo: config.titulo, esNumero: config.esNumero });
    }
    for (const colEsp of COLUMNAS_ESPECIFICAS) {
        const tieneDatos = ordenes.some(ord => { const valor = ord[colEsp.key]; return valor !== null && valor !== undefined; });
        if (tieneDatos) columnasConDatos.push({ key: colEsp.key, titulo: colEsp.titulo, esNumero: false });
    }
    return columnasConDatos;
}

function actualizarReemplazosEnOrdenes() {
    if (SistemaInventario.ordenes.length === 0) return;
    let actualizados = 0;
    SistemaInventario.ordenes.forEach(orden => {
        if (orden.codigoExtraido) {
            const reemplazo = SistemaInventario.catalogoReemplazos[orden.codigoExtraido] || null;
            orden.reemplazo = reemplazo;
            if (reemplazo) actualizados++;
        }
    });
    if (actualizados > 0) { log(`✓ Se actualizaron ${actualizados} reemplazos en las órdenes`, 'info'); actualizarTablaOrdenes(); }
}

// ─── Función reutilizable: expande datosCrudosOrdenes en órdenes multi-corte ───
function expandirOrdenesDesdeExcel() {
    const filaEncabezado = detectarFilaEncabezado(SistemaInventario.datosCrudosOrdenes);
    const encabezados = SistemaInventario.datosCrudosOrdenes[filaEncabezado];
    const mapeoColumnas = detectarColumnasExcel(encabezados);

    // Detección dinámica de columna TUBO
    const _normH = h => String(h || '').replace(/[\s\u00A0\t]+/g, ' ').trim().toUpperCase();
    let idxTubo = encabezados.findIndex(h => _normH(h) === 'TUBO');
    if (idxTubo === -1) {
        idxTubo = encabezados.findIndex(h => {
            const n = _normH(h);
            return n.includes('TUBO') && !n.includes('TUBERIA');
        });
    }
    console.log("📍 Columna TUBO detectada dinámicamente en índice:", idxTubo);
    log(`🔍 Columna TUBO en índice: ${idxTubo !== -1 ? idxTubo : 'NO ENCONTRADA (usando fallback)'}`, 'info');

    // ─── MULTI-CORTE: Columnas de componentes a procesar ───
    const COLUMNAS_CORTE = [
        'TUBO', 'PESO', 'CENEFA OVALADA', 'PESO U', 'PESO INTERNO',
        'PLETINA', 'PERFIL [CORTINA VERTICAL]', 'VARILLA [CORTINA VERTICAL]',
        'PERFIL (IZQ) INT', 'PERFIL (DER) INT', 'PERFIL BASE',
        'CENEFA DELANTERA', 'CENEFA TRASERA', 'PESO SOFT LIGHT'
    ];

    // Traductor: nombre columna en Libro4 → nombre exacto en el Catálogo
    const MAPA_NOMBRES_CATALOGO = {
        'PERFIL [CORTINA VERTICAL]': 'CABEZAL VERTICAL',
        'VARILLA [CORTINA VERTICAL]': 'VARILLA VERTICAL 4 PUNTAS',
        'PERFIL (IZQ) INT': 'PERFIL IZQUIERDO INTERNO',
        'PERFIL (DER) INT': 'PERFIL DERECHO INTERNO',
        'PESO U': 'PESO INFERIOR DE DÚO LÁGRIMA',
        'PESO': 'PESO INFERIOR ROLLER'
    };

    // Mapa explícito: componente → columna de color que le corresponde
    const MAPA_COLUMNAS_COLOR = {
        'CENEFA OVALADA': 'COLOR ACCESORIOS',
        'PESO U': 'COLOR ACCESORIOS',
        'PESO INTERNO': 'COLOR ACCESORIOS',
        'PLETINA': 'COLOR ACCESORIOS',
        'PERFIL [CORTINA VERTICAL]': 'COLOR PERFIL',
        'VARILLA [CORTINA VERTICAL]': 'COLOR PERFIL',
        'PERFIL (IZQ) INT': 'COLOR PERFIL',
        'PERFIL (DER) INT': 'COLOR PERFIL',
        'PERFIL BASE': 'COLOR PERFIL',
        'CENEFA DELANTERA': 'COLOR PERFIL',
        'CENEFA TRASERA': 'COLOR PERFIL',
        'PESO SOFT LIGHT': 'COLOR PESO INF. SOFT LIGHT',
        'PESO': 'COLOR ACCESORIOS'
    };

    // Mapear cada nombre de COLUMNAS_CORTE a su índice real en el Excel
    const idxCorte = {};
    for (const nombreCol of COLUMNAS_CORTE) {
        const idx = encabezados.findIndex(h => _normH(h) === nombreCol);
        if (idx !== -1) idxCorte[nombreCol] = idx;
    }
    const columnasCorteDetectadas = Object.keys(idxCorte);
    log(`🔧 Multi-corte: ${columnasCorteDetectadas.length} columnas detectadas: ${columnasCorteDetectadas.join(', ')}`, 'info');

    // Detectar columnas de color únicas referenciadas por MAPA_COLUMNAS_COLOR
    const nombresColorUnicos = [...new Set(Object.values(MAPA_COLUMNAS_COLOR))];
    const idxColumnaColor = {};
    for (const nombreColor of nombresColorUnicos) {
        let idx = encabezados.findIndex(h => _normH(h) === nombreColor);
        if (idx === -1) idx = encabezados.findIndex(h => _normH(h).includes(nombreColor));
        if (idx !== -1) idxColumnaColor[nombreColor] = idx;
    }
    log(`🎨 Columnas de color detectadas: ${Object.keys(idxColumnaColor).join(', ') || 'ninguna'}`, 'info');

    SistemaInventario.ordenes = [];
    let filasConMultiCorte = 0;

    for (let i = filaEncabezado + 1; i < SistemaInventario.datosCrudosOrdenes.length; i++) {
        const fila = SistemaInventario.datosCrudosOrdenes[i];
        if (!fila || fila.every(c => c === null || c === undefined || String(c).trim() === '')) continue;

        // Extraer datos compartidos de la fila (OT, UBIC, TUBERIA, COLOR, etc.)
        const datosCompartidos = {};
        for (const [nombre, config] of Object.entries(COLUMNAS_ESPECIALES)) {
            const idx = mapeoColumnas[nombre];
            if (idx !== undefined && idx !== null) {
                let valorCelda = fila[idx];
                if (config.esNumero) valorCelda = formatearNumero(valorCelda);
                datosCompartidos[nombre] = valorCelda;
            }
        }
        for (const colEsp of COLUMNAS_ESPECIFICAS) {
            const idx = mapeoColumnas[colEsp.key];
            if (idx !== undefined && idx !== null) datosCompartidos[colEsp.key] = fila[idx];
        }

        // Código del tubo (de TUBERIA)
        let codTubo = null;
        if (datosCompartidos.tuberia) {
            codTubo = extraerCodigoDesdeTuberia(datosCompartidos.tuberia);
        }

        let ordenesDeEstaFila = 0;

        // ─── Iterar sobre cada columna de corte ───
        for (const nombreCol of COLUMNAS_CORTE) {
            const idxCol = idxCorte[nombreCol];
            if (idxCol === undefined) continue; // columna no existe en este Excel

            const valorCrudo = fila[idxCol];
            // Protección: ignorar celdas vacías, texto puro o valores no numéricos
            if (valorCrudo === undefined || valorCrudo === null || valorCrudo === '') continue;
            const medidaNum = limpiarNumero(valorCrudo);
            if (isNaN(medidaNum) || medidaNum <= 0) continue;

            const orden = {
                id: SistemaInventario.ordenes.length + 1,
                medida_mm: Math.round(medidaNum * 10),
                medida_cm: formatearNumero(medidaNum),
                componente: nombreCol // Trazabilidad: qué componente es
            };

            // Copiar datos compartidos de la fila
            Object.assign(orden, datosCompartidos);

            if (nombreCol === 'TUBO') {
                // ─── Lógica original del TUBO ───
                if (codTubo) {
                    orden.codigoExtraido = codTubo;
                    orden.reemplazo = SistemaInventario.catalogoReemplazos[codTubo] || null;
                    orden.cod = codTubo || orden.codSec || 'TUBO-' + orden.id;
                } else {
                    orden.cod = orden.codSec || 'TUBO-' + orden.id;
                }

                // Override nuclear de medida para TUBO
                const _valTubo = idxTubo !== -1 ? limpiarNumero(fila[idxTubo]) : null;
                const _medidaFinal = (_valTubo !== null && _valTubo > 0)
                    ? _valTubo
                    : limpiarNumero(orden['medida']);
                orden['medida'] = _medidaFinal;
                orden.medida_cm = _medidaFinal;
                orden.medida_mm = Math.round(_medidaFinal * 10);
            } else {
                // ─── Lógica de accesorios (no-TUBO) ───
                // Buscar color usando el mapa explícito componente → columna de color
                const nombreColumnaColor = MAPA_COLUMNAS_COLOR[nombreCol];
                let colorDeseado = '';
                if (nombreColumnaColor && idxColumnaColor[nombreColumnaColor] !== undefined) {
                    const valColor = fila[idxColumnaColor[nombreColumnaColor]];
                    colorDeseado = valColor ? String(valColor).toUpperCase().trim() : '';
                }
                // Fallback general: si la columna específica no existe en este Excel, usar COLOR ACCESORIOS
                if (!colorDeseado && idxColumnaColor['COLOR ACCESORIOS'] !== undefined) {
                    const valFallback = fila[idxColumnaColor['COLOR ACCESORIOS']];
                    colorDeseado = valFallback ? String(valFallback).toUpperCase().trim() : '';
                }

                // Excepción dura: PESO INTERNO siempre es E13, sin importar color
                if (nombreCol === 'PESO INTERNO') {
                    orden.cod = 'E13';
                    orden.codigoExtraido = 'E13';
                    orden.reemplazo = SistemaInventario.catalogoReemplazos['E13'] || null;
                    orden.tuberia = 'PESO INTERNO';
                    SistemaInventario.ordenes.push(orden);
                    ordenesDeEstaFila++;
                    continue;
                }

                // Traducir nombre de columna al nombre exacto del catálogo
                const nombreCatalogo = (MAPA_NOMBRES_CATALOGO[nombreCol] || nombreCol).toUpperCase().trim();
                let llaveDiccionario = `${nombreCatalogo}|${colorDeseado}`;
                let codigoAccesorio = null;
                if (colorDeseado) {
                    codigoAccesorio = SistemaInventario.catalogoAccesorios[llaveDiccionario] || null;
                }
                // Fallback 1: buscar en color ALUMINIO por defecto (común en rieles/varillas)
                if (!codigoAccesorio) {
                    codigoAccesorio = SistemaInventario.catalogoAccesorios[`${nombreCatalogo}|ALUMINIO`] || null;
                }
                // Fallback 2: buscar sin color (por si en el catálogo está en blanco)
                if (!codigoAccesorio) {
                    codigoAccesorio = SistemaInventario.catalogoAccesorios[`${nombreCatalogo}|`] || null;
                }

                if (!codigoAccesorio) {
                    // Hard fail: no inventar códigos, saltar hasta que el usuario mapee en el catálogo
                    console.warn(`⚠️ Accesorio no encontrado en catálogo (ni fallbacks): ${llaveDiccionario}`);
                    continue;
                }
                orden.cod = codigoAccesorio;
                orden.codigoExtraido = codigoAccesorio;
                orden.reemplazo = SistemaInventario.catalogoReemplazos[codigoAccesorio] || null;

                // Sobrescribir propiedades heredadas para dar contexto al operario
                orden.tuberia = `${nombreCol} - ${colorDeseado || 'SIN COLOR'}`;
                // codSec se mantiene del datosCompartidos (sistema de cortina original)
            }

            SistemaInventario.ordenes.push(orden);
            ordenesDeEstaFila++;
        }

        if (ordenesDeEstaFila > 1) filasConMultiCorte++;
    }

    // Ordenar por OT y luego por Ubicación para agrupar como kit de armado
    SistemaInventario.ordenes.sort((a, b) => {
        const otA = String(a.ot || '');
        const otB = String(b.ot || '');
        if (otA !== otB) return otA.localeCompare(otB, undefined, { numeric: true });
        return String(a.ubic || '').localeCompare(String(b.ubic || ''), undefined, { numeric: true });
    });
    // Reasignar IDs secuenciales tras el ordenamiento
    SistemaInventario.ordenes.forEach((ord, idx) => { ord.id = idx + 1; });

    return filasConMultiCorte;
}

async function cargarOrdenes(event) {
    const file = event.target.files[0];
    if (!file) return;
    procesandoExcel = true;
    try {
        SistemaInventario.datosCrudosOrdenes = await leerExcelCompleto(file);

        const filasConMultiCorte = expandirOrdenesDesdeExcel();

        actualizarTablaOrdenes();
        document.getElementById('estadoOrdenes').textContent = `✓ ${SistemaInventario.ordenes.length} órdenes`;
        document.getElementById('estadoOrdenes').className = 'estado-archivo estado-ok';
        verificarListo();
        log(`📋 Órdenes cargadas: ${SistemaInventario.ordenes.length} (${filasConMultiCorte} filas expandidas a multi-corte)`, 'success');
        guardarSistema();
        guardarEnFirestore();
    } catch (e) { alert('Error: ' + e.message); console.error(e); }
    finally { procesandoExcel = false; }
}

async function cargarColmenas(event) {
    const file = event.target.files[0];
    if (!file) return;
    procesandoExcel = true;
    try {
        const datos = await leerExcelCompleto(file);
        const filaEncabezado = detectarFilaEncabezado(datos);
        const encabezados = datos[filaEncabezado];
        let columnaMedida = -1, columnaCod = -1, columnaNColmena = -1;
        for (let i = 0; i < encabezados.length; i++) {
            const enc = String(encabezados[i] || '').toLowerCase();
            if (enc.includes('medida') && columnaMedida === -1) columnaMedida = i;
            if (enc.includes('cod') && columnaCod === -1) columnaCod = i;
            if ((enc.includes('n_colmena') || enc.includes('n° colmena') || enc === 'n colmena' || enc.includes('n°') || enc.includes('numero')) && columnaNColmena === -1) columnaNColmena = i;
        }
        SistemaInventario.colmenas = [];
        for (let i = filaEncabezado + 1; i < datos.length; i++) {
            const fila = datos[i];
            if (!fila) continue;
            const medida = fila[columnaMedida];
            // Usar limpiarNumero para manejar tanto comas como puntos
            let medidaNum = limpiarNumero(medida);
            if (medidaNum !== null && medidaNum > 0) {
                // ── Auto-detección de unidades ──
                // Si la medida es ≤ 10, probablemente está en metros (ej: 5.78m = 578cm)
                // Los tubos reales nunca miden menos de 10 cm, así que es seguro convertir
                if (medidaNum <= 10) {
                    log(`⚠️ Medida ${medidaNum} detectada como metros → convertida a ${medidaNum * 100} cm`, 'warn');
                    medidaNum = medidaNum * 100;
                }
                const cod = fila[columnaCod];
                let nColmena;
                if (columnaNColmena !== -1 && fila[columnaNColmena] !== null && fila[columnaNColmena] !== undefined) {
                    nColmena = fila[columnaNColmena];
                } else {
                    nColmena = SistemaInventario.colmenas.length + 1;
                }
                SistemaInventario.colmenas.push({ n_colmena: nColmena, medida_mm: Math.round(medidaNum * 10), medida_cm: formatearNumero(medidaNum), cod: cod || `TUBO-${SistemaInventario.colmenas.length + 1}` });
            }
        }
        SistemaInventario.colmenas.sort((a, b) => {
            const cmpColmena = ordenarColmena(a.n_colmena, b.n_colmena);
            if (cmpColmena !== 0) return cmpColmena;
            return String(a.cod || '').localeCompare(String(b.cod || ''));
        });
        actualizarTablaColmenas();
        document.getElementById('estadoColmenas').textContent = `✓ ${SistemaInventario.colmenas.length} colmenas`;
        document.getElementById('estadoColmenas').className = 'estado-archivo estado-ok';
        verificarListo();
        log(`📦 Colmenas cargadas: ${SistemaInventario.colmenas.length}`, 'success');
        // Marcar como manual y actualizar indicador
        usandoColmenaManual = true;
        actualizarIndicadorFuente(true);
        // Guardar copia inmutable de colmenas para poder recalcular sin recargar Excel
        SistemaInventario.colmenaCruda = JSON.parse(JSON.stringify(SistemaInventario.colmenas));
        guardarSistema();
    } catch (e) { alert('Error: ' + e.message); }
    finally { procesandoExcel = false; }
}

async function cargarCatalogo(event) {
    const file = event.target.files[0];
    if (!file) return;
    procesandoExcel = true;
    try {
        const datos = await leerExcelCompleto(file);
        const filaEncabezado = detectarFilaEncabezado(datos);
        const encabezados = datos[filaEncabezado];
        let colCodigo = -1, colReemplazo = -1, colColor = -1, colMedidaReal = -1, colDescripcion = -1;
        for (let i = 0; i < encabezados.length; i++) {
            const enc = String(encabezados[i] || '').trim().toUpperCase();
            if (enc.includes('CODIGO') || enc === 'COD') colCodigo = i;
            if (enc.includes('REEMPLAZ')) colReemplazo = i;
            if (enc === 'COLOR') colColor = i;  // Match exacto para evitar falsos positivos
            if (enc.includes('MEDIDA REAL') || enc.includes('MEDIDA_REAL')) colMedidaReal = i;
            if (enc.includes('DESCRIPCI')) colDescripcion = i;
        }
        if (colCodigo === -1 || colReemplazo === -1) { alert('No se encontraron columnas CODIGO y REEMPLAZO'); return; }
        SistemaInventario.catalogoReemplazos = {};
        SistemaInventario.catalogoColores = {};
        SistemaInventario.catalogoMedidas = {};
        SistemaInventario.catalogoAccesorios = {};
        for (let i = filaEncabezado + 1; i < datos.length; i++) {
            const fila = datos[i];
            if (!fila) continue;
            const codigo = fila[colCodigo];
            const reemplazo = fila[colReemplazo];
            if (codigo && reemplazo) {
                const codigoLimpio = String(codigo).trim().toUpperCase();
                SistemaInventario.catalogoReemplazos[codigoLimpio] = String(reemplazo).trim();
                // Guardar color si la columna existe y tiene valor
                if (colColor !== -1 && fila[colColor] !== null && fila[colColor] !== undefined) {
                    const colorStr = String(fila[colColor]).trim();
                    if (colorStr) SistemaInventario.catalogoColores[codigoLimpio] = colorStr;
                }
                // Guardar medida real del tubo si la columna existe y tiene valor
                if (colMedidaReal !== -1 && fila[colMedidaReal] !== null && fila[colMedidaReal] !== undefined) {
                    const medida = limpiarNumero(fila[colMedidaReal]);
                    if (medida > 0) SistemaInventario.catalogoMedidas[codigoLimpio] = medida;
                }
                // Super-Catálogo de Accesorios: llave compuesta DESCRIPCIÓN|COLOR → código
                if (colDescripcion !== -1 && colColor !== -1) {
                    const desc = fila[colDescripcion] ? String(fila[colDescripcion]).toUpperCase().trim() : '';
                    const colorAcc = fila[colColor] ? String(fila[colColor]).toUpperCase().trim() : '';
                    if (desc && colorAcc) {
                        const llave = `${desc}|${colorAcc}`;
                        SistemaInventario.catalogoAccesorios[llave] = codigoLimpio;
                    }
                }
            }
        }
        actualizarReemplazosEnOrdenes();
        actualizarTablaCatalogo();
        const nColores = Object.keys(SistemaInventario.catalogoColores).length;
        const nMedidas = Object.keys(SistemaInventario.catalogoMedidas).length;
        const nAccesorios = Object.keys(SistemaInventario.catalogoAccesorios).length;
        log(`✓ Catálogo cargado: ${Object.keys(SistemaInventario.catalogoReemplazos).length} reemplazos, ${nColores} colores, ${nMedidas} medidas reales, ${nAccesorios} accesorios`, 'success');
        document.getElementById('estadoCatalogo').textContent = `✓ ${Object.keys(SistemaInventario.catalogoReemplazos).length} reemplazos`;
        document.getElementById('estadoCatalogo').className = 'estado-archivo estado-ok';
        guardarSistema();
        actualizarTablaOrdenes(); // Refrescar tabla para mostrar colores en órdenes ya cargadas
    } catch (e) { alert('Error: ' + e.message); }
    finally { procesandoExcel = false; }
}

// Devuelve el color de un código desde el catálogo, normalizando el código antes de buscar
function obtenerColorDeCatalogo(codigo) {
    if (!codigo) return '';
    const cod = String(codigo).trim().toUpperCase();
    return SistemaInventario.catalogoColores[cod] || '';
}

function formatearValor(valor) {
    if (valor === null || valor === undefined) return '';
    if (typeof valor === 'number') { if (valor === 0) return ''; if (Number.isInteger(valor)) return valor; return Math.round(valor * 100) / 100; }
    return valor;
}

function actualizarTablaOrdenes() {
    console.log("🚨 REVISIÓN DE RENDERIZADO - Primera orden:", SistemaInventario.ordenes[0]);

    let columnasVisibles = detectarColumnasConDatos(SistemaInventario.ordenes);

    // Columna Color: posición fija después de Tubería (se reconstruye desde cero en cada llamada)
    const posInsertar = columnasVisibles.findIndex(c => c.key === 'tuberia' || c.key === 'codigoExtraido');
    if (posInsertar !== -1) columnasVisibles.splice(posInsertar + 1, 0, { key: '_colorCatalogo', titulo: 'Color' });
    else columnasVisibles.push({ key: '_colorCatalogo', titulo: 'Color' });

    const headerRow = document.getElementById('headerOrdenes');
    headerRow.innerHTML = columnasVisibles.map(col => `<th>${col.titulo}</th>`).join('');
    const tbody = document.getElementById('tbodyOrdenes');
    tbody.innerHTML = SistemaInventario.ordenes.map(orden => {
        const celdas = columnasVisibles.map(col => {
            if (col.key === '_colorCatalogo') {
                return `<td>${obtenerColorDeCatalogo(orden.cod || orden.codigoExtraido || '')}</td>`;
            }
            // Columna 'medida': forzar siempre orden.medida_cm (fuente de verdad post-override)
            if (col.key === 'medida') {
                return `<td>${formatearValor(orden.medida_cm)}</td>`;
            }
            return `<td>${formatearValor(orden[col.key])}</td>`;
        }).join('');
        return `<tr>${celdas}</tr>`;
    }).join('');
}

function actualizarTablaColmenas() {
    const colmenasOrdenadas = [...SistemaInventario.colmenas].sort((a, b) => {
        const cmpColmena = ordenarColmena(a.n_colmena, b.n_colmena);
        if (cmpColmena !== 0) return cmpColmena;
        return String(a.cod || '').localeCompare(String(b.cod || ''));
    });
    document.getElementById('tbodyColmenas').innerHTML = colmenasOrdenadas.map(c => `<tr><td>${c.n_colmena}</td><td>${c.cod}</td><td>${formatearValor(c.medida_cm)}</td></tr>`).join('');
}

function actualizarTablaCatalogo() {
    const items = Object.entries(SistemaInventario.catalogoReemplazos).map(([codigo, reemplazo]) => `<tr><td>${codigo}</td><td>${reemplazo}</td></tr>`).join('');
    document.getElementById('tbodyReemplazos').innerHTML = items || '<tr><td>Sin datos</td></tr>';
}

function verificarListo() {
    const tieneOrdenes = SistemaInventario.ordenes.length > 0;
    const tieneColmenas = SistemaInventario.colmenas.length > 0 || (colmenaActual && colmenaActual.length > 0);
    const tieneCatalogo = Object.keys(SistemaInventario.catalogoReemplazos).length > 0;
    const tieneSeriales = SistemaInventario.seriales.length > 0;

    const btnEjecutar = document.getElementById('btnEjecutar');
    btnEjecutar.disabled = !(tieneOrdenes && tieneColmenas);

    // ── Barra de progreso ──
    const archivos = [
        { ok: tieneOrdenes, nombre: 'Órdenes' },
        { ok: tieneColmenas, nombre: 'Colmenas' },
        { ok: tieneCatalogo, nombre: 'Catálogo' },
        { ok: tieneSeriales, nombre: 'Inventario' }
    ];
    const cargados = archivos.filter(a => a.ok).length;
    const faltantes = archivos.filter(a => !a.ok).map(a => a.nombre);
    const pct = Math.round((cargados / 4) * 100);

    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');
    const progressFaltantes = document.getElementById('progressFaltantes');

    if (progressFill) progressFill.style.width = pct + '%';
    if (progressText) progressText.textContent = `${cargados} de 4 archivos cargados`;
    if (progressFaltantes) {
        progressFaltantes.textContent = faltantes.length > 0
            ? `Faltan: ${faltantes.join(', ')}`
            : 'Todo listo';
        progressFaltantes.style.color = faltantes.length === 0 ? '#27ae60' : '#aaa';
    }

    // ── Tooltip en botón deshabilitado ──
    if (btnEjecutar.disabled && faltantes.length > 0) {
        btnEjecutar.title = `Falta cargar: ${faltantes.join(', ')}`;
    } else {
        btnEjecutar.title = '';
    }
}

function log(mensaje, tipo) {
    SistemaInventario.logs.push({ mensaje, tipo });
    const logDiv = document.getElementById('logs');
    const entry = document.createElement('div');
    entry.className = 'log-entry';
    entry.textContent = mensaje;
    logDiv.appendChild(entry);
    logDiv.scrollTop = logDiv.scrollHeight;
}

function agregarPaso(numero, titulo, descripcion) {
    const procesoDiv = document.getElementById('proceso');
    const paso = document.createElement('div');
    paso.style.cssText = 'background:#0a1628; padding:8px; margin:3px 0; border-radius:3px; border-left:3px solid #3498db; font-size:12px; color:#ddd;';
    paso.innerHTML = `<strong style="color:#00d2ff;">${numero}. ${titulo}</strong> <span style="color:#aaa;">- ${descripcion}</span>`;
    procesoDiv.appendChild(paso);
}

// ── Umbrales configurables de sobrante (en mm) ──
// MERMA_MAX:         ≤ este valor → se desecha como merma (no vale guardarlo)
// ZONA_PROHIBIDA_MIN/MAX: sobrantes en esta zona son demasiado cortos para ser útiles
//                    pero demasiado largos para desechar — el optimizador evita estos cortes.
const MERMA_MAX_MM = 100;              // ≤ 10 cm → merma (desechar)
const ZONA_PROHIBIDA_MIN_MM = 101;     // 10.1 cm
const ZONA_PROHIBIDA_MAX_MM = 1299;    // 129.9 cm

function evaluarSobrante(sobrante) {
    if (sobrante < 0) return { estado: 'prohibido' };
    if (sobrante <= MERMA_MAX_MM) return { estado: 'merma' };
    if (sobrante >= ZONA_PROHIBIDA_MIN_MM && sobrante <= ZONA_PROHIBIDA_MAX_MM) return { estado: 'prohibido' };
    return { estado: 'colmena' };
}

function normalizarCodigo(codigo) { return String(codigo).trim().toUpperCase(); }

function buscarReemplazos(codigo) {
    const codNormalizado = normalizarCodigo(codigo);
    for (const [key, value] of Object.entries(SistemaInventario.catalogoReemplazos)) {
        if (normalizarCodigo(key) === codNormalizado) return String(value).split(';').map(r => r.trim()).filter(r => r);
    }
    return null;
}

function buscarTubosParaOrden(codigoBuscado, medidaRequerida, codigoOriginal = null) {
    const codNormalizado = normalizarCodigo(codigoBuscado);

    // ── BEST FIT: buscar el tubo que genere MENOS desperdicio ──
    // Prioridad: 1) sobrantes existentes (menor desperdicio), 2) tubos nuevos (mayor desperdicio)
    // Esto evita gastar tubos nuevos cuando hay sobrantes aprovechables.

    let mejorResultado = null;
    let mejorSobrante = Infinity;

    // Función interna para evaluar un candidato
    function evaluarCandidato(col, indice, esReemplazo) {
        if (!col || !col.medida_mm || col.medida_mm <= 0) return;
        if (col.medida_mm < medidaRequerida) return;

        const sobrante = col.medida_mm - medidaRequerida - MM_KERF;
        const clasificacion = evaluarSobrante(sobrante);
        if (clasificacion.estado === 'prohibido') return;

        // Priorizar: menor sobrante = mejor ajuste (menos desperdicio)
        if (sobrante < mejorSobrante) {
            mejorSobrante = sobrante;
            mejorResultado = {
                colmena: col,
                sobrante_mm: sobrante,
                indice: indice,
                clasificacion: clasificacion,
                medidaOriginal: col.medida_mm,
                esReemplazo: esReemplazo
            };
        }
    }

    // Paso 1: Buscar con código original (si se proporcionó)
    if (codigoOriginal) {
        const codOrigNormalizado = normalizarCodigo(codigoOriginal);
        for (let i = 0; i < SistemaInventario.colmenasDisponibles.length; i++) {
            const col = SistemaInventario.colmenasDisponibles[i];
            if (normalizarCodigo(col.cod) === codOrigNormalizado) {
                evaluarCandidato(col, i, false);
            }
        }
        if (mejorResultado) return mejorResultado;
    }

    // Paso 2: Buscar con código normalizado (incluye reemplazos)
    for (let i = 0; i < SistemaInventario.colmenasDisponibles.length; i++) {
        const col = SistemaInventario.colmenasDisponibles[i];
        if (normalizarCodigo(col.cod) === codNormalizado) {
            const esReemp = codigoOriginal ? codNormalizado !== normalizarCodigo(codigoOriginal) : false;
            evaluarCandidato(col, i, esReemp);
        }
    }

    return mejorResultado;
}

function buscarColmenaDisponibleConCodigo(codigo) {
    const codNormalizado = normalizarCodigo(codigo);

    for (let i = 0; i < SistemaInventario.colmenasHistorico.length; i++) {
        const c = SistemaInventario.colmenasHistorico[i];

        if (c.estado !== 'disponible') continue;
        if (normalizarCodigo(c.cod) !== codNormalizado) continue;
        if (normalizarCodigo(c.codigo_original) !== codNormalizado) continue;

        return c;
    }

    return null;
}

function formatearResultado(orden, resultado) {
    const fuente = resultado.fuente;
    let descripcion = '';
    switch(fuente) {
        case 'exacta': descripcion = `✓ Exacta: Colmena ${resultado.colmena}`; break;
        case 'colmena': descripcion = `📦 Colmena ${resultado.colmena} (sobrante: ${resultado.sobrante_cm}cm)`; break;
        case 'misma_codigo': descripcion = `📦 Colmena ${resultado.colmena} (sobrante: ${resultado.sobrante_cm}cm)`; break;
        case 'merma': descripcion = `✂️ Merma: Colmena ${resultado.colmena} (sobrante: ${resultado.sobrante_cm}cm)`; break;
        case 'reemplazo': descripcion = `🔄 REEMPLAZO: ${resultado.codigo_original} → ${resultado.codigo_reemplazo} (Colmena ${resultado.colmena}, sobrante: ${resultado.sobrante_cm}cm)`; break;
        case 'tubo_nuevo': descripcion = `🆕 Tubo nuevo (sobrante: ${resultado.sobrante_cm}cm)`; break;
        default: descripcion = 'Desconocido';
    }
    return descripcion;
}

function actualizarTablaColmenasResultado() {
    const tbody = document.getElementById('tbodyColmenasResultado');
    if (!tbody) return;
    tbody.innerHTML = SistemaInventario.resultadosOptimizacion.map(item => {
        const r = item.resultado;
        const ord = SistemaInventario.ordenes.find(o => o.id === r.orden) || {};
        const color = r.color || obtenerColorDeCatalogo(r.codigo || r.codigo_original || '');
        const accion = (ord.componente && ord.componente !== 'TUBO') ? `CORTAR ${ord.componente}` : 'CORTAR';
        let filaHtml = `<tr class="fila-cortar">
            <td>${r.colmena || r.nombreMaterialNuevo || 'TUBO NUEVO'}</td>
            <td>${r.codigo || '-'}</td>
            <td>${color}</td>
            <td>${Number(r.medida_cm).toFixed(1)} cm</td>
            <td>${r.medida_origen !== undefined ? Number(r.medida_origen).toFixed(1) + ' cm' : '-'}</td>
            <td><span class="tag-accion tag-cortar">${accion}</span></td>
        </tr>`;
        if (r.sobrante_cm > 0) {
            if (r.es_intermedio) {
                filaHtml += `<tr class="fila-mesa">
                    <td>-</td><td>${r.codigo || '-'}</td><td>${color}</td>
                    <td>${Number(r.sobrante_cm).toFixed(1)} cm</td><td>-</td>
                    <td><span class="tag-accion tag-mesa">RESERVAR EN MESA</span></td>
                </tr>`;
            } else if (r.es_desecho) {
                filaHtml += `<tr class="fila-merma">
                    <td>BASURERO</td><td>${r.codigo || '-'}</td><td>${color}</td>
                    <td>${Number(r.sobrante_cm).toFixed(1)} cm</td><td>-</td>
                    <td><span class="tag-accion tag-merma">DESECHAR MERMA</span></td>
                </tr>`;
            } else {
                filaHtml += `<tr class="fila-guardar">
                    <td>${r.colmena_sobrante || r.colmena || '-'}</td><td>${r.codigo || '-'}</td><td>${color}</td>
                    <td>${Number(r.sobrante_cm).toFixed(1)} cm</td><td>-</td>
                    <td><span class="tag-accion tag-guardar">GUARDAR SOBRANTE</span></td>
                </tr>`;
            }
        }
        return filaHtml;
    }).join('');
}

function verificarAlertasStock() {
    const conteo = {};
    // Contar solo los tubos enteros disponibles por código
    SistemaInventario.seriales.forEach(s => {
        if (s.estado === 'disponible') {
            conteo[s.codigo] = (conteo[s.codigo] || 0) + 1;
        }
    });

    // Filtrar los que están en peligro (≤ STOCK_MINIMO)
    const alertas = Object.keys(conteo)
        .filter(cod => conteo[cod] <= STOCK_MINIMO)
        .map(cod => ({ codigo: cod, cantidad: conteo[cod] }))
        .sort((a, b) => a.cantidad - b.cantidad); // Más críticos primero

    renderizarAlertasStock(alertas);
}

function renderizarAlertasStock(alertas) {
    const panel = document.getElementById('panelAlertasStock');
    if (!panel) return;

    if (alertas.length === 0) {
        panel.innerHTML = '';
        panel.style.display = 'none';
        return;
    }

    const items = alertas.map(a => {
        const esCritico = a.cantidad <= 3;
        const bgColor = esCritico ? '#e74c3c' : '#f39c12';
        const icon = esCritico ? '🔴' : '🟡';
        return `<span style="display:inline-block; background:${bgColor}; color:white; padding:4px 10px; border-radius:4px; margin:3px 2px; font-size:12px; font-weight:bold;">${icon} ${a.codigo}: ${a.cantidad} tubo${a.cantidad !== 1 ? 's' : ''}</span>`;
    }).join('');

    panel.style.display = 'block';
    panel.innerHTML = `
        <div style="background:#1a1a2e; border:1px solid #e67e22; border-left:4px solid #f39c12; border-radius:6px; padding:12px 16px; margin:10px 0;">
            <strong style="color:#f39c12; font-size:13px;">ALERTA DE STOCK MÍNIMO — Materiales con ≤${STOCK_MINIMO} tubos disponibles:</strong>
            <div style="margin-top:8px; line-height:2;">${items}</div>
        </div>
    `;
}

function actualizarTablaMermas() {
    const tbody = document.getElementById('tbodyMermas');
    if (!tbody) {
        console.warn("⚠️ No se encontró el elemento 'tbodyMermas' en el HTML.");
        return;
    }
    if (SistemaInventario.mermas.length === 0) {
        tbody.innerHTML = '<tr><td colspan="4">Sin mermas</td></tr>';
        return;
    }
    tbody.innerHTML = SistemaInventario.mermas.map(m => {
        return `<tr><td>${m.orden}</td><td>${m.espacioOriginal}</td><td>${m.codigoOriginal}</td><td>${formatearValor(m.valor)} cm</td></tr>`;
    }).join('');
}

// Función principal para ejecutar el algoritmo de optimización
function ejecutarAlgoritmoOptimización(ordenes, colmenas) {
    // Validar que los arrays no estén vacíos
    if (!Array.isArray(ordenes) || ordenes.length === 0) {
        console.error("No hay órdenes para optimizar");
        return { resultados: [], totalMerma: 0 };
    }
    
    if (!Array.isArray(colmenas) || colmenas.length === 0) {
        console.error("No hay colmenas para asignar");
        return { resultados: [], totalMerma: 0 };
    }

    // Ordenar las órdenes de mayor a menor (First-Fit Decreasing)
    const ordenesOrdenadas = [...ordenes].sort((a, b) => {
        const medidaA = typeof a.medida === 'number' ? a.medida : (parseFloat(a.medida) || 0);
        const medidaB = typeof b.medida === 'number' ? b.medida : (parseFloat(b.medida) || 0);
        return medidaB - medidaA;
    });

    // Crear una copia de las colmenas para manipulate durante la asignación
    const colmenasDisponibles = colmenas.map(c => ({
        ...c,
        estado: 'Disponible',
        medidaOriginal: typeof c['Medida (cm)'] === 'number' ? c['Medida (cm)'] : (parseFloat(c['Medida (cm)']) || 0)
    }));

    const resultados = [];
    let totalMerma = 0;

    // Recorrer cada orden y buscar la primera colmena disponible
    for (let i = 0; i < ordenesOrdenadas.length; i++) {
        const orden = ordenesOrdenadas[i];
        
        // Obtener la medida de la orden
        const medidaOrden = typeof orden.medida === 'number' ? orden.medida : (parseFloat(orden.medida) || 0);
        
        // Buscar la primera colmena donde quepa la orden
        let colmenaAsignada = null;
        let indiceColmena = -1;
        
        for (let j = 0; j < colmenasDisponibles.length; j++) {
            const colmena = colmenasDisponibles[j];
            
            if (colmena.estado === 'Disponible' && colmena.medidaOriginal >= medidaOrden) {
                colmenaAsignada = colmena;
                indiceColmena = j;
                break;
            }
        }

        if (colmenaAsignada) {
            // Marcar la colmena como ocupada
            colmenasDisponibles[indiceColmena].estado = 'Ocupada';
            
            // Calcular la merma
            const merma = colmenaAsignada.medidaOriginal - medidaOrden;
            totalMerma += merma;

            // Guardar la asignación en el resultado
            resultados.push({
                orden: orden,
                colmena: colmenaAsignada,
                medidaOrden: medidaOrden,
                medidaColmena: colmenaAsignada.medidaOriginal,
                merma: merma,
                asignada: true,
                mensaje: `Orden asignada a colmena ${colmenaAsignada['N° Colmena'] || colmenaAsignada.n_colmena}`
            });
        } else {
            // La orden no cabe en ninguna colmena
            resultados.push({
                orden: orden,
                colmena: null,
                medidaOrden: medidaOrden,
                medidaColmena: null,
                merma: 0,
                asignada: false,
                mensaje: 'No Asignada: No hay colmena disponible con medida suficiente'
            });
        }
    }

    console.log(`✅ Optimización completada: ${resultados.filter(r => r.asignada).length} órdenes asignadas, ${resultados.filter(r => !r.asignada).length} no asignadas`);
    console.log(`📊 Total merma: ${totalMerma.toFixed(2)} cm`);

    return {
        resultados: resultados,
        totalMerma: totalMerma
    };
}

// Función asíncrona para ejecutar el flujo completo desde el botón "Ejecutar"
async function ejecutarFlujoCompleto() {
    console.log("🚀 Iniciando flujo completo de optimización...");
    
    // Obtener datos de las variables globales
    const ordenes = SistemaInventario.ordenes;
    const colmenas = SistemaInventario.colmenas;

    // Validar que existan datos
    if (!ordenes || ordenes.length === 0) {
        const mensajeError = "❌ Error: No hay órdenes cargadas";
        console.error(mensajeError);
        log(mensajeError, 'error');
        return;
    }

    if (!colmenas || colmenas.length === 0) {
        const mensajeError = "❌ Error: No hay colmenas cargadas";
        console.error(mensajeError);
        log(mensajeError, 'error');
        return;
    }

    // Limpiar el panel de logs y resultados
    document.getElementById('logs').innerHTML = '';
    document.getElementById('proceso').innerHTML = '';
    document.getElementById('resultados').innerHTML = '';

    log("📋 Validando datos de órdenes...", 'info');
    
    // Validar y limpiar órdenes
    const ordenesValidadas = validarYLimpiarDatos(ordenes, 'orden');
    
    if (ordenesValidadas.length === 0) {
        const mensajeError = "❌ Error: No hay órdenes válidas después de la validación";
        console.error(mensajeError);
        log(mensajeError, 'error');
        return;
    }

    log(`✓ Órdenes validadas: ${ordenesValidadas.length} registros válidos`, 'success');

    log("📦 Validando datos de colmenas...", 'info');
    
    // Validar y limpiar colmenas
    const colmenasValidadas = validarYLimpiarDatos(colmenas, 'colmena');
    
    if (colmenasValidadas.length === 0) {
        const mensajeError = "❌ Error: No hay colmenas válidas después de la validación";
        console.error(mensajeError);
        log(mensajeError, 'error');
        return;
    }

    log(`✓ Colmenas validadas: ${colmenasValidadas.length} registros válidos`, 'success');

    // Ejecutar el algoritmo de optimización
    log("⚙️ Ejecutando algoritmo de optimización...", 'info');
    
    const resultadoOptimizacion = ejecutarAlgoritmoOptimización(ordenesValidadas, colmenasValidadas);
    
    if (!resultadoOptimizacion || resultadoOptimizacion.resultados.length === 0) {
        const mensajeError = "❌ Error: No se pudieron generar resultados de optimización";
        console.error(mensajeError);
        log(mensajeError, 'error');
        return;
    }

    const asignadas = resultadoOptimizacion.resultados.filter(r => r.asignada).length;
    const noAsignadas = resultadoOptimizacion.resultados.filter(r => !r.asignada).length;
    
    log(`✅ Optimización completada: ${asignadas} órdenes asignadas, ${noAsignadas} no asignadas`, 'success');
    log(`📊 Total merma: ${resultadoOptimizacion.totalMerma.toFixed(2)} cm`, 'info');

    // Guardar resultados en Firebase
    log("💾 Guardando resultados en Firebase...", 'info');
    
    try {
        await guardarResultadoOptimizacion(resultadoOptimizacion.resultados);
        log("✅ Resultados guardados en historial de Firebase", 'success');
    } catch (error) {
        console.error("Error al guardar en Firebase:", error);
        log("⚠️ Warning: No se pudieron guardar los resultados en Firebase", 'warn');
    }

    // Actualizar el DOM con los resultados
    const divResultados = document.getElementById('resultados');
    divResultados.innerHTML = `
        <div class="resumen-resultado">
            <h3>✓ Optimización Completada</h3>
            <p><strong>Total de órdenes:</strong> ${resultadoOptimizacion.resultados.length}</p>
            <p><strong>Órdenes asignadas:</strong> ${asignadas}</p>
            <p><strong>Órdenes no asignadas:</strong> ${noAsignadas}</p>
            <p><strong>Total Merma:</strong> ${resultadoOptimizacion.totalMerma.toFixed(2)} cm</p>
        </div>
    `;

    // Llenar la tabla de mermas con las asignaciones
    const tbodyMermas = document.getElementById('tbodyMermas');
    if (tbodyMermas) {
        tbodyMermas.innerHTML = resultadoOptimizacion.resultados
            .filter(r => r.asignada)
            .map(r => {
                const pedido = r.orden.Pedido || r.orden.id || 'N/A';
                const nColmena = r.colmena['N° Colmena'] || r.colmena.n_colmena || 'N/A';
                return `<tr>
                    <td>${pedido}</td>
                    <td>${nColmena}</td>
                    <td>${r.medidaOrden}</td>
                    <td>${r.medidaColmena}</td>
                    <td>${r.merma.toFixed(2)}</td>
                </tr>`;
            }).join('');
    }

    // Habilitar botón de exportar si existe
    const btnExportar = document.getElementById('btnExportar');
    if (btnExportar) {
        btnExportar.disabled = false;
    }

    const btnExportarDisponibles = document.getElementById('btnExportarDisponibles');
    if (btnExportarDisponibles) {
        btnExportarDisponibles.disabled = false;
    }

    log("🎉 Flujo completo ejecutado exitosamente", 'success');
    console.log("🚀 Flujo completo terminado", resultadoOptimizacion);
}

function ejecutarOptimizacion() {
    // ── RESETEO MANDATORIO: limpiar TODO el estado residual de optimizaciones previas ──
    SistemaInventario.colmenasDisponibles = [];
    SistemaInventario.colmenasHistorico = [];
    SistemaInventario.resultadosOptimizacion = [];
    SistemaInventario.mermas = [];
    SistemaInventario.logs = [];

    // ── Determinar fuente de colmenas: deep-copy fresca desde la fuente autoritativa ──
    // CRÍTICO: siempre clonar para romper cualquier referencia a estados anteriores
    let colmenasAUsar;
    if (usandoColmenaManual) {
        colmenasAUsar = JSON.parse(JSON.stringify(SistemaInventario.colmenas));
    } else if (colmenaActual && colmenaActual.length > 0) {
        colmenasAUsar = JSON.parse(JSON.stringify(colmenaActual));
    } else {
        colmenasAUsar = JSON.parse(JSON.stringify(SistemaInventario.colmenas));
    }

    if (SistemaInventario.ordenes.length === 0 || colmenasAUsar.length === 0) { alert('Cargue órdenes y colmenas'); return; }

    // ── ALERTA DE OT DUPLICADA: detectar si estas órdenes ya fueron procesadas ──
    const otsEnOrden = [...new Set(SistemaInventario.ordenes.map(o => String(o.ot || '')).filter(ot => ot && ot !== '-'))];
    if (otsEnOrden.length > 0) {
        // Revisar en el historial local si alguna OT ya fue confirmada hoy
        const hoy = new Date().toISOString().split('T')[0];
        const otsProcesadas = JSON.parse(localStorage.getItem('ots_procesadas_hoy') || '{}');
        const duplicadas = otsEnOrden.filter(ot => otsProcesadas[ot] === hoy);
        if (duplicadas.length > 0) {
            const continuar = confirm(
                `⚠️ ALERTA: Las siguientes OT ya fueron procesadas hoy:\n\n` +
                duplicadas.map(ot => `• OT ${ot}`).join('\n') +
                `\n\n¿Estás seguro de que quieres procesarlas de nuevo?\n` +
                `(Si confirmas, el inventario se descontará doble)`
            );
            if (!continuar) return;
        }
    }

    // Guardar copias inmutables para poder recalcular sin recargar Excel
    if (SistemaInventario.ordenesCrudas.length === 0) {
        SistemaInventario.ordenesCrudas = JSON.parse(JSON.stringify(SistemaInventario.ordenes));
    }
    if (SistemaInventario.serialesCrudos.length === 0 && SistemaInventario.seriales.length > 0) {
        SistemaInventario.serialesCrudos = JSON.parse(JSON.stringify(SistemaInventario.seriales));
    }

    log(`ℹ️ Fuente de colmenas: ${usandoColmenaManual ? 'Archivo manual' : 'Firebase (sincronizado)'}`, 'info');
    log(`ℹ️ Colmenas cargadas para esta optimización: ${colmenasAUsar.length}`, 'info');

    // ── BEST FIT DECREASING: ordenar órdenes de mayor a menor medida ──
    // Esto reduce el desperdicio total al asignar primero los cortes más grandes
    // a los tubos más ajustados, dejando sobrantes útiles para cortes pequeños.
    // Se preserva el orden original por OT/Ubicación como desempate.
    const ordenOriginal = SistemaInventario.ordenes.map((o, i) => ({ ...o, _ordenOriginal: i }));
    SistemaInventario.ordenes = ordenOriginal.sort((a, b) => {
        // Primero por medida descendente (cortes grandes primero)
        if (b.medida_mm !== a.medida_mm) return b.medida_mm - a.medida_mm;
        // Desempate: mantener orden original (OT + Ubicación)
        return a._ordenOriginal - b._ordenOriginal;
    });
    log(`📐 Órdenes reordenadas: mayor a menor (Best Fit Decreasing)`, 'info');

    document.getElementById('logs').innerHTML = '';
    document.getElementById('proceso').innerHTML = '';
    document.getElementById('resultados').innerHTML = '';
    SistemaInventario.colmenasDisponibles = JSON.parse(JSON.stringify(colmenasAUsar));

    // ── Ordenar colmenas: sobrantes primero, tubos nuevos después ──
    // Junto con Best Fit, esto asegura que los sobrantes existentes se
    // evalúen primero y se prefieran sobre tubos nuevos cuando ajustan.
    SistemaInventario.colmenasDisponibles.sort((a, b) => a.medida_mm - b.medida_mm);
    SistemaInventario.colmenasHistorico = [];
    SistemaInventario.resultadosOptimizacion = [];
    SistemaInventario.mermas = [];

    colmenasAUsar.forEach((col) => {
        // Incluir serial para que los sobrantes hereden la trazabilidad del tubo original
        SistemaInventario.colmenasHistorico.push({ n_colmena: col.n_colmena, medida_cm: col.medida_cm, medida_mm: col.medida_mm, cod: col.cod, codigo_original: col.cod, estado: 'disponible', origen: 'Original', posicionOriginal: col.n_colmena, serial: col.serial || null });
    });

    log('=== INICIO OPTIMIZACIÓN ===', 'info');
    const resultados = [];

    SistemaInventario.ordenes.forEach((orden, idx) => {
        const numPaso = idx + 1;
        const codOrden = orden.cod;
        let resultado = null;
        
        const tuboEncontrado = buscarTubosParaOrden(codOrden, orden.medida_mm, codOrden);
        
        if (tuboEncontrado && tuboEncontrado.sobrante_mm === 0) {
            resultado = {
                orden: orden.id,
                medida_cm: orden.medida_cm,
                fuente: 'exacta',
                colmena: tuboEncontrado.colmena.n_colmena,
                codigo: tuboEncontrado.colmena.cod,
                sobrante_cm: 0,
                medida_origen: tuboEncontrado.colmena.medida_cm,
                serial: tuboEncontrado.colmena.serial || orden.serial || null
            };
            const idxHistorico = SistemaInventario.colmenasHistorico.findIndex(c => c.n_colmena === tuboEncontrado.colmena.n_colmena && c.cod === tuboEncontrado.colmena.cod && c.medida_mm === tuboEncontrado.colmena.medida_mm);
            if (idxHistorico !== -1) { SistemaInventario.colmenasHistorico[idxHistorico].estado = 'usada'; SistemaInventario.colmenasHistorico[idxHistorico].origen = 'Orden ' + orden.id; }
            SistemaInventario.colmenasDisponibles.splice(tuboEncontrado.indice, 1);
        } else if (tuboEncontrado) {
            const esMerma = tuboEncontrado.clasificacion && tuboEncontrado.clasificacion.estado === 'merma';
            const esReemplazo = tuboEncontrado.esReemplazo;
            let fuente = esMerma ? 'merma' : (esReemplazo ? 'reemplazo' : 'colmena');

            resultado = { orden: orden.id, medida_cm: orden.medida_cm, fuente: fuente, colmena: tuboEncontrado.colmena.n_colmena, codigo: tuboEncontrado.colmena.cod, codigo_original: codOrden, sobrante_cm: tuboEncontrado.sobrante_mm / 10, medida_origen: tuboEncontrado.medidaOriginal / 10, serial: tuboEncontrado.colmena.serial || null, es_desecho: esMerma };
            const idxHistorico = SistemaInventario.colmenasHistorico.findIndex(c => c.n_colmena === tuboEncontrado.colmena.n_colmena && c.cod === tuboEncontrado.colmena.cod && c.medida_mm === tuboEncontrado.colmena.medida_mm);
            const sobrante = tuboEncontrado.medidaOriginal - orden.medida_mm - MM_KERF;
            const clasificacion = evaluarSobrante(sobrante);

            if (clasificacion.estado === 'merma') {
                SistemaInventario.mermas.push({
                    orden: orden.id,
                    espacioOriginal: tuboEncontrado.colmena.n_colmena,
                    codigoOriginal: codOrden,
                    medidaRequerida: orden.medida_cm,
                    tipo: 'MERMA',
                    valor: sobrante / 10,
                    codigoUsado: tuboEncontrado.colmena.cod
                });
                // CORRECCIÓN: eliminar el tubo de colmenasDisponibles (el slot queda vacío)
                SistemaInventario.colmenasDisponibles.splice(tuboEncontrado.indice, 1);
                if (idxHistorico !== -1) {
                    // Colmena vacía tras merma: 'usada' para que NO vuelva al inventario
                    SistemaInventario.colmenasHistorico[idxHistorico].estado = 'usada';
                    SistemaInventario.colmenasHistorico[idxHistorico].medida_mm = 0;
                    SistemaInventario.colmenasHistorico[idxHistorico].medida_cm = 0;
                    SistemaInventario.colmenasHistorico[idxHistorico].cod = '';
                    SistemaInventario.colmenasHistorico[idxHistorico].serial = null;
                    SistemaInventario.colmenasHistorico[idxHistorico].origen = 'MERMA desechada (tubo consumido)';
                }
            } else if (clasificacion.estado !== 'prohibido') {
                if (idxHistorico !== -1) { SistemaInventario.colmenasHistorico[idxHistorico].estado = 'usada'; SistemaInventario.colmenasHistorico[idxHistorico].origen = 'Orden ' + orden.id; }

                // CORRECCIÓN TRAZABILIDAD: buscar colmena destino correcta para el sobrante
                const colmenaDestinoSobrante = buscarColmenaDisponibleConCodigo(tuboEncontrado.colmena.cod);
                const nColmenaDestino = colmenaDestinoSobrante ? colmenaDestinoSobrante.n_colmena : tuboEncontrado.colmena.n_colmena;

                // Guardar en resultado la colmena destino real del sobrante
                resultado.colmena_sobrante = nColmenaDestino;

                if (colmenaDestinoSobrante) {
                    log(`📦 Sobrante reubicado en colmena ${nColmenaDestino} (código ${tuboEncontrado.colmena.cod}). Medida sobrante: ${sobrante / 10}cm`, 'info');
                }

                const idxInsertar = SistemaInventario.colmenasHistorico.findIndex(c => c.n_colmena === nColmenaDestino && c.cod === tuboEncontrado.colmena.cod && c.medida_mm === tuboEncontrado.colmena.medida_mm);
                if (idxInsertar !== -1) {
                SistemaInventario.colmenasHistorico.splice(idxInsertar + 1, 0, {
                    n_colmena: nColmenaDestino,
                    medida_cm: sobrante / 10,
                    medida_mm: sobrante,
                    cod: tuboEncontrado.colmena.cod,
                    codigo_original: tuboEncontrado.colmena.cod,
                    estado: 'disponible',
                    origen: 'Sobrante orden ' + orden.id,
                    posicionOriginal: nColmenaDestino,
                    serial: tuboEncontrado.colmena.serial || orden.serial || null,
                    fecha: tuboEncontrado.colmena.serial ? tuboEncontrado.colmena.serial.fecha : null
                });
                }
                SistemaInventario.colmenasDisponibles.splice(tuboEncontrado.indice, 1);
                SistemaInventario.colmenasDisponibles.push({
                    n_colmena: nColmenaDestino,
                    medida_mm: sobrante,
                    medida_cm: sobrante / 10,
                    cod: tuboEncontrado.colmena.cod,
                    serial: tuboEncontrado.colmena.serial || orden.serial || null
                });
            }
        } else {
            const listaReemplazos = buscarReemplazos(codOrden);
            if (listaReemplazos) {
                for (const codReemplazo of listaReemplazos) {
                    const tuboReemplazo = buscarTubosParaOrden(codReemplazo, orden.medida_mm, codOrden);
                    if (tuboReemplazo) {
                        const esMerma = tuboReemplazo.clasificacion && tuboReemplazo.clasificacion.estado === 'merma';
                        let fuente = esMerma ? 'merma' : 'reemplazo';

                        resultado = {
                            orden: orden.id,
                            medida_cm: orden.medida_cm,
                            fuente: fuente,
                            colmena: tuboReemplazo.colmena.n_colmena,
                            codigo: tuboReemplazo.colmena.cod,
                            codigo_original: codOrden,
                            codigo_reemplazo: tuboReemplazo.colmena.cod,
                            sobrante_cm: tuboReemplazo.sobrante_mm / 10,
                            medida_origen: tuboReemplazo.medidaOriginal / 10,
                            serial: tuboReemplazo.colmena.serial || null,
                            es_desecho: esMerma
                        };
                        const idxHistorico = SistemaInventario.colmenasHistorico.findIndex(c => c.n_colmena === tuboReemplazo.colmena.n_colmena && c.cod === tuboReemplazo.colmena.cod && c.medida_mm === tuboReemplazo.colmena.medida_mm);
                        const sobrante = tuboReemplazo.medidaOriginal - orden.medida_mm - MM_KERF;
                        const clasificacion = evaluarSobrante(sobrante);

                        if (clasificacion.estado === 'merma') {
                            SistemaInventario.mermas.push({
                                orden: orden.id,
                                espacioOriginal: tuboReemplazo.colmena.n_colmena,
                                codigoOriginal: codOrden,
                                medidaRequerida: orden.medida_cm,
                                tipo: 'MERMA',
                                valor: sobrante / 10,
                                codigoUsado: codReemplazo
                            });
                            // CORRECCIÓN: eliminar el tubo de colmenasDisponibles (el slot queda vacío)
                            SistemaInventario.colmenasDisponibles.splice(tuboReemplazo.indice, 1);
                            if (idxHistorico !== -1) {
                                // Colmena vacía tras merma: 'usada' para que NO vuelva al inventario
                                SistemaInventario.colmenasHistorico[idxHistorico].estado = 'usada';
                                SistemaInventario.colmenasHistorico[idxHistorico].medida_mm = 0;
                                SistemaInventario.colmenasHistorico[idxHistorico].medida_cm = 0;
                                SistemaInventario.colmenasHistorico[idxHistorico].cod = '';
                                SistemaInventario.colmenasHistorico[idxHistorico].serial = null;
                                SistemaInventario.colmenasHistorico[idxHistorico].origen = 'MERMA desechada (tubo consumido)';
                            }
                        } else if (clasificacion.estado !== 'prohibido') {
                            SistemaInventario.colmenasHistorico[idxHistorico].codigo_original = tuboReemplazo.colmena.cod;
                            if (idxHistorico !== -1) { SistemaInventario.colmenasHistorico[idxHistorico].estado = 'usada'; SistemaInventario.colmenasHistorico[idxHistorico].origen = 'Orden ' + orden.id + ' (Reemplazo ' + codReemplazo + ')';  }
                            
                            const colmenaExistente = buscarColmenaDisponibleConCodigo(codReemplazo);
                            // CORRECCIÓN TRAZABILIDAD: guardar colmena destino real del sobrante
                            resultado.colmena_sobrante = colmenaExistente ? colmenaExistente.n_colmena : tuboReemplazo.colmena.n_colmena;

                            if (colmenaExistente) {
                                log(`📦 Sobrante agregado como nueva fila para colmena ${colmenaExistente.n_colmena} (código ${codReemplazo}). Medida sobrante: ${sobrante / 10}cm`, 'info');
                                const idxHistoricoExistente = SistemaInventario.colmenasHistorico.findIndex(c => c.n_colmena === colmenaExistente.n_colmena);
                                if (idxHistoricoExistente !== -1) {
                                    SistemaInventario.colmenasHistorico.splice(idxHistoricoExistente + 1, 0, { n_colmena: colmenaExistente.n_colmena, medida_cm: sobrante / 10, medida_mm: sobrante, cod: codReemplazo, codigo_original: tuboReemplazo.colmena.cod, estado: 'disponible', origen: 'Sobrante reemplazo orden ' + orden.id, posicionOriginal: colmenaExistente.n_colmena, serial: tuboReemplazo.colmena.serial || null, fecha: tuboReemplazo.colmena.serial ? tuboReemplazo.colmena.serial.fecha : null });
                                }
                                SistemaInventario.colmenasDisponibles.splice(tuboReemplazo.indice, 1);
                                SistemaInventario.colmenasDisponibles.push({
                                    n_colmena: colmenaExistente.n_colmena,
                                    medida_mm: sobrante,
                                    medida_cm: sobrante / 10,
                                    cod: codReemplazo,
                                    serial: tuboReemplazo.colmena.serial || null
                                });
                            } else {
                                const idxInsertar = SistemaInventario.colmenasHistorico.findIndex(c => c.n_colmena === tuboReemplazo.colmena.n_colmena && c.cod === tuboReemplazo.colmena.cod && c.medida_mm === tuboReemplazo.colmena.medida_mm);
                                if (idxInsertar !== -1) {
                                    SistemaInventario.colmenasHistorico.splice(idxInsertar + 1, 0, { n_colmena: tuboReemplazo.colmena.n_colmena, medida_cm: sobrante / 10, medida_mm: sobrante, cod: tuboReemplazo.colmena.cod, codigo_original: tuboReemplazo.colmena.cod, estado: 'disponible', origen: 'Sobrante reemplazo orden ' + orden.id, posicionOriginal: tuboReemplazo.colmena.n_colmena, serial: tuboReemplazo.colmena.serial || null, fecha: tuboReemplazo.colmena.serial ? tuboReemplazo.colmena.serial.fecha : null });
                                }
                                SistemaInventario.colmenasDisponibles.splice(tuboReemplazo.indice, 1);
                                SistemaInventario.colmenasDisponibles.push({
                                    n_colmena: tuboReemplazo.colmena.n_colmena,
                                    medida_mm: sobrante,
                                    medida_cm: sobrante / 10,
                                    cod: codReemplazo,
                                    serial: tuboReemplazo.colmena.serial || null
                                });
                            }
                            SistemaInventario.colmenasDisponibles.splice(tuboReemplazo.indice, 1);
                        }
                        break;
                    }
                }
            }
            if (!resultado) {
                let serialDisponible = buscarSerialDisponible(codOrden);
                let codParaTuboNuevo = codOrden;
                if (!serialDisponible) {
                    const reemplazosParaNuevo = buscarReemplazos(codOrden);
                    if (reemplazosParaNuevo && reemplazosParaNuevo.length > 0) {
                        for (const codReemp of reemplazosParaNuevo) {
                            const serialReemp = buscarSerialDisponible(codReemp);
                            if (serialReemp) {
                                serialDisponible = serialReemp;
                                codParaTuboNuevo = codReemp;
                                log(`🔄 Serial no disponible para ${codOrden} — abriendo tubo nuevo de reemplazo ${codReemp}`, 'info');
                                break;
                            }
                        }
                    }
                    if (!serialDisponible) log(`⚠️ No hay serial para ${codOrden} ni sus reemplazos`, 'warn');
                }
                const codBuscarMedida = String(codParaTuboNuevo || '').trim().toUpperCase();
                let medidaNuevoCm = SistemaInventario.catalogoMedidas[codBuscarMedida] || (MM_TUBO_ORIGINAL / 10);

                // Rectificación al vuelo: si hay un override de medida real, usarlo
                if (SistemaInventario.overridesNuevos[codBuscarMedida] && SistemaInventario.overridesNuevos[codBuscarMedida].length > 0) {
                    medidaNuevoCm = SistemaInventario.overridesNuevos[codBuscarMedida].shift();
                    log(`📐 Rectificación aplicada para ${codBuscarMedida}: ${medidaNuevoCm} cm (medida real del tubo)`, 'info');
                }

                const medidaNuevoMm = medidaNuevoCm * 10;

                const sobranteNuevo = medidaNuevoMm - orden.medida_mm - MM_KERF;
                
                // Calcular posición nueva ANTES de crear el resultado para poder incluirla en res.colmena
                const posicionesOcupadas = new Set(SistemaInventario.colmenasHistorico.map(c => c.n_colmena));
                let posicionNueva = null;
                for (let i = 1; i <= colmenasAUsar.length + 100; i++) {
                    const pos = 'A' + i;
                    if (!posicionesOcupadas.has(pos)) { posicionNueva = pos; break; }
                }
                
                // Nomenclatura dinámica: TUBO NUEVO vs CENEFA OVALADA NUEVA, etc.
                let nombreMaterialNuevo = 'TUBO NUEVO';
                if (orden.componente && orden.componente !== 'TUBO') {
                    nombreMaterialNuevo = `${orden.componente} NUEVO`;
                }

                const esReemplazoNuevo = codParaTuboNuevo !== codOrden;
                if (serialDisponible) {
                    resultado = {
                        orden: orden.id,
                        medida_cm: orden.medida_cm,
                        fuente: esReemplazoNuevo ? 'reemplazo' : 'tubo_nuevo',
                        colmena: posicionNueva,
                        codigo: codParaTuboNuevo,
                        codigo_original: codOrden,
                        codigo_reemplazo: esReemplazoNuevo ? codParaTuboNuevo : undefined,
                        sobrante_cm: sobranteNuevo / 10,
                        medida_origen: medidaNuevoCm,
                        serial: serialDisponible,
                        nombreMaterialNuevo: nombreMaterialNuevo
                    };
                    marcarSerialComoOcupado(serialDisponible, orden.id);
                    log(`🏷️ Serial asignado: ${serialDisponible.codigo} - Lote: ${serialDisponible.lote} - Paquete: ${serialDisponible.paquete} - Serial: ${serialDisponible.serial}`, 'info');
                } else {
                    resultado = {
                        orden: orden.id,
                        medida_cm: orden.medida_cm,
                        fuente: 'tubo_nuevo',
                        colmena: posicionNueva,
                        codigo: codParaTuboNuevo,
                        codigo_original: codOrden,
                        sobrante_cm: sobranteNuevo / 10,
                        medida_origen: medidaNuevoCm,
                        nombreMaterialNuevo: nombreMaterialNuevo
                    };
                }
                const codigoTuboNuevo = codParaTuboNuevo || 'TUBO-NUEVO';
                
                // Añadir información del serial al histórico si está disponible
                const infoSerial = serialDisponible ? 
                    ` (Serial: ${serialDisponible.lote}-${serialDisponible.paquete}-${serialDisponible.serial})` : '';
                
                SistemaInventario.colmenasHistorico.push({ 
                    n_colmena: posicionNueva, 
                    medida_cm: medidaNuevoCm,
                    medida_mm: medidaNuevoMm,
                    cod: codigoTuboNuevo, 
                    codigo_original: codOrden, 
                    estado: 'usada', 
                    origen: `Orden ${orden.id} (Tubo nuevo)${infoSerial}`, 
                    posicionOriginal: posicionNueva,
                    serial: serialDisponible || null
                });
                
                const clasificacion = evaluarSobrante(sobranteNuevo);
                if (clasificacion.estado !== 'prohibido' && clasificacion.estado !== 'merma') {
                    const colmenaExistente = buscarColmenaDisponibleConCodigo(codigoTuboNuevo);
                    // CORRECCIÓN TRAZABILIDAD: guardar colmena destino real del sobrante
                    resultado.colmena_sobrante = colmenaExistente ? colmenaExistente.n_colmena : posicionNueva;

                    if (colmenaExistente) {
                        log(`📦 Sobrante agregado como nueva fila para colmena ${colmenaExistente.n_colmena} (código ${codigoTuboNuevo}). Medida sobrante: ${sobranteNuevo / 10}cm`, 'info');
                        const idxHistoricoExistente = SistemaInventario.colmenasHistorico.findIndex(c => c.n_colmena === colmenaExistente.n_colmena);
                        if (idxHistoricoExistente !== -1) {
                            SistemaInventario.colmenasHistorico.splice(idxHistoricoExistente + 1, 0, { 
                                n_colmena: colmenaExistente.n_colmena, 
                                medida_cm: sobranteNuevo / 10, 
                                medida_mm: sobranteNuevo, 
                                cod: codigoTuboNuevo, 
                                codigo_original: codigoTuboNuevo, 
                                estado: 'disponible', 
                                origen: `Sobrante tubo nuevo orden ${orden.id}${infoSerial}`, 
                                posicionOriginal: colmenaExistente.n_colmena,
                                serial: serialDisponible || null,
                                fecha: serialDisponible ? serialDisponible.fecha : null
                            });
                        }
                        SistemaInventario.colmenasDisponibles.push({
                            n_colmena: colmenaExistente.n_colmena,
                            medida_mm: sobranteNuevo,
                            medida_cm: sobranteNuevo / 10,
                            cod: codigoTuboNuevo,
                            serial: serialDisponible || null
                        });
                    } else {
                        SistemaInventario.colmenasHistorico.push({ 
                            n_colmena: posicionNueva, 
                            medida_cm: sobranteNuevo / 10, 
                            medida_mm: sobranteNuevo, 
                            cod: codigoTuboNuevo, 
                            codigo_original: codigoTuboNuevo, 
                            estado: 'disponible', 
                            origen: `Sobrante tubo nuevo orden ${orden.id}${infoSerial}`, 
                            posicionOriginal: posicionNueva,
                            serial: serialDisponible || null,
                            fecha: serialDisponible ? serialDisponible.fecha : null
                        });
                        SistemaInventario.colmenasDisponibles.push({
                            n_colmena: posicionNueva,
                            medida_mm: sobranteNuevo,
                            medida_cm: sobranteNuevo / 10,
                            cod: codigoTuboNuevo,
                            serial: serialDisponible || null
                        });
                    }
                }
            }
        }
        
        // Agregar color del catálogo al resultado para trazabilidad
        if (resultado) {
            resultado.color = obtenerColorDeCatalogo(resultado.codigo || resultado.codigo_original || '');
        }
        resultados.push(resultado);
        SistemaInventario.resultadosOptimizacion.push({ orden: orden, resultado: resultado });
        const infoResultado = formatearResultado(orden, resultado);
        agregarPaso(numPaso, `Orden ${orden.id}: ${orden.medida_cm}cm`, infoResultado);
    });

    // ─── Limpieza de sobrantes intermedios en memoria (anti inventario fantasma) ───
    // Recorrer resultados de atrás hacia adelante: si un sobrante fue reutilizado
    // como tubo origen en un corte posterior, es intermedio y debe eliminarse de
    // resultadosOptimizacion Y de colmenasHistorico ANTES de persistir.
    const _tubosConsumidos = []; // Cada elemento: { llave, ri } para rastrear el consumidor
    for (let ri = SistemaInventario.resultadosOptimizacion.length - 1; ri >= 0; ri--) {
        const item = SistemaInventario.resultadosOptimizacion[ri];
        const res = item.resultado;
        if (!res) continue;
        const codigoRes = res.codigo || res.codigo_original || '';

        // Registrar cada tubo origen consumido por un corte (con índice del consumidor)
        const origenNum = Number(res.medida_origen);
        if (!isNaN(origenNum) && origenNum > 0) {
            _tubosConsumidos.push({ llave: `${codigoRes}|${origenNum.toFixed(1)}`, ri: ri });
        }

        // Si este resultado generó un sobrante (no desecho), verificar si fue consumido más abajo
        if (res.sobrante_cm > 0 && !res.es_desecho) {
            const llaveSobrante = `${codigoRes}|${Number(res.sobrante_cm).toFixed(1)}`;
            const idxConsumo = _tubosConsumidos.findIndex(t => t.llave === llaveSobrante);
            if (idxConsumo !== -1) {
                // Re-etiquetar el resultado consumidor: su tubo vino de MESA, no de la colmena original
                const consumidorIdx = _tubosConsumidos[idxConsumo].ri;
                const resConsumidor = SistemaInventario.resultadosOptimizacion[consumidorIdx].resultado;
                if (resConsumidor) {
                    resConsumidor.colmena = 'MESA';
                }

                // Sobrante intermedio: fue reutilizado → limpiar
                _tubosConsumidos.splice(idxConsumo, 1); // balancear

                // Purgar el fantasma de colmenasHistorico (buscar entrada disponible con esa medida y código)
                const medidaFantasma = Math.round(res.sobrante_cm * 10);
                const idxFantasma = SistemaInventario.colmenasHistorico.findIndex(c =>
                    c.estado === 'disponible' &&
                    c.cod === codigoRes &&
                    c.medida_mm === medidaFantasma
                );
                if (idxFantasma !== -1) {
                    SistemaInventario.colmenasHistorico.splice(idxFantasma, 1);
                    log(`🧹 Sobrante intermedio eliminado de colmenas: ${codigoRes} ${res.sobrante_cm}cm`, 'info');
                }

                // También purgar de colmenasDisponibles
                const idxDispFantasma = SistemaInventario.colmenasDisponibles.findIndex(c =>
                    c.cod === codigoRes &&
                    c.medida_mm === medidaFantasma
                );
                if (idxDispFantasma !== -1) {
                    SistemaInventario.colmenasDisponibles.splice(idxDispFantasma, 1);
                }

                // Marcar como intermedio (visible para el operario, invisible para la BD)
                res.es_intermedio = true;
            }
        }
    }

    actualizarTablaColmenasResultado();
    actualizarTablaMermas();

    // ── Re-ordenar resultados por OT + Ubicación para el operario ──
    // La optimización procesó de mayor a menor (Best Fit Decreasing),
    // pero el plan de corte se muestra agrupado por kit de armado.
    SistemaInventario.resultadosOptimizacion.sort((a, b) => {
        const otA = String(a.orden.ot || a.orden.otAsignada || '');
        const otB = String(b.orden.ot || b.orden.otAsignada || '');
        if (otA !== otB) return otA.localeCompare(otB, undefined, { numeric: true });
        return String(a.orden.ubic || '').localeCompare(String(b.orden.ubic || ''), undefined, { numeric: true });
    });
    // También re-ordenar las órdenes para que el staging las muestre por kit
    SistemaInventario.ordenes.sort((a, b) => {
        const otA = String(a.ot || a.otAsignada || '');
        const otB = String(b.ot || b.otAsignada || '');
        if (otA !== otB) return otA.localeCompare(otB, undefined, { numeric: true });
        return String(a.ubic || '').localeCompare(String(b.ubic || ''), undefined, { numeric: true });
    });

    // ── RESUMEN DE AHORRO: métricas de eficiencia ──
    const resumen = { tubosNuevos: 0, sobrantesUsados: 0, reemplazos: 0, mermaTotal: 0, materialTotal: 0 };
    SistemaInventario.resultadosOptimizacion.forEach(item => {
        const r = item.resultado;
        if (!r) return;
        resumen.materialTotal += r.medida_cm || 0;
        if (r.fuente === 'tubo_nuevo') resumen.tubosNuevos++;
        else if (r.fuente === 'reemplazo') resumen.reemplazos++;
        else resumen.sobrantesUsados++;
        if (r.es_desecho && r.sobrante_cm > 0) resumen.mermaTotal += r.sobrante_cm;
    });
    const totalCortes = SistemaInventario.resultadosOptimizacion.length;
    const pctReutilizado = totalCortes > 0 ? Math.round((resumen.sobrantesUsados / totalCortes) * 100) : 0;
    log(`📊 RESUMEN: ${totalCortes} cortes | ${resumen.tubosNuevos} tubos nuevos | ${resumen.sobrantesUsados} sobrantes reutilizados (${pctReutilizado}%) | ${resumen.reemplazos} reemplazos | Merma desechada: ${resumen.mermaTotal.toFixed(1)} cm`, 'success');

    log('=== CÁLCULO COMPLETADO — Revise la Vista Previa ===', 'success');

    // Renderizar panel de staging (vista previa)
    renderizarStaging(resultados);

    // Actualizar alertas de stock (informativo, aún no guardado)
    verificarAlertasStock();
}

// ═══ STAGING: VISTA PREVIA ANTES DE GUARDAR ═══════════════════════════════

function renderizarStaging(resultados) {
    const panel = document.getElementById('panel-vista-previa');
    panel.style.display = 'block';
    panel.scrollIntoView({ behavior: 'smooth' });

    // ── Resumen de Ahorro visual (tarjetas) ──
    const resumenDiv = document.getElementById('resumenAhorro');
    if (resumenDiv) {
        const stats = { nuevos: 0, reutilizados: 0, reemplazos: 0, merma: 0 };
        resultados.forEach(r => {
            if (!r) return;
            if (r.fuente === 'tubo_nuevo') stats.nuevos++;
            else if (r.fuente === 'reemplazo') stats.reemplazos++;
            else stats.reutilizados++;
            if (r.es_desecho && r.sobrante_cm > 0) stats.merma += r.sobrante_cm;
        });
        const total = resultados.length;
        const pct = total > 0 ? Math.round((stats.reutilizados / total) * 100) : 0;
        resumenDiv.innerHTML = `
            <div class="stat-card"><div class="num">${total}</div><div class="lbl">Cortes totales</div></div>
            <div class="stat-card"><div class="num" style="color:#f39c12">${stats.nuevos}</div><div class="lbl">Tubos nuevos</div></div>
            <div class="stat-card"><div class="num" style="color:#27ae60">${stats.reutilizados}</div><div class="lbl">Sobrantes usados</div></div>
            <div class="stat-card"><div class="num" style="color:#3498db">${stats.reemplazos}</div><div class="lbl">Reemplazos</div></div>
            <div class="stat-card"><div class="num" style="color:#e74c3c">${stats.merma.toFixed(1)}</div><div class="lbl">Merma (cm)</div></div>
            <div class="stat-card"><div class="num" style="color:${pct >= 50 ? '#27ae60' : '#e67e22'}">${pct}%</div><div class="lbl">Reutilización</div></div>
        `;
    }

    // ── Sección A: Órdenes de Entrada ──
    const tbodyOrdenes = document.getElementById('tbodyStagingOrdenes');
    tbodyOrdenes.innerHTML = SistemaInventario.ordenes.map((o, i) => `<tr>
        <td>${i + 1}</td>
        <td>${o.ot || '-'}</td>
        <td>${o.ubic || '-'}</td>
        <td>${o.cod || '-'}</td>
        <td>${o.medida_cm || '-'}</td>
        <td>${o.componente || '-'}</td>
    </tr>`).join('');

    // ── Sección B: Inventario a Consumir ──
    const tbodyConsumo = document.getElementById('tbodyStagingConsumo');
    const consumidos = SistemaInventario.colmenasHistorico.filter(c => c.estado === 'usada' || c.origen.includes('MERMA'));
    tbodyConsumo.innerHTML = consumidos.map(c => {
        let tipo = '';
        if (c.origen.includes('MERMA')) tipo = '<span style="color:#e74c3c;">MERMA</span>';
        else tipo = '<span style="color:#00d2ff;">CORTE</span>';
        return `<tr>
            <td>${c.n_colmena || '-'}</td>
            <td>${c.cod || '-'}</td>
            <td>${c.medida_cm || '-'}</td>
            <td>${c.origen || '-'}</td>
            <td>${tipo}</td>
        </tr>`;
    }).join('') || '<tr><td colspan="5">Ningún tubo consumido</td></tr>';

    // ── Sección C: Resultado del Plan (con filas color-coded) ──
    const tbodyPlan = document.getElementById('tbodyStagingPlan');
    tbodyPlan.innerHTML = resultados.map(r => {
        const ordenOriginal = SistemaInventario.ordenes[r.orden - 1] || {};
        const codigoReal = r.codigo || r.codigo_original || '';
        const codigoDisplay = (r.codigo_original && r.codigo && r.codigo_original !== r.codigo)
            ? `${r.codigo_original} → ${r.codigo}`
            : codigoReal;
        const color = obtenerColorDeCatalogo(codigoReal) || ordenOriginal.color || '';

        // Acción con tag de color
        let accionTag = '';
        if (r.fuente === 'tubo_nuevo') accionTag = '<span class="tag-accion" style="background:#f39c12;color:#fff;">TUBO NUEVO</span>';
        else if (r.fuente === 'reemplazo') accionTag = '<span class="tag-accion" style="background:#3498db;color:#fff;">REEMPLAZO</span>';
        else if (r.fuente === 'merma') accionTag = '<span class="tag-accion" style="background:#e74c3c;color:#fff;">MERMA</span>';
        else accionTag = '<span class="tag-accion" style="background:#534ab7;color:#fff;">COLMENA</span>';

        let filas = `<tr class="fila-cortar">
            <td>${r.colmena || r.nombreMaterialNuevo || '-'}</td>
            <td>${codigoDisplay}</td>
            <td>${color}</td>
            <td>${formatearValor(r.medida_cm)}</td>
            <td>${formatearValor(r.medida_origen)}</td>
            <td>${formatearValor(r.sobrante_cm)}</td>
            <td>${accionTag}</td>
        </tr>`;

        // Fila de sobrante si aplica (color-coded)
        if (r.sobrante_cm > 0) {
            let accionSobrante, claseRow;
            if (r.es_intermedio) {
                accionSobrante = '<span class="tag-accion tag-mesa">RESERVAR EN MESA</span>';
                claseRow = 'fila-mesa';
            } else if (r.es_desecho) {
                accionSobrante = '<span class="tag-accion tag-merma">DESECHAR MERMA</span>';
                claseRow = 'fila-merma';
            } else {
                accionSobrante = '<span class="tag-accion tag-guardar">GUARDAR SOBRANTE</span>';
                claseRow = 'fila-guardar';
            }
            filas += `<tr class="${claseRow}">
                <td>${r.colmena_sobrante || '-'}</td>
                <td>${codigoDisplay}</td>
                <td></td>
                <td colspan="2">${formatearValor(r.sobrante_cm)} cm</td>
                <td></td>
                <td>${accionSobrante}</td>
            </tr>`;
        }
        return filas;
    }).join('');

    // ── Resumen rápido en el panel principal de resultados ──
    let html = '<div class="resumen-resultado"><h3>Vista Previa Generada</h3>';
    html += `<p>Se calcularon <strong>${resultados.length}</strong> cortes. Revise el panel de vista previa arriba y confirme para guardar.</p>`;
    html += '</div>';
    document.getElementById('resultados').innerHTML = html;
}

// ── PASO 1: Capturar snapshot intacto ANTES de cualquier escritura ──
// Devuelve un objeto con las colmenas y seriales tal como están en Firebase en este instante.
// Se ejecuta ANTES de guardarSistema() y guardarColmenaFinalEnFirestore() para evitar race conditions.
async function capturarSnapshotPreOptimizacion() {
    const user = window.firebaseAuth.currentUser;
    if (!user || !user.email) return null;

    const db = window.firebaseDb;
    let snapshotColmenas = [];
    let snapshotSeriales = [];

    // ── Colmenas: leer directamente del servidor (bypass de caché) ──
    try {
        const colmenaRef = window.firebaseDoc(db, "usuarios", user.email, "colmena_final", "datos");
        const colmenaSnap = await window.firebaseGetDocFromServer(colmenaRef);
        if (colmenaSnap && colmenaSnap.exists()) {
            const colData = colmenaSnap.data();
            if (colData && colData.data) {
                snapshotColmenas = JSON.parse(colData.data) || [];
            }
        }
        console.log(`📸 Snapshot colmenas capturado desde servidor: ${snapshotColmenas.length} registros`);
    } catch (fbErr) {
        console.warn("getDocFromServer falló para snapshot de colmenas, usando copias inmutables:", fbErr.message);
        // Fallback: copias inmutables tomadas al inicio de ejecutarOptimizacion (estado pre-corte)
        snapshotColmenas = SistemaInventario.colmenaCruda.length > 0
            ? JSON.parse(JSON.stringify(SistemaInventario.colmenaCruda))
            : (colmenaActual ? JSON.parse(JSON.stringify(colmenaActual)) : []);
    }

    // ── Seriales: usar la copia inmutable pre-optimización ──
    snapshotSeriales = SistemaInventario.serialesCrudos.length > 0
        ? JSON.parse(JSON.stringify(SistemaInventario.serialesCrudos))
        : JSON.parse(JSON.stringify(SistemaInventario.seriales));

    return { colmenas: snapshotColmenas, seriales: snapshotSeriales };
}

// ── PASO 2: Persistir el snapshot ya capturado en historial_operaciones ──
// Recibe el snapshot pre-capturado para garantizar que NUNCA lea estado post-guardado.
async function guardarPuntoRestauracion(snapshotIntacto) {
    try {
        const user = window.firebaseAuth.currentUser;
        if (!user) return;
        const db = window.firebaseDb;

        // Construir array de resultados para re-dibujar la vista previa
        const resultadosGuardar = SistemaInventario.resultadosOptimizacion.map(item => {
            const res = item.resultado;
            const ord = item.orden || {};
            return { ...res, ot: ord.ot || '', ubic: ord.ubic || '', componente: ord.componente || '', cod_orden: ord.cod || '' };
        });

        const nombreExcel = `plan_corte_${new Date().toISOString().slice(0,10)}.xlsx`;

        const registro = {
            fecha: new Date().toISOString(),
            nombre_excel: nombreExcel,
            usuario: user.email,
            resultados: JSON.stringify(resultadosGuardar),
            snapshot_inventario: JSON.stringify(snapshotIntacto.colmenas),
            snapshot_seriales: JSON.stringify(snapshotIntacto.seriales),
            total_cortes: resultadosGuardar.length
        };

        const colRef = window.firebaseCollection(db, "usuarios", user.email, "historial_operaciones");
        await window.firebaseAddDoc(colRef, registro);

        log('📸 Punto de restauración guardado en historial.', 'info');
    } catch (error) {
        console.error("Error guardando punto de restauración:", error);
        log('⚠️ No se pudo guardar el punto de restauración: ' + error.message, 'warn');
    }
}

// ── Confirmar: guardar en Firebase + descargar Excel ──
async function confirmarYGuardarStaging() {
    log('💾 Confirmado. Guardando en Firebase...', 'info');

    // ── Bloqueo preventivo: deshabilitar "Calcular" mientras se persiste ──
    const btnEjecutar = document.getElementById('btnEjecutar');
    if (btnEjecutar) btnEjecutar.disabled = true;

    // ── PASO 1: Capturar snapshot INTACTO desde Firebase (antes de cualquier escritura) ──
    const snapshotSeguro = await capturarSnapshotPreOptimizacion();

    // ── PASO 2: Persistir el punto de restauración con el snapshot puro ──
    if (snapshotSeguro) {
        await guardarPuntoRestauracion(snapshotSeguro);
    } else {
        log('⚠️ No se pudo capturar snapshot pre-optimización. El historial puede estar incompleto.', 'warn');
    }

    // ── PASO 3: AHORA sí persistir los cambios (descontar tubos) ──
    guardarSistema();
    await guardarColmenaFinalEnFirestore();

    // ── PASO 3b: Marcar OTs como procesadas hoy (protección anti-duplicados) ──
    const hoy = new Date().toISOString().split('T')[0];
    const otsProcesadas = JSON.parse(localStorage.getItem('ots_procesadas_hoy') || '{}');
    // Limpiar OTs de días anteriores
    Object.keys(otsProcesadas).forEach(ot => { if (otsProcesadas[ot] !== hoy) delete otsProcesadas[ot]; });
    SistemaInventario.ordenes.forEach(o => {
        const ot = String(o.ot || '').trim();
        if (ot && ot !== '-') otsProcesadas[ot] = hoy;
    });
    localStorage.setItem('ots_procesadas_hoy', JSON.stringify(otsProcesadas));

    // ── Capturar resumen ANTES de limpiar el estado temporal ──
    const resultados = SistemaInventario.resultadosOptimizacion.map(item => item.resultado);

    // ── Guardar último plan en Firebase para acceso multi-dispositivo ──
    try {
        const user = window.firebaseAuth.currentUser;
        if (user) {
            const planData = {
                resultados: JSON.stringify(SistemaInventario.resultadosOptimizacion.map(item => ({
                    resultado: item.resultado,
                    orden: {
                        id: item.orden.id, ot: item.orden.ot, ubic: item.orden.ubic,
                        cod: item.orden.cod, medida_cm: item.orden.medida_cm,
                        componente: item.orden.componente, color: item.orden.color
                    }
                }))),
                ordenes: JSON.stringify(SistemaInventario.ordenes.map(o => ({
                    id: o.id, ot: o.ot, ubic: o.ubic, cod: o.cod,
                    medida_cm: o.medida_cm, componente: o.componente
                }))),
                fecha: new Date().toISOString()
            };
            await window.firebaseSetDoc(
                window.firebaseDoc(window.firebaseDb, "usuarios", user.email, "ultimo_plan", "datos"),
                planData
            );
            log('☁️ Último plan guardado en la nube (accesible desde cualquier dispositivo)', 'info');
        }
    } catch (e) {
        console.warn("No se pudo guardar último plan en Firebase:", e.message);
    }

    // ── Exportar Excel ANTES de limpiar (usa resultadosOptimizacion que aún tiene datos) ──
    exportarResultados();

    // ── Limpieza de estado temporal para que el siguiente Excel parta limpio ──
    SistemaInventario.resultadosOptimizacion = [];
    SistemaInventario.colmenasHistorico = [];
    SistemaInventario.colmenasDisponibles = [];
    SistemaInventario.mermas = [];
    SistemaInventario.ordenesCrudas = [];
    SistemaInventario.serialesCrudos = [];
    SistemaInventario.colmenaCruda = [];
    SistemaInventario.overridesNuevos = {};

    // ── Limpieza visual de inputs de archivo: forzar re-selección ──
    ['fileOrdenes', 'fileColmenas', 'fileCatalogo', 'fileEstructura'].forEach(id => {
        const input = document.getElementById(id);
        if (input) input.value = '';
    });
    ['estadoOrdenes', 'estadoColmenas', 'estadoCatalogo', 'estadoEstructura'].forEach(id => {
        const el = document.getElementById(id);
        if (el) { el.textContent = ''; el.className = 'estado-archivo'; }
    });

    // ── Limpieza de caché localStorage: forzar regeneración desde cero ──
    localStorage.removeItem('sistemaInventario');

    // ── Refresco forzado: re-leer colmena_final desde Firebase (bypass de caché) ──
    try {
        const user = window.firebaseAuth.currentUser;
        if (user && user.email) {
            const db = window.firebaseDb;
            const colmenaRef = window.firebaseDoc(db, "usuarios", user.email, "colmena_final", "datos");
            const freshSnap = await window.firebaseGetDocFromServer(colmenaRef);
            if (freshSnap && freshSnap.exists()) {
                const freshData = freshSnap.data();
                if (freshData && freshData.data) {
                    const colmenasFrescas = JSON.parse(freshData.data) || [];
                    colmenaActual = colmenasFrescas;
                    SistemaInventario.colmenas = colmenasFrescas;
                    actualizarTablaColmenas();
                    log(`🔄 Inventario refrescado desde Firebase: ${colmenasFrescas.length} colmenas`, 'info');
                }
            }
        }
    } catch (refreshErr) {
        console.warn("Refresco forzado desde servidor falló, onSnapshot mantendrá la sincronía:", refreshErr.message);
    }

    // Persistir el estado limpio en localStorage
    guardarSistema();

    log('🔄 Memoria local sincronizada. Listo para el siguiente archivo.', 'info');

    // Habilitar botones de exportación
    document.getElementById('btnExportar').disabled = false;
    document.getElementById('btnExportarDisponibles').disabled = false;

    // Ocultar panel de staging
    document.getElementById('panel-vista-previa').style.display = 'none';

    // Mostrar resumen final (usa la copia capturada antes de la limpieza)
    let html = '<div class="resumen-resultado"><h3>✓ Optimización Confirmada y Guardada</h3>';
    resultados.forEach(r => {
        let badge = 'badge';
        let clasificacion = r.fuente.toUpperCase();
        if (r.fuente === 'tubo_nuevo') badge += ' badge-nueva';
        if (r.fuente === 'reemplazo') badge += ' badge-reemplazo';

        const ordenOriginal = SistemaInventario.ordenes[r.orden - 1] || {};
        const ot = ordenOriginal.ot || '';
        const ubic = ordenOriginal.ubic || '';
        const color = ordenOriginal.color || '';
        const tituloOrden = ot ? `OT ${ot}` : `Orden ${r.orden}`;
        const codigo = r.codigo || r.codigo_original || '';

        let infoReemplazo = '';
        if (r.fuente === 'reemplazo' && r.codigo_original && r.codigo_reemplazo) {
            infoReemplazo = ` <span style="color: #3498db;">[${r.codigo_original} → ${r.codigo_reemplazo}]</span>`;
        }

        html += `<p><strong>${tituloOrden}</strong>${ubic ? ` [${ubic}]` : ''}${color ? ` [${color}]` : ''}: ${formatearValor(r.medida_cm)}cm ${codigo ? `(${codigo})` : ''}${infoReemplazo} → <span class="${badge}">${clasificacion}</span> | Sobrante: ${formatearValor(r.sobrante_cm)}cm</p>`;
    });
    html += '</div>';
    document.getElementById('resultados').innerHTML = html;

    // ── Re-habilitar "Calcular" tras sincronización completa ──
    if (btnEjecutar) btnEjecutar.disabled = false;

    log('✅ Inventario guardado y Excel descargado.', 'success');
}

// ── Descartar: volver al estado previo sin guardar ──
function descartarStaging() {
    if (!confirm('¿Descartar la vista previa? Los cambios calculados se perderán.')) return;

    document.getElementById('panel-vista-previa').style.display = 'none';
    document.getElementById('resultados').innerHTML = '';
    document.getElementById('btnExportar').disabled = true;
    document.getElementById('btnExportarDisponibles').disabled = true;

    // Restaurar estado pre-optimización
    if (SistemaInventario.ordenesCrudas.length > 0) {
        SistemaInventario.ordenes = JSON.parse(JSON.stringify(SistemaInventario.ordenesCrudas));
    }
    if (SistemaInventario.serialesCrudos.length > 0) {
        SistemaInventario.seriales = JSON.parse(JSON.stringify(SistemaInventario.serialesCrudos));
    }
    if (SistemaInventario.colmenaCruda.length > 0) {
        SistemaInventario.colmenas = JSON.parse(JSON.stringify(SistemaInventario.colmenaCruda));
    } else if (colmenaActual && colmenaActual.length > 0) {
        SistemaInventario.colmenas = JSON.parse(JSON.stringify(colmenaActual));
    }

    SistemaInventario.resultadosOptimizacion = [];
    SistemaInventario.colmenasHistorico = [];
    SistemaInventario.colmenasDisponibles = [];
    SistemaInventario.mermas = [];
    SistemaInventario.ordenesCrudas = [];
    SistemaInventario.serialesCrudos = [];
    SistemaInventario.colmenaCruda = [];
    SistemaInventario.overridesNuevos = {};

    // ── Limpieza visual de inputs de archivo: forzar re-selección ──
    ['fileOrdenes', 'fileColmenas', 'fileCatalogo', 'fileEstructura'].forEach(id => {
        const input = document.getElementById(id);
        if (input) input.value = '';
    });
    ['estadoOrdenes', 'estadoColmenas', 'estadoCatalogo', 'estadoEstructura'].forEach(id => {
        const el = document.getElementById(id);
        if (el) { el.textContent = ''; el.className = 'estado-archivo'; }
    });

    actualizarTablaColmenas();
    actualizarTablaSeriales();
    log('🗑️ Vista previa descartada. Estado restaurado.', 'info');
}

// ── Cargar último plan desde Firebase (para acceso multi-dispositivo) ──
async function cargarUltimoPlanDesdeFirebase() {
    try {
        const user = window.firebaseAuth.currentUser;
        if (!user) return false;

        const db = window.firebaseDb;
        const docRef = window.firebaseDoc(db, "usuarios", user.email, "ultimo_plan", "datos");
        const snap = await window.firebaseGetDoc(docRef);

        if (!snap.exists()) return false;

        const data = snap.data();
        if (!data.resultados) return false;

        const items = JSON.parse(data.resultados);
        const ordenes = data.ordenes ? JSON.parse(data.ordenes) : [];

        SistemaInventario.resultadosOptimizacion = items;
        if (ordenes.length > 0) SistemaInventario.ordenes = ordenes;

        const fecha = data.fecha ? new Date(data.fecha).toLocaleString('es-CL') : 'desconocida';
        log(`☁️ Último plan cargado desde Firebase (${fecha}) — ${items.length} cortes`, 'info');
        return true;
    } catch (e) {
        console.warn("Error cargando último plan:", e.message);
        return false;
    }
}

async function exportarResultados() {
    // ── Si no hay datos locales, intentar cargar último plan desde Firebase ──
    if (SistemaInventario.resultadosOptimizacion.length === 0) {
        const cargado = await cargarUltimoPlanDesdeFirebase();
        if (!cargado) {
            alert('No hay resultados para exportar. Ejecute una optimización primero, o el último plan no está disponible.');
            return;
        }
    }
    // COLOR va después de Código (columna 5, índice 5)
    const datosExcel = [['OT', 'Ubicación', 'Acción', 'Colmena', 'Código', 'Color', 'Medida a Cortar (cm)', 'Tubo Origen (cm)', 'Lote', 'Paquete', 'Serial', 'Fecha Serial']];

    SistemaInventario.resultadosOptimizacion.forEach(item => {
        const res = item.resultado;
        const ord = SistemaInventario.ordenes.find(o => o.id === res.orden) || {};

        const s = res.serial || ord.serial || {};
        const fechaFormateada = s.fecha ? formatearFecha(s.fecha) : '-';
        const codigoReal = res.codigo || res.codigo_original || ord.cod || '-';
        const codigoExcel = (res.codigo_original && res.codigo && res.codigo_original !== res.codigo)
            ? `${res.codigo_original} → ${res.codigo}`
            : codigoReal;
        // res.color ya se setea en ejecutarOptimizacion; fallback a lookup directo
        const color = res.color || obtenerColorDeCatalogo(codigoReal);

        const _comp = (ord.componente && ord.componente !== 'TUBO') ? ord.componente : '';
        const _esTuboNuevoReemplazo = res.fuente === 'reemplazo'
            && res.codigo_original && res.codigo && res.codigo !== res.codigo_original
            && res.codigo_reemplazo;
        const _esTuboNuevo = res.fuente === 'tubo_nuevo';
        let accionCortar;
        if (_esTuboNuevoReemplazo) {
            accionCortar = `TUBO NUEVO (REEMPLAZO ${res.codigo_original} \u2192 ${res.codigo})`;
        } else if (_esTuboNuevo) {
            accionCortar = _comp ? `${_comp} NUEVO` : 'TUBO NUEVO';
        } else if (_comp) {
            accionCortar = `CORTAR ${_comp}`;
        } else {
            accionCortar = 'CORTAR';
        }
        const _colmenaExcel = (_esTuboNuevo || _esTuboNuevoReemplazo)
            ? (res.nombreMaterialNuevo || 'TUBO NUEVO')
            : (res.colmena || '-');

        // Fila CORTAR
        datosExcel.push([
            ord.ot || '-',
            ord.ubic || '-',
            accionCortar,
            _colmenaExcel,
            codigoExcel,
            color,
            res.medida_cm,
            res.medida_origen || '-',
            s.lote || '-',
            s.paquete || '-',
            s.serial || '-',
            fechaFormateada
        ]);

        // Si existe un sobrante (> 0), insertar una fila según su tipo
        if (res.sobrante_cm > 0) {
            let accionSobrante, colmenaDestino;

            if (res.es_intermedio) {
                // Sobrante intermedio: el operario lo dejará en mesa, se reutiliza en el siguiente corte
                accionSobrante = 'RESERVAR EN MESA';
                colmenaDestino = '-';
            } else if (res.es_desecho) {
                // Merma ≤ 10 cm
                accionSobrante = 'DESECHAR MERMA';
                colmenaDestino = 'BASURERO';
            } else {
                // Sobrante final real (> 10 cm)
                accionSobrante = 'GUARDAR SOBRANTE';
                colmenaDestino = res.colmena_sobrante || res.colmena || '-';
            }

            datosExcel.push([
                ord.ot || '-',
                '',
                accionSobrante,
                colmenaDestino,
                codigoExcel,
                color,
                res.sobrante_cm,
                '-',
                s.lote || ord.lote || '-',
                s.paquete || ord.paquete || '-',
                s.serial || ord.serial || '-',
                s.fecha || ord.fecha || '-'
            ]);
        }
    });

    // ─── Paso 2: procesar GUARDAR SOBRANTE reutilizados ───
    // Los RESERVAR EN MESA ya están correctamente posicionados junto a su corte
    // (generados en el loop anterior cuando es_intermedio=true). No tocarlos.
    // Solo procesar GUARDAR SOBRANTE que fueron consumidos más abajo como tubo origen.
    const tubosConsumidos = [];
    for (let f = datosExcel.length - 1; f >= 1; f--) {
        const fila = datosExcel[f];
        const accion = String(fila[2] || '').toUpperCase();
        const codigo = fila[4];

        if (accion.includes('CORTAR')) {
            const origen = Number(fila[7]);
            if (!isNaN(origen) && origen > 0) {
                tubosConsumidos.push({ llave: `${codigo}|${origen.toFixed(1)}`, filaIdx: f });
            }
        } else if (accion === 'GUARDAR SOBRANTE') {
            // Solo GUARDAR SOBRANTE — nunca tocar RESERVAR EN MESA ya colocados
            const medidaSob = Number(fila[6]);
            if (!isNaN(medidaSob) && medidaSob > 0) {
                const llaveSobrante = `${codigo}|${medidaSob.toFixed(1)}`;
                const idxConsumo = tubosConsumidos.findIndex(t => t.llave === llaveSobrante);
                if (idxConsumo !== -1) {
                    // Re-etiquetar el CORTAR consumidor: su tubo vino de MESA
                    const consumidorIdx = tubosConsumidos[idxConsumo].filaIdx;
                    const filaConsumidor = datosExcel[consumidorIdx];
                    if (filaConsumidor) filaConsumidor[3] = 'MESA';

                    // Convertir GUARDAR SOBRANTE → RESERVAR EN MESA
                    const filaReserva = [...datosExcel[f]];
                    filaReserva[2] = 'RESERVAR EN MESA';
                    filaReserva[3] = '-';

                    // Eliminar de posición original
                    datosExcel.splice(f, 1);

                    // Recalcular índice consumidor tras splice
                    const consumidorIdxAjustado = f < consumidorIdx ? consumidorIdx - 1 : consumidorIdx;

                    // Insertar RESERVAR EN MESA justo antes del CORTAR MESA
                    datosExcel.splice(consumidorIdxAjustado, 0, filaReserva);

                    tubosConsumidos.splice(idxConsumo, 1);
                }
            }
        }
    }

    const ws = XLSX.utils.aoa_to_sheet(datosExcel);

    // Aplicar estilos visuales a filas especiales
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = 1; R <= range.e.r; ++R) {
        const accionCell = ws[XLSX.utils.encode_cell({r: R, c: 2})]; // Columna "Acción"
        if (!accionCell) continue;

        let fillColor = null, fontColor = null;
        if (accionCell.v === 'DESECHAR MERMA') {
            fillColor = "FFFF9999"; fontColor = "FF990000"; // Rojo claro / rojo oscuro
        } else if (accionCell.v === 'RESERVAR EN MESA') {
            fillColor = "FFFFF3E0"; fontColor = "FFE65100"; // Naranja claro / naranja oscuro
        }

        if (fillColor) {
            for (let C = 0; C <= range.e.c; ++C) {
                const cell = ws[XLSX.utils.encode_cell({r: R, c: C})];
                if (cell) {
                    if (!cell.s) cell.s = {};
                    cell.s.fill = {fgColor: {rgb: fillColor}};
                    cell.s.font = {color: {rgb: fontColor}, bold: true};
                }
            }
        }
    }
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Plan de Corte');
    XLSX.writeFile(wb, `plan_corte_${new Date().toISOString().slice(0,10)}.xlsx`);
}

function exportarInventarioActualizado() {
    const disponibles = SistemaInventario.seriales.filter(s => s.estado === 'disponible');

    if (disponibles.length === 0) {
        alert("No hay inventario disponible para exportar.");
        return;
    }

    // Preparar las filas con sus cabeceras
    const filas = [['Fecha', 'Código', 'Lote', 'Paquete', 'Serial', 'Estado']];

    // Agrupar por Fecha+Código+Lote+Paquete y contar cuántos tubos quedan disponibles
    const agrupado = {};
    disponibles.forEach(s => {
        const key = `${s.fecha}|${s.codigo}|${s.lote}|${s.paquete}`;
        if (!agrupado[key]) {
            agrupado[key] = { ...s, cantidadRestante: 0 };
        }
        agrupado[key].cantidadRestante++;
    });

    Object.values(agrupado).forEach(g => {
        filas.push([
            g.fecha || '-',
            g.codigo || '-',
            g.lote || '-',
            g.paquete || '-',
            g.cantidadRestante.toString(), // Cantidad restante de tubos en ese paquete
            'disponible'
        ]);
    });

    // Crear el libro y la hoja usando SheetJS (XLSX)
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(filas);

    // Ajustar el ancho de las columnas para que se vea ordenado
    ws['!cols'] = [
        { wch: 15 }, // Fecha
        { wch: 12 }, // Código
        { wch: 10 }, // Lote
        { wch: 10 }, // Paquete
        { wch: 10 }, // Serial
        { wch: 15 }  // Estado
    ];

    XLSX.utils.book_append_sheet(wb, ws, "Inventario Actual");

    // Exportar el archivo como .xlsx real
    const nombreArchivo = `inventario_actualizado_${new Date().toISOString().split('T')[0]}.xlsx`;
    XLSX.writeFile(wb, nombreArchivo);

    log(`📥 Inventario actualizado exportado: ${disponibles.length} seriales disponibles`, 'success');
}

function exportarColmenasDisponibles() {
    const encabezado = ['N° Colmena (Ubicación)', 'Código', 'Medida (cm)', 'Estado', 'Fecha Registro', 'Lote', 'Paquete', 'Serial'];

    // ── Fuente de datos: colmenasHistorico si hay optimización activa,
    //    sino colmenaActual (Firebase) como fallback robusto ──
    let fuente = [];
    if (SistemaInventario.colmenasHistorico.length > 0) {
        fuente = SistemaInventario.colmenasHistorico.filter(c => c.estado === 'disponible');
    } else if (colmenaActual && colmenaActual.length > 0) {
        fuente = colmenaActual.map(c => ({
            n_colmena: c.n_colmena || '-',
            cod: c.cod || '-',
            medida_cm: c.medida_cm || (c.medida_mm ? (c.medida_mm / 10) : 0),
            estado: 'disponible',
            serial: c.serial || null
        }));
    } else if (SistemaInventario.colmenas.length > 0) {
        fuente = SistemaInventario.colmenas.map(c => ({
            n_colmena: c.n_colmena || '-',
            cod: c.cod || '-',
            medida_cm: c.medida_cm || (c.medida_mm ? (c.medida_mm / 10) : 0),
            estado: 'disponible',
            serial: c.serial || null
        }));
    }

    if (fuente.length === 0) {
        alert("No hay colmenas disponibles para exportar. Verifique que el inventario esté sincronizado.");
        return;
    }

    const filas = fuente.map(c => {
        const s = c.serial || {};
        return [
            c.n_colmena || '-',
            c.cod || '-',
            c.medida_cm || 0,
            c.estado || 'disponible',
            formatearFecha(s.fecha || c.fecha || '-'),
            s.lote || '-',
            s.paquete || '-',
            s.serial || '-'
        ];
    });
    const ws = XLSX.utils.aoa_to_sheet([encabezado, ...filas]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "UBICACION_COLMENAS");
    XLSX.writeFile(wb, `colmenas_disponibles_${new Date().toISOString().slice(0,10)}.xlsx`);
    log(`📦 Colmenas exportadas: ${filas.length} disponibles`, 'success');
}

function mostrarLogin() {
  document.body.innerHTML = `
    <div style="
      display:flex;
      justify-content:center;
      align-items:center;
      height:100vh;
      font-family:Arial;
      flex-direction:column;
      gap:10px;
    ">
      <h2>Login Sistema Inventario</h2>
      <input type="email" id="email" placeholder="Correo" />
      <input type="password" id="password" placeholder="Contraseña" />
      <button id="btnLogin">Ingresar</button>
      <p id="error" style="color:red;"></p>
    </div>
  `;

  document.getElementById("btnLogin")
    .addEventListener("click", iniciarSesion);
}
function iniciarSesion() {
  const email = document.getElementById("email").value;
  const password = document.getElementById("password").value;

  window.firebaseSignIn(window.firebaseAuth, email, password)
    .then(() => {
      location.reload();
    })
    .catch((error) => {
      document.getElementById("error").innerText = error.message;
    });
}
function cerrarSesion() {
  window.firebaseSignOut(window.firebaseAuth)
    .then(() => {
      location.reload();
    })
    .catch((error) => {
      console.error("Error al cerrar sesión:", error);
    });
}

window.cerrarSesion = cerrarSesion;

// ====== RECTIFICACIÓN AL VUELO DE TUBOS NUEVOS ======

function recalcularPlan() {
    // Restaurar colmenas desde la copia cruda
    if (SistemaInventario.colmenaCruda.length > 0) {
        SistemaInventario.colmenas = JSON.parse(JSON.stringify(SistemaInventario.colmenaCruda));
    }

    // Restaurar seriales al estado original (todos 'disponible', sin ordenId/fechaUso)
    if (SistemaInventario.serialesCrudos.length > 0) {
        SistemaInventario.seriales = JSON.parse(JSON.stringify(SistemaInventario.serialesCrudos));
    } else {
        SistemaInventario.seriales.forEach(s => {
            s.estado = 'disponible';
            delete s.ordenId;
            delete s.fechaUso;
        });
    }

    // Limpiar estado previo
    SistemaInventario.colmenasHistorico = [];
    SistemaInventario.resultadosOptimizacion = [];
    SistemaInventario.mermas = [];
    SistemaInventario.logs = [];

    // Re-expandir órdenes desde los datos crudos del Excel (FASE 3 completa)
    // Esto reconstruye tubos + accesorios + pesos desde cero
    SistemaInventario.ordenesCrudas = []; // Resetear para que ejecutarOptimizacion guarde la nueva copia
    expandirOrdenesDesdeExcel();

    // Re-ejecutar optimización con los overrides activos
    ejecutarOptimizacion();

    log('📐 Plan recalculado con medidas rectificadas', 'success');
}

document.addEventListener('DOMContentLoaded', function() {
    const btnRectificar = document.getElementById('btnRectificar');
    if (btnRectificar) {
        btnRectificar.addEventListener('click', function() {
            const input = prompt(
                'Ingrese código y medida real separados por coma (ej. E02,579.2).\n' +
                'Si son varios, sepárelos por punto y coma (ej. E02,579.2 ; E02,579.5 ; E53,598).'
            );
            if (!input || input.trim() === '') return;

            // Limpiar overrides previos
            SistemaInventario.overridesNuevos = {};

            const pares = input.split(';');
            let totalOverrides = 0;
            for (const par of pares) {
                const partes = par.split(',').map(s => s.trim());
                if (partes.length < 2) continue;
                const codigo = partes[0].toUpperCase();
                const medida = parseFloat(partes[1].replace(',', '.'));
                if (isNaN(medida) || medida <= 0) continue;

                if (!SistemaInventario.overridesNuevos[codigo]) {
                    SistemaInventario.overridesNuevos[codigo] = [];
                }
                SistemaInventario.overridesNuevos[codigo].push(medida);
                totalOverrides++;
            }

            if (totalOverrides === 0) {
                alert('No se detectaron rectificaciones válidas. Formato: E02,579.2 ; E53,598');
                return;
            }

            const resumen = Object.entries(SistemaInventario.overridesNuevos)
                .map(([cod, medidas]) => `${cod}: ${medidas.join(', ')} cm`)
                .join('\n');

            if (confirm(`Se aplicarán ${totalOverrides} rectificación(es):\n\n${resumen}\n\n¿Recalcular el plan?`)) {
                recalcularPlan();
            }
        });
    }

    // Botones del panel de staging (vista previa)
    const btnConfirmar = document.getElementById('btnConfirmarGuardar');
    if (btnConfirmar) btnConfirmar.addEventListener('click', confirmarYGuardarStaging);

    const btnDescartar = document.getElementById('btnDescartarPreview');
    if (btnDescartar) btnDescartar.addEventListener('click', descartarStaging);

    // ── Botón: Corregir Corte Erróneo ──
    const btnCorregir = document.getElementById('btnCorregirCorte');
    if (btnCorregir) btnCorregir.addEventListener('click', corregirCorteErroneo);
});

// ====== CORRECCIÓN DE CORTE ERRÓNEO (sin rollback) ======
// Permite al operario corregir un sobrante que quedó con medida diferente
// a la calculada, sin necesidad de revertir toda la operación.
//
// Flujo:
//   1. El operario indica la colmena del sobrante y la medida real
//   2. Se actualiza ese tubo específico en colmena_final (Firebase)
//   3. El resto del inventario no se toca

async function corregirCorteErroneo() {
    // Verificar que hay colmenas cargadas
    const colmenas = colmenaActual || SistemaInventario.colmenas;
    if (!colmenas || colmenas.length === 0) {
        alert('No hay inventario cargado. Espere a que se sincronice desde Firebase.');
        return;
    }

    // Paso 1: Pedir la colmena a corregir
    const listaColmenas = [...new Set(colmenas.map(c => c.n_colmena))].sort((a, b) => {
        const na = String(a).match(/(\D*)(\d+)/);
        const nb = String(b).match(/(\D*)(\d+)/);
        if (na && nb) {
            if (na[1] !== nb[1]) return na[1].localeCompare(nb[1]);
            return parseInt(na[2]) - parseInt(nb[2]);
        }
        return String(a).localeCompare(String(b));
    });

    const inputColmena = prompt(
        'CORREGIR CORTE ERRÓNEO\n\n' +
        'Paso 1: ¿En qué colmena está el sobrante que quedó con medida incorrecta?\n\n' +
        'Colmenas con tubos:\n' +
        listaColmenas.slice(0, 30).join(', ') +
        (listaColmenas.length > 30 ? '...' : '') +
        '\n\nEscribe el nombre de la colmena (ej: A27, L03):'
    );
    if (!inputColmena || inputColmena.trim() === '') return;
    const colmenaBuscada = inputColmena.trim().toUpperCase();

    // Buscar tubos en esa colmena
    const tubosEnColmena = colmenas
        .map((c, idx) => ({ ...c, _idx: idx }))
        .filter(c => String(c.n_colmena).toUpperCase() === colmenaBuscada);

    if (tubosEnColmena.length === 0) {
        alert(`No se encontraron tubos en la colmena "${colmenaBuscada}".`);
        return;
    }

    // Mostrar qué hay en esa colmena
    const detalle = tubosEnColmena.map((t, i) =>
        `  ${i + 1}) ${t.cod} — ${t.medida_cm} cm`
    ).join('\n');

    // Paso 2: Seleccionar cuál tubo y dar la medida real
    let seleccion = 1;
    if (tubosEnColmena.length > 1) {
        const inputSel = prompt(
            `Colmena ${colmenaBuscada} tiene ${tubosEnColmena.length} tubos:\n\n${detalle}\n\n` +
            `¿Cuál quieres corregir? (ingresa el número):`
        );
        if (!inputSel) return;
        seleccion = parseInt(inputSel);
        if (isNaN(seleccion) || seleccion < 1 || seleccion > tubosEnColmena.length) {
            alert('Número inválido.');
            return;
        }
    }

    const tuboSeleccionado = tubosEnColmena[seleccion - 1];

    const inputMedida = prompt(
        `Corrigiendo: ${tuboSeleccionado.cod} en ${colmenaBuscada}\n` +
        `Medida registrada: ${tuboSeleccionado.medida_cm} cm\n\n` +
        `¿Cuál es la medida REAL del sobrante? (en cm):`
    );
    if (!inputMedida) return;

    const medidaReal = limpiarNumero(inputMedida);
    if (medidaReal <= 0) {
        alert('Medida inválida.');
        return;
    }

    // Confirmar
    const confirmar = confirm(
        `¿Confirmar corrección?\n\n` +
        `Colmena: ${colmenaBuscada}\n` +
        `Código: ${tuboSeleccionado.cod}\n` +
        `Medida anterior: ${tuboSeleccionado.medida_cm} cm\n` +
        `Medida real: ${medidaReal} cm\n\n` +
        `Esto actualizará el inventario en Firebase.`
    );
    if (!confirmar) return;

    // Paso 3: Aplicar la corrección
    document.getElementById('loading-overlay').style.display = 'flex';

    try {
        const idxReal = tuboSeleccionado._idx;

        // Actualizar en el array local
        colmenas[idxReal].medida_cm = medidaReal;
        colmenas[idxReal].medida_mm = Math.round(medidaReal * 10);

        // Si la medida real es ≤ 0 o muy pequeña (merma), eliminar el tubo
        if (medidaReal <= 1) {
            colmenas.splice(idxReal, 1);
            log(`🗑️ Tubo eliminado de ${colmenaBuscada} (medida real: ${medidaReal} cm — descartado como merma)`, 'info');
        } else {
            log(`✏️ Corregido: ${tuboSeleccionado.cod} en ${colmenaBuscada}: ${tuboSeleccionado.medida_cm} → ${medidaReal} cm`, 'success');
        }

        // Guardar en Firebase
        colmenaActual = colmenas;
        SistemaInventario.colmenas = colmenas;

        const user = window.firebaseAuth.currentUser;
        if (user) {
            const datos = {
                data: JSON.stringify(colmenas),
                fechaActualizacion: new Date().toISOString(),
                ultimaCorreccion: `${colmenaBuscada} ${tuboSeleccionado.cod}: ${tuboSeleccionado.medida_cm}→${medidaReal}cm`
            };
            await window.firebaseSetDoc(
                window.firebaseDoc(window.firebaseDb, "usuarios", user.email, "colmena_final", "datos"),
                datos
            );
        }

        actualizarTablaColmenas();
        alert(`Corrección guardada exitosamente.\n${tuboSeleccionado.cod} en ${colmenaBuscada}: ${medidaReal} cm`);

    } catch (error) {
        console.error("Error corrigiendo corte:", error);
        alert("Error al guardar la corrección: " + error.message);
    } finally {
        document.getElementById('loading-overlay').style.display = 'none';
    }
}