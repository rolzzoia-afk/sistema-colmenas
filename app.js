const MM_TUBO_ORIGINAL = 5780;
const MM_KERF = 3;
const STOCK_MINIMO = 10; // Umbral de alerta: códigos con ≤ este número de tubos enteros disponibles

const SistemaInventario = {
    ordenes: [],
    colmenas: [],
    catalogoReemplazos: {},
    logs: [],
    colmenasDisponibles: [],
    datosCrudosOrdenes: [],
    resultadosOptimizacion: [],
    colmenasHistorico: [],
    mermas: [],
    seriales: []  // Nuevo array para almacenar los seriales disponibles
};

// Variables globales para manejo de colmena desde Firebase
let colmenaActual = null;        // Colmena descargada de Firebase al inicio de sesión
let usandoColmenaManual = false; // true si el usuario subió un archivo manualmente
let serialesDisponibles = [];    // Seriales disponibles cargados desde Firebase

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
// Verificar sesión activa
document.addEventListener("DOMContentLoaded", () => {

  const esperarFirebase = setInterval(() => {
    if (window.firebaseAuth) {
      clearInterval(esperarFirebase);

      window.firebaseOnAuth(window.firebaseAuth, async (user) => {
        if (!user) {
          mostrarLogin();
        } else {
          console.log("Usuario logueado:", user.email);

          // Cargar datos desde Firestore
          await cargarDesdeFirestore();
          actualizarTablaOrdenes();
          actualizarTablaColmenas();
          actualizarTablaCatalogo();
          
          // Cargar seriales desde Firestore
          await cargarSerialesDesdeFirestore();
          actualizarTablaSeriales();

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

// Función para cargar desde Firestore
async function cargarDesdeFirestore() {
    // Mostrar el overlay de carga
    document.getElementById('loading-overlay').style.display = 'flex';
    
    try {
        const user = window.firebaseAuth.currentUser;
        if (!user) {
            console.log("No hay usuario logueado para cargar desde Firestore");
            return;
        }

        const db = window.firebaseDb;

        // Crear la referencia al documento
        const docRef = window.firebaseDoc(db, "usuarios", user.email, "inventario", "datos");
        
        try {
            // Obtener el documento usando firebaseGetDoc (singular)
            const docSnap = await window.firebaseGetDoc(docRef);
            
            if (docSnap && docSnap.exists()) {
                const docData = docSnap.data();
                if (docData && docData.data) {
                    // Parsear el JSON guardado para restaurar los arrays originales
                    const datos = JSON.parse(docData.data);
                    
                    // Actualizar el objeto global con los datos recuperados
                    Object.assign(SistemaInventario, datos);
                    
                    console.log("✅ Inventario cargado correctamente desde Firestore");
                } else {
                    console.log("El documento existe pero no contiene datos válidos");
                }
            } else {
                console.log("No hay inventario guardado aún en Firestore");
            }
        } catch (docError) {
            console.error("Error específico al obtener el documento:", docError);
            
            // Intento alternativo usando la sintaxis antigua si la nueva falla
            try {
                const docSnap = await window.firebaseGetDoc(
                    window.firebaseDoc(db, "usuarios", user.email, "inventario", "datos")
                );
                
                if (docSnap && docSnap.exists()) {
                    const docData = docSnap.data();
                    if (docData && docData.data) {
                        // Parsear el JSON guardado para restaurar los arrays originales
                        const datos = JSON.parse(docData.data);
                        
                        // Actualizar el objeto global con los datos recuperados
                        Object.assign(SistemaInventario, datos);
                        
                        console.log("✅ Inventario cargado correctamente desde Firestore (método alternativo)");
                    }
                }
            } catch (altError) {
                console.error("Error en método alternativo:", altError);
                throw altError; // Re-lanzar para el catch exterior
            }
        }

        // Cargar colmena_final específica desde Firebase (no bloquea si falla)
        await cargarColmenaFinalDesdeFirestore();
    } catch (error) {
        console.error("Error cargando desde Firestore:", error);
        log("❌ Error cargando desde Firestore: " + error.message, "error");
    } finally {
        // Ocultar el overlay de carga independientemente del resultado
        document.getElementById('loading-overlay').style.display = 'none';
    }
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
    try {
        const user = window.firebaseAuth.currentUser;
        if (!user) return;

        const db = window.firebaseDb;
        const docRef = window.firebaseDoc(db, "usuarios", user.email, "colmena_final", "datos");
        const docSnap = await window.firebaseGetDoc(docRef);

        if (docSnap && docSnap.exists()) {
            const docData = docSnap.data();
            if (docData && docData.data) {
                colmenaActual = JSON.parse(docData.data);
                console.log(`✅ Colmena final cargada desde Firebase: ${colmenaActual.length} registros`);
                actualizarIndicadorFuente(false);
                verificarListo();
            }
        } else {
            // Sin colmena_final guardada: usar SistemaInventario.colmenas como fallback
            if (SistemaInventario.colmenas && SistemaInventario.colmenas.length > 0) {
                colmenaActual = SistemaInventario.colmenas;
                console.log(`ℹ️ Sin colmena_final en Firebase. Usando colmenas del inventario (${colmenaActual.length})`);
                actualizarIndicadorFuente(false);
                verificarListo();
            } else {
                console.log("ℹ️ No hay colmena_final ni colmenas en inventario aún.");
            }
        }
    } catch (error) {
        console.error("Error cargando colmena_final desde Firestore:", error);
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
            .filter(c => c.estado === 'disponible')
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
        SistemaInventario.logs = datos.logs || [];
        SistemaInventario.colmenasDisponibles = datos.colmenasDisponibles || [];
        SistemaInventario.datosCrudosOrdenes = datos.datosCrudosOrdenes || [];
        SistemaInventario.resultadosOptimizacion = datos.resultadosOptimizacion || [];
        SistemaInventario.colmenasHistorico = datos.colmenasHistorico || [];
        SistemaInventario.mermas = datos.mermas || [];
    }
}

// Cargar automáticamente al inicio
cargarSistema();

// ─── FUNCIONES DE SERIALES (Estructura de Inventario) ───────────────────────────

// Función para cargar el archivo de estructura de inventario
async function cargarEstructuraInventario(event) {
    const file = event.target.files[0];
    if (!file) return;
    
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

        // Usar la instancia db global
        const db = window.firebaseDb;
        
        if (!db) {
            throw new Error("Firestore no está inicializado");
        }
        
        log("Guardando seriales en Firestore...", "info");
        
       // Antes: const batch = writeBatch(window.firebaseDb);
const batch = window.firebaseWriteBatch(window.firebaseDb);
        
        // Ruta: usuarios/{email}/maestro_seriales
        const coleccionRef = window.firebaseCollection(db, "usuarios", user.email, "maestro_seriales");
        
        // Primero, eliminar los documentos existentes
        const docsExistentes = await window.firebaseGetDocs(coleccionRef);
        docsExistentes.forEach(doc => {
            batch.delete(doc.ref);
        });
        
        // Luego, agregar los nuevos documentos
        SistemaInventario.seriales.forEach(serial => {
            const docId = `${serial.codigo}_${serial.lote}_${serial.paquete}_${serial.serial}`;
            const docRef = window.firebaseDoc(coleccionRef, docId);
            batch.set(docRef, {
                ...serial,
                usuario: user.email,
                fechaActualizacion: new Date().toISOString()
            });
        });
        
        // Ejecutar el batch
        await batch.commit();
        
        console.log(`✅ ${SistemaInventario.seriales.length} seriales guardados en Firestore`);
        log(`✅ ${SistemaInventario.seriales.length} seriales guardados en Firestore`, "success");
        
    } catch (error) {
        console.error("Error guardando seriales en Firestore:", error);
        
        // Verificar errores de permisos
        if (error.code === 'permission-denied' || error.message.includes('permission')) {
            log("❌ Error de permisos: No tienes acceso a Firestore. Verifica las reglas de Firebase.", "error");
        } else {
            log("❌ Error guardando seriales en Firestore: " + error.message, "error");
        }
    } finally {
        document.getElementById('loading-overlay').style.display = 'none';
    }
}


// Función para cargar los seriales desde Firestore
async function cargarSerialesDesdeFirestore() {
    try {
        const user = window.firebaseAuth.currentUser;
        if (!user) return;

        const db = window.firebaseDb;
        
        try {
            const coleccionRef = window.firebaseCollection(db, "usuarios", user.email, "maestro_seriales");
            const querySnapshot = await window.firebaseGetDocs(coleccionRef);
            
            if (!querySnapshot.empty) {
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
                
                // Ordenar los seriales
                SistemaInventario.seriales.sort((a, b) => {
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
                
                console.log(`✅ ${SistemaInventario.seriales.length} seriales cargados desde Firestore`);
                
                // Actualizar la tabla, el estado y las alertas de stock
                actualizarTablaSeriales();
                verificarAlertasStock();
                const estadoEl = document.getElementById('estadoEstructura');
                if (estadoEl) {
                    estadoEl.textContent = `✓ ${SistemaInventario.seriales.length} seriales (sincronizados)`;
                    estadoEl.className = 'estado-archivo estado-ok';
                }
            } else {
                console.log("No hay seriales guardados en Firestore");
            }
        } catch (error) {
            console.error("Error cargando seriales desde Firestore:", error);
        }
    } catch (error) {
        console.error("Error general cargando seriales:", error);
    }
}

// Función para buscar un serial disponible para un código específico
function buscarSerialDisponible(codigo) {
    // Normalizar el código
    const codigoNormalizado = normalizarCodigo(codigo);
    
    // Filtrar seriales disponibles para este código
    const serialesDisponibles = SistemaInventario.seriales.filter(s => 
        normalizarCodigo(s.codigo) === codigoNormalizado && s.estado === 'disponible'
    );
    
    // Si no hay seriales disponibles, retornar null
    if (serialesDisponibles.length === 0) return null;
    
    // Los seriales ya están ordenados por FECHA, LOTE, PAQUETE, SERIAL
    // Retornar el primero (FIFO)
    return serialesDisponibles[0];
}

// Función para marcar un serial como ocupado
function marcarSerialComoOcupado(serial, ordenId) {
    // Buscar el serial en la lista
    const idx = SistemaInventario.seriales.findIndex(s => 
        s.codigo === serial.codigo && 
        s.lote === serial.lote && 
        s.paquete === serial.paquete && 
        s.serial === serial.serial
    );
    
    if (idx !== -1) {
        // Actualizar el estado
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
    for (let i = 0; i < Math.min(datos.length, 10); i++) {
        if (!datos[i]) continue;
        const tieneTexto = datos[i].some(cell => typeof cell === 'string' && cell.trim().length > 0);
        if (tieneTexto) return i;
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
    if (typeof valor === 'number') return Math.round(valor * 100) / 100;
    if (typeof valor === 'string') {
        const num = parseFloat(valor.replace(',', '.'));
        return isNaN(num) ? null : Math.round(num * 100) / 100;
    }
    return null;
}

  function parsearNumero(valor) {
      if (valor === null || valor === undefined || valor === '') return null;
      if (typeof valor === 'number') return valor;
      if (typeof valor === 'string') {
          const num = parseFloat(valor.replace(',', '.'));
          return isNaN(num) ? null : num;
      }
      return null;
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

              // Convertir Medida a número (float)
              let medidaNum = null;
              if (medidaRaw !== null && medidaRaw !== undefined && medidaRaw !== '') {
                  if (typeof medidaRaw === 'number') {
                      medidaNum = parseFloat(medidaRaw);
                  } else if (typeof medidaRaw === 'string') {
                      medidaNum = parseFloat(String(medidaRaw).replace(',', '.'));
                  }
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

              // Convertir Medida (cm) a número
              let medidaNum = null;
              if (medidaRaw !== null && medidaRaw !== undefined && medidaRaw !== '') {
                  if (typeof medidaRaw === 'number') {
                      medidaNum = parseFloat(medidaRaw);
                  } else if (typeof medidaRaw === 'string') {
                      medidaNum = parseFloat(String(medidaRaw).replace(',', '.'));
                  }
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

function detectarColumnasExcel(encabezados) {
    const mapeo = {};
    for (let i = 0; i < encabezados.length; i++) {
        const enc = String(encabezados[i] || '').trim().toLowerCase();
        for (const [nombre, config] of Object.entries(COLUMNAS_ESPECIALES)) {
            if (mapeo[nombre] !== undefined) continue;
            if (config.buscar && config.buscar.length > 0) {
                for (const busqueda of config.buscar) {
                    if (enc.includes(busqueda)) { mapeo[nombre] = i; break; }
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
                    if (enc.includes(busqueda)) { mapeo[key] = i; break; }
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

async function cargarOrdenes(event) {
    const file = event.target.files[0];
    if (!file) return;
    try {
        SistemaInventario.datosCrudosOrdenes = await leerExcelCompleto(file);
        const filaEncabezado = detectarFilaEncabezado(SistemaInventario.datosCrudosOrdenes);
        const encabezados = SistemaInventario.datosCrudosOrdenes[filaEncabezado];
        const mapeoColumnas = detectarColumnasExcel(encabezados);
        SistemaInventario.ordenes = [];
        for (let i = filaEncabezado + 1; i < SistemaInventario.datosCrudosOrdenes.length; i++) {
            const fila = SistemaInventario.datosCrudosOrdenes[i];
            if (!fila) continue;
            const colMedida = mapeoColumnas['medida'];
            const valor = colMedida !== undefined ? fila[colMedida] : null;
            const medidaNum = parsearNumero(valor);
            if (medidaNum !== null && medidaNum > 0) {
                const orden = { id: SistemaInventario.ordenes.length + 1, medida_mm: Math.round(medidaNum * 10), medida_cm: formatearNumero(medidaNum) };
                for (const [nombre, config] of Object.entries(COLUMNAS_ESPECIALES)) {
                    const idx = mapeoColumnas[nombre];
                    if (idx !== undefined && idx !== null) {
                        let valorCelda = fila[idx];
                        if (config.esNumero) valorCelda = formatearNumero(valorCelda);
                        orden[nombre] = valorCelda;
                    }
                }
                for (const colEsp of COLUMNAS_ESPECIFICAS) {
                    const idx = mapeoColumnas[colEsp.key];
                    if (idx !== undefined && idx !== null) orden[colEsp.key] = fila[idx];
                }
                if (orden.tuberia) {
                    orden.codigoExtraido = extraerCodigoDesdeTuberia(orden.tuberia);
                    orden.reemplazo = SistemaInventario.catalogoReemplazos[orden.codigoExtraido] || null;
                    orden.cod = orden.codigoExtraido || orden.codSec || 'TUBO-' + orden.id;
                } else {
                    orden.cod = orden.codSec || 'TUBO-' + orden.id;
                }
                SistemaInventario.ordenes.push(orden);
            }
        }
        actualizarTablaOrdenes();
        document.getElementById('estadoOrdenes').textContent = `✓ ${SistemaInventario.ordenes.length} órdenes`;
        document.getElementById('estadoOrdenes').className = 'estado-archivo estado-ok';
        verificarListo();
        log(`📋 Órdenes cargadas: ${SistemaInventario.ordenes.length}`, 'success');
        guardarSistema();
        guardarEnFirestore();
    } catch (e) { alert('Error: ' + e.message); console.error(e); }
}

async function cargarColmenas(event) {
    const file = event.target.files[0];
    if (!file) return;
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
            const medidaNum = parsearNumero(medida);
            if (medidaNum !== null && medidaNum > 0) {
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
        guardarSistema();
    } catch (e) { alert('Error: ' + e.message); }
}

async function cargarCatalogo(event) {
    const file = event.target.files[0];
    if (!file) return;
    try {
        const datos = await leerExcelCompleto(file);
        const filaEncabezado = detectarFilaEncabezado(datos);
        const encabezados = datos[filaEncabezado];
        let colCodigo = -1, colReemplazo = -1;
        for (let i = 0; i < encabezados.length; i++) {
            const enc = String(encabezados[i] || '').trim().toUpperCase();
            if (enc.includes('CODIGO') || enc === 'COD') colCodigo = i;
            if (enc.includes('REEMPLAZ')) colReemplazo = i;
        }
        if (colCodigo === -1 || colReemplazo === -1) { alert('No se encontraron columnas CODIGO y REEMPLAZO'); return; }
        SistemaInventario.catalogoReemplazos = {};
        for (let i = filaEncabezado + 1; i < datos.length; i++) {
            const fila = datos[i];
            if (!fila) continue;
            const codigo = fila[colCodigo];
            const reemplazo = fila[colReemplazo];
            if (codigo && reemplazo) {
                const codigoLimpio = String(codigo).trim().toUpperCase();
                SistemaInventario.catalogoReemplazos[codigoLimpio] = String(reemplazo).trim();
            }
        }
        actualizarReemplazosEnOrdenes();
        actualizarTablaCatalogo();
        log(`✓ Catálogo cargado: ${Object.keys(SistemaInventario.catalogoReemplazos).length} reemplazos`, 'success');
        document.getElementById('estadoCatalogo').textContent = `✓ ${Object.keys(SistemaInventario.catalogoReemplazos).length} reemplazos`;
        document.getElementById('estadoCatalogo').className = 'estado-archivo estado-ok';
        guardarSistema();
    } catch (e) { alert('Error: ' + e.message); }
}

function formatearValor(valor) {
    if (valor === null || valor === undefined) return '';
    if (typeof valor === 'number') { if (valor === 0) return ''; if (Number.isInteger(valor)) return valor; return Math.round(valor * 100) / 100; }
    return valor;
}

function actualizarTablaOrdenes() {
    const columnasVisibles = detectarColumnasConDatos(SistemaInventario.ordenes);
    const headerRow = document.getElementById('headerOrdenes');
    headerRow.innerHTML = columnasVisibles.map(col => `<th>${col.titulo}</th>`).join('');
    const tbody = document.getElementById('tbodyOrdenes');
    tbody.innerHTML = SistemaInventario.ordenes.map(orden => {
        const celdas = columnasVisibles.map(col => `<td>${formatearValor(orden[col.key])}</td>`).join('');
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
    document.getElementById('btnEjecutar').disabled = !(tieneOrdenes && tieneColmenas);
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
    paso.style.cssText = 'background:#ecf0f1; padding:8px; margin:3px 0; border-radius:3px; border-left:3px solid #3498db; font-size:12px;';
    paso.innerHTML = `<strong>${numero}. ${titulo}</strong> - ${descripcion}`;
    procesoDiv.appendChild(paso);
}

function evaluarSobrante(sobrante) {
    if (sobrante < 0) return { estado: 'prohibido' };
    if (sobrante <= 100) return { estado: 'merma' };
    if (sobrante > 100 && sobrante < 1300) return { estado: 'prohibido' };
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
    
    if (codigoOriginal) {
        const codOrigNormalizado = normalizarCodigo(codigoOriginal);
        for (let i = 0; i < SistemaInventario.colmenasDisponibles.length; i++) {
            const col = SistemaInventario.colmenasDisponibles[i];
            if (normalizarCodigo(col.cod) === codOrigNormalizado && col.medida_mm >= medidaRequerida) {
                const sobrante = col.medida_mm - medidaRequerida - MM_KERF;
                const clasificacion = evaluarSobrante(sobrante);
                if (clasificacion.estado !== 'prohibido') {
                    return { colmena: col, sobrante_mm: sobrante, indice: i, clasificacion: clasificacion, medidaOriginal: col.medida_mm, esReemplazo: false };
                }
            }
        }
    }
    
    for (let i = 0; i < SistemaInventario.colmenasDisponibles.length; i++) {
        const col = SistemaInventario.colmenasDisponibles[i];
        if (normalizarCodigo(col.cod) === codNormalizado && col.medida_mm >= medidaRequerida) {
            const sobrante = col.medida_mm - medidaRequerida - MM_KERF;
            const clasificacion = evaluarSobrante(sobrante);
            if (clasificacion.estado !== 'prohibido') {
                const esReemp = codigoOriginal ? codNormalizado !== normalizarCodigo(codigoOriginal) : false;
                return { colmena: col, sobrante_mm: sobrante, indice: i, clasificacion: clasificacion, medidaOriginal: col.medida_mm, esReemplazo: esReemp };
            }
        }
    }
    return null;
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
        return `<tr>
            <td>${r.colmena || 'TUBO NUEVO'}</td>
            <td>${r.codigo || '-'}</td>
            <td>${r.medida_cm} cm</td>
            <td><span class="tag-accion">CORTAR</span></td>
        </tr>`;
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
        <div style="background:#fff3cd; border:1px solid #ffc107; border-left:4px solid #f39c12; border-radius:6px; padding:12px 16px; margin:10px 0;">
            <strong style="color:#856404; font-size:13px;">⚠️ ALERTA DE STOCK MÍNIMO — Materiales con ≤${STOCK_MINIMO} tubos disponibles:</strong>
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

// Función para exportar resultados a Excel usando SheetJS (XLSX)
function exportarResultadosExcel(resultados) {
    if (!resultados || resultados.length === 0) {
        const mensajeError = "❌ Error: No hay resultados para exportar";
        console.error(mensajeError);
        log(mensajeError, 'error');
        alert("No hay resultados para exportar. Ejecute la optimización primero.");
        return;
    }

    console.log("📊 Exportando resultados a Excel...", resultados);

    // Crear un nuevo libro de trabajo
    const wb = XLSX.utils.book_new();

    // Convertir el array de resultados en formato para SheetJS
    const datosParaExcel = resultados.map(r => {
        return {
            'Pedido': r.orden && r.orden.Pedido ? r.orden.Pedido : (r.orden && r.orden.id ? r.orden.id : 'N/A'),
            'Colmena': r.colmena && r.colmena['N° Colmena'] ? r.colmena['N° Colmena'] : (r.colmena && r.colmena.n_colmena ? r.colmena.n_colmena : 'N/A'),
            'Medida Orden (cm)': r.medidaOrden || 0,
            'Medida Colmena (cm)': r.medidaColmena || 0,
            'Merma (cm)': r.asignada ? (r.merma ? r.merma.toFixed(2) : '0.00') : 'No Asignada',
            'Estado': r.asignada ? 'Asignada' : 'No Asignada'
        };
    });

    // Crear la hoja con los datos
    const ws = XLSX.utils.json_to_sheet(datosParaExcel);

    // Dar formato a las columnas para que sean legibles
    const colWidths = [
        { wch: 20 },  // Pedido
        { wch: 15 },  // Colmena
        { wch: 18 },  // Medida Orden
        { wch: 18 },  // Medida Colmena
        { wch: 15 },  // Merma
        { wch: 15 }   // Estado
    ];
    ws['!cols'] = colWidths;

    // Agregar la hoja al libro
    XLSX.utils.book_append_sheet(wb, ws, 'Plan de Optimización');

    // Generar el nombre del archivo con la fecha actual
    const fechaActual = new Date();
    const año = fechaActual.getFullYear();
    const mes = String(fechaActual.getMonth() + 1).padStart(2, '0');
    const dia = String(fechaActual.getDate()).padStart(2, '0');
    const fechaStr = `${año}${mes}${dia}`;
    
    const nombreArchivo = `Optimizacion_Rolzzo_${fechaStr}.xlsx`;

    // Generar y descargar el archivo
    XLSX.writeFile(wb, nombreArchivo);

    // Añadir log de éxito
    const mensajeExito = `✅ Archivo ${nombreArchivo} descargado exitosamente`;
    console.log(mensajeExito);
    log(mensajeExito, 'success');
    alert(`Resultados exportados exitosamente a: ${nombreArchivo}`);
}

function ejecutarOptimizacion() {
    // Determinar fuente de colmenas: archivo manual o colmena sincronizada de Firebase
    const colmenasAUsar = usandoColmenaManual
        ? SistemaInventario.colmenas
        : (colmenaActual && colmenaActual.length > 0 ? colmenaActual : SistemaInventario.colmenas);

    if (SistemaInventario.ordenes.length === 0 || colmenasAUsar.length === 0) { alert('Cargue órdenes y colmenas'); return; }

    log(`ℹ️ Fuente de colmenas: ${usandoColmenaManual ? 'Archivo manual' : 'Firebase (sincronizado)'}`, 'info');

    document.getElementById('logs').innerHTML = '';
    document.getElementById('proceso').innerHTML = '';
    document.getElementById('resultados').innerHTML = '';
    SistemaInventario.logs = [];
    SistemaInventario.colmenasDisponibles = JSON.parse(JSON.stringify(colmenasAUsar));
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
                serial: tuboEncontrado.colmena.serial || orden.serial || null
            };
            const idxHistorico = SistemaInventario.colmenasHistorico.findIndex(c => c.n_colmena === tuboEncontrado.colmena.n_colmena);
            if (idxHistorico !== -1) { SistemaInventario.colmenasHistorico[idxHistorico].estado = 'usada'; SistemaInventario.colmenasHistorico[idxHistorico].origen = 'Orden ' + orden.id; }
            SistemaInventario.colmenasDisponibles.splice(tuboEncontrado.indice, 1);
        } else if (tuboEncontrado) {
            const esMerma = tuboEncontrado.clasificacion && tuboEncontrado.clasificacion.estado === 'merma';
            const esReemplazo = tuboEncontrado.esReemplazo;
            let fuente = esMerma ? 'merma' : (esReemplazo ? 'reemplazo' : 'colmena');

            resultado = { orden: orden.id, medida_cm: orden.medida_cm, fuente: fuente, colmena: tuboEncontrado.colmena.n_colmena, codigo: tuboEncontrado.colmena.cod, codigo_original: codOrden, sobrante_cm: tuboEncontrado.sobrante_mm / 10, serial: tuboEncontrado.colmena.serial || null };
            const idxHistorico = SistemaInventario.colmenasHistorico.findIndex(c => c.n_colmena === tuboEncontrado.colmena.n_colmena);
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
                SistemaInventario.colmenasDisponibles.push({
                    n_colmena: tuboEncontrado.colmena.n_colmena,
                    medida_mm: tuboEncontrado.colmena.medida_mm,
                    medida_cm: tuboEncontrado.colmena.medida_cm,
                    cod: tuboEncontrado.colmena.cod
                });
                if (idxHistorico !== -1) {
                    SistemaInventario.colmenasHistorico[idxHistorico].estado = 'disponible';
                    SistemaInventario.colmenasHistorico[idxHistorico].origen = 'Original (reusado por merma)';
                }
            } else if (clasificacion.estado !== 'prohibido') {
                if (idxHistorico !== -1) { SistemaInventario.colmenasHistorico[idxHistorico].estado = 'usada'; SistemaInventario.colmenasHistorico[idxHistorico].origen = 'Orden ' + orden.id; }
                
                const idxInsertar = SistemaInventario.colmenasHistorico.findIndex(c => c.n_colmena === tuboEncontrado.colmena.n_colmena);
                if (idxInsertar !== -1) {
                SistemaInventario.colmenasHistorico.splice(idxInsertar + 1, 0, { 
                    n_colmena: tuboEncontrado.colmena.n_colmena, 
                    medida_cm: sobrante / 10, 
                    medida_mm: sobrante, 
                    cod: tuboEncontrado.colmena.cod, 
                    codigo_original: tuboEncontrado.colmena.cod, 
                    estado: 'disponible', 
                    origen: 'Sobrante orden ' + orden.id, 
                    posicionOriginal: tuboEncontrado.colmena.n_colmena,
                    serial: tuboEncontrado.colmena.serial || orden.serial || null,
                    fecha: tuboEncontrado.colmena.serial ? tuboEncontrado.colmena.serial.fecha : null
                });
                }
                SistemaInventario.colmenasDisponibles.push({
                    n_colmena: tuboEncontrado.colmena.n_colmena,
                    medida_mm: sobrante,
                    medida_cm: sobrante / 10,
                    cod: tuboEncontrado.colmena.cod,
                    serial: tuboEncontrado.colmena.serial || orden.serial || null
                });
                SistemaInventario.colmenasDisponibles.splice(tuboEncontrado.indice, 1);
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
                            codigo_original: tuboReemplazo.colmena.cod, 
                            codigo_reemplazo: codReemplazo, 
                            sobrante_cm: tuboReemplazo.sobrante_mm / 10,
                            serial: tuboReemplazo.colmena.serial || null
                        };
                        const idxHistorico = SistemaInventario.colmenasHistorico.findIndex(c => c.n_colmena === tuboReemplazo.colmena.n_colmena);
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

                            SistemaInventario.colmenasDisponibles.push({
                                n_colmena: tuboReemplazo.colmena.n_colmena,
                                medida_mm: tuboReemplazo.colmena.medida_mm,
                                medida_cm: tuboReemplazo.colmena.medida_cm,
                                cod: tuboReemplazo.colmena.cod
                            });
                            if (idxHistorico !== -1) {
                                SistemaInventario.colmenasHistorico[idxHistorico].estado = 'disponible';
                                SistemaInventario.colmenasHistorico[idxHistorico].origen = 'Original (reusado por merma)';
                            }
                        } else if (clasificacion.estado !== 'prohibido') {
                            SistemaInventario.colmenasHistorico[idxHistorico].codigo_original = tuboReemplazo.colmena.cod;
                            if (idxHistorico !== -1) { SistemaInventario.colmenasHistorico[idxHistorico].estado = 'usada'; SistemaInventario.colmenasHistorico[idxHistorico].origen = 'Orden ' + orden.id + ' (Reemplazo ' + codReemplazo + ')';  }
                            
                            const colmenaExistente = buscarColmenaDisponibleConCodigo(codReemplazo);
                            
                            if (colmenaExistente) {
                                log(`📦 Sobrante agregado como nueva fila para colmena ${colmenaExistente.n_colmena} (código ${codReemplazo}). Medida sobrante: ${sobrante / 10}cm`, 'info');
                                const idxHistoricoExistente = SistemaInventario.colmenasHistorico.findIndex(c => c.n_colmena === colmenaExistente.n_colmena);
                                if (idxHistoricoExistente !== -1) {
                                    SistemaInventario.colmenasHistorico.splice(idxHistoricoExistente + 1, 0, { n_colmena: colmenaExistente.n_colmena, medida_cm: sobrante / 10, medida_mm: sobrante, cod: codReemplazo, codigo_original: tuboReemplazo.colmena.cod, estado: 'disponible', origen: 'Sobrante reemplazo orden ' + orden.id, posicionOriginal: colmenaExistente.n_colmena });
                                }
                                SistemaInventario.colmenasDisponibles.push({
                                    n_colmena: colmenaExistente.n_colmena,
                                    medida_mm: sobrante,
                                    medida_cm: sobrante / 10,
                                    cod: codReemplazo
                                });
                            } else {
                                const idxInsertar = SistemaInventario.colmenasHistorico.findIndex(c => c.n_colmena === tuboReemplazo.colmena.n_colmena);
                                if (idxInsertar !== -1) {
                                    SistemaInventario.colmenasHistorico.splice(idxInsertar + 1, 0, { n_colmena: tuboReemplazo.colmena.n_colmena, medida_cm: sobrante / 10, medida_mm: sobrante, cod: tuboReemplazo.colmena.cod, codigo_original: tuboReemplazo.colmena.cod, estado: 'disponible', origen: 'Sobrante reemplazo orden ' + orden.id, posicionOriginal: tuboReemplazo.colmena.n_colmena });
                                }
                                SistemaInventario.colmenasDisponibles.push({
                                    n_colmena: tuboReemplazo.colmena.n_colmena,
                                    medida_mm: sobrante,
                                    medida_cm: sobrante / 10,
                                    cod: codReemplazo
                                });
                            }
                            if (!esMerma) {
                                SistemaInventario.colmenasDisponibles.splice(tuboReemplazo.indice, 1);
                            }
                        }
                        break;
                    }
                }
            }
            if (!resultado) {
                // Buscar un serial disponible para este código
                const serialDisponible = buscarSerialDisponible(codOrden);
                
                const sobranteNuevo = MM_TUBO_ORIGINAL - orden.medida_mm - MM_KERF;
                
                // Calcular posición nueva ANTES de crear el resultado para poder incluirla en res.colmena
                const posicionesOcupadas = new Set(SistemaInventario.colmenasHistorico.map(c => c.n_colmena));
                let posicionNueva = null;
                for (let i = 1; i <= colmenasAUsar.length + 100; i++) {
                    const pos = 'A' + i;
                    if (!posicionesOcupadas.has(pos)) { posicionNueva = pos; break; }
                }
                
                // Crear el resultado con información del serial y colmena asignada
                if (serialDisponible) {
                    resultado = { 
                        orden: orden.id, 
                        medida_cm: orden.medida_cm, 
                        fuente: 'tubo_nuevo',
                        colmena: posicionNueva,
                        codigo_original: codOrden, 
                        sobrante_cm: sobranteNuevo / 10,
                        serial: serialDisponible
                    };
                    
                    // Marcar el serial como ocupado en la lista local
                    marcarSerialComoOcupado(serialDisponible, orden.id);
                    
                    log(`🏷️ Serial asignado: ${serialDisponible.codigo} - Lote: ${serialDisponible.lote} - Paquete: ${serialDisponible.paquete} - Serial: ${serialDisponible.serial}`, 'info');
                } else {
                    resultado = { 
                        orden: orden.id, 
                        medida_cm: orden.medida_cm, 
                        fuente: 'tubo_nuevo',
                        colmena: posicionNueva,
                        codigo_original: codOrden, 
                        sobrante_cm: sobranteNuevo / 10
                    };
                    
                    log(`⚠️ No se encontró serial disponible para el código ${codOrden}`, 'warn');
                }
                
                const codigoTuboNuevo = codOrden || 'TUBO-NUEVO';
                
                // Añadir información del serial al histórico si está disponible
                const infoSerial = serialDisponible ? 
                    ` (Serial: ${serialDisponible.lote}-${serialDisponible.paquete}-${serialDisponible.serial})` : '';
                
                SistemaInventario.colmenasHistorico.push({ 
                    n_colmena: posicionNueva, 
                    medida_cm: MM_TUBO_ORIGINAL / 10, 
                    medida_mm: MM_TUBO_ORIGINAL, 
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
        
        resultados.push(resultado);
        SistemaInventario.resultadosOptimizacion.push({ orden: orden, resultado: resultado });
        const infoResultado = formatearResultado(orden, resultado);
        agregarPaso(numPaso, `Orden ${orden.id}: ${orden.medida_cm}cm`, infoResultado);
    });

    actualizarTablaColmenasResultado();
    actualizarTablaMermas();

    document.getElementById('btnExportarDisponibles').disabled = false;
    log('=== COMPLETADO ===', 'success');
    
    let html = '<div class="resumen-resultado"><h3>✓ Optimización Completada</h3>';
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
    document.getElementById('btnExportar').disabled = false;
    guardarSistema();

    // Guardar colmena final en Firebase para que sea la base del día siguiente
    log('💾 Guardando colmena final en Firebase...', 'info');
    guardarColmenaFinalEnFirestore();

    // Actualizar alertas de stock en tiempo real tras consumir tubos
    verificarAlertasStock();
}

function exportarResultados() {
    if (SistemaInventario.resultadosOptimizacion.length === 0) return alert('No hay resultados');
    const datosExcel = [['OT', 'Ubicación', 'Acción', 'Colmena', 'Código', 'Medida (cm)', 'Lote', 'Paquete', 'Serial', 'Fecha Serial']];
    
    SistemaInventario.resultadosOptimizacion.forEach(item => {
        const res = item.resultado;
        const ord = SistemaInventario.ordenes.find(o => o.id === res.orden) || {};
        
        // Recuperar serial si existe en la orden o en el resultado
        const s = res.serial || ord.serial || {};
        const fechaFormateada = s.fecha ? formatearFecha(s.fecha) : '-';
        
        // Fila CORTAR
        datosExcel.push([
            ord.ot || '-', 
            ord.ubic || '-', 
            'CORTAR', 
            res.fuente === 'tubo_nuevo' ? 'TUBO NUEVO' : (res.colmena || '-'), 
            res.codigo || ord.cod || '-', 
            res.medida_cm,
            s.lote || '-', 
            s.paquete || '-', 
            s.serial || '-', 
            fechaFormateada
        ]);
        
        // Si existe un sobrante (> 0), insertar una fila de "GUARDAR SOBRANTE" inmediatamente después
        if (res.sobrante_cm > 0) {
            datosExcel.push([
                ord.ot || '-', 
                '', 
                'GUARDAR SOBRANTE', 
                res.colmena || '-', 
                res.codigo || ord.cod || '-', 
                res.sobrante_cm,
                s.lote || ord.lote || '-', 
                s.paquete || ord.paquete || '-', 
                s.serial || ord.serial || '-', 
                s.fecha || ord.fecha || '-'
            ]);
        }
    });
    
    const ws = XLSX.utils.aoa_to_sheet(datosExcel);
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
    const filas = SistemaInventario.colmenasHistorico
        .filter(c => c.estado === 'disponible')
        .map(c => {
            // Recuperar información del serial si existe
            const s = c.serial || {};
            return [
                c.n_colmena || '-',
                c.cod || '-',
                c.medida_cm || 0,
                c.estado,
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
