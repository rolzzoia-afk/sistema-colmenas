const MM_TUBO_ORIGINAL = 5780;
const MM_KERF = 3;

const SistemaInventario = {
    ordenes: [],
    colmenas: [],
    catalogoReemplazos: {},
    logs: [],
    colmenasDisponibles: [],
    datosCrudosOrdenes: [],
    resultadosOptimizacion: [],
    colmenasHistorico: [],
    mermas: []
};
// Verificar sesión activa
document.addEventListener("DOMContentLoaded", () => {

  const esperarFirebase = setInterval(() => {
    if (window.firebaseAuth) {
      clearInterval(esperarFirebase);

      window.firebaseOnAuth(window.firebaseAuth, (user) => {
        if (!user) {
          mostrarLogin();
        } else {
          console.log("Usuario logueado:", user.email);

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

function verificarListo() { document.getElementById('btnEjecutar').disabled = !(SistemaInventario.ordenes.length > 0 && SistemaInventario.colmenas.length > 0); }

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
    if (SistemaInventario.colmenasHistorico.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6">Sin datos</td></tr>';
        return;
    }
    const colmenasOrdenadas = [...SistemaInventario.colmenasHistorico].sort((a, b) => {
        const cmpColmena = ordenarColmena(a.n_colmena, b.n_colmena);
        if (cmpColmena !== 0) return cmpColmena;
        return String(a.cod || '').localeCompare(String(b.cod || ''));
    });
    tbody.innerHTML = colmenasOrdenadas.map(c => {
        let estadoBadge = '';
        const esReemplazo = c.codigo_original && c.cod && normalizarCodigo(c.cod) !== normalizarCodigo(c.codigo_original);
        const esLeftover = c.n_colmena && c.n_colmena.includes('-S');
        
        switch(c.estado) {
            case 'usada': 
                estadoBadge = esReemplazo 
                    ? '<span style="background-color: #3498db; color: white; padding: 2px 6px; border-radius: 3px; font-size: 10px;">REEMPLAZO</span>' 
                    : '<span style="background-color: #27ae60; color: white; padding: 2px 6px; border-radius: 3px; font-size: 10px;">USADA</span>'; 
                break;
            case 'disponible': 
                estadoBadge = esReemplazo 
                    ? '<span style="background-color: #3498db; color: white; padding: 2px 6px; border-radius: 3px; font-size: 10px;">REEMPLAZO</span>' 
                    : '<span style="background-color: #3498db; color: white; padding: 2px 6px; border-radius: 3px; font-size: 10px;">DISPONIBLE</span>'; 
                break;
            case 'nueva': estadoBadge = '<span style="background-color: #f39c12; color: white; padding: 2px 6px; border-radius: 3px; font-size: 10px;">NUEVA</span>'; break;
            case 'merma': estadoBadge = '<span style="background-color: #e67e22; color: white; padding: 2px 6px; border-radius: 3px; font-size: 10px;">MERMA</span>'; break;
            default: estadoBadge = c.estado || '';
        }
        
        let agrupacionBadge = '';
        if (esLeftover) {
            agrupacionBadge = '<span style="background-color: #9b59b6; color: white; padding: 1px 4px; border-radius: 2px; font-size: 9px; margin-left: 4px;">SOBRANTE</span>';
        }
        
        return `<tr${esLeftover ? ' style="background-color: #f8f9fa;"' : ''}><td>${c.n_colmena}</td><td>${c.cod}${agrupacionBadge}</td><td>${c.codigo_original || ''}</td><td>${formatearValor(c.medida_cm)}</td><td>${estadoBadge}</td><td>${c.origen}</td></tr>`;
    }).join('');
}

function actualizarTablaMermas() {
    const tbody = document.getElementById('tbodyMermas');
    if (SistemaInventario.mermas.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7">Sin mermas</td></tr>';
        return;
    }
    tbody.innerHTML = SistemaInventario.mermas.map(m => {
        let tipoBadge = '';
        if (m.tipo === 'MERMA') {
            tipoBadge = '<span style="background-color: #e67e22; color: white; padding: 2px 6px; border-radius: 3px; font-size: 10px;">MERMA</span>';
        }
        return `<tr><td>${m.orden}</td><td>${m.espacioOriginal}</td><td>${m.codigoOriginal}</td><td>${formatearValor(m.medidaRequerida)}</td><td>${tipoBadge}</td><td>${formatearValor(m.valor)} cm</td><td>${m.codigoUsado}</td></tr>`;
    }).join('');
}

function ejecutarOptimizacion() {
    if (SistemaInventario.ordenes.length === 0 || SistemaInventario.colmenas.length === 0) { alert('Cargue órdenes y colmenas'); return; }

    document.getElementById('logs').innerHTML = '';
    document.getElementById('proceso').innerHTML = '';
    document.getElementById('resultados').innerHTML = '';
    SistemaInventario.logs = [];
    SistemaInventario.colmenasDisponibles = JSON.parse(JSON.stringify(SistemaInventario.colmenas));
    SistemaInventario.colmenasHistorico = [];
    SistemaInventario.resultadosOptimizacion = [];
    SistemaInventario.mermas = [];

    SistemaInventario.colmenas.forEach((col) => {
        SistemaInventario.colmenasHistorico.push({ n_colmena: col.n_colmena, medida_cm: col.medida_cm, medida_mm: col.medida_mm, cod: col.cod, codigo_original: col.cod, estado: 'disponible', origen: 'Original', posicionOriginal: col.n_colmena });
    });

    log('=== INICIO OPTIMIZACIÓN ===', 'info');
    const resultados = [];

    SistemaInventario.ordenes.forEach((orden, idx) => {
        const numPaso = idx + 1;
        const codOrden = orden.cod;
        let resultado = null;
        
        const tuboEncontrado = buscarTubosParaOrden(codOrden, orden.medida_mm, codOrden);
        
        if (tuboEncontrado && tuboEncontrado.sobrante_mm === 0) {
            resultado = { orden: orden.id, medida_cm: orden.medida_cm, fuente: 'exacta', colmena: tuboEncontrado.colmena.n_colmena, codigo: tuboEncontrado.colmena.cod, sobrante_cm: 0 };
            const idxHistorico = SistemaInventario.colmenasHistorico.findIndex(c => c.n_colmena === tuboEncontrado.colmena.n_colmena);
            if (idxHistorico !== -1) { SistemaInventario.colmenasHistorico[idxHistorico].estado = 'usada'; SistemaInventario.colmenasHistorico[idxHistorico].origen = 'Orden ' + orden.id; }
            SistemaInventario.colmenasDisponibles.splice(tuboEncontrado.indice, 1);
        } else if (tuboEncontrado) {
            const esMerma = tuboEncontrado.clasificacion && tuboEncontrado.clasificacion.estado === 'merma';
            const esReemplazo = tuboEncontrado.esReemplazo;
            let fuente = esMerma ? 'merma' : (esReemplazo ? 'reemplazo' : 'colmena');

            resultado = { orden: orden.id, medida_cm: orden.medida_cm, fuente: fuente, colmena: tuboEncontrado.colmena.n_colmena, codigo: tuboEncontrado.colmena.cod, codigo_original: codOrden, sobrante_cm: tuboEncontrado.sobrante_mm / 10 };
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
                    SistemaInventario.colmenasHistorico.splice(idxInsertar + 1, 0, { n_colmena: tuboEncontrado.colmena.n_colmena, medida_cm: sobrante / 10, medida_mm: sobrante, cod: tuboEncontrado.colmena.cod, codigo_original: tuboEncontrado.colmena.cod, estado: 'disponible', origen: 'Sobrante orden ' + orden.id, posicionOriginal: tuboEncontrado.colmena.n_colmena });
                }
                SistemaInventario.colmenasDisponibles.push({
                    n_colmena: tuboEncontrado.colmena.n_colmena,
                    medida_mm: sobrante,
                    medida_cm: sobrante / 10,
                    cod: tuboEncontrado.colmena.cod
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
                            sobrante_cm: tuboReemplazo.sobrante_mm / 10 
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
                const sobranteNuevo = MM_TUBO_ORIGINAL - orden.medida_mm - MM_KERF;
                resultado = { orden: orden.id, medida_cm: orden.medida_cm, fuente: 'tubo_nuevo', codigo_original: codOrden, sobrante_cm: sobranteNuevo / 10 };
                const posicionesOcupadas = new Set(SistemaInventario.colmenasHistorico.map(c => c.n_colmena));
                let posicionNueva = null;
                for (let i = 1; i <= SistemaInventario.colmenas.length + 100; i++) {
                    const pos = 'A' + i;
                    if (!posicionesOcupadas.has(pos)) { posicionNueva = pos; break; }
                }
                const codigoTuboNuevo = codOrden || 'TUBO-NUEVO';
                SistemaInventario.colmenasHistorico.push({ n_colmena: posicionNueva, medida_cm: MM_TUBO_ORIGINAL / 10, medida_mm: MM_TUBO_ORIGINAL, cod: codigoTuboNuevo, codigo_original: codOrden, estado: 'usada', origen: 'Orden ' + orden.id + ' (Tubo nuevo)', posicionOriginal: posicionNueva });
                const clasificacion = evaluarSobrante(sobranteNuevo);
                if (clasificacion.estado !== 'prohibido' && clasificacion.estado !== 'merma') {
                    const colmenaExistente = buscarColmenaDisponibleConCodigo(codigoTuboNuevo);
                    if (colmenaExistente) {
                        log(`📦 Sobrante agregado como nueva fila para colmena ${colmenaExistente.n_colmena} (código ${codigoTuboNuevo}). Medida sobrante: ${sobranteNuevo / 10}cm`, 'info');
                        const idxHistoricoExistente = SistemaInventario.colmenasHistorico.findIndex(c => c.n_colmena === colmenaExistente.n_colmena);
                        if (idxHistoricoExistente !== -1) {
                            SistemaInventario.colmenasHistorico.splice(idxHistoricoExistente + 1, 0, { n_colmena: colmenaExistente.n_colmena, medida_cm: sobranteNuevo / 10, medida_mm: sobranteNuevo, cod: codigoTuboNuevo, codigo_original: codigoTuboNuevo, estado: 'disponible', origen: 'Sobrante tubo nuevo orden ' + orden.id, posicionOriginal: colmenaExistente.n_colmena });
                        }
                        SistemaInventario.colmenasDisponibles.push({
                            n_colmena: colmenaExistente.n_colmena,
                            medida_mm: sobranteNuevo,
                            medida_cm: sobranteNuevo / 10,
                            cod: codigoTuboNuevo
                        });
                    } else {
                        SistemaInventario.colmenasHistorico.push({ n_colmena: posicionNueva, medida_cm: sobranteNuevo / 10, medida_mm: sobranteNuevo, cod: codigoTuboNuevo, codigo_original: codigoTuboNuevo, estado: 'disponible', origen: 'Sobrante tubo nuevo orden ' + orden.id, posicionOriginal: posicionNueva });
                        SistemaInventario.colmenasDisponibles.push({
                            n_colmena: posicionNueva,
                            medida_mm: sobranteNuevo,
                            medida_cm: sobranteNuevo / 10,
                            cod: codigoTuboNuevo
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
}

function exportarResultados() {
    if (SistemaInventario.resultadosOptimizacion.length === 0) {
        alert('No hay resultados para exportar.');
        return;
    }

    const datosExcel = [];
    datosExcel.push(['OT', 'Ubicación', 'Acción', 'Colmena', 'Código', 'Medida (cm)']);

    SistemaInventario.resultadosOptimizacion.forEach(item => {
        const orden = item.orden;
        const resultado = item.resultado;
        const ordenOriginal = SistemaInventario.ordenes[resultado.orden - 1] || {};

        const ot = ordenOriginal.ot || '';
        const ubic = ordenOriginal.ubic || '';
        const codigoUsado = resultado.codigo || resultado.codigo_original || '';
        const textoColmena = resultado.colmena && resultado.colmena !== '' ? resultado.colmena : 'TUBO NUEVO';
        const medida = resultado.medida_cm;

        datosExcel.push([ot, ubic, 'CORTAR', textoColmena, codigoUsado, medida]);

        if (resultado.sobrante_cm && resultado.sobrante_cm > 0) {
            const sobranteEncontrado = SistemaInventario.colmenasHistorico.find(c =>
                c.estado === 'disponible' &&
                c.origen &&
                c.origen.toLowerCase().includes('orden ' + resultado.orden)
            );

            if (sobranteEncontrado) {
                datosExcel.push([
                    ot,
                    '',
                    'GUARDAR SOBRANTE',
                    sobranteEncontrado.n_colmena,
                    sobranteEncontrado.cod,
                    sobranteEncontrado.medida_cm
                ]);
            }
        }
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(datosExcel);

    ws['!cols'] = [
        { wch: 12 },
        { wch: 15 },
        { wch: 20 },
        { wch: 12 },
        { wch: 12 },
        { wch: 15 }
    ];

    XLSX.utils.book_append_sheet(wb, ws, 'Plan de Corte');

    const nombreArchivo = `plan_corte_${new Date().toISOString().slice(0,10)}.xlsx`;
    XLSX.writeFile(wb, nombreArchivo);

    alert('Plan de corte exportado correctamente.');
}

function exportarColmenasDisponibles() {
    if (SistemaInventario.colmenasHistorico.length === 0) { alert('No hay resultados para exportar. Ejecute la optimización primero.'); return; }

    const colmenasDisponiblesFiltradas = SistemaInventario.colmenasHistorico.filter(c => c.estado === 'disponible');

    if (colmenasDisponiblesFiltradas.length === 0) { alert('No hay colmenas disponibles para exportar.'); return; }

    const wb = XLSX.utils.book_new();

    const estiloTitulo = { 
        font: { bold: true, sz: 14, color: { rgb: '000000' } }, 
        fill: { fgColor: { rgb: 'E8E8E8' } },
        border: { 
            top: { style: 'medium', color: { rgb: '000000' } },
            bottom: { style: 'medium', color: { rgb: '000000' } },
            left: { style: 'medium', color: { rgb: '000000' } },
            right: { style: 'medium', color: { rgb: '000000' } }
        }
    };
    
    const estiloEncabezado = { 
        font: { bold: true, sz: 12, color: { rgb: '000000' } }, 
        fill: { fgColor: { rgb: 'D0D0D0' } },
        border: { 
            top: { style: 'thin', color: { rgb: '000000' } },
            bottom: { style: 'thin', color: { rgb: '000000' } },
            left: { style: 'thin', color: { rgb: '000000' } },
            right: { style: 'thin', color: { rgb: '000000' } }
        },
        alignment: { horizontal: 'center' }
    };
    
    const estiloNegrita = { font: { bold: true, sz: 11 } };
    
    const estiloContenido = { 
        font: { bold: false, sz: 11 },
        border: { 
            top: { style: 'thin', color: { rgb: 'CCCCCC' } },
            bottom: { style: 'thin', color: { rgb: 'CCCCCC' } },
            left: { style: 'thin', color: { rgb: 'CCCCCC' } },
            right: { style: 'thin', color: { rgb: 'CCCCCC' } }
        }
    };
    
    const estiloNumero = { 
        font: { bold: false, sz: 11 },
        border: { 
            top: { style: 'thin', color: { rgb: 'CCCCCC' } },
            bottom: { style: 'thin', color: { rgb: 'CCCCCC' } },
            left: { style: 'thin', color: { rgb: 'CCCCCC' } },
            right: { style: 'thin', color: { rgb: 'CCCCCC' } }
        },
        alignment: { horizontal: 'right' }
    };
    
    const estiloSeparador = {
        border: { bottom: { style: 'medium', color: { rgb: '000000' } } }
    };
    
    const estiloTotal = { 
        font: { bold: true, sz: 11 },
        fill: { fgColor: { rgb: 'E0E0E0' } },
        border: { 
            top: { style: 'thin', color: { rgb: '000000' } },
            bottom: { style: 'medium', color: { rgb: '000000' } },
            left: { style: 'thin', color: { rgb: '000000' } },
            right: { style: 'thin', color: { rgb: '000000' } }
        }
    };

    const fechaActual = new Date().toLocaleString('es-AR', { 
        year: 'numeric', 
        month: '2-digit', 
        day: '2-digit', 
        hour: '2-digit', 
        minute: '2-digit' 
    });
    
    let numeroPagina = 1;

    let datosUbicacion = [];
    
    datosUbicacion.push(['REPORTE DE UBICACIÓN DE COLMENAS']);
    datosUbicacion.push(['Fecha de generación: ' + fechaActual]);
    datosUbicacion.push([]);
    datosUbicacion.push([]);
    
    const agrupadasPorColmena = {};
    colmenasDisponiblesFiltradas.forEach(c => {
        const nColmena = c.n_colmena;
        if (!agrupadasPorColmena[nColmena]) {
            agrupadasPorColmena[nColmena] = [];
        }
        agrupadasPorColmena[nColmena].push(c);
    });

    const clavesColmena = Object.keys(agrupadasPorColmena).sort((a, b) => ordenarColmena(a, b));

    clavesColmena.forEach((nColmena, idx) => {
        const tubosEnColmena = agrupadasPorColmena[nColmena];
        const totalTubos = tubosEnColmena.length;
        
        datosUbicacion.push(['COLMENA: ' + nColmena + '   |   TOTAL TUBOS: ' + totalTubos]);
        datosUbicacion.push(['Código', 'Medida (cm)', 'Estado']);
        
        const tubosOrdenados = tubosEnColmena.sort((a, b) => (b.medida_cm || 0) - (a.medida_cm || 0));
        
        tubosOrdenados.forEach(c => {
            datosUbicacion.push([c.cod, c.medida_cm, c.estado]);
        });
        
        if (idx < clavesColmena.length - 1) {
            datosUbicacion.push([]);
        }
    });
    
    datosUbicacion.push([]);
    datosUbicacion.push(['Página ' + numeroPagina]);

    const ws1 = XLSX.utils.aoa_to_sheet(datosUbicacion);
    
    ws1['A1'].s = estiloTitulo;
    ws1['A2'].s = estiloNegrita;
    
    let filaActual = 5;
    clavesColmena.forEach((nColmena, idx) => {
        if (ws1['A' + filaActual]) {
            ws1['A' + filaActual].s = estiloTitulo;
        }
        filaActual++;
        
        for (let col = 0; col < 3; col++) {
            const cellAddr = XLSX.utils.encode_cell({ r: filaActual, c: col });
            if (ws1[cellAddr]) {
                ws1[cellAddr].s = estiloEncabezado;
            }
        }
        filaActual++;
        
        const totalTubos = agrupadasPorColmena[nColmena].length;
        for (let i = 0; i < totalTubos; i++) {
            if (ws1[XLSX.utils.encode_cell({ r: filaActual, c: 0 })]) {
                ws1[XLSX.utils.encode_cell({ r: filaActual, c: 0 })].s = estiloContenido;
            }
            if (ws1[XLSX.utils.encode_cell({ r: filaActual, c: 1 })]) {
                ws1[XLSX.utils.encode_cell({ r: filaActual, c: 1 })].s = estiloNumero;
            }
            if (ws1[XLSX.utils.encode_cell({ r: filaActual, c: 2 })]) {
                ws1[XLSX.utils.encode_cell({ r: filaActual, c: 2 })].s = estiloContenido;
            }
            filaActual++;
        }
        
        if (idx < clavesColmena.length - 1) {
            for (let col = 0; col < 3; col++) {
                const cellAddr = XLSX.utils.encode_cell({ r: filaActual, c: col });
                if (ws1[cellAddr]) {
                    ws1[cellAddr].s = estiloSeparador;
                }
            }
            filaActual++;
        }
    });

    ws1['!cols'] = [{ wch: 12 }, { wch: 14 }, { wch: 14 }];
    XLSX.utils.book_append_sheet(wb, ws1, 'UBICACION_COLMENAS');

    numeroPagina++;
    let datosBusqueda = [];
    
    datosBusqueda.push(['REPORTE DE BÚSQUEDA POR CÓDIGO']);
    datosBusqueda.push(['Fecha de generación: ' + fechaActual]);
    datosBusqueda.push([]);
    datosBusqueda.push([]);
    
    const agrupadasPorCodigo = {};
    colmenasDisponiblesFiltradas.forEach(c => {
        const cod = c.cod || 'SIN-CODIGO';
        if (!agrupadasPorCodigo[cod]) {
            agrupadasPorCodigo[cod] = [];
        }
        agrupadasPorCodigo[cod].push(c);
    });

    const clavesCodigo = Object.keys(agrupadasPorCodigo).sort();

    clavesCodigo.forEach((cod, idx) => {
        const tubosEnCodigo = agrupadasPorCodigo[cod];
        const totalTubos = tubosEnCodigo.length;
        
        datosBusqueda.push(['CÓDIGO: ' + cod + '   |   TOTAL: ' + totalTubos]);
        datosBusqueda.push(['Colmena', 'Medida (cm)', 'Estado']);
        
        const tubosOrdenados = tubosEnCodigo.sort((a, b) => (b.medida_cm || 0) - (a.medida_cm || 0));
        
        tubosOrdenados.forEach(c => {
            datosBusqueda.push([c.n_colmena, c.medida_cm, c.estado]);
        });
        
        if (idx < clavesCodigo.length - 1) {
            datosBusqueda.push([]);
        }
    });
    
    datosBusqueda.push([]);
    datosBusqueda.push(['Página ' + numeroPagina]);

    const ws2 = XLSX.utils.aoa_to_sheet(datosBusqueda);
    
    ws2['A1'].s = estiloTitulo;
    ws2['A2'].s = estiloNegrita;
    
    filaActual = 5;
    clavesCodigo.forEach((cod, idx) => {
        if (ws2['A' + filaActual]) {
            ws2['A' + filaActual].s = estiloTitulo;
        }
        filaActual++;
        
        for (let col = 0; col < 3; col++) {
            const cellAddr = XLSX.utils.encode_cell({ r: filaActual, c: col });
            if (ws2[cellAddr]) {
                ws2[cellAddr].s = estiloEncabezado;
            }
        }
        filaActual++;
        
        const totalTubos = agrupadasPorCodigo[cod].length;
        for (let i = 0; i < totalTubos; i++) {
            if (ws2[XLSX.utils.encode_cell({ r: filaActual, c: 0 })]) {
                ws2[XLSX.utils.encode_cell({ r: filaActual, c: 0 })].s = estiloContenido;
            }
            if (ws2[XLSX.utils.encode_cell({ r: filaActual, c: 1 })]) {
                ws2[XLSX.utils.encode_cell({ r: filaActual, c: 1 })].s = estiloNumero;
            }
            if (ws2[XLSX.utils.encode_cell({ r: filaActual, c: 2 })]) {
                ws2[XLSX.utils.encode_cell({ r: filaActual, c: 2 })].s = estiloContenido;
            }
            filaActual++;
        }
        
        if (idx < clavesCodigo.length - 1) {
            for (let col = 0; col < 3; col++) {
                const cellAddr = XLSX.utils.encode_cell({ r: filaActual, c: col });
                if (ws2[cellAddr]) {
                    ws2[cellAddr].s = estiloSeparador;
                }
            }
            filaActual++;
        }
    });

    ws2['!cols'] = [{ wch: 12 }, { wch: 14 }, { wch: 14 }];
    XLSX.utils.book_append_sheet(wb, ws2, 'BUSQUEDA_POR_CODIGO');

    numeroPagina++;
    let datosResumen = [];
    
    datosResumen.push(['RESUMEN GENERAL DE COLMENAS']);
    datosResumen.push(['Fecha de generación: ' + fechaActual]);
    datosResumen.push([]);
    datosResumen.push(['Colmena', 'Total Tubos']);
    
    const totalesPorColmena = [];
    clavesColmena.forEach(nColmena => {
        totalesPorColmena.push({
            colmena: nColmena,
            total: agrupadasPorColmena[nColmena].length
        });
    });
    
    totalesPorColmena.forEach(item => {
        datosResumen.push([item.colmena, item.total]);
    });
    
    const totalGeneral = colmenasDisponiblesFiltradas.length;
    datosResumen.push([]);
    datosResumen.push(['TOTAL GENERAL', totalGeneral]);
    
    datosResumen.push([]);
    datosResumen.push(['Página ' + numeroPagina]);

    const ws3 = XLSX.utils.aoa_to_sheet(datosResumen);
    
    ws3['A1'].s = estiloTitulo;
    ws3['A2'].s = estiloNegrita;
    
    for (let col = 0; col < 2; col++) {
        const cellAddr = XLSX.utils.encode_cell({ r: 3, c: col });
        if (ws3[cellAddr]) {
            ws3[cellAddr].s = estiloEncabezado;
        }
    }
    
    for (let r = 4; r < datosResumen.length - 4; r++) {
        for (let col = 0; col < 2; col++) {
            const cellAddr = XLSX.utils.encode_cell({ r: r, c: col });
            if (ws3[cellAddr]) {
                ws3[cellAddr].s = estiloContenido;
            }
        }
    }
    
    const filaTotal = datosResumen.length - 4;
    for (let col = 0; col < 2; col++) {
        const cellAddr = XLSX.utils.encode_cell({ r: filaTotal, c: col });
        if (ws3[cellAddr]) {
            ws3[cellAddr].s = estiloTotal;
        }
    }

    ws3['!cols'] = [{ wch: 20 }, { wch: 14 }];
    XLSX.utils.book_append_sheet(wb, ws3, 'RESUMEN_GENERAL');

    const nombreArchivo = `colmenas_disponibles_${new Date().toISOString().slice(0,10)}.xlsx`;
    XLSX.writeFile(wb, nombreArchivo);
    log(`✅ Colmenas disponibles exportadas a ${nombreArchivo}`, 'success');
    alert(`Colmenas disponibles exportadas exitosamente a: ${nombreArchivo}`);
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
async function pruebaFirestore() {
    const db = window.firebaseDB;

    const docRef = await window.fbAddDoc(
        window.fbCollection(db, "test"),
        { mensaje: "Firestore funcionando 🚀", fecha: new Date() }
    );

    console.log("Documento creado:", docRef.id);
}

pruebaFirestore();







