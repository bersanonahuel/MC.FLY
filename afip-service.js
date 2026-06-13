/**
 * afip-service.js
 * Integración directa con AFIP/ARCA - WSAA (Autenticación) + WSFE (Facturación Electrónica)
 * Sin intermediarios ni SaaS. Usa soap + node-forge.
 *
 * Contribuyente: QUAGLIA MARIA CAROLINA | CUIT: 27303503256 | Pto. Vta: 2
 * Condición: Responsable Inscripto → Emite Facturas A (tipo 1) y B (tipo 6)
 */

const soap    = require('soap');
const forge   = require('node-forge');
const fs      = require('fs');
const path    = require('path');


// Node.js v17+ rechaza claves DH pequeñas de los servidores AFIP.
// Resolvemos bajando el nivel de seguridad TLS a nivel de proceso.
const tls = require('tls');
const origCreateSecureContext = tls.createSecureContext;
tls.createSecureContext = (options = {}) => {
    // SECLEVEL=0 permite DH keys de cualquier tamaño (AFIP usa 1024-bit)
    options.ciphers = options.ciphers || 'DEFAULT@SECLEVEL=0';
    if (!options.ciphers.includes('SECLEVEL')) {
        options.ciphers += ':DEFAULT@SECLEVEL=0';
    }
    return origCreateSecureContext(options);
};

// WSFE de AFIP usa certificados auto-firmados en algunos entornos
process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';


// ─── CONFIGURACIÓN ────────────────────────────────────────────────────────────
const CUIT        = '27303503256';
const CERT_PATH   = path.join(__dirname, 'certificado.crt');
const KEY_PATH    = path.join(__dirname, 'privada.key');

// URLs de producción de AFIP
const WSAA_WSDL   = 'https://wsaa.afip.gov.ar/ws/services/LoginCms?wsdl';
const WSFE_WSDL   = 'https://servicios1.afip.gov.ar/wsfev1/service.asmx?WSDL';

// Cache del Ticket de Autorización (TA)
let cachedTA = null;

// ─── WSAA: GENERAR TOKEN/SIGN ─────────────────────────────────────────────────
function buildCMS(service) {
    const certPem = fs.readFileSync(CERT_PATH, 'utf8');
    const keyPem  = fs.readFileSync(KEY_PATH,  'utf8');

    const cert = forge.pki.certificateFromPem(certPem);
    const key  = forge.pki.privateKeyFromPem(keyPem);

    const now       = new Date();
    const genTime   = new Date(now.getTime() - 60000);   // 1 min antes
    const expTime   = new Date(now.getTime() + 43200000); // +12 horas

    // Formato ISO 8601 UTC que acepta AFIP (sin milisegundos, con Z)
    function toAfipISO(d) {
        return d.toISOString().replace(/\.\d{3}Z$/, 'Z');
    }

    // TRA (Ticket de Requerimiento de Acceso) en XML
    const tra = `<?xml version="1.0" encoding="UTF-8"?>
<loginTicketRequest version="1.0">
  <header>
    <uniqueId>${Math.floor(Date.now() / 1000)}</uniqueId>
    <generationTime>${toAfipISO(genTime)}</generationTime>
    <expirationTime>${toAfipISO(expTime)}</expirationTime>
  </header>
  <service>${service}</service>
</loginTicketRequest>`;

    // Firmar el TRA como CMS (PKCS#7)
    const p7 = forge.pkcs7.createSignedData();
    p7.content = forge.util.createBuffer(tra, 'utf8');
    p7.addCertificate(cert);
    p7.addSigner({
        key:         key,
        certificate: cert,
        digestAlgorithm: forge.pki.oids.sha256,
        authenticatedAttributes: [
            { type: forge.pki.oids.contentType,   value: forge.pki.oids.data },
            { type: forge.pki.oids.messageDigest                              },
            { type: forge.pki.oids.signingTime,   value: now                 }
        ]
    });
    p7.sign();

    const der    = forge.asn1.toDer(p7.toAsn1()).getBytes();
    const base64 = forge.util.encode64(der);
    return base64;
}

// Ruta del cache del TA en disco (para reutilizar entre reinicios del proceso)
const TA_CACHE_PATH = path.join(__dirname, '.ta_cache.json');

function loadCachedTA() {
    try {
        if (fs.existsSync(TA_CACHE_PATH)) {
            const data = JSON.parse(fs.readFileSync(TA_CACHE_PATH, 'utf8'));
            if (data && new Date(data.expiration) > new Date(Date.now() + 60000)) {
                cachedTA = data;
                console.log('[AFIP] Ticket de Autorización reutilizado del cache. Vence:', data.expiration);
                return data;
            }
        }
    } catch (e) { /* ignorar errores de cache */ }
    return null;
}

async function getTicketAcceso(service = 'wsfe') {
    // 1. Reusar ticket en memoria
    if (cachedTA && new Date(cachedTA.expiration) > new Date(Date.now() + 60000)) {
        return cachedTA;
    }

    // 2. Reusar ticket en disco
    const fromDisk = loadCachedTA();
    if (fromDisk) return fromDisk;

    // 3. Pedir nuevo ticket a WSAA
    const cms = buildCMS(service);

    const client = await soap.createClientAsync(WSAA_WSDL);
    const response = await client.loginCmsAsync({ in0: cms });
    // soap devuelve [result, rawResponse, soapHeader, rawRequest]
    const result = Array.isArray(response) ? response[0] : response;
    const xmlTA = result.loginCmsReturn;

    // Parsear el XML de respuesta del TA
    const token = xmlTA.match(/<token>([\s\S]*?)<\/token>/)[1];
    const sign  = xmlTA.match(/<sign>([\s\S]*?)<\/sign>/)[1];
    const expiration = xmlTA.match(/<expirationTime>([\s\S]*?)<\/expirationTime>/)[1];

    cachedTA = { token, sign, expiration };

    // Guardar en disco para reutilizar
    try { fs.writeFileSync(TA_CACHE_PATH, JSON.stringify(cachedTA)); } catch (e) { /* ok */ }

    console.log('[AFIP] Ticket de Autorización renovado. Vence:', expiration);
    return cachedTA;
}

// ─── WSFE: OBTENER ÚLTIMO COMPROBANTE ─────────────────────────────────────────
async function getUltimoComprobante(ptoVta, cbteTipo) {
    const ta     = await getTicketAcceso();
    const client = await soap.createClientAsync(WSFE_WSDL);

    const args = {
        Auth: { Token: ta.token, Sign: ta.sign, Cuit: CUIT },
        PtoVta: ptoVta,
        CbteTipo: cbteTipo
    };

    const methodName = typeof client.FECompUltimoAutorizadoAsync === 'function'
        ? 'FECompUltimoAutorizadoAsync' : 'FECompUltimoAutorizado_V1Async';
    const response = await client[methodName](args);
    // soap devuelve [result, rawResponse, soapHeader, rawRequest]
    const result = Array.isArray(response) ? response[0] : response;

    const res = result.FECompUltimoAutorizadoResult || result;
    if (res.Errors && res.Errors.Err) {
        const errMsg = Array.isArray(res.Errors.Err) ? res.Errors.Err[0].Msg : res.Errors.Err.Msg;
        throw new Error('WSFE Error: ' + errMsg);
    }
    return parseInt(res.CbteNro) || 0;
}

// ─── WSFE: CREAR COMPROBANTE ───────────────────────────────────────────────────
async function crearComprobante({ total, ptoVta, cbteTipo, docTipo, docNro }) {
    const ta     = await getTicketAcceso();
    const client = await soap.createClientAsync(WSFE_WSDL);

    const ultimoNro  = await getUltimoComprobante(ptoVta, cbteTipo);
    const nroComp    = ultimoNro + 1;

    // Fecha en formato AFIP: YYYYMMDD
    const hoy   = new Date();
    const fecha = `${hoy.getFullYear()}${String(hoy.getMonth()+1).padStart(2,'0')}${String(hoy.getDate()).padStart(2,'0')}`;

    // Desglosa IVA 21% incluido en el precio final
    const impTotal  = +parseFloat(total).toFixed(2);
    let   impNeto   = +(impTotal / 1.21).toFixed(2);
    let   impIVA    = +(impTotal - impNeto).toFixed(2);

    // Corrección mínima de redondeo
    if (Math.abs(impNeto + impIVA - impTotal) > 0.01) {
        impNeto = +(impTotal - impIVA).toFixed(2);
    }

    const args = {
        Auth: { Token: ta.token, Sign: ta.sign, Cuit: CUIT },
        FeCAEReq: {
            FeCabReq: {
                CantReg:  1,
                PtoVta:   ptoVta,
                CbteTipo: cbteTipo
            },
            FeDetReq: {
                FECAEDetRequest: {
                    Concepto:   1,
                    DocTipo:    docTipo || 99,
                    DocNro:     docNro  || 0,
                    CbteDesde:  nroComp,
                    CbteHasta:  nroComp,
                    CbteFch:    fecha,
                    ImpTotal:   impTotal,
                    ImpTotConc: 0,
                    ImpNeto:    impNeto,
                    ImpOpEx:    0,
                    ImpIVA:     impIVA,
                    ImpTrib:    0,
                    MonId:      'PES',
                    MonCotiz:   1,
                    Iva: {
                        AlicIva: {
                            Id:      5,   // 21%
                            BaseImp: impNeto,
                            Importe: impIVA
                        }
                    }
                }
            }
        }
    };

    const solicitarMethod = typeof client.FECAESolicitarAsync === 'function'
        ? 'FECAESolicitarAsync' : 'FECAESolicitar_V1Async';
    const solicitarResp = await client[solicitarMethod](args);
    const result = Array.isArray(solicitarResp) ? solicitarResp[0] : solicitarResp;

    const res = result.FECAESolicitarResult || result;

    // Verificar errores AFIP
    if (res.FeDetResp && res.FeDetResp.FECAEDetResponse) {
        const det = res.FeDetResp.FECAEDetResponse;
        if (det.Resultado === 'R') {
            const obs = det.Observaciones && det.Observaciones.Obs
                ? (Array.isArray(det.Observaciones.Obs) ? det.Observaciones.Obs : [det.Observaciones.Obs])
                      .map(o => `[${o.Code}] ${o.Msg}`).join(' | ')
                : 'Sin observaciones';
            throw new Error('AFIP rechazó el comprobante: ' + obs);
        }
        const cae           = det.CAE;
        const caeVto        = det.CAEFchVto;
        const invoiceNumber = String(ptoVta).padStart(5,'0') + '-' + String(nroComp).padStart(8,'0');
        return { cae, caeVto, invoiceNumber, nroComp };
    }

    throw new Error('Respuesta inesperada de AFIP: ' + JSON.stringify(res));
}

module.exports = { crearComprobante, getUltimoComprobante };
