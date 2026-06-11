// Proxy AEAT VeriFactu con TLS mutuo
// GAS llama a este Worker con el SOAP XML ya firmado;
// el Worker añade el certificado cliente en el handshake TLS hacia AEAT.

const SOAP_ACTION = 'https://www2.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/es/aeat/tikeV1/cont/ws/SistemaFacturacion/RegFactuSistemaFacturacion';

const ENDPOINTS = {
  TEST: 'https://prewww1.aeat.es/wlpl/TIKE-CONT/ws/SistemaFacturacion/VerifactuSOAP',
  PROD: 'https://www1.aeat.es/wlpl/TIKE-CONT/ws/SistemaFacturacion/VerifactuSOAP',
};

export default {
  async fetch(request, env) {
    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: corsHeaders() });
    }
    if (request.method !== 'POST') {
      return new Response('Method not allowed', { status: 405 });
    }

    const body    = await request.text();
    const entorno = (request.headers.get('X-Entorno') || 'TEST').toUpperCase();
    const url     = ENDPOINTS[entorno] || ENDPOINTS.TEST;

    const resp = await env.AEAT_CERT.fetch(url, {
      method:  'POST',
      headers: {
        'Content-Type': 'text/xml; charset=UTF-8',
        'SOAPAction':   SOAP_ACTION,
      },
      body,
    });

    const text = await resp.text();
    return new Response(text, {
      status:  resp.status,
      headers: {
        'Content-Type': resp.headers.get('Content-Type') || 'text/xml',
        ...corsHeaders(),
      },
    });
  },
};

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin':  '*',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, X-Entorno',
  };
}
