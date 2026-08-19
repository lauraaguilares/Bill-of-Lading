/**
 * mailTransport.js
 * Envío de correo vía la API HTTP de Resend (https://resend.com) en vez de SMTP directo.
 *
 * SMTP (puertos 465 y 587) está bloqueado en el plan gratuito de Render — es una
 * restricción anti-spam del hosting, no algo arreglable cambiando configuración de
 * conexión. Una API de correo por HTTPS (como Resend) no tiene ese problema porque no
 * es una conexión SMTP tradicional.
 *
 * Variable de entorno requerida: RESEND_API_KEY
 */

const RESEND_API_URL = 'https://api.resend.com/emails';

/**
 * Envía un correo con adjuntos vía Resend.
 * @param {object} opciones
 * @param {string} opciones.apiKey - RESEND_API_KEY
 * @param {string} opciones.from - Remitente. Sin dominio propio verificado en Resend, debe
 *   ser 'onboarding@resend.dev' (el dominio de pruebas que Resend da por default).
 * @param {string} opciones.to - Destinatario.
 * @param {string} opciones.subject
 * @param {string} opciones.text
 * @param {Array<{filename: string, path: string}>} opciones.attachments - igual formato
 *   que nodemailer, para no tener que tocar el código que arma los adjuntos.
 */
async function enviarCorreo({ apiKey, from, to, subject, text, attachments = [] }) {
  if (!apiKey) throw new Error('Falta RESEND_API_KEY en las variables de entorno.');

  const fs = require('fs');
  const attachmentsBase64 = attachments.map((a) => ({
    filename: a.filename,
    content: fs.readFileSync(a.path).toString('base64'),
  }));

  const res = await fetch(RESEND_API_URL, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${apiKey}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      from,
      to: [to],
      subject,
      text,
      attachments: attachmentsBase64,
    }),
  });

  if (!res.ok) {
    const detalle = await res.text().catch(() => '');
    throw new Error(`Resend respondió ${res.status}: ${detalle}`);
  }

  return res.json();
}

module.exports = { enviarCorreo };
