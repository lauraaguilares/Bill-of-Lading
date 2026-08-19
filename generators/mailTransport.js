/**
 * mailTransport.js
 * Crea el transportador de nodemailer para Gmail, resolviendo la IP a mano y forzando
 * IPv4 — Render no soporta salida por IPv6, y ni el flag --dns-result-order=ipv4first
 * ni interceptar dns.lookup fueron suficientes para que nodemailer/SMTP lo respetara
 * (sigue intentando IPv6 primero internamente). Conectar directo a la IP ya resuelta
 * evita que la librería tenga que decidir nada.
 */

const dns = require('dns').promises;
const nodemailer = require('nodemailer');

const GMAIL_HOST = 'smtp.gmail.com';

async function crearTransportadorGmail(gmailUser, gmailAppPassword) {
  const { address } = await dns.lookup(GMAIL_HOST, { family: 4 });

  return nodemailer.createTransport({
    host: address,
    port: 587, // STARTTLS — más compatible que 465 (SSL directo) en redes de hosting con restricciones
    secure: false,
    requireTLS: true,
    auth: { user: gmailUser, pass: gmailAppPassword },
    tls: {
      // Necesario al conectar por IP directa: el certificado de Gmail es para el
      // hostname, no para la IP, así que hay que decirle explícitamente cuál validar.
      servername: GMAIL_HOST,
    },
    connectionTimeout: 20000, // 20s — para que falle claro en vez de colgarse mucho tiempo
  });
}

module.exports = { crearTransportadorGmail };
