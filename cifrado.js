/*
 * cifrado.js
 * Cifra y descifra el data.json con una contraseña usando Web Crypto:
 * PBKDF2-SHA256 (310.000 iteraciones, sal aleatoria) para derivar la clave y AES-GCM de 256 bits.
 * Sin la contraseña, el archivo publicado es ilegible.
 */
(function (root) {
  'use strict';
  var ITERACIONES = 310000;
  var subtle = (root.crypto || (typeof globalThis !== 'undefined' && globalThis.crypto)).subtle;

  function aB64(bytes) {
    var s = '';
    for (var i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
    return btoa(s);
  }
  function deB64(t) {
    var s = atob(t), b = new Uint8Array(s.length);
    for (var i = 0; i < s.length; i++) b[i] = s.charCodeAt(i);
    return b;
  }
  function aleatorio(n) {
    var b = new Uint8Array(n);
    (root.crypto || globalThis.crypto).getRandomValues(b);
    return b;
  }

  function derivar(contrasena, sal, iter) {
    return subtle.importKey('raw', new TextEncoder().encode(contrasena), 'PBKDF2', false, ['deriveKey'])
      .then(function (base) {
        return subtle.deriveKey(
          { name: 'PBKDF2', salt: sal, iterations: iter, hash: 'SHA-256' },
          base, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
      });
  }

  function cifrar(objeto, contrasena) {
    var sal = aleatorio(16), iv = aleatorio(12);
    return derivar(contrasena, sal, ITERACIONES).then(function (clave) {
      return subtle.encrypt({ name: 'AES-GCM', iv: iv }, clave, new TextEncoder().encode(JSON.stringify(objeto)));
    }).then(function (buf) {
      return {
        cifrado: true,
        formato: 'PBKDF2-SHA256/AES-GCM-256',
        iteraciones: ITERACIONES,
        sal: aB64(sal),
        iv: aB64(iv),
        datos: aB64(new Uint8Array(buf))
      };
    });
  }

  // Rechaza con Error('CONTRASENA') si la clave no es la correcta
  function descifrar(paquete, contrasena) {
    return derivar(contrasena, deB64(paquete.sal), paquete.iteraciones || ITERACIONES)
      .then(function (clave) {
        return subtle.decrypt({ name: 'AES-GCM', iv: deB64(paquete.iv) }, clave, deB64(paquete.datos));
      })
      .then(function (buf) { return JSON.parse(new TextDecoder().decode(buf)); },
        function () { throw new Error('CONTRASENA'); });
  }

  var api = { cifrar: cifrar, descifrar: descifrar };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.CifradoSary = api;
})(typeof window !== 'undefined' ? window : globalThis);
