/*
 * github.js
 * Lee y publica data.json en el repositorio usando la API de GitHub.
 * El token se guarda solo en el navegador de este computador (localStorage), nunca en el repositorio.
 */
(function (root) {
  'use strict';
  var CLAVE = 'sary-github';

  function config() { try { return JSON.parse(localStorage.getItem(CLAVE) || 'null'); } catch (e) { return null; } }
  function guardarConfig(c) { localStorage.setItem(CLAVE, JSON.stringify(c)); }
  function borrarConfig() { localStorage.removeItem(CLAVE); }

  // Deduce usuario y repositorio a partir de la dirección de GitHub Pages
  function sugerencia() {
    var m = /^([^.]+)\.github\.io$/i.exec(location.hostname);
    var partes = location.pathname.split('/').filter(Boolean);
    var repo = partes.length > 1 || (partes[0] && !/\.html?$/i.test(partes[0])) ? partes[0] : '';
    return { owner: m ? m[1] : '', repo: repo || '' };
  }

  function base(c) { return 'https://api.github.com/repos/' + encodeURIComponent(c.owner) + '/' + encodeURIComponent(c.repo); }
  function cabeceras(c) {
    return { 'Authorization': 'Bearer ' + c.token, 'Accept': 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28' };
  }

  function falla(res) {
    return res.json().catch(function () { return {}; }).then(function (j) {
      var msg;
      if (res.status === 401) msg = 'El token no es válido o ya venció. Configura uno nuevo.';
      else if (res.status === 403) msg = 'El token no tiene permiso para escribir en este repositorio. Debe tener "Contents: Read and write".';
      else if (res.status === 404) msg = 'No se encontró el repositorio, o el token no tiene acceso a él.';
      else if (res.status === 409 || res.status === 422) msg = 'Los datos cambiaron en GitHub mientras trabajabas. Vuelve a abrir los datos publicados y repite la carga.';
      else msg = 'GitHub respondió con error ' + res.status + (j.message ? ': ' + j.message : '') + '.';
      var e = new Error(msg); e.status = res.status; throw e;
    });
  }

  function b64aTexto(b) {
    var bin = atob(String(b).replace(/\s/g, ''));
    var bytes = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return new TextDecoder('utf-8').decode(bytes);
  }
  function textoAB64(t) {
    var bytes = new TextEncoder().encode(t), s = '';
    for (var i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(s);
  }

  // Comprueba que el repositorio existe y el token lo puede ver
  function probar(c) {
    return fetch(base(c), { headers: cabeceras(c), cache: 'no-store' }).then(function (res) {
      if (!res.ok) return falla(res);
      return res.json();
    });
  }

  // Devuelve { texto, sha } del archivo; texto null si no existe
  function leer(c, ruta) {
    return fetch(base(c) + '/contents/' + ruta + '?t=' + Date.now(), { headers: cabeceras(c), cache: 'no-store' }).then(function (res) {
      if (res.status === 404) return { texto: null, sha: null };
      if (!res.ok) return falla(res);
      return res.json().then(function (j) { return { texto: b64aTexto(j.content), sha: j.sha }; });
    });
  }

  // Crea o reemplaza el archivo; sha = versión que se leyó (evita pisar cambios de otra persona)
  function guardar(c, ruta, texto, sha, mensaje) {
    var cuerpo = { message: mensaje, content: textoAB64(texto) };
    if (sha) cuerpo.sha = sha;
    return fetch(base(c) + '/contents/' + ruta, {
      method: 'PUT',
      headers: Object.assign({ 'Content-Type': 'application/json' }, cabeceras(c)),
      body: JSON.stringify(cuerpo)
    }).then(function (res) {
      if (!res.ok) return falla(res);
      return res.json().then(function (j) { return { sha: j.content && j.content.sha }; });
    });
  }

  root.GitHubSary = { config: config, guardarConfig: guardarConfig, borrarConfig: borrarConfig, sugerencia: sugerencia, probar: probar, leer: leer, guardar: guardar };
})(window);
