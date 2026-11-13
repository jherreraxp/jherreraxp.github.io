/* Página de proyecto: lee ./proyecto.json y arma las tres opciones.
   1. Usar la app · 2. Usar el prompt en IAs populares · 3. Instalar tú mismo */
(function () {
  'use strict';
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var esc = function (t) {
    return String(t == null ? '' : t).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  };
  var GUIA = {
    claude: ['Abra Claude y pegue el prompt (Ctrl+V).', 'Claude entrega cada archivo como artefacto descargable.', 'Siga el LEEME que genera para instalarlo en Google Apps Script.'],
    chatgpt: ['Abra ChatGPT y pegue el prompt (Ctrl+V).', 'Si aparece Canvas, cada archivo se abre allí; si la respuesta se corta, escriba «continúa».', 'Siga el LEEME que genera para instalarlo en Google Apps Script.'],
    gemini: ['Abra Gemini y pegue el prompt (Ctrl+V).', 'Active Canvas para obtener cada archivo completo.', 'Siga el LEEME que genera para instalarlo en Google Apps Script.'],
    copilot: ['Abra Copilot y pegue el prompt (Ctrl+V).', 'Recibirá un archivo por respuesta: escriba «siguiente» para el próximo.', 'Siga el LEEME que genera para instalarlo en Google Apps Script.'],
    deepseek: ['Abra DeepSeek, active DeepThink y pegue el prompt (Ctrl+V).', 'Si un archivo llega en partes numeradas, únalas en orden.', 'Siga el LEEME que genera para instalarlo en Google Apps Script.'],
    kimi: ['Abra Kimi y pegue el prompt (Ctrl+V).', 'Cada archivo llega en su propio bloque; si se corta, escriba «continúa».', 'Siga el LEEME que genera para instalarlo en Google Apps Script.']
  };
  var textos = {}, actual = null;

  function toast(msg) {
    var t = $('#toast'); t.textContent = msg; t.classList.add('on');
    clearTimeout(toast.h); toast.h = setTimeout(function () { t.classList.remove('on'); }, 2600);
  }
  function copiar(txt) {
    if (navigator.clipboard && window.isSecureContext) {
      return navigator.clipboard.writeText(txt).then(function () { return true; }, function () { return viejo(txt); });
    }
    return Promise.resolve(viejo(txt));
  }
  function viejo(txt) {
    var a = document.createElement('textarea'); a.value = txt; a.setAttribute('readonly', '');
    a.style.position = 'fixed'; a.style.opacity = '0'; document.body.appendChild(a); a.select();
    var ok = false; try { ok = document.execCommand('copy'); } catch (e) {}
    document.body.removeChild(a); return ok;
  }
  function cargarTexto(p) {
    if (textos[p.id]) return Promise.resolve(textos[p.id]);
    return fetch(p.archivo, { cache: 'no-cache' }).then(function (r) {
      if (!r.ok) throw new Error(r.status); return r.text();
    }).then(function (t) { textos[p.id] = t; return t; });
  }
  function elegir(d, p) {
    actual = p;
    document.querySelectorAll('.tab').forEach(function (b) { b.setAttribute('aria-selected', b.dataset.id === p.id ? 'true' : 'false'); });
    $('#guia').innerHTML = ((d.guia && d.guia[p.id]) || GUIA[p.id] || ['Copie el prompt y péguelo en la IA.']).map(function (g) { return '<li>' + esc(g) + '</li>'; }).join('');
    $('#abrir').textContent = 'Copiar y abrir ' + p.nombre;
    $('#bajar').href = p.archivo; $('#bajar').setAttribute('download', d.id + '-prompt-' + p.id + '.txt');
    $('#prompt').textContent = 'Cargando…';
    cargarTexto(p).then(function (t) { if (actual === p) $('#prompt').textContent = t; },
      function () { $('#prompt').textContent = 'No se pudo cargar el prompt. Use el botón «Descargar .txt».'; });
  }

  function pintar(d) {
    document.title = d.titulo + ' · toolsHXP';
    $('#titulo').textContent = d.titulo;
    $('#resumen').textContent = d.resumen;
    $('#tags').innerHTML = (d.etiquetas || []).map(function (e) { return '<li>' + esc(e) + '</li>'; }).join('');
    $('#feat').innerHTML = (d.caracteristicas || []).map(function (c) {
      return '<div class="card"><h3>' + esc(c[0]) + '</h3><p>' + esc(c[1]) + '</p></div>';
    }).join('');

    if (d.app && d.app.url) {
      $('#app-btn').href = d.app.url; $('#app-btn').textContent = d.app.etiqueta || 'Abrir la aplicación';
      $('#app-nota').textContent = d.app.nota || '';
    } else { $('#op1').hidden = true; }

    var ps = d.prompts || [];
    if (ps.length) {
      $('#tabs').innerHTML = ps.map(function (p) {
        return '<button class="tab" role="tab" type="button" data-id="' + esc(p.id) + '" aria-selected="false">' + esc(p.nombre) + '</button>';
      }).join('');
      document.querySelectorAll('.tab').forEach(function (b) {
        b.addEventListener('click', function () { elegir(d, ps.filter(function (p) { return p.id === b.dataset.id; })[0]); });
      });
      $('#copiar').addEventListener('click', function () {
        cargarTexto(actual).then(copiar).then(function (ok) {
          toast(ok ? 'Prompt para ' + actual.nombre + ' copiado.' : 'Seleccione el texto y cópielo manualmente.');
        });
      });
      $('#abrir').addEventListener('click', function () {
        var p = actual, t = textos[p.id];
        var fin = function (ok) { toast(ok ? 'Copiado. Péguelo en ' + p.nombre + ' con Ctrl+V.' : 'Copie el texto manualmente y péguelo en ' + p.nombre + '.'); };
        if (t) {                       // copiar antes de abrir: el portapapeles exige la pestaña enfocada
          var pr = copiar(t);
          window.open(p.url, '_blank', 'noopener');
          pr.then(fin);
        } else {
          cargarTexto(p).then(copiar).then(function (ok) { window.open(p.url, '_blank', 'noopener'); fin(ok); });
        }
      });
      elegir(d, ps[0]);
    } else { $('#op2').hidden = true; }

    var ins = d.instalar;
    if (ins && ins.zip) {
      $('#zip').href = ins.zip; $('#zip').setAttribute('download', ins.zip.split('/').pop());
      if (ins.leeme) $('#leeme').href = ins.leeme; else $('#leeme').hidden = true;
      $('#min').textContent = ins.minutos ? ' (unos ' + ins.minutos + ' minutos)' : '';
      $('#pasos').innerHTML = (ins.pasos || []).map(function (s) { return '<li>' + esc(s) + '</li>'; }).join('');
    } else { $('#op3').hidden = true; }
  }

  fetch('proyecto.json', { cache: 'no-cache' }).then(function (r) { return r.json(); }).then(pintar, function () {
    $('#resumen').textContent = 'No se pudo cargar la información del proyecto.';
  });
})();
