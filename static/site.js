(function () {
  var base = document.body.getAttribute('data-base') || '';
  var lang = document.documentElement.lang || '';
  var T = lang === 'art-x-tora' ? {
    results: function (q, n) { return '«' + q + '»: ' + n + ' řɛżʊłṅǎ'; },
    prompt: 'Ʒůƥğǎ ðįʃŧǎ ğřǎƒħő.',
    none: 'Ňœł ğřǎƒŧɔł ðǎʃ ʃɛ.'
  } : lang.indexOf('en') === 0 ? {
    results: function (q, n) { return n + (n === 1 ? ' result' : ' results') + ' for “' + q + '”'; },
    prompt: 'Type a search term in the box above.',
    none: 'There are no published articles yet.'
  } : {
    results: function (q, n) { return '「' + q + '」共有 ' + n + ' 筆結果'; },
    prompt: '請在上方輸入關鍵字。',
    none: '目前還沒有已發布的條目。'
  };
  var idx = null;
  function load() {
    if (idx) return Promise.resolve(idx);
    return fetch(base + 'search-index.json').then(function (r) { return r.json(); }).then(function (d) { idx = d; return d; });
  }
  function norm(s) { return (s || '').toLowerCase(); }
  function search(q, list) {
    q = norm(q.trim());
    if (!q) return [];
    var hits = [];
    list.forEach(function (e) {
      // t = 本語言標題，a = 其他語言的標題（以「 / 」分隔；用任何一種語言的名稱都能搜到）
      var t = norm(e.t), a = norm(e.a), s = norm(e.s), c = norm(e.c), score = 0;
      if (t === q || a === q) score = 100; else if (t.indexOf(q) === 0 || a.indexOf(q) === 0) score = 60;
      else if (t.indexOf(q) >= 0 || a.indexOf(q) >= 0) score = 40;
      else if (c.indexOf(q) >= 0) score = 15; else if (s.indexOf(q) >= 0) score = 10;
      if (score) hits.push({ e: e, score: score });
    });
    hits.sort(function (x, y) { return y.score - x.score || x.e.t.length - y.e.t.length; });
    return hits.map(function (h) { return h.e; });
  }
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

  // 頂欄即時建議
  var input = document.getElementById('q');
  var box = document.querySelector('.suggest');
  if (input && box) {
    input.addEventListener('input', function () {
      var v = input.value;
      if (!v.trim()) { box.hidden = true; return; }
      load().then(function (list) {
        var r = search(v, list).slice(0, 8);
        if (!r.length) { box.hidden = true; return; }
        box.innerHTML = r.map(function (e) {
          return '<li><a href="' + base + e.u + '">' + esc(e.t) + (e.c ? '<small>' + esc(e.c) + '</small>' : '') + '</a></li>';
        }).join('');
        box.hidden = false;
      });
    });
    input.addEventListener('keydown', function (ev) {
      if (ev.key === 'Escape') box.hidden = true;
    });
    document.addEventListener('click', function (ev) { if (!box.contains(ev.target) && ev.target !== input) box.hidden = true; });
  }

  // 語言選單：點選單外面或按 Esc 時關閉
  var langsw = document.querySelector('.langsw');
  if (langsw) {
    document.addEventListener('click', function (ev) { if (langsw.open && !langsw.contains(ev.target)) langsw.open = false; });
    document.addEventListener('keydown', function (ev) {
      if (ev.key === 'Escape' && langsw.open) { langsw.open = false; langsw.querySelector('summary').focus(); }
    });
  }

  // 搜尋結果頁
  if (document.body.classList.contains('search-page')) {
    var q = new URLSearchParams(location.search).get('q') || '';
    if (input) input.value = q;
    load().then(function (list) {
      var k = q.trim();
      var exact = list.filter(function (e) { return e.t === k || (e.a || '').split(' / ').indexOf(k) >= 0; })[0];
      if (exact) { location.replace(base + exact.u); return; }
      var r = search(q, list);
      document.getElementById('search-status').textContent = q ? T.results(q, r.length) : T.prompt;
      document.getElementById('search-results').innerHTML = r.map(function (e) {
        return '<li><a href="' + base + e.u + '"><b>' + esc(e.t) + '</b></a>' + (e.c ? '<small>' + esc(e.c) + '</small>' : '') + '<div>' + esc(e.s) + '</div></li>';
      }).join('');
    });
  }

  // 隨機條目
  if (document.body.classList.contains('random-page')) {
    load().then(function (list) {
      if (!list.length) { document.getElementById('random-status').textContent = T.none; return; }
      location.replace(base + list[Math.floor(Math.random() * list.length)].u);
    });
  }
})();
