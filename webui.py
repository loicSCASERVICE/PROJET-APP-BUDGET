#!/usr/bin/env python3
"""
webui.py
--------
Interface web locale (Flask) pour catégoriser les transactions dans un
navigateur au lieu du terminal. Lancée automatiquement par budget_processor.py
en mode interactif par défaut.
"""

import json
import socket
import threading
import webbrowser

from flask import Flask, jsonify, render_template_string, request
from werkzeug.serving import make_server

UNCATEGORIZED_LABEL = "À catégoriser"

PAGE_TEMPLATE = """
<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="utf-8">
<title>Catégorisation du budget</title>
<style>
  :root { --bleu: #1F4E78; --bleu-clair: #D9E1F2; --vert: #70AD47; --rouge: #C0504D; }
  * { box-sizing: border-box; }
  body { font-family: -apple-system, "Segoe UI", Arial, sans-serif; margin: 0; background: #f4f6f8; color: #222; }
  header { background: var(--bleu); color: white; padding: 16px 24px; position: sticky; top: 0; z-index: 10; }
  header h1 { margin: 0; font-size: 18px; }
  header p { margin: 4px 0 0; font-size: 13px; opacity: .85; }
  .wrap { padding: 16px 24px 100px; }
  table { border-collapse: collapse; width: 100%; background: white; box-shadow: 0 1px 3px rgba(0,0,0,.1); }
  th, td { padding: 8px 10px; border-bottom: 1px solid #eee; font-size: 13px; text-align: left; vertical-align: middle; }
  th { background: var(--bleu-clair); position: sticky; top: 68px; }
  td.montant { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
  td.montant.neg { color: var(--rouge); }
  td.montant.pos { color: var(--vert); }
  .libelle { max-width: 420px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .star { color: var(--vert); font-weight: bold; margin-right: 4px; }
  footer { position: fixed; bottom: 0; left: 0; right: 0; background: white; border-top: 1px solid #ddd;
           padding: 12px 24px; display: flex; align-items: center; gap: 16px; box-shadow: 0 -2px 6px rgba(0,0,0,.08); }
  footer label { font-size: 13px; display: flex; align-items: center; gap: 6px; }
  button { background: var(--bleu); color: white; border: none; padding: 10px 22px; font-size: 14px;
           border-radius: 4px; cursor: pointer; }
  button:hover { background: #163a5c; }
  .count { font-size: 13px; color: #555; }
  .combo { position: relative; }
  .cat-input { width: 100%; padding: 4px 6px; font-size: 13px; border: 1px solid #ccc; border-radius: 3px; }
  tr.suggested td.cat .cat-input { border-color: var(--vert); }
  tr.none td.cat .cat-input { border-color: var(--rouge); }
  .cat-list { position: absolute; top: 100%; left: 0; right: 0; margin-top: 2px; background: white;
              border: 1px solid #ccc; border-radius: 3px; max-height: 220px; overflow-y: auto; z-index: 20;
              box-shadow: 0 2px 8px rgba(0,0,0,.18); }
  .cat-item { padding: 6px 8px; font-size: 13px; cursor: pointer; }
  .cat-item:hover, .cat-item.highlight { background: var(--bleu-clair); }
  .cat-item.empty-msg { color: #888; cursor: default; }
  .cat-item.new-cat-option { font-weight: bold; color: var(--bleu); }
  .cat-item.clear-option { color: #888; border-top: 1px solid #eee; }
  #done { display: none; text-align: center; padding: 80px 20px; }
  #done h2 { color: var(--vert); }
</style>
</head>
<body>
<header>
  <h1>Catégorisation des transactions — {{ month_label }}</h1>
  <p>{{ n }} transaction(s). Les lignes avec <span class="star">★</span> ont déjà une suggestion (mot-clé mémorisé) — vérifie-la et corrige si besoin.</p>
</header>

<div class="wrap" id="main-view">
<table>
  <thead>
    <tr><th>Date</th><th>Libellé</th><th>Montant</th><th>Catégorie</th></tr>
  </thead>
  <tbody id="tbody"></tbody>
</table>
</div>

<div id="done">
  <h2>✅ Terminé !</h2>
  <p>Le fichier Excel a été généré. Tu peux fermer cette fenêtre.</p>
</div>

<footer id="footer">
  <span class="count"><span id="remaining">0</span> ligne(s) sans catégorie choisie (resteront "À catégoriser")</span>
  <label><input type="checkbox" id="remember" checked> Mémoriser mes corrections pour les prochains relevés</label>
  <button id="submit-btn" onclick="submitAll()">Valider et générer le fichier</button>
</footer>

<script>
const DATA = {{ data_json|safe }};
const CATEGORIES = DATA.categories.slice(); // copie independante, pour detecter les nouvelles categories
sortCategories();
const TX = DATA.transactions;
const selected = TX.map(tx => tx.suggestion || '');

function sortCategories() {
  CATEGORIES.sort((a, b) => a.localeCompare(b, 'fr', {sensitivity: 'base'}));
}

function escapeHtml(s) {
  const div = document.createElement('div');
  div.textContent = s;
  return div.innerHTML;
}

function closeAllDropdowns(except) {
  document.querySelectorAll('.cat-list').forEach(el => { if (el !== except) el.hidden = true; });
}

function filterCategories(query) {
  const q = query.trim().toLowerCase();
  if (!q) return CATEGORIES.slice();
  return CATEGORIES.filter(c => c.toLowerCase().includes(q));
}

function renderDropdown(idx, query) {
  const wrap = document.querySelector(`.combo[data-idx="${idx}"]`);
  const list = wrap.querySelector('.cat-list');
  const matches = filterCategories(query);
  let html = '';
  if (matches.length === 0) {
    html += `<div class="cat-item empty-msg">Aucune catégorie trouvée</div>`;
  } else {
    for (const cat of matches) {
      html += `<div class="cat-item" data-cat="${escapeHtml(cat)}">${escapeHtml(cat)}</div>`;
    }
  }
  const trimmed = query.trim();
  if (trimmed && !CATEGORIES.some(c => c.toLowerCase() === trimmed.toLowerCase())) {
    html += `<div class="cat-item new-cat-option" data-newcat="${escapeHtml(trimmed)}">➕ Créer « ${escapeHtml(trimmed)} »</div>`;
  }
  html += `<div class="cat-item clear-option" data-clear="1">— À catégoriser —</div>`;
  list.innerHTML = html;
  list.querySelectorAll('.cat-item[data-cat]').forEach(el => {
    el.addEventListener('mousedown', (e) => { e.preventDefault(); selectCategory(idx, el.dataset.cat); });
  });
  const newEl = list.querySelector('.cat-item[data-newcat]');
  if (newEl) newEl.addEventListener('mousedown', (e) => { e.preventDefault(); createAndSelect(idx, newEl.dataset.newcat); });
  const clearEl = list.querySelector('.cat-item[data-clear]');
  if (clearEl) clearEl.addEventListener('mousedown', (e) => { e.preventDefault(); clearCategory(idx); });
}

function openDropdown(idx) {
  const wrap = document.querySelector(`.combo[data-idx="${idx}"]`);
  const input = wrap.querySelector('.cat-input');
  const list = wrap.querySelector('.cat-list');
  renderDropdown(idx, input.value);
  list.hidden = false;
  closeAllDropdowns(list);
}

function selectCategory(idx, cat) {
  selected[idx] = cat;
  const wrap = document.querySelector(`.combo[data-idx="${idx}"]`);
  const input = wrap.querySelector('.cat-input');
  input.value = cat;
  wrap.querySelector('.cat-list').hidden = true;
  updateRemaining();
}

function clearCategory(idx) {
  selected[idx] = '';
  const wrap = document.querySelector(`.combo[data-idx="${idx}"]`);
  const input = wrap.querySelector('.cat-input');
  input.value = '';
  wrap.querySelector('.cat-list').hidden = true;
  updateRemaining();
}

function createAndSelect(idx, name) {
  const clean = name.trim();
  if (!clean) return;
  if (!CATEGORIES.some(c => c.toLowerCase() === clean.toLowerCase())) {
    CATEGORIES.push(clean);
    sortCategories();
  }
  selectCategory(idx, clean);
}

function handleInput(idx, inputEl) {
  const wrap = document.querySelector(`.combo[data-idx="${idx}"]`);
  wrap.querySelector('.cat-list').hidden = false;
  renderDropdown(idx, inputEl.value);
  closeAllDropdowns(wrap.querySelector('.cat-list'));
}

function handleBlur(idx, inputEl) {
  setTimeout(() => {
    const wrap = document.querySelector(`.combo[data-idx="${idx}"]`);
    wrap.querySelector('.cat-list').hidden = true;
    inputEl.value = selected[idx] || '';
  }, 150);
}

function handleKeydown(idx, inputEl, e) {
  if (e.key === 'Escape') {
    inputEl.value = selected[idx] || '';
    document.querySelector(`.combo[data-idx="${idx}"] .cat-list`).hidden = true;
    inputEl.blur();
  } else if (e.key === 'Enter') {
    e.preventDefault();
    const matches = filterCategories(inputEl.value);
    if (matches.length >= 1) {
      selectCategory(idx, matches[0]);
    } else if (inputEl.value.trim()) {
      createAndSelect(idx, inputEl.value);
    }
    inputEl.blur();
  }
}

document.addEventListener('click', (e) => {
  if (!e.target.closest('.combo')) closeAllDropdowns();
});

function fmtMontant(m) {
  const s = m.toLocaleString('fr-FR', {minimumFractionDigits: 2, maximumFractionDigits: 2});
  return s + ' €';
}

function render() {
  const tbody = document.getElementById('tbody');
  tbody.innerHTML = '';
  TX.forEach((tx, idx) => {
    const tr = document.createElement('tr');
    tr.className = tx.suggestion ? 'suggested' : 'none';
    const montantClass = tx.montant < 0 ? 'neg' : 'pos';
    tr.innerHTML = `
      <td>${tx.date}</td>
      <td class="libelle" title="${tx.libelle}">${tx.suggestion ? '<span class="star">★</span>' : ''}${tx.libelle}</td>
      <td class="montant ${montantClass}">${fmtMontant(tx.montant)}</td>
      <td class="cat">
        <div class="combo" data-idx="${idx}">
          <input type="text" class="cat-input" autocomplete="off" placeholder="— À catégoriser —" value="${escapeHtml(tx.suggestion || '')}">
          <div class="cat-list" hidden></div>
        </div>
      </td>
    `;
    const input = tr.querySelector('.cat-input');
    input.addEventListener('focus', () => openDropdown(idx));
    input.addEventListener('input', () => handleInput(idx, input));
    input.addEventListener('blur', () => handleBlur(idx, input));
    input.addEventListener('keydown', (e) => handleKeydown(idx, input, e));
    tbody.appendChild(tr);
  });
  updateRemaining();
}

function updateRemaining() {
  let n = 0;
  selected.forEach(v => { if (!v) n++; });
  document.getElementById('remaining').textContent = n;
}

function submitAll() {
  const categorized = selected.map(v => v || null);

  const newCategories = CATEGORIES.filter(c => !DATA.categories.includes(c));

  document.getElementById('submit-btn').disabled = true;
  document.getElementById('submit-btn').textContent = 'Génération en cours...';

  fetch('/submit', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({
      categorized: categorized,
      remember: document.getElementById('remember').checked,
      new_categories: newCategories
    })
  }).then(r => r.json()).then(() => {
    document.getElementById('main-view').style.display = 'none';
    document.getElementById('footer').style.display = 'none';
    document.getElementById('done').style.display = 'block';
  });
}

render();
</script>
</body>
</html>
"""

DONE_NOTE = "Fenêtre à fermer manuellement — le script continue dans le terminal."


def _find_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


class _ServerThread(threading.Thread):
    def __init__(self, app, port):
        super().__init__(daemon=True)
        self.server = make_server("127.0.0.1", port, app)

    def run(self):
        self.server.serve_forever()

    def shutdown(self):
        self.server.shutdown()


def run_web_categorizer(transactions: list[dict], cfg: dict, month_label: str,
                         suggest_fn, extract_signature_fn) -> tuple[list[str], bool]:
    """Ouvre une page web locale listant toutes les transactions avec une
    catégorie suggérée pré-sélectionnée. Retourne (liste des catégories
    choisies, True si la config a été enrichie de nouveaux mots-clés)."""

    categories = cfg["categories"]
    tx_payload = []
    for tx in transactions:
        suggestion = suggest_fn(tx, cfg)
        tx_payload.append({
            "date": tx["date"].strftime("%d/%m/%Y") if tx["date"] else "?",
            "libelle": tx["libelle"],
            "montant": round(tx["montant"], 2),
            "suggestion": suggestion,
        })

    app = Flask(__name__)
    result_holder = {}
    done_event = threading.Event()

    @app.route("/")
    def index():
        data_json = json.dumps({"categories": categories, "transactions": tx_payload}, ensure_ascii=False)
        return render_template_string(PAGE_TEMPLATE, month_label=month_label, n=len(transactions), data_json=data_json)

    @app.route("/submit", methods=["POST"])
    def submit():
        payload = request.get_json()
        result_holder["categorized"] = payload["categorized"]
        result_holder["remember"] = bool(payload.get("remember", True))
        result_holder["new_categories"] = payload.get("new_categories", [])
        done_event.set()
        return jsonify({"status": "ok"})

    port = _find_free_port()
    server = _ServerThread(app, port)
    server.start()

    url = f"http://127.0.0.1:{port}"
    print(f"\n🌐 Ouverture du navigateur : {url}")
    print("   (si la fenêtre ne s'ouvre pas automatiquement, ouvre ce lien manuellement)")
    try:
        webbrowser.open(url)
    except Exception:
        pass

    done_event.wait()
    server.shutdown()

    raw_categorized = result_holder["categorized"]
    remember = result_holder["remember"]
    new_categories = result_holder.get("new_categories", [])

    categorized = []
    config_changed = False

    for new_cat in new_categories:
        if new_cat and new_cat not in cfg["categories"]:
            cfg["categories"].append(new_cat)
            config_changed = True

    for tx, chosen in zip(transactions, raw_categorized):
        if not chosen:
            categorized.append(UNCATEGORIZED_LABEL)
            continue
        categorized.append(chosen)
        suggestion = suggest_fn(tx, cfg)
        if remember and chosen != suggestion:
            signature = extract_signature_fn(tx["libelle"])
            if signature:
                cfg.setdefault("keywords", {}).setdefault(chosen, [])
                if signature not in cfg["keywords"][chosen]:
                    cfg["keywords"][chosen].append(signature)
                    cfg["_keywords_norm"].setdefault(chosen, []).append(signature)
                    config_changed = True

    print(f"   {DONE_NOTE}")
    return categorized, config_changed
