#!/usr/bin/env python3
"""
Coletor das "Programações de Contratação" do módulo Contratações do SEI (leitura apenas).

Roda na ESTAÇÃO (WSL/Windows), com a sessão do próprio usuário: o navegador abre visível, você faz o
login (senha/MFA) e deixa a lista na tela; o script lê cada tela "Consultar Programação de Contratação"
e grava um JSON que o plugin PCA importa (PCA > Reconciliação SEI > Carregar do SEI).

Uso:
  python3 coletar_programacoes.py --url https://sei.cjf.jus.br/sei/ [--saida pasta] [--limite N] [--dump]

  --dump   salva o HTML da lista e da 1ª tela de detalhe (para ajustar seletores) e para.
Nenhuma credencial é guardada; o estado do navegador fica em <saida>/estado_navegador.json (apague ao terminar).
Nada é gravado no SEI: o script só abre telas de consulta (ícone "Consultar").
"""
import argparse
import json
import re
import sys
import time
from datetime import datetime
from pathlib import Path

from playwright.sync_api import TimeoutError as PWTimeout, sync_playwright

RE_ID = re.compile(r"^PC\s*\d+\s*/\s*\d{2,4}$", re.I)

JS_LISTA = """
() => {
  const tabelas = [...document.querySelectorAll('table')];
  for (const t of tabelas) {
    const linhas = [...t.querySelectorAll('tr')];
    const cab = linhas.find(r => /Identificador/i.test(r.textContent) && r.querySelectorAll('th,td').length >= 4);
    if (!cab) continue;
    const cols = [...cab.querySelectorAll('th,td')].map(c => c.textContent.replace(/\\s+/g,' ').trim());
    const itens = [];
    for (const r of linhas) {
      if (r === cab) continue;
      const cel = [...r.querySelectorAll('td')];
      if (cel.length < 4) continue;
      const reg = {};
      cel.forEach((c, i) => { reg[cols[i] || ('col' + i)] = c.textContent.replace(/\\s+/g,' ').trim(); });
      const acoes = [...r.querySelectorAll('a')].map(a => ({
        href: a.getAttribute('href') || '', onclick: a.getAttribute('onclick') || '',
        titulo: (a.getAttribute('title') || (a.querySelector('img') && (a.querySelector('img').getAttribute('title') || a.querySelector('img').getAttribute('alt'))) || '').trim(),
        abs: a.href || ''
      }));
      itens.push({reg, acoes});
    }
    return {cols, itens};
  }
  return null;
}
"""

JS_DETALHE = """
() => {
  const out = {};
  const ctrls = [...document.querySelectorAll('input,select,textarea')].filter(e => e.type !== 'hidden' && e.type !== 'button' && e.type !== 'submit');
  const valor = el => {
    if (!el) return null;
    if (el.tagName === 'SELECT' && (el.multiple || el.size > 1)) {
      // lista de vinculações (selVinculacoes*): os itens vinculados são TODAS as opções, não as selecionadas
      return [...el.options].map(o => o.text.replace(/\\s+/g, ' ').trim()).filter(Boolean).join('\\n');
    }
    if (el.tagName === 'SELECT') { const o = el.options[el.selectedIndex]; return o ? o.text.trim() : ''; }
    if (el.type === 'checkbox') return el.checked ? '1' : '0';
    if (el.type === 'radio') return el.checked ? (el.value || '1') : null;
    return (el.value || '').trim();
  };
  for (const l of document.querySelectorAll('label')) {
    const rotulo = l.textContent.replace(/\\s+/g, ' ').trim().replace(/:$/, '');
    if (!rotulo) continue;
    let c = null;
    if (l.htmlFor) c = document.getElementById(l.htmlFor);
    if (!c) c = l.querySelector('input,select,textarea');
    if (!c) {
      const seguintes = ctrls.filter(e => l.compareDocumentPosition(e) & Node.DOCUMENT_POSITION_FOLLOWING);
      const anteriores = ctrls.filter(e => l.compareDocumentPosition(e) & Node.DOCUMENT_POSITION_PRECEDING);
      const ant = anteriores[anteriores.length - 1];
      // caixa de seleção vem ANTES do rótulo ("[x] Serviço"); os demais campos vêm depois
      c = (ant && ant.type === 'checkbox' && !seguintes.length) ? ant : (seguintes[0] || ant || null);
      if (ant && ant.type === 'checkbox' && seguintes[0] && seguintes[0].type !== 'checkbox') {
        const dA = l.compareDocumentPosition(ant), r = l.getBoundingClientRect(), ra = ant.getBoundingClientRect();
        if (Math.abs(ra.top - r.top) < 12 && ra.left < r.left) c = ant;
      }
    }
    const v = valor(c);
    if (v !== null && !(rotulo in out)) out[rotulo] = v;
  }
  return out;
}
"""


def achar_quadro(page, js, teste):
    """Roda `js` em cada frame até `teste(resultado)` aceitar; devolve (frame, resultado)."""
    for f in page.frames:
        try:
            r = f.evaluate(js)
        except Exception:  # noqa: BLE001
            continue
        if teste(r):
            return f, r
    return None, None


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--url", required=True, help="URL raiz do SEI (ex.: https://sei.cjf.jus.br/sei/)")
    ap.add_argument("--saida", default="saida-sei", help="pasta de saída")
    ap.add_argument("--limite", type=int, default=0, help="só as N primeiras (teste)")
    ap.add_argument("--dump", action="store_true", help="salva HTML da lista e do 1º detalhe e para")
    ap.add_argument("--canal", default="msedge", help="navegador instalado: msedge (padrão), chrome, ou chromium (do Playwright)")
    ap.add_argument("--pausa", type=float, default=1.2, help="segundos entre telas (não sobrecarregar o SEI)")
    a = ap.parse_args()

    saida = Path(a.saida)
    saida.mkdir(parents=True, exist_ok=True)
    estado = saida / "estado_navegador.json"

    with sync_playwright() as pw:
        try:
            nav = pw.chromium.launch(headless=False) if a.canal == "chromium" \
                else pw.chromium.launch(headless=False, channel=a.canal)
        except Exception as e:  # noqa: BLE001
            sys.exit(f"Não consegui abrir o navegador '{a.canal}': {str(e).splitlines()[0]}\n"
                     "Tente --canal chrome, ou instale o do Playwright: python -m playwright install chromium")
        ctx = nav.new_context(storage_state=str(estado) if estado.exists() else None, viewport={"width": 1400, "height": 900})
        page = ctx.new_page()
        page.goto(a.url)
        print("\n1) Faça o login no SEI (senha/MFA).")
        print("2) Abra: Contratações > Programações de Contratação.")
        print("3) Filtre Plano (ex.: 2027) e Unidade Requisitante (CEGOVE) e clique em Pesquisar.")
        input("4) Com a LISTA na tela, volte aqui e pressione ENTER... ")
        ctx.storage_state(path=str(estado))

        quadro, lista = achar_quadro(page, JS_LISTA, lambda r: r and r["itens"])
        if not lista:
            for p in ctx.pages:  # a lista pode estar em outra aba
                quadro, lista = achar_quadro(p, JS_LISTA, lambda r: r and r["itens"])
                if lista:
                    page = p
                    break
        if not lista:
            sys.exit("Não encontrei a tabela com a coluna 'Identificador'. Rode com --dump e envie o HTML salvo.")

        itens = [i for i in lista["itens"] if any(RE_ID.match(v) for v in i["reg"].values())]
        print(f"Lista: {len(itens)} programações (colunas: {', '.join(lista['cols'])})")
        if a.dump:
            (saida / "lista.html").write_text(quadro.content(), encoding="utf-8")
            (saida / "lista.json").write_text(json.dumps(lista, ensure_ascii=False, indent=1), encoding="utf-8")
        if a.limite:
            itens = itens[: a.limite]

        resultado, erros = [], []
        for n, it in enumerate(itens, 1):
            ident = next((v for v in it["reg"].values() if RE_ID.match(v)), "?")
            # ícone de consulta: título "Consult…", senão o primeiro link com href
            ac = next((x for x in it["acoes"] if re.search(r"consult", x["titulo"], re.I)), None) or \
                 next((x for x in it["acoes"] if x["href"] and not x["href"].startswith("javascript")), None)
            det = None
            try:
                if ac and ac["abs"] and not ac["href"].startswith("javascript"):
                    dp = ctx.new_page()
                    dp.goto(ac["abs"])
                    dp.wait_for_load_state("networkidle")
                    _, det = achar_quadro(dp, JS_DETALHE, lambda r: r and len(r) >= 5)
                    if a.dump and n == 1:
                        (saida / "detalhe.html").write_text(dp.content(), encoding="utf-8")
                    dp.close()
                else:
                    # sem href: clica no ícone da linha e lê o diálogo/popup que abrir
                    linha = quadro.locator("tr", has_text=ident).first
                    with ctx.expect_page(timeout=4000) as np:
                        linha.locator("a").first.click()
                    dp = np.value
                    dp.wait_for_load_state("networkidle")
                    _, det = achar_quadro(dp, JS_DETALHE, lambda r: r and len(r) >= 5)
                    dp.close()
            except (PWTimeout, Exception) as e:  # noqa: BLE001
                erros.append((ident, str(e).split("\n")[0]))
            if not det:
                erros.append((ident, "detalhe não lido"))
            else:
                reg = dict(det)
                # a lista traz a unidade; a tela de detalhe, não
                for k, v in it["reg"].items():
                    if re.match(r"unidade", k, re.I):
                        reg["Unidade"] = v
                    elif re.fullmatch(r"PT", k.strip(), re.I):
                        reg["PT"] = v  # Programa de Trabalho (a tela de detalhe usa um seletor sem rótulo)
                reg.setdefault("Identificador", ident)
                resultado.append(reg)
            print(f"  [{n}/{len(itens)}] {ident}: {'ok' if det else 'FALHOU'}")
            if a.dump and n == 1:
                break
            time.sleep(a.pausa)

        if resultado:
            arq = saida / f"programacoes_{datetime.now():%Y%m%d-%H%M%S}.json"
            arq.write_text(json.dumps(resultado, ensure_ascii=False, indent=1), encoding="utf-8")
            print(f"\nGravado: {arq} ({len(resultado)} programações)")
        if erros:
            print("Pendentes:", "; ".join(f"{i}: {m}" for i, m in erros))
        nav.close()


if __name__ == "__main__":
    main()
