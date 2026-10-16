#!/usr/bin/env python3
"""
SEI Programações - coletor em executável único (somente leitura).

Abre o Edge (ou Chrome) visível, você faz o login (senha/MFA), deixa a lista de
"Programações de Contratação" na tela e pressiona ENTER. O programa lê a tela
"Consultar" de cada programação e grava uma planilha Excel (.xls ou .xlsx) ao lado do .exe.
Nada é gravado no SEI e nenhuma credencial/cookie é guardado em disco.

Interface gráfica: URL, formato (.xls/.xlsx/.json) e pasta de destino são escolhidos na janela.
"""
import json
import os
import queue
import re
import threading
import time
import traceback
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from playwright.sync_api import TimeoutError as PWTimeout, sync_playwright

try:  # gerado por gerar-exe.bat; ausente ao rodar direto do fonte
    from _versao import COMPILACAO, VERSAO
except ImportError:
    VERSAO = "1.1"
    COMPILACAO = "fonte (" + datetime.fromtimestamp(Path(__file__).stat().st_mtime).strftime("%d/%m/%Y %H:%M") + ")"
RODAPE = "STI/CJF - Secretaria de Tecnologia da Informação - Conselho da Justiça Federal"
NOME_LOG = "sei_programacoes.log"

FORMATOS = {"Excel 97-2003 (.xls)": "xls", "Excel (.xlsx)": "xlsx", "JSON (.json)": "json"}
CANAIS = {"Microsoft Edge": "msedge", "Google Chrome": "chrome", "Chromium (Playwright)": "chromium"}
PRIMEIRAS = ["Identificador", "Unidade", "PT"]

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



def colunas(registros):
    cols = []
    for r in registros:
        for k in r:
            if k not in cols:
                cols.append(k)
    return [c for c in PRIMEIRAS if c in cols] + [c for c in cols if c not in PRIMEIRAS]


def gravar(registros, arq, formato):
    cols = colunas(registros)
    if formato == "json":
        arq.write_text(json.dumps(registros, ensure_ascii=False, indent=1), encoding="utf-8")
    elif formato == "xlsx":
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font
        wb = Workbook()
        ws = wb.active
        ws.title = "Programações"
        ws.append(cols)
        for r in registros:
            ws.append([r.get(c, "") for c in cols])
        for c in ws[1]:
            c.font = Font(bold=True)
        for col in ws.columns:
            larg = max(len(str(c.value or "").split("\n")[0]) for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(max(larg + 2, 10), 60)
            for c in col[1:]:
                c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.freeze_panes = "A2"
        wb.save(str(arq))
    else:  # xls (Excel 97-2003)
        import xlwt
        wb = xlwt.Workbook(encoding="utf-8")
        ws = wb.add_sheet("Programações")
        neg = xlwt.easyxf("font: bold on")
        qb = xlwt.easyxf("alignment: wrap on, vert top")
        for j, c in enumerate(cols):
            ws.write(0, j, c[:32767], neg)
            ws.col(j).width = 256 * 28
        for i, r in enumerate(registros, 1):
            for j, c in enumerate(cols):
                ws.write(i, j, str(r.get(c, ""))[:32767], qb)
        ws.set_panes_frozen(True)
        ws.horz_split_pos = 1
        wb.save(str(arq))


def coletar(cfg, log, progresso, esperar_lista, cancelado):
    """Roda no thread de trabalho. cfg: url, formato, pasta, limite, canal, pausa."""
    with sync_playwright() as pw:
        try:
            nav = pw.chromium.launch(headless=False) if cfg["canal"] == "chromium" \
                else pw.chromium.launch(headless=False, channel=cfg["canal"])
        except Exception as e:  # noqa: BLE001
            raise RuntimeError("Não consegui abrir o navegador '%s': %s\nTente outro navegador na lista."
                               % (cfg["canal"], str(e).splitlines()[0]))
        ctx = nav.new_context(viewport={"width": 1400, "height": 900})
        page = ctx.new_page()
        page.goto(cfg["url"])
        log("Navegador aberto. Faça o login no SEI, abra Contratações > Programações de Contratação,")
        log("filtre Plano/Unidade, clique em Pesquisar e, com a LISTA na tela, clique em 'Continuar' aqui.")
        esperar_lista()
        if cancelado():
            nav.close()
            return None

        quadro, lista = None, None
        limite_t = time.time() + 20  # a lista pode ainda estar carregando: tenta por até 20 s
        while not lista:
            for p in [page] + [x for x in ctx.pages if x is not page]:
                quadro, lista = achar_quadro(p, JS_LISTA, lambda r: r and r["itens"])
                if lista:
                    page = p
                    break
            if lista or time.time() > limite_t or cancelado():
                break
            time.sleep(1.5)
        if not lista:
            diag = Path(cfg["pasta"]) / "diagnostico"
            diag.mkdir(parents=True, exist_ok=True)
            log("Não achei a lista. Abas abertas e conteúdo visto:")
            for i, p in enumerate(ctx.pages):
                try:
                    log("  aba %d: %s | %s" % (i + 1, p.title(), p.url))
                    for k, f in enumerate(p.frames):
                        try:
                            txt = f.evaluate("() => [document.querySelectorAll('table').length, "
                                             "/Identificador/i.test(document.body.innerText), "
                                             "document.body.innerText.length]")
                            log("    quadro %d: tabelas=%s, tem 'Identificador'=%s, texto=%s car. (%s)"
                                % (k + 1, txt[0], txt[1], txt[2], f.url[:90]))
                            (diag / ("aba%d_quadro%d.html" % (i + 1, k + 1))).write_text(f.content(), encoding="utf-8")
                        except Exception:  # noqa: BLE001
                            pass
                except Exception:  # noqa: BLE001
                    pass
            log("HTML das telas salvo em: %s" % diag)
            nav.close()
            raise RuntimeError("Não encontrei a tabela com a coluna 'Identificador'.\n"
                               "Confira se a lista tem resultados (clique em Pesquisar) e se está numa aba do "
                               "navegador aberto pelo programa.\nDiagnóstico salvo em:\n%s" % diag)

        itens = [i for i in lista["itens"] if any(RE_ID.match(v) for v in i["reg"].values())]
        log("Lista: %d programações." % len(itens))
        if cfg["limite"]:
            itens = itens[: cfg["limite"]]
            log("Limite de teste aplicado: %d programações." % len(itens))

        resultado, erros = [], []
        for n, it in enumerate(itens, 1):
            if cancelado():
                log("Cancelado pelo usuário.")
                break
            ident = next((v for v in it["reg"].values() if RE_ID.match(v)), "?")
            ac = next((x for x in it["acoes"] if re.search(r"consult", x["titulo"], re.I)), None) or \
                 next((x for x in it["acoes"] if x["href"] and not x["href"].startswith("javascript")), None)
            det = None
            try:
                if ac and ac["abs"] and not ac["href"].startswith("javascript"):
                    dp = ctx.new_page()
                    dp.goto(ac["abs"])
                    dp.wait_for_load_state("networkidle")
                    _, det = achar_quadro(dp, JS_DETALHE, lambda r: r and len(r) >= 5)
                    dp.close()
                else:
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
                for k, v in it["reg"].items():
                    if re.match(r"unidade", k, re.I):
                        reg["Unidade"] = v
                    elif re.fullmatch(r"PT", k.strip(), re.I):
                        reg["PT"] = v
                reg.setdefault("Identificador", ident)
                resultado.append(reg)
            log("[%d/%d] %s: %s" % (n, len(itens), ident, "ok" if det else "FALHOU"))
            progresso(n, len(itens))
            time.sleep(cfg["pausa"])
        nav.close()

    if not resultado:
        log("Nenhuma programação lida.")
        return None
    pasta = Path(cfg["pasta"])
    pasta.mkdir(parents=True, exist_ok=True)
    arq = pasta / ("programacoes_%s.%s" % (datetime.now().strftime("%Y%m%d-%H%M%S"), cfg["formato"]))
    gravar(resultado, arq, cfg["formato"])
    log("\nGravado: %s (%d programações)" % (arq, len(resultado)))
    if erros:
        log("Pendentes: " + "; ".join("%s: %s" % e for e in erros))
    return arq


INSTRUCOES = (
    "Como usar\n"
    "1. Confira a URL do SEI e escolha o formato e a pasta onde a planilha será salva.\n"
    "2. Clique em 'Abrir SEI'. O Edge abre em outra janela.\n"
    "3. Faça o login (senha e MFA), abra Contratações > Programações de Contratação,\n"
    "    filtre o Plano (ex.: 2027) e a Unidade Requisitante (ex.: CEGOVE) e clique em Pesquisar.\n"
    "4. Com a LISTA na tela, volte a esta janela e clique em 'Continuar'.\n"
    "5. Aguarde a leitura (cerca de 2 minutos para 57 programações). Não feche o navegador.\n"
    "6. Ao final, a planilha é gravada na pasta escolhida.\n"
    "O programa só abre telas de consulta; nada é alterado no SEI e nenhuma senha ou cookie é guardado."
)
EXEMPLO_URL = "https://sei.cjf.jus.br/sei/"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SEI - Programações de Contratação  v%s" % VERSAO)
        self.geometry("760x640")
        self.minsize(640, 560)
        self.fila = queue.Queue()
        self.evt_lista = threading.Event()
        self.cancel = threading.Event()
        self.trabalhando = False
        self.cfg = {}
        self.protocol("WM_DELETE_WINDOW", self.sair)

        f = ttk.Frame(self, padding=12)
        f.pack(fill="both", expand=True)
        f.columnconfigure(1, weight=1)
        f.rowconfigure(6, weight=1)

        ttk.Label(f, text=INSTRUCOES, justify="left", background="#f3f6fa", relief="groove",
                  padding=8).grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))

        ttk.Label(f, text="URL do SEI:").grid(row=1, column=0, sticky="w")
        self.url = ttk.Entry(f, foreground="gray")
        self.url.insert(0, EXEMPLO_URL)
        self._ph = True
        self.url.bind("<FocusIn>", self._limpar_ph)
        self.url.bind("<FocusOut>", self._por_ph)
        self.url.grid(row=1, column=1, columnspan=2, sticky="ew", padx=6, pady=3)
        ttk.Label(f, text="Modelo: " + EXEMPLO_URL + "  (endereço raiz do SEI, terminando em /sei/)",
                  foreground="gray").grid(row=2, column=1, columnspan=2, sticky="w", padx=6)

        ttk.Label(f, text="Exportar como:").grid(row=3, column=0, sticky="w")
        self.fmt = ttk.Combobox(f, state="readonly", width=38, values=list(FORMATOS))
        self.fmt.set("Excel (.xlsx)")
        self.fmt.grid(row=3, column=1, sticky="w", padx=6, pady=3)

        ttk.Label(f, text="Salvar em:").grid(row=4, column=0, sticky="w")
        self.pasta = ttk.Entry(f)
        self.pasta.insert(0, str(Path.home() / "Documents"))
        self.pasta.grid(row=4, column=1, sticky="ew", padx=6, pady=3)
        ttk.Button(f, text="Escolher...", command=self.escolher).grid(row=4, column=2, pady=3)

        op = ttk.Frame(f)
        op.grid(row=5, column=0, columnspan=3, sticky="w", pady=3)
        ttk.Label(op, text="Navegador:").pack(side="left")
        self.canal = ttk.Combobox(op, state="readonly", width=16, values=list(CANAIS))
        self.canal.current(0)
        self.canal.pack(side="left", padx=6)
        ttk.Label(op, text="Limite (só para teste, vazio = todas):").pack(side="left", padx=(12, 0))
        self.limite = ttk.Entry(op, width=5)
        self.limite.pack(side="left", padx=6)

        self.log = scrolledtext.ScrolledText(f, height=10, state="disabled", wrap="word")
        self.log.grid(row=6, column=0, columnspan=3, sticky="nsew", pady=8)
        self.barra = ttk.Progressbar(f, mode="determinate")
        self.barra.grid(row=7, column=0, columnspan=3, sticky="ew")

        ttk.Label(f, text="Versão %s  |  compilação: %s  |  %s" % (VERSAO, COMPILACAO, RODAPE),
                  foreground="gray").grid(row=9, column=0, columnspan=3, pady=(10, 0))

        bt = ttk.Frame(f)
        bt.grid(row=8, column=0, columnspan=3, pady=(10, 0))
        self.b_abrir = ttk.Button(bt, text="Abrir SEI", command=self.iniciar)
        self.b_abrir.pack(side="left", padx=4)
        self.b_cont = ttk.Button(bt, text="Continuar (a lista está na tela)", state="disabled",
                                 command=self.continuar)
        self.b_cont.pack(side="left", padx=4)
        self.b_cancel = ttk.Button(bt, text="Cancelar", state="disabled", command=self.cancelar)
        self.b_cancel.pack(side="left", padx=4)
        self.after(150, self.drenar)

    def _limpar_ph(self, _):
        if self._ph:
            self.url.delete(0, "end")
            self.url.configure(foreground="black")
            self._ph = False

    def _por_ph(self, _):
        if not self.url.get().strip():
            self.url.insert(0, EXEMPLO_URL)
            self.url.configure(foreground="gray")
            self._ph = True

    def escolher(self):
        d = filedialog.askdirectory(initialdir=self.pasta.get() or str(Path.home()), title="Pasta para salvar a planilha")
        if d:
            self.pasta.delete(0, "end")
            self.pasta.insert(0, d)

    def escrever(self, txt):
        self.log.configure(state="normal")
        self.log.insert("end", txt + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def drenar(self):
        try:
            while True:
                tipo, a = self.fila.get_nowait()
                if tipo == "log":
                    self.escrever(a)
                elif tipo == "prog":
                    self.barra.configure(maximum=a[1], value=a[0])
                elif tipo == "espera":
                    self.b_cont.configure(state="normal")
                elif tipo == "fim":
                    self.fim(a)
        except queue.Empty:
            pass
        self.after(150, self.drenar)

    def iniciar(self):
        url = "" if self._ph else self.url.get().strip()
        url = url or EXEMPLO_URL
        if not re.match(r"^https?://", url):
            messagebox.showerror("URL inválida", "A URL deve começar com https://\nModelo: " + EXEMPLO_URL)
            return
        lim = self.limite.get().strip()
        if lim and not lim.isdigit():
            messagebox.showerror("Limite inválido", "Digite apenas números ou deixe vazio.")
            return
        cfg = {"url": url, "formato": FORMATOS[self.fmt.get()], "pasta": self.pasta.get().strip(),
               "limite": int(lim) if lim else 0, "canal": CANAIS[self.canal.get()], "pausa": 1.2}
        if not cfg["pasta"]:
            messagebox.showerror("Pasta", "Escolha a pasta onde salvar.")
            return
        self.cfg = cfg
        self.evt_lista.clear()
        self.cancel.clear()
        self.trabalhando = True
        self.b_abrir.configure(state="disabled")
        self.b_cancel.configure(state="normal")
        self.barra.configure(value=0)
        threading.Thread(target=self.trabalho, args=(cfg,), daemon=True).start()

    def trabalho(self, cfg):
        arq_log = None
        try:
            pasta = Path(cfg["pasta"])
            pasta.mkdir(parents=True, exist_ok=True)
            arq_log = pasta / NOME_LOG
        except OSError:
            pass

        def log(m):
            self.fila.put(("log", m))
            if arq_log:
                try:
                    with open(arq_log, "a", encoding="utf-8") as fh:
                        fh.write("%s %s\n" % (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), m))
                except OSError:
                    pass

        log("=== SEI-Programacoes v%s, compilação %s | %s | formato=%s navegador=%s limite=%s ===" % (
            VERSAO, COMPILACAO, cfg["url"], cfg["formato"], cfg["canal"], cfg["limite"] or "todas"))
        if arq_log:
            self.fila.put(("log", "Log gravado em: %s" % arq_log))

        def esperar():
            self.fila.put(("espera", None))
            while not (self.evt_lista.is_set() or self.cancel.is_set()):
                time.sleep(0.2)
        try:
            arq = coletar(cfg, log, lambda n, t: self.fila.put(("prog", (n, t))),
                          esperar, self.cancel.is_set)
            self.fila.put(("fim", ("ok", arq)))
        except Exception as e:  # noqa: BLE001
            log("ERRO: " + str(e))
            log(traceback.format_exc())
            self.fila.put(("fim", ("erro", str(e))))

    def continuar(self):
        """O limite vale o que estiver na caixa neste momento (pode ter sido digitado após 'Abrir SEI')."""
        lim = self.limite.get().strip()
        if lim and not lim.isdigit():
            messagebox.showerror("Limite inválido", "Digite apenas números ou deixe vazio.")
            return
        self.cfg["limite"] = int(lim) if lim else 0
        self.fila.put(("log", "Limite: %s" % (lim or "todas")))
        self.evt_lista.set()

    def cancelar(self):
        self.cancel.set()

    def sair(self):
        """Fechar a janela encerra o processo (o Playwright/Edge em andamento não pode deixá-lo órfão)."""
        self.cancel.set()
        self.evt_lista.set()
        try:
            self.destroy()
        finally:
            os._exit(0)

    def fim(self, res):
        self.trabalhando = False
        self.b_abrir.configure(state="normal")
        self.b_cont.configure(state="disabled")
        self.b_cancel.configure(state="disabled")
        if res[0] == "erro":
            self.escrever("ERRO: " + res[1])
            messagebox.showerror("Erro", res[1])
        elif res[1]:
            if messagebox.askyesno("Concluído", "Planilha gravada:\n%s\n\nAbrir a pasta?" % res[1]):
                try:
                    os.startfile(str(Path(res[1]).parent))  # Windows
                except AttributeError:
                    pass


if __name__ == "__main__":
    App().mainloop()
