"""
Janela para converter POP em BPMN, sem usar o terminal.
Abra pelo arquivo abrir.bat (dois cliques).
"""
import json
import os
import queue
import threading
import tkinter as tk
import traceback
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext

import pop_para_bpmn as core

PASTA = Path(__file__).resolve().parent
SAIDA = PASTA / "diagramas"
core.carregar_env()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("POP → BPMN")
        self.geometry("760x640")
        self.minsize(640, 520)
        self.arquivo = None
        self.modo_json = False
        self.fila = queue.Queue()   # a thread de fundo só pede; quem mexe na tela é a janela

        pad = {"padx": 14, "pady": 6}

        tk.Label(self, text="Conversor de POP para BPMN", font=("Segoe UI", 16, "bold")).pack(anchor="w", **pad)

        # 1) documento
        quadro1 = tk.LabelFrame(self, text=" 1. Escolha o arquivo ", padx=10, pady=8)
        quadro1.pack(fill="x", **pad)
        linha = tk.Frame(quadro1)
        linha.pack(fill="x")
        tk.Button(linha, text="Escolher POP (.docx / .pdf)…", command=self.escolher_pop,
                  font=("Segoe UI", 10, "bold"), padx=8, pady=4).pack(side="left")
        tk.Button(linha, text="Usar JSON pronto (sem IA)…", command=self.escolher_json,
                  padx=8, pady=4).pack(side="left", padx=8)
        self.lbl_arquivo = tk.Label(quadro1, text="Nenhum arquivo escolhido", fg="#666", anchor="w")
        self.lbl_arquivo.pack(fill="x", pady=(6, 0))

        # 2) IA e chave
        quadro2 = tk.LabelFrame(self, text=" 2. IA e chave da API (só para ler POP com IA) ", padx=10, pady=8)
        quadro2.pack(fill="x", **pad)
        topo2 = tk.Frame(quadro2)
        topo2.pack(fill="x", pady=(0, 6))
        tk.Label(topo2, text="IA:").pack(side="left")
        self.var_prov = tk.StringVar(value="Gemini (grátis)")
        tk.OptionMenu(topo2, self.var_prov, *core.PROVEDORES.keys(), command=self._trocou_provedor).pack(side="left", padx=6)
        self.var_chave = tk.StringVar(value=os.environ.get("GEMINI_API_KEY", ""))
        self.var_lembrar = tk.BooleanVar(value=False)
        linha_chave = tk.Frame(quadro2)
        linha_chave.pack(fill="x")
        tk.Checkbutton(linha_chave, text="Lembrar neste computador", variable=self.var_lembrar).pack(side="right", padx=(8, 0))
        tk.Entry(linha_chave, textvariable=self.var_chave, show="•").pack(side="left", fill="x", expand=True)
        tk.Label(quadro2, text="Atenção: no plano grátis o Google pode usar o conteúdo enviado para melhorar seus produtos. "
                 "Não use POPs confidenciais ou com dados pessoais.",
                 fg="#a05a00", font=("Segoe UI", 8), wraplength=660, justify="left").pack(anchor="w", pady=(6, 0))

        # 3) opções e botão
        quadro3 = tk.Frame(self)
        quadro3.pack(fill="x", **pad)
        self.var_raias = tk.BooleanVar(value=True)
        tk.Checkbutton(quadro3, text="Separar por raias (responsáveis)", variable=self.var_raias).pack(side="left")
        self.btn_gerar = tk.Button(quadro3, text="Gerar diagrama", command=self.gerar,
                                   font=("Segoe UI", 11, "bold"), bg="#1f3864", fg="white", padx=16, pady=6)
        self.btn_gerar.pack(side="right")

        # log
        self.log_box = scrolledtext.ScrolledText(self, height=14, state="disabled", font=("Consolas", 9))
        self.log_box.pack(fill="both", expand=True, **pad)
        self.log("Escolha um arquivo e clique em \"Gerar diagrama\".")
        self.after(100, self._processar_fila)

    # ---------- utilidades ----------
    def agendar(self, fn):
        """Pode ser chamado de qualquer thread: executa fn na thread da janela."""
        self.fila.put(fn)

    def _processar_fila(self):
        while True:
            try:
                fn = self.fila.get_nowait()
            except queue.Empty:
                break
            fn()
        self.after(100, self._processar_fila)

    def log(self, msg):
        def _escrever():
            self.log_box.configure(state="normal")
            self.log_box.insert("end", msg + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        self.agendar(_escrever)

    def _var_chave(self):
        return core.PROVEDORES[self.var_prov.get()][1]

    def _trocou_provedor(self, _=None):
        self.var_chave.set(os.environ.get(self._var_chave(), ""))

    def escolher_pop(self):
        caminho = filedialog.askopenfilename(
            title="Escolha o POP", filetypes=[("Documentos", "*.docx *.pdf"), ("Todos", "*.*")])
        if caminho:
            self.arquivo, self.modo_json = Path(caminho), False
            self.lbl_arquivo.config(text=f"POP: {self.arquivo.name}", fg="#000")

    def escolher_json(self):
        caminho = filedialog.askopenfilename(
            title="Escolha o JSON do processo", filetypes=[("JSON", "*.json"), ("Todos", "*.*")])
        if caminho:
            self.arquivo, self.modo_json = Path(caminho), True
            self.lbl_arquivo.config(text=f"JSON (sem IA): {self.arquivo.name}", fg="#000")

    # ---------- etapa 1: ler + IA (em segundo plano) ----------
    def gerar(self):
        if not self.arquivo:
            messagebox.showwarning("Falta o arquivo", "Escolha um POP ou um JSON primeiro.")
            return
        if not self.modo_json and not self.var_chave.get().strip():
            messagebox.showwarning("Falta a chave",
                                   "Para ler um POP com IA é preciso colar a chave da API.\n"
                                   "Para testar sem chave, use \"Usar JSON pronto\".")
            return
        self.btn_gerar.config(state="disabled")
        threading.Thread(target=self._etapa1, daemon=True).start()

    def _etapa1(self):
        try:
            if self.modo_json:
                processo = json.loads(self.arquivo.read_text(encoding="utf-8"))
            else:
                self.log(f"Lendo {self.arquivo.name}...")
                texto = core.ler_documento(str(self.arquivo))
                self.log(f"Texto extraído: {len(texto)} caracteres.")
                chave = self.var_chave.get().strip()
                prov_id, var = core.PROVEDORES[self.var_prov.get()]
                os.environ[var] = chave
                if self.var_lembrar.get():
                    core.salvar_chave(chave, var)
                self.log(f"Interpretando com IA ({prov_id})... (pode levar alguns segundos)")
                processo = core.interpretar_com_ia(texto, prov_id)

            SAIDA.mkdir(exist_ok=True)
            nome = self.arquivo.stem
            (SAIDA / f"{nome}.json").write_text(json.dumps(processo, ensure_ascii=False, indent=2), encoding="utf-8")

            erros = core.validar_json(processo)
            if erros:
                self.log("\nProblemas encontrados:")
                for e in erros:
                    self.log(f" - {e}")
                self.log(f"O JSON ficou salvo em diagramas\\{nome}.json para você corrigir.")
                self.agendar(lambda: self.btn_gerar.config(state="normal"))
                return
            for a in core.avisos_json(processo):
                self.log(f"Aviso: {a}")
            self.agendar(lambda: self._revisar(processo, nome))
        except Exception as e:
            self.log("\nERRO: " + str(e))
            self.log(traceback.format_exc())
            self.agendar(lambda: self.btn_gerar.config(state="normal"))

    # ---------- revisão humana ----------
    def _revisar(self, processo, nome):
        usar_raias = self.var_raias.get()
        top = tk.Toplevel(self)
        top.title("Revise o processo")
        top.geometry("640x520")
        top.transient(self)
        top.grab_set()
        tk.Label(top, text="A IA entendeu o processo assim. Confere com o POP?",
                 font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=12, pady=8)
        caixa = scrolledtext.ScrolledText(top, font=("Consolas", 10))
        caixa.pack(fill="both", expand=True, padx=12)
        caixa.insert("end", core.resumo_texto(processo, usar_raias))
        caixa.configure(state="disabled")
        botoes = tk.Frame(top)
        botoes.pack(fill="x", padx=12, pady=10)

        def ok():
            top.destroy()
            self._etapa2(processo, nome, usar_raias)

        def cancelar():
            top.destroy()
            self.log(f"Cancelado. Corrija diagramas\\{nome}.json e use \"Usar JSON pronto\".")
            self.btn_gerar.config(state="normal")

        tk.Button(botoes, text="Gerar diagrama", command=ok, bg="#1f3864", fg="white",
                  padx=14, pady=4).pack(side="right")
        tk.Button(botoes, text="Cancelar", command=cancelar, padx=14, pady=4).pack(side="right", padx=8)
        top.protocol("WM_DELETE_WINDOW", cancelar)

    # ---------- etapa 2: gerar arquivos ----------
    def _etapa2(self, processo, nome, usar_raias):
        try:
            xml = core.gerar_bpmn(processo, usar_raias)
            core.ET.fromstring(xml.encode("utf-8"))  # confere se o XML é bem formado
            bpmn = SAIDA / f"{nome}.bpmn"
            pagina = SAIDA / f"{nome}.html"
            bpmn.write_text(xml, encoding="utf-8")
            pagina.write_text(core.gerar_html(xml, processo.get("processo", nome)), encoding="utf-8")
            self.log(f"\nPronto! Arquivos em: {SAIDA}")
            self.log(f"  {bpmn.name}  (abra no bpmn.io ou Camunda Modeler)")
            self.log(f"  {pagina.name}  (visualização no navegador)")
            webbrowser.open(pagina.as_uri())
        except Exception as e:
            self.log("\nERRO ao gerar: " + str(e))
            self.log(traceback.format_exc())
        finally:
            self.btn_gerar.config(state="normal")


def _console(visivel: bool):
    """No Windows, mostra/esconde a janela preta do abrir.bat (só quando aberto por ele)."""
    if not os.environ.get("LEITOR_ESCONDER"):
        return
    try:
        import ctypes
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 5 if visivel else 0)   # 5 = mostrar, 0 = esconder
    except Exception:
        pass


if __name__ == "__main__":
    try:
        app = App()
        _console(False)          # tudo certo: esconde a janela preta
        app.mainloop()
    except Exception:
        _console(True)           # deu erro: mostra a janela preta com a mensagem
        messagebox.showerror("Erro", traceback.format_exc())
        raise
