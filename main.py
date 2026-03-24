import tkinter as tk
from tkinter import filedialog, ttk
import threading
import os
import datetime
from collections import defaultdict, Counter

try:
    import openpyxl
except ImportError:
    import subprocess, sys
    subprocess.call([sys.executable, "-m", "pip", "install", "openpyxl", "--quiet"])
    import openpyxl

# ── Cores ─────────────────────────────────────────────────────────────────────
C = {
    "bg":      "#0D1F13",
    "card":    "#152A1A",
    "input":   "#1A3320",
    "green":   "#2D7A45",
    "glight":  "#3DAF60",
    "gaccent": "#56D97F",
    "white":   "#F0F4F1",
    "muted":   "#8BA898",
    "label":   "#C2D9C8",
    "red":     "#E05555",
    "yellow":  "#E0B455",
    "border":  "#253D2C",
}

FT = ("Helvetica", 22, "bold")
FB = ("Helvetica", 10, "bold")
FN = ("Helvetica", 10)
FS = ("Helvetica", 9)


# ── Helpers ───────────────────────────────────────────────────────────────────

def normalizar(v):
    return str(v).strip().upper() if v is not None else ""

def to_float(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None

def eh_mao_de_obra(el):
    el = normalizar(el)
    return any(p in el for p in [
        "MAO DE OBRA", "MÃO DE OBRA", "M.O.", "DIARISTA",
        "DIARIA", "DIÁRIA", "MEEIRO", "CONTRATADA",
        "FUNCIONARIO", "FUNCIONÁRIO", "TRABALHADOR",
    ])

ATIV_MO = {
    "CONDUÇÃO DE LAVOURA", "CONDUCAO DE LAVOURA",
    "ADUBAÇÃO VIA SOLO",   "ADUBACAO VIA SOLO",
    "ADUBAÇÃO VIA FOLHA",  "ADUBACAO VIA FOLHA",
    "CONTROLE DE PRAGAS E DOENÇAS", "CONTROLE DE PRAGAS E DOENCAS",
    "CONTROLE DE PLANTAS DANINHAS",
    "COLHEITA",
    "PÓS-COLHEITA", "POS-COLHEITA", "PÓS COLHEITA", "POS COLHEITA",
}


# ── Análise ───────────────────────────────────────────────────────────────────

def analisar(filepath):
    issues = []

    try:
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    except Exception as e:
        return [{"aba": "ERRO", "linha": "-", "tipo": "CRITICO",
                 "desc": f"Nao foi possivel abrir o arquivo: {e}"}]

    abas = wb.sheetnames

    def rows(aba):
        if aba not in abas:
            return [], []
        ws = wb[aba]
        data = list(ws.iter_rows(values_only=True))
        if not data:
            return [], []
        hdr = [normalizar(c) for c in data[0]]
        return hdr, [r for r in data[1:] if any(r)]

    def col(hdr, *nomes):
        for n in nomes:
            for i, h in enumerate(hdr):
                if n in h:
                    return i
        return None

    def get(row, i):
        try:
            return row[i] if i is not None and i < len(row) else None
        except IndexError:
            return None

    def add(aba, linha, tipo, desc):
        issues.append({"aba": aba, "linha": linha, "tipo": tipo, "desc": desc})

    # ── TALHAO ────────────────────────────────────────────────────────────────
    talhoes_producao = set()
    talhoes_todos    = set()

    hdr, data = rows("TALHAO")
    if not hdr:
        add("TALHAO", "-", "CRITICO", "Aba TALHAO nao encontrada.")
    else:
        c_nome = col(hdr, "TALHAO")
        c_est  = col(hdr, "ESTÁGIO", "ESTAGIO")
        for i, row in enumerate(data, 2):
            nome = normalizar(get(row, c_nome))
            est  = normalizar(get(row, c_est))
            if nome:
                talhoes_todos.add(nome)
                if est in ("PRODUÇÃO", "PRODUCAO"):
                    talhoes_producao.add(nome)

    # ── INVENTARIO ────────────────────────────────────────────────────────────
    hdr, data = rows("INVENTARIO")
    if not hdr:
        add("INVENTARIO", "-", "CRITICO", "Aba INVENTARIO nao encontrada.")
    else:
        c_desc  = col(hdr, "DESCRIÇÃO", "DESCRICAO")
        c_fab   = col(hdr, "DATA DE FABRICAÇÃO", "DATA DE FABRICACAO")
        c_acq   = col(hdr, "DATA DE AQUISIÇÃO",  "DATA DE AQUISICAO")
        c_vpago = col(hdr, "VALOR PAGO")
        c_vnovo = col(hdr, "VALOR DO ITEM NOVO")

        for i, row in enumerate(data, 2):
            desc  = get(row, c_desc)
            fab   = get(row, c_fab)
            acq   = get(row, c_acq)
            vpago = to_float(get(row, c_vpago))
            vnovo = to_float(get(row, c_vnovo))
            ds    = str(desc) if desc else f"Linha {i}"

            for lbl, val in [("Valor Pago", vpago), ("Valor Novo", vnovo)]:
                if val is None:
                    continue
                if val < 100:
                    add("INVENTARIO", i, "ALERTA",
                        f"{ds} | {lbl} R$ {val:,.2f} abaixo de R$ 100,00 — verificar.")
                elif val > 500000:
                    add("INVENTARIO", i, "ALERTA",
                        f"{ds} | {lbl} R$ {val:,.2f} acima de R$ 500.000,00 — confirmar.")

            if isinstance(fab, datetime.datetime) and isinstance(acq, datetime.datetime):
                if fab > acq:
                    add("INVENTARIO", i, "ERRO",
                        f"{ds} | Data fabricacao ({fab.strftime('%d/%m/%Y')}) "
                        f"posterior a aquisicao ({acq.strftime('%d/%m/%Y')}).")

    # ── PRODUCAO ──────────────────────────────────────────────────────────────
    hdr, data = rows("PRODUCAO")
    if not hdr:
        add("PRODUCAO", "-", "CRITICO", "Aba PRODUCAO nao encontrada.")
    else:
        c_tal  = col(hdr, "TALHÃO", "TALHAO")
        c_rat  = col(hdr, "RATEIO")
        c_ptot = col(hdr, "PRODUÇÃO TOTAL", "PRODUCAO TOTAL")
        c_ptal = col(hdr, "PRODUÇÃO TALHAO", "PRODUCAO TALHAO")
        c_mes  = col(hdr, "MÊS", "MES")

        grupos = defaultdict(list)
        for i, row in enumerate(data, 2):
            rat = normalizar(get(row, c_rat))
            mes = get(row, c_mes)
            if rat == "SIM" and mes:
                mk = str(mes)[:7]
                grupos[(mk, to_float(get(row, c_ptot)))].append({
                    "linha": i,
                    "talhao": normalizar(get(row, c_tal)),
                    "ptal": to_float(get(row, c_ptal)),
                })

        for (mk, _), grp in grupos.items():
            tg   = {g["talhao"] for g in grp}
            falt = talhoes_producao - tg
            if falt:
                add("PRODUCAO", grp[0]["linha"], "ERRO",
                    f"Rateio {mk} | Talhoes faltando: {', '.join(sorted(falt))}.")
            vals = [g["ptal"] for g in grp if g["ptal"] is not None]
            if vals and len(set(round(v, 2) for v in vals)) > 1:
                add("PRODUCAO", grp[0]["linha"], "ERRO",
                    f"Rateio {mk} | Valores divergentes: {[round(v,2) for v in vals]}.")

    # ── DESPESAS ──────────────────────────────────────────────────────────────
    hdr, data = rows("DESPESAS")
    if not hdr:
        add("DESPESAS", "-", "CRITICO", "Aba DESPESAS nao encontrada.")
    else:
        c_tal  = col(hdr, "TALHÃO", "TALHAO")
        c_mes  = col(hdr, "MÊS", "MES")
        c_atv  = col(hdr, "ATIVIDADE")
        c_elm  = col(hdr, "ELEMENTO")
        c_rat  = col(hdr, "RATEIO")
        c_uni  = col(hdr, "UNIDADE")
        c_vun  = col(hdr, "VALOR UNITÁRIO", "VALOR UNITARIO")
        c_vtot = col(hdr, "VALOR TOTAL (R$)")
        c_rsha = col(hdr, "R$/HA", "VALOR TOTAL (R$/HA)")

        desp_rat    = defaultdict(list)
        desp_iguais = []
        adm_mes_elm = defaultdict(list)

        for i, row in enumerate(data, 2):
            tal  = normalizar(get(row, c_tal))
            mes  = get(row, c_mes)
            atv  = normalizar(get(row, c_atv))
            elm  = normalizar(get(row, c_elm))
            rat  = normalizar(get(row, c_rat))
            uni  = normalizar(get(row, c_uni))
            vun  = to_float(get(row, c_vun))
            vtot = to_float(get(row, c_vtot))
            rsha = to_float(get(row, c_rsha))
            mk   = str(mes)[:7] if mes else "?"

            # 1. Atividades que exigem M.O.
            if any(a in atv for a in ATIV_MO) and not eh_mao_de_obra(elm):
                add("DESPESAS", i, "ALERTA",
                    f"[{atv}] '{elm}' nao e M.O. — verificar lancamento de "
                    f"mao de obra no mes {mk} para '{tal}'.")

            # 2. Manutencao de maquinas como Administracao
            if "MANUTENÇÃO" in elm and "MÁQUINAS" in elm and atv == "ADMINISTRAÇÃO":
                add("DESPESAS", i, "ERRO",
                    f"Manutencao de Maquinas lancada como ADMINISTRACAO "
                    f"(talhao: {tal}, mes: {mk}). Corrigir atividade.")

            # 3. R$/ha acima de 5.000
            if rsha is not None and rsha > 5000:
                add("DESPESAS", i, "ALERTA",
                    f"[{atv}] '{elm}' | R$/ha R$ {rsha:,.2f} acima de R$ 5.000 "
                    f"(talhao: {tal}, mes: {mk}).")

            # 4. Valor unitario acima de 5.000
            if vun is not None and vun > 5000:
                add("DESPESAS", i, "ALERTA",
                    f"[{atv}] '{elm}' | Valor unitario R$ {vun:,.2f} acima de "
                    f"R$ 5.000 (talhao: {tal}, mes: {mk}).")

            # 5. Valor unitario abaixo de R$1
            if vun is not None and 0 < vun < 1:
                add("DESPESAS", i, "ALERTA",
                    f"[{atv}] '{elm}' | Valor unitario R$ {vun:,.4f} abaixo de "
                    f"R$ 1,00 — possivel erro (talhao: {tal}, mes: {mk}).")

            # 6. Kg ou Litros com unitario acima de R$200
            if uni in ("KG", "LITROS", "L", "LITRO") and vun is not None and vun > 200:
                add("DESPESAS", i, "ALERTA",
                    f"[{atv}] '{elm}' | {uni} unitario R$ {vun:,.2f} acima de "
                    f"R$ 200 (talhao: {tal}, mes: {mk}).")

            # 7. Coletar rateio
            if rat == "SIM":
                desp_rat[(mk, atv, elm)].append(
                    {"linha": i, "talhao": tal, "vtot": vtot})

            # 8. Coletar duplicados
            desp_iguais.append((mk, atv, elm, round(vtot or 0, 4), tal, i))

            # 9. Coletar admin
            if atv == "ADMINISTRAÇÃO":
                adm_mes_elm[(mk, elm)].append(i)

        # Rateio: talhoes + valores
        for (mk, atv, elm), grp in desp_rat.items():
            tg   = {g["talhao"] for g in grp}
            falt = talhoes_todos - tg
            if falt:
                add("DESPESAS", grp[0]["linha"], "ERRO",
                    f"[{atv}] '{elm}' rateio {mk} | Talhoes faltando: "
                    f"{', '.join(sorted(falt))}.")
            vals = [g["vtot"] for g in grp if g["vtot"] is not None]
            if vals and abs(max(vals) - min(vals)) > 0.05:
                add("DESPESAS", grp[0]["linha"], "ERRO",
                    f"[{atv}] '{elm}' rateio {mk} | Valores divergentes "
                    f"(min R$ {min(vals):,.2f} / max R$ {max(vals):,.2f}).")

        # Lancamentos identicos
        cnt_ig  = Counter((mk, atv, elm, vt, tal)
                          for mk, atv, elm, vt, tal, _ in desp_iguais)
        lns_ig  = defaultdict(list)
        for mk, atv, elm, vt, tal, ln in desp_iguais:
            lns_ig[(mk, atv, elm, vt, tal)].append(ln)
        for k, cnt in cnt_ig.items():
            if cnt > 1:
                mk, atv, elm, vt, tal = k
                add("DESPESAS", lns_ig[k][0], "ALERTA",
                    f"[{atv}] '{elm}' | R$ {vt:,.2f} identico {cnt}x no mes "
                    f"{mk} para '{tal}' (linhas: {', '.join(map(str, lns_ig[k]))}) "
                    f"— possivel duplicidade.")

        # Recorrencia admin
        meses_adm = set(k[0] for k in adm_mes_elm)
        for elm in set(k[1] for k in adm_mes_elm):
            com = {k[0] for k in adm_mes_elm if k[1] == elm}
            sem = meses_adm - com
            if len(com) >= 2 and sem:
                add("DESPESAS", "-", "INFO",
                    f"Administracao | '{elm}' lancado em {sorted(com)} "
                    f"mas falta em {sorted(sem)} — verificar recorrencia.")

    # ── VENDAS ────────────────────────────────────────────────────────────────
    hdr, data = rows("VENDAS")
    if not hdr:
        add("VENDAS", "-", "CRITICO", "Aba VENDAS nao encontrada.")
    else:
        c_tal   = col(hdr, "TALHÃO", "TALHAO")
        c_mes   = col(hdr, "MÊS", "MES")
        c_preco = col(hdr, "PREÇO", "PRECO")

        for i, row in enumerate(data, 2):
            tal   = normalizar(get(row, c_tal))
            mes   = get(row, c_mes)
            preco = to_float(get(row, c_preco))
            mk    = str(mes)[:7] if mes else "?"
            if preco is not None and preco > 100:
                add("VENDAS", i, "ALERTA",
                    f"Talhao '{tal}' | Preco venda R$ {preco:,.2f}/sc acima de "
                    f"R$ 100,00 (mes: {mk}).")

    return issues


# ── Interface ─────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Analise de Consistencia - Labor Rural")
        self.geometry("1100x720")
        self.minsize(900, 580)
        self.configure(bg=C["bg"])
        self._path   = None
        self._issues = []
        self._ui()
        self.update_idletasks()
        x = (self.winfo_screenwidth()  - 1100) // 2
        y = (self.winfo_screenheight() -  720) // 2
        self.geometry(f"1100x720+{x}+{y}")

    def _ui(self):
        # Cabecalho
        hdr = tk.Frame(self, bg=C["card"])
        hdr.pack(fill="x")
        tk.Frame(hdr, bg=C["gaccent"], height=3).pack(fill="x")
        hi = tk.Frame(hdr, bg=C["card"], padx=28, pady=16)
        hi.pack(fill="x")
        lf = tk.Frame(hi, bg=C["card"])
        lf.pack(side="left")
        tk.Label(lf, text="⬡", font=("Helvetica", 26), fg=C["gaccent"],
                 bg=C["card"]).pack(side="left", padx=(0, 10))
        tf = tk.Frame(lf, bg=C["card"])
        tf.pack(side="left")
        tk.Label(tf, text="Analise de Consistencia",
                 font=FT, fg=C["white"], bg=C["card"]).pack(anchor="w")
        tk.Label(tf, text="Labor Rural  |  Plataforma MIMC",
                 font=FS, fg=C["muted"], bg=C["card"]).pack(anchor="w")
        self.bdg = tk.Frame(hi, bg=C["card"])
        self.bdg.pack(side="right")

        # Importacao
        imp = tk.Frame(self, bg=C["bg"], padx=28, pady=14)
        imp.pack(fill="x")
        fc = tk.Frame(imp, bg=C["input"])
        fc.pack(fill="x")
        tk.Frame(fc, bg=C["green"], width=4).pack(side="left", fill="y")
        fi = tk.Frame(fc, bg=C["input"], padx=14, pady=12)
        fi.pack(side="left", fill="x", expand=True)
        tk.Label(fi, text="PLANILHA SELECIONADA", font=FS,
                 fg=C["muted"], bg=C["input"]).pack(anchor="w")
        self.lbl_f = tk.Label(fi, text="Nenhum arquivo selecionado",
                               font=FN, fg=C["label"], bg=C["input"])
        self.lbl_f.pack(anchor="w")
        ba = tk.Frame(fc, bg=C["input"], padx=14, pady=12)
        ba.pack(side="right")
        self._btn(ba, "Importar Planilha", self._imp,
                  C["green"], C["white"]).pack(side="left", padx=6)
        self.b_run = self._btn(ba, "  Analisar  ", self._run,
                               C["gaccent"], C["bg"], state="disabled")
        self.b_run.pack(side="left", padx=6)

        # Filtros
        fb = tk.Frame(self, bg=C["bg"], padx=28, pady=4)
        fb.pack(fill="x")
        tk.Label(fb, text="FILTRAR:", font=FS,
                 fg=C["muted"], bg=C["bg"]).pack(side="left", padx=(0, 6))
        self.fv = tk.StringVar(value="TODOS")
        for f in ["TODOS", "ERRO", "ALERTA", "INFO", "CRITICO"]:
            tk.Radiobutton(fb, text=f, variable=self.fv, value=f,
                           command=self._filtrar, font=FS,
                           fg=self._cor(f), selectcolor=C["bg"],
                           bg=C["bg"], activebackground=C["bg"],
                           bd=0, highlightthickness=0).pack(side="left", padx=5)
        self.fa = tk.StringVar(value="TODAS")
        self.cmb = ttk.Combobox(fb, textvariable=self.fa, state="readonly",
                                 width=20, font=FS)
        self.cmb["values"] = ["TODAS"]
        self.cmb.pack(side="left", padx=8)
        self.cmb.bind("<<ComboboxSelected>>", lambda _: self._filtrar())
        s = ttk.Style()
        s.theme_use("default")
        s.configure("TCombobox", fieldbackground=C["input"],
                    background=C["input"], foreground=C["white"],
                    arrowcolor=C["gaccent"])

        # Tabela
        tbl = tk.Frame(self, bg=C["bg"], padx=28, pady=6)
        tbl.pack(fill="both", expand=True)
        th = tk.Frame(tbl, bg=C["card"], pady=7)
        th.pack(fill="x")
        for lbl, w, anc in [("ABA", 10, "w"), ("LINHA", 6, "center"),
                              ("TIPO", 8, "center"), ("DESCRICAO", 60, "w")]:
            tk.Label(th, text=lbl, font=FB, fg=C["muted"],
                     bg=C["card"], width=w, anchor=anc, padx=8).pack(side="left")

        sc = tk.Frame(tbl, bg=C["bg"])
        sc.pack(fill="both", expand=True, pady=(1, 0))
        sb = tk.Scrollbar(sc, orient="vertical", width=7,
                           bg=C["border"], troughcolor=C["bg"])
        sb.pack(side="right", fill="y")
        self.cv = tk.Canvas(sc, bg=C["bg"],
                             yscrollcommand=sb.set, highlightthickness=0)
        self.cv.pack(side="left", fill="both", expand=True)
        sb.config(command=self.cv.yview)
        self.inn = tk.Frame(self.cv, bg=C["bg"])
        self.cw  = self.cv.create_window((0, 0), window=self.inn, anchor="nw")
        self.inn.bind("<Configure>", lambda _: self.cv.configure(
            scrollregion=self.cv.bbox("all")))
        self.cv.bind("<Configure>", lambda e: self.cv.itemconfig(
            self.cw, width=e.width))
        self.cv.bind_all("<MouseWheel>", lambda e: self.cv.yview_scroll(
            int(-1 * (e.delta / 120)), "units"))
        tk.Label(self.inn,
                 text="Importe uma planilha e clique em Analisar.",
                 font=FN, fg=C["muted"], bg=C["bg"], pady=40).pack()

        # Status bar
        sb2 = tk.Frame(self, bg=C["card"])
        sb2.pack(fill="x", side="bottom")
        tk.Frame(sb2, bg=C["border"], height=1).pack(fill="x")
        self.lbl_st = tk.Label(sb2, text="Pronto.", font=FS,
                                fg=C["muted"], bg=C["card"], padx=14, pady=4)
        self.lbl_st.pack(side="left")
        self.lbl_tot = tk.Label(sb2, text="", font=FB,
                                 fg=C["gaccent"], bg=C["card"], padx=14)
        self.lbl_tot.pack(side="right")

    def _btn(self, p, txt, cmd, bg, fg, state="normal"):
        return tk.Button(p, text=txt, command=cmd, font=FB,
                          bg=bg, fg=fg, relief="flat", bd=0,
                          padx=12, pady=7, cursor="hand2", state=state,
                          activebackground=C["glight"],
                          activeforeground=C["bg"])

    def _cor(self, t):
        return {"ERRO": C["red"], "CRITICO": C["red"],
                "ALERTA": C["yellow"], "INFO": C["muted"],
                "TODOS": C["label"]}.get(t, C["label"])

    def _bbg(self, t):
        return {"ERRO": "#3A1A1A", "CRITICO": "#3A1A1A",
                "ALERTA": "#3A2A10", "INFO": "#1A2530"}.get(t, C["input"])

    def _badges(self):
        for w in self.bdg.winfo_children():
            w.destroy()
        cnt = Counter(i["tipo"] for i in self._issues)
        for tipo, n in cnt.items():
            f = tk.Frame(self.bdg, bg=C["bg"], padx=8, pady=3)
            f.pack(side="left", padx=3)
            tk.Label(f, text=f"● {tipo}", font=FS,
                     fg=self._cor(tipo), bg=C["bg"]).pack(side="left")
            tk.Label(f, text=f"  {n}", font=FB,
                     fg=C["white"], bg=C["bg"]).pack(side="left")

    def _imp(self):
        fp = filedialog.askopenfilename(
            title="Selecionar planilha MIMC",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")])
        if fp:
            self._path = fp
            self.lbl_f.config(text=os.path.basename(fp), fg=C["gaccent"])
            self.b_run.config(state="normal")
            self.lbl_st.config(text=f"Arquivo: {os.path.basename(fp)}")

    def _run(self):
        if not self._path:
            return
        self.b_run.config(state="disabled", text=" Analisando... ")
        self.lbl_st.config(text="Analisando planilha...")
        for w in self.inn.winfo_children():
            w.destroy()

        def task():
            res = analisar(self._path)
            self.after(0, lambda: self._exibir(res))

        threading.Thread(target=task, daemon=True).start()

    def _exibir(self, issues):
        self._issues = issues
        self._badges()
        abas = ["TODAS"] + sorted(set(i["aba"] for i in issues))
        self.cmb["values"] = abas
        self.fa.set("TODAS")
        self._filtrar()
        self.b_run.config(state="normal", text="  Analisar  ")
        self.lbl_tot.config(text=f"Total: {len(issues)} ocorrencias")
        self.lbl_st.config(
            text=f"Concluido. {len(issues)} inconsistencias encontradas.")

    def _filtrar(self):
        tf  = self.fv.get()
        af  = self.fa.get()
        lst = [i for i in self._issues
               if (tf == "TODOS" or i["tipo"] == tf)
               and (af == "TODAS" or i["aba"] == af)]
        for w in self.inn.winfo_children():
            w.destroy()
        if not lst:
            tk.Label(self.inn,
                     text="Nenhuma ocorrencia com os filtros selecionados.",
                     font=FN, fg=C["muted"], bg=C["bg"], pady=40).pack()
            return
        for idx, iss in enumerate(lst):
            bg  = C["card"] if idx % 2 == 0 else C["bg"]
            cor = self._cor(iss["tipo"])
            rf  = tk.Frame(self.inn, bg=bg)
            rf.pack(fill="x")
            tk.Frame(rf, bg=cor, width=3).pack(side="left", fill="y")
            tk.Label(rf, text=iss["aba"], font=FS, fg=C["muted"],
                     bg=bg, width=12, anchor="w",
                     padx=7, pady=7).pack(side="left")
            tk.Label(rf, text=str(iss["linha"]), font=FS, fg=C["muted"],
                     bg=bg, width=6, anchor="center").pack(side="left")
            bf = tk.Frame(rf, bg=self._bbg(iss["tipo"]), padx=5, pady=2)
            bf.pack(side="left", padx=7)
            tk.Label(bf, text=iss["tipo"], font=FS,
                     fg=cor, bg=self._bbg(iss["tipo"])).pack()
            tk.Label(rf, text=iss["desc"], font=FS, fg=C["label"],
                     bg=bg, wraplength=680, justify="left",
                     anchor="w", padx=7).pack(side="left", fill="x", expand=True)


if __name__ == "__main__":
    App().mainloop()
