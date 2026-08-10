#!/usr/bin/env python3
"""
Gera análise de frequência por grupo de NÍVEL INFERIDO pelas respostas
(mesma estrutura do analise_frequencia_por_grupo.xlsx, mas usando a
classificação inferida em vez do autoposicionamento Q10).
"""

import re, unicodedata
from collections import defaultdict, Counter
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

CAMINHO_BASE  = "/root/.claude/uploads/9923a1d6-2dc6-5622-bc3f-41df5cfa84af/f9c4405b-Base_da_pesquisa_e_l_gica_revisado15.07_1.xlsx"
CAMINHO_CLASS = "/home/user/arandu/classificacao_individual.xlsx"
SAIDA         = "/home/user/arandu/frequencia_por_nivel_inferido.xlsx"

ESTAGIOS = ["Novo", "Adaptando", "Inserindo", "Comprometido",
            "Ativo no Motivo", "Embaixador", "Cocriador"]

# Cores por estágio (mesmo padrão da classificação individual)
COR_ESTAGIO = {
    "Novo":           "F4CCCC",
    "Adaptando":      "FCE5CD",
    "Inserindo":      "FFF2CC",
    "Comprometido":   "D9EAD3",
    "Ativo no Motivo":"C9DAF8",
    "Embaixador":     "D9D2E9",
    "Cocriador":      "E6B8A2",
}

# ─── Normalização ─────────────────────────────────────────────────────────────

def norm(s):
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s.lower().strip())

def extrai_colchetes(s):
    m = re.search(r"\[([^\[\]]+)\]\s*$", str(s))
    return m.group(1).strip() if m else None

def casa_textos(a, b):
    na, nb = norm(a), norm(b)
    if na == nb: return True
    ca, cb = extrai_colchetes(a), extrai_colchetes(b)
    if ca and cb and norm(ca) == norm(cb): return True
    return False

# ─── Carregamento ─────────────────────────────────────────────────────────────

def carrega_base():
    wb = openpyxl.load_workbook(CAMINHO_BASE, read_only=True, data_only=True)

    def aba_df(nome):
        ws = wb[nome]
        rows = list(ws.iter_rows(values_only=True))
        if not rows: return pd.DataFrame()
        header = [str(c) if c is not None else f"_col{i}" for i, c in enumerate(rows[0])]
        data = [list(r) for r in rows[1:] if any(v is not None for v in r)]
        return pd.DataFrame(data, columns=header)

    return aba_df("Respostas ao formulário 1"), aba_df("Perguntas_Dimensoes")

def carrega_inferidos():
    """Retorna dict {nome_norm → nível_inferido}."""
    wb = openpyxl.load_workbook(CAMINHO_CLASS, read_only=True, data_only=True)
    ws = wb["Classificação Individual"]
    rows = list(ws.iter_rows(values_only=True))
    header = list(rows[0])
    col_nome = header.index("Nome")
    col_inf  = header.index("Nível inferido pelas respostas")
    resultado = {}
    for r in rows[1:]:
        if not any(r): continue
        nome = norm(str(r[col_nome] or ""))
        nivel = r[col_inf]
        if nome and nivel:
            resultado[nome] = str(nivel).strip()
    return resultado

# ─── Mapeamento esfera → colunas ──────────────────────────────────────────────

def mapeia_esferas(df_pd, resp_cols):
    esfera_cols = {}
    for _, row in df_pd.iterrows():
        perg  = row.iloc[0]
        esfera = row.get("Esfera")
        if pd.isna(perg) or (esfera is None) or (isinstance(esfera, float) and pd.isna(esfera)):
            continue
        for rc in resp_cols:
            if casa_textos(str(perg), rc):
                esfera_cols.setdefault(str(esfera).strip(), []).append(rc)
                break
    return esfera_cols

# ─── Frequências ──────────────────────────────────────────────────────────────

def freq_col(df_grupo, col):
    """Retorna Counter de respostas normalizadas para uma coluna."""
    vals = df_grupo[col].dropna().astype(str).map(str.strip).map(norm)
    vals = vals[vals != ""]
    return Counter(vals), len(vals)

def label_curto(col_name):
    """Extrai o conteúdo entre colchetes ou primeiros 80 chars."""
    c = extrai_colchetes(col_name)
    if c: return c.strip()
    return col_name[:80].strip()

# ─── Estilos ──────────────────────────────────────────────────────────────────

COR_HEADER   = "1F3864"
COR_SUBHEADER= "2E5090"
COR_ALT      = "EBF3FB"
COR_TOTAL_ROW= "D6E4F0"

def hdr(ws, row_idx):
    for cell in ws[row_idx]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=COR_HEADER)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

def subhdr(ws, row_idx):
    for cell in ws[row_idx]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=COR_SUBHEADER)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

def auto_w(ws, max_w=40):
    for col in ws.iter_cols():
        mx = max((len(str(c.value or "")) for c in col), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(mx + 3, max_w)

def pct(v):
    if v is None: return ""
    return round(float(v), 4)

def titulo_aba(ws, texto):
    ws.cell(1, 1, texto).font = Font(bold=True, size=12)
    ws.cell(1, 1).alignment = Alignment(horizontal="left")

# ─── Aba Resumo ───────────────────────────────────────────────────────────────

def aba_resumo(wb_out, grupos, esfera_cols):
    ws = wb_out.create_sheet("Resumo")
    titulo_aba(ws, "Resumo — Nível Inferido pelas Respostas (Classificação Individual)")
    ws.append([])

    # Distribuição dos grupos
    ws.append(["Nível inferido pelas respostas", "N", "%"])
    hdr(ws, ws.max_row)
    total = sum(len(v) for v in grupos.values())
    for est in ESTAGIOS:
        if est not in grupos: continue
        n = len(grupos[est])
        row = [est, n, round(n / total, 4)]
        ws.append(row)
        cor = COR_ESTAGIO.get(est, "FFFFFF")
        for ci in range(1, 4):
            ws.cell(ws.max_row, ci).fill = PatternFill("solid", fgColor=cor)
    ws.append([])

    # Tabela resumida de concordância por esfera × grupo
    grupos_presentes = [e for e in ESTAGIOS if e in grupos]
    n_row = [""] + [f"N = {len(grupos[g])}" for g in grupos_presentes]
    header_row = ["Esfera / Indicador"] + grupos_presentes

    ws.append(["Resumo — % Concordância e Participação por Nível Inferido"])
    ws.cell(ws.max_row, 1).font = Font(bold=True)
    ws.append(n_row)
    ws.append(header_row)
    hdr(ws, ws.max_row)
    ws.row_dimensions[ws.max_row].height = 30

    def pooled_pct_pos(esfera, grupo_df, resps_pos):
        cols = esfera_cols.get(esfera, [])
        pos = tot = 0
        for col in cols:
            for v in grupo_df[col].dropna():
                sv = norm(str(v))
                if not sv or sv == "nan": continue
                tot += 1
                if sv in resps_pos: pos += 1
        return pos / tot if tot else None

    def pooled_pct_resp(esfera, grupo_df, resp_norm):
        cols = esfera_cols.get(esfera, [])
        match = tot = 0
        for col in cols:
            for v in grupo_df[col].dropna():
                sv = norm(str(v))
                if not sv or sv == "nan": continue
                tot += 1
                if sv == resp_norm: match += 1
        return match / tot if tot else None

    esfera_rows = [
        ("Conhecimento Institucional — % Sim",
         "Conhecimento institucional",
         {"sim"}, None),
        ("Pertencimento Cultural — % Concorda (total)",
         "Pertencimento cultural",
         {"concordo", "concordo parcialmente"}, None),
        ("   Concordo (pleno)",
         "Pertencimento cultural",
         {"concordo"}, None),
        ("   Não sei opinar",
         "Pertencimento cultural",
         {"nao sei opinar"}, None),
        ("   Discordo",
         "Pertencimento cultural",
         {"discordo", "discordo parcialmente"}, None),
        ("Acordos Sociais — % Concorda (total)",
         "Acordos sociais",
         {"concordo", "concordo parcialmente"}, None),
        ("   Concordo (pleno)",
         "Acordos sociais",
         {"concordo"}, None),
        ("   Não sei opinar",
         "Acordos sociais",
         {"nao sei opinar"}, None),
        ("   Discordo",
         "Acordos sociais",
         {"discordo", "discordo parcialmente"}, None),
        ("Prática da Ec. Fraterna — % Concorda (total)",
         "Prática da Economia Fraterna",
         {"concordo", "concordo parcialmente"}, None),
        ("   Concordo (pleno)",
         "Prática da Economia Fraterna",
         {"concordo"}, None),
        ("   Não sei opinar",
         "Prática da Economia Fraterna",
         {"nao sei opinar"}, None),
        ("   Discordo",
         "Prática da Economia Fraterna",
         {"discordo", "discordo parcialmente"}, None),
        ("Engajamento — % Sou atualmente (algum papel)",
         "Engajamento",
         {"sou atualmente"}, None),
        ("   Já fui (algum papel)",
         "Engajamento",
         {"ja fui"}, None),
        ("   Nunca fui (todos)",
         "Engajamento",
         {"nunca fui"}, None),
    ]

    for rotulo, esfera, resps_pos, _ in esfera_rows:
        vals = [rotulo]
        for g in grupos_presentes:
            df_g = grupos[g]
            v = pooled_pct_pos(esfera, df_g, resps_pos)
            vals.append(pct(v) if v is not None else "—")
        ws.append(vals)
        ri = ws.max_row
        if not rotulo.startswith("   "):
            for ci in range(1, len(vals) + 1):
                ws.cell(ri, ci).fill = PatternFill("solid", fgColor=COR_TOTAL_ROW)
                ws.cell(ri, ci).font = Font(bold=True)
        else:
            ws.cell(ri, 1).alignment = Alignment(indent=2)

    ws.append([])
    nota = (f"Grupos definidos pela inferência de nível a partir das respostas "
            f"(metodologia da classificação individual). N total = {total} respondentes.")
    ws.append([nota])
    ws.cell(ws.max_row, 1).font = Font(italic=True, size=9, color="606060")

    ws.freeze_panes = "B4"
    ws.column_dimensions["A"].width = 52
    for col_idx in range(2, len(grupos_presentes) + 2):
        ws.column_dimensions[get_column_letter(col_idx)].width = 16
    return ws

# ─── Aba por esfera de concordância ───────────────────────────────────────────

OPCOES_CONCORDANCIA = ["Concordo", "Concordo parcialmente", "Não sei opinar",
                       "Discordo", "Discordo totalmente"]
OPCOES_SIM_NAO      = ["Sim", "Não"]
OPCOES_ENGAJAMENTO  = ["Sou atualmente", "Já fui", "Nunca fui"]

def short_q(col_name):
    """Retorna label curto da pergunta (conteúdo entre colchetes ou início)."""
    c = extrai_colchetes(col_name)
    if c: return c[:100].strip()
    # Tira prefixo de bloco e retorna o que sobra
    s = re.sub(r"^[^.]+\.\s*", "", col_name).strip()
    return s[:100] if s else col_name[:100]

def aba_esfera(wb_out, nome_aba, titulo_completo, esfera, cols_esfera,
               grupos, grupos_presentes, opcoes, tem_total_col=True):
    ws = wb_out.create_sheet(nome_aba)
    titulo_aba(ws, f"Distribuição de Frequência por Nível Inferido — {titulo_completo}")
    ws.append([])

    # Cabeçalho
    n_row = [None, None] + [f"N = {len(grupos[g])}" for g in grupos_presentes]
    header = ["Pergunta / Afirmação", "Resposta"] + grupos_presentes
    if tem_total_col:
        header += ["% Concorda\n(total)"]
        n_row += [None]

    ws.append(n_row)
    ws.append(header)
    hdr(ws, ws.max_row)
    ws.row_dimensions[ws.max_row].height = 32

    total = sum(len(grupos[g]) for g in grupos_presentes)
    df_all = pd.concat([grupos[g] for g in grupos_presentes], ignore_index=True)

    alt = 0
    for col in cols_esfera:
        label = short_q(col)
        first_row = True
        for opt in opcoes:
            opt_n = norm(opt)
            row_vals = [label if first_row else None, opt]
            first_row = False

            # % por grupo
            for g in grupos_presentes:
                df_g = grupos[g]
                vals = df_g[col].dropna().astype(str).map(str.strip).map(norm)
                vals = vals[vals.map(bool)]
                n_tot = len(vals)
                n_opt = (vals == opt_n).sum()
                row_vals.append(pct(n_opt / n_tot) if n_tot else "—")

            # % total (pooled para o arquivo todo)
            if tem_total_col:
                all_vals = df_all[col].dropna().astype(str).map(str.strip).map(norm)
                all_vals = all_vals[all_vals.map(bool)]
                n_tot_all = len(all_vals)
                n_opt_all = (all_vals == opt_n).sum()
                if opt_n in {"concordo", "concordo parcialmente", "sim", "sou atualmente", "ja fui"}:
                    row_vals.append(pct(n_opt_all / n_tot_all) if n_tot_all else "—")
                else:
                    row_vals.append(None)

            ws.append(row_vals)
            ri = ws.max_row
            cor_linha = COR_ALT if alt % 2 == 1 else "FFFFFF"
            for ci in range(1, len(row_vals) + 1):
                cell = ws.cell(ri, ci)
                if cell.fill.fgColor.rgb in ("00000000", "FFFFFFFF", "00FFFFFF"):
                    cell.fill = PatternFill("solid", fgColor=cor_linha)
                # Color dos grupos no header
                if ri > 3 and ci > 2 and ci <= 2 + len(grupos_presentes):
                    g_idx = ci - 3
                    g_nome = grupos_presentes[g_idx] if g_idx < len(grupos_presentes) else None
                    if g_nome:
                        v = ws.cell(ri, ci).value
                        if isinstance(v, (int, float)) and not isinstance(v, bool):
                            ws.cell(ri, ci).number_format = "0%"
        alt += 1
        ws.append([None])  # linha separadora entre perguntas

    ws.freeze_panes = "C4"
    ws.column_dimensions["A"].width = 50
    ws.column_dimensions["B"].width = 22
    for col_idx in range(3, 3 + len(grupos_presentes) + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 15

    return ws

# ─── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("Carregando dados...")
    df_resp, df_pd = carrega_base()
    inferidos = carrega_inferidos()

    resp_cols = list(df_resp.columns)
    esfera_cols = mapeia_esferas(df_pd, resp_cols)

    print("Esferas mapeadas:")
    for k, v in esfera_cols.items():
        print(f"  {k}: {len(v)} colunas")

    # Coluna nome no df_resp
    col_nome = next((c for c in resp_cols if "nome completo" in norm(c)), None)

    # Adicionar coluna de nível inferido ao df_resp
    def get_nivel(row):
        nome = norm(str(row[col_nome] or "")) if col_nome else ""
        return inferidos.get(nome, None)

    df_resp["__nivel_inferido__"] = df_resp.apply(get_nivel, axis=1)

    grupos_presentes = [e for e in ESTAGIOS if e in df_resp["__nivel_inferido__"].values]
    grupos = {g: df_resp[df_resp["__nivel_inferido__"] == g].copy()
              for g in grupos_presentes}

    print("\nDistribuição por nível inferido:")
    for g in grupos_presentes:
        print(f"  {g}: {len(grupos[g])}")

    cols_conhec = esfera_cols.get("Conhecimento institucional", [])
    cols_pertenc = esfera_cols.get("Pertencimento cultural", [])
    cols_acordos = esfera_cols.get("Acordos sociais", [])
    cols_eco     = esfera_cols.get("Prática da Economia Fraterna", [])
    cols_engaj   = esfera_cols.get("Engajamento", [])

    # ── Export ────────────────────────────────────────────────────────────────
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    print("\nGerando abas...")
    aba_resumo(wb_out, grupos, esfera_cols)

    aba_esfera(wb_out, "Conhecimento", "Conhecimento institucional",
               "Conhecimento institucional", cols_conhec,
               grupos, grupos_presentes, OPCOES_SIM_NAO, tem_total_col=False)

    aba_esfera(wb_out, "Pertencimento", "Pertencimento cultural",
               "Pertencimento cultural", cols_pertenc,
               grupos, grupos_presentes, OPCOES_CONCORDANCIA, tem_total_col=True)

    aba_esfera(wb_out, "Acordos", "Acordos sociais",
               "Acordos sociais", cols_acordos,
               grupos, grupos_presentes, OPCOES_CONCORDANCIA, tem_total_col=True)

    aba_esfera(wb_out, "Ec. Fraterna", "Prática da Economia Fraterna",
               "Prática da Economia Fraterna", cols_eco,
               grupos, grupos_presentes, OPCOES_CONCORDANCIA, tem_total_col=True)

    aba_esfera(wb_out, "Engajamento", "Engajamento",
               "Engajamento", cols_engaj,
               grupos, grupos_presentes, OPCOES_ENGAJAMENTO, tem_total_col=False)

    wb_out.save(SAIDA)
    print(f"\n✓ Salvo: {SAIDA}")


if __name__ == "__main__":
    main()
