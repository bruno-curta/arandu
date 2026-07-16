#!/usr/bin/env python3
"""
Classificação individual v2 — nível atual inferido considerando:
  - % concordo / concordo parc / não sei opinar / discordo por esfera
  - engajamento ativo / já foi / nunca
  - tempo de escola (veteranos sem comprometimento = anomalia)
"""

import re, unicodedata
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

CAMINHO = "/root/.claude/uploads/9923a1d6-2dc6-5622-bc3f-41df5cfa84af/f9c4405b-Base_da_pesquisa_e_l_gica_revisado15.07_1.xlsx"
SAIDA = "/home/user/arandu/classificacao_individual.xlsx"
ANO_REF = 2025

ESTAGIOS = ["Novo", "Adaptando", "Inserindo", "Comprometido",
            "Ativo no Motivo", "Embaixador", "Cocriador"]
ESTAGIO_NOME = {i+1: e for i, e in enumerate(ESTAGIOS)}

PAPEIS = {
    1: "Recém-chegado — em reconhecimento",
    2: "Observador — assimilando como a escola funciona",
    3: "Explorador — processando dúvidas, próximo à decisão",
    4: "Membro comprometido — escolheu a Arandu",
    5: "Colaborador ativo — tem papel definido na escola",
    6: "Embaixador — referência e representante da escola",
    7: "Cocriador — co-constrói o futuro da escola",
}

PROXIMA_ACAO = {
    1: "Participar de apresentações sobre propósito, missão e valores",
    2: "Aprofundar conhecimento sobre pilares e pedagogia; tirar dúvidas",
    3: "Decidir: dizer SIM para a Arandu e engajar como Comprometido",
    4: "Assumir um papel ativo (representante, voluntário, APG ou GT)",
    5: "Ampliar atuação para 2+ papéis e fortalecer prática da Ec. Fraterna",
    6: "Engajar na co-criação: grupos estratégicos, mentoria, novos projetos",
    7: "Manter e ampliar a co-criação — já no topo da curva",
}

# ─── Texto ────────────────────────────────────────────────────────────────────

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

def carrega():
    wb = openpyxl.load_workbook(CAMINHO, read_only=True, data_only=True)
    def aba_df(nome):
        ws = wb[nome]
        rows = list(ws.iter_rows(values_only=True))
        if not rows: return pd.DataFrame()
        header = [str(c) if c is not None else f"_col{i}" for i, c in enumerate(rows[0])]
        data = [list(r) for r in rows[1:] if any(v is not None for v in r)]
        return pd.DataFrame(data, columns=header)
    return aba_df("Respostas ao formulário 1"), aba_df("Perguntas_Dimensoes")

# ─── Mapeamento esfera → colunas ──────────────────────────────────────────────

def mapeia_esferas(df_pd, resp_cols):
    esfera_cols = {}
    for _, row in df_pd.iterrows():
        perg = row.iloc[0]
        esfera = row.get("Esfera")
        if pd.isna(perg) or (esfera is None) or (isinstance(esfera, float) and pd.isna(esfera)):
            continue
        for rc in resp_cols:
            if casa_textos(str(perg), rc):
                esfera_cols.setdefault(str(esfera).strip(), []).append(rc)
                break
    return esfera_cols

# ─── Stats por tipo de resposta ───────────────────────────────────────────────

def stats_sim_nao(row, cols):
    n_sim = n = 0
    for c in cols:
        v = row.get(c)
        if v is None or (isinstance(v, float) and pd.isna(v)): continue
        sv = norm(str(v))
        if not sv: continue
        n += 1
        if sv == "sim": n_sim += 1
    return {"sim": n_sim/n, "n": n} if n else None

def stats_concordancia(row, cols):
    n_pleno = n_parc = n_nao = n_disc = n = 0
    for c in cols:
        v = row.get(c)
        if v is None or (isinstance(v, float) and pd.isna(v)): continue
        sv = norm(str(v))
        if not sv: continue
        n += 1
        if sv == "concordo": n_pleno += 1
        elif sv == "concordo parcialmente": n_parc += 1
        elif sv == "nao sei opinar": n_nao += 1
        elif "discordo" in sv: n_disc += 1
    if n == 0: return None
    ppl = n_pleno/n; ppa = n_parc/n; pno = n_nao/n; pdi = n_disc/n
    # Net signal: pleno=+1.0, parc=+0.4, nao=-0.2, disc=-0.8
    # Normalizado de [-0.8,1.0] → [0,1] via (raw+0.8)/1.8
    raw = ppl*1.0 + ppa*0.4 + pno*(-0.2) + pdi*(-0.8)
    net = max(0.0, min(1.0, (raw + 0.8) / 1.8))
    return {"pleno": ppl, "parc": ppa, "nao": pno, "disc": pdi,
            "pos": ppl+ppa, "net": net, "n": n}

def stats_engajamento(row, cols):
    n_ativo = n_foi = n = 0
    for c in cols:
        v = row.get(c)
        if v is None or (isinstance(v, float) and pd.isna(v)): continue
        sv = norm(str(v))
        if not sv: continue
        n += 1
        if "atualmente" in sv: n_ativo += 1
        elif "ja fui" in sv or sv == "ja fui": n_foi += 1
    return {"ativo": n_ativo, "foi": n_foi, "n": n} if n else None

# ─── Inferência de estágio ────────────────────────────────────────────────────

THRESH_YES = 0.60     # pertencimento net ≥ threshold → "disse SIM"
THRESH_DISC = 0.25    # discordância acima disso é sinal relevante

def inferred_stage(c_sim, p, a, ef, eng, entry_year):
    years = (ANO_REF - entry_year) if entry_year else 0

    p_net  = p["net"]  if p  else 0.5;  p_pos  = p["pos"]  if p  else 0.5
    p_plno = p["pleno"] if p else 0.5;  p_disc = p["disc"] if p  else 0.0
    p_nao  = p["nao"]  if p  else 0.0
    a_net  = a["net"]  if a  else 0.5;  a_disc = a["disc"] if a  else 0.0
    a_nao  = a["nao"]  if a  else 0.0
    ef_net = ef["net"] if ef else 0.5;  ef_disc= ef["disc"] if ef else 0.0
    ef_nao = ef["nao"] if ef else 0.0

    e_ativo = eng["ativo"] if eng else 0
    e_foi   = eng["foi"]   if eng else 0
    active  = e_ativo > 0
    multi   = e_ativo >= 2
    many    = e_ativo >= 3

    if c_sim is None: c_sim = 0.5
    overall = (p_net + a_net + ef_net) / 3

    # "Disse SIM": Pertencimento net ≥ threshold E discordância < threshold
    said_yes = (p_net >= THRESH_YES) and (p_disc < THRESH_DISC)

    # ── Modificador de tempo ──────────────────────────────────────────────────
    if years >= 5 and not said_yes:
        if p_disc >= 0.25 or a_disc >= 0.30:
            return 2, (
                f"Tempo na escola: {years} anos. Discordância em Pertencimento ({p_disc:.0%}) "
                f"e/ou Acordos ({a_disc:.0%}) → tensão com a escola"
            )
        return 3, (
            f"Tempo na escola: {years} anos sem comprometimento firme "
            f"(Pertencimento net={p_net:.0%}) → no ponto de virada há muito tempo"
        )

    # ── Lado esquerdo (não disse SIM) ─────────────────────────────────────────
    if not said_yes:
        if p_disc >= 0.40:
            return 1, f"Alta discordância em Pertencimento ({p_disc:.0%}) → resistência/desalinhamento"
        if p_nao >= 0.50:
            return 1, f"Muitas respostas 'Não sei' em Pertencimento ({p_nao:.0%}) → não reconhece a escola ainda"
        if c_sim >= 0.75 and p_pos >= 0.50:
            return 3, f"Conhecimento alto ({c_sim:.0%} Sim) + Pertencimento parcial ({p_pos:.0%} positivo)"
        if c_sim >= 0.50 or p_pos >= 0.35:
            return 2, f"Conhecimento ({c_sim:.0%}) ou Pertencimento ({p_pos:.0%}) em construção"
        return 1, f"Conhecimento ({c_sim:.0%}) e Pertencimento ({p_pos:.0%}) ainda baixos"

    # ── Lado direito (disse SIM) ──────────────────────────────────────────────
    if a_disc >= THRESH_DISC and ef_disc >= THRESH_DISC:
        return 4, (
            f"Diz SIM, mas discordância em Acordos ({a_disc:.0%}) "
            f"e Ec. Fraterna ({ef_disc:.0%}) → tensões nas práticas"
        )
    if a_disc >= THRESH_DISC:
        return 4, f"Diz SIM, mas discordância em Acordos Sociais ({a_disc:.0%}) → tensão nos acordos"
    if ef_disc >= THRESH_DISC:
        return 4, f"Diz SIM, mas discordância em Ec. Fraterna ({ef_disc:.0%}) → tensão econômica"
    if a_nao >= 0.40 and ef_nao >= 0.40:
        return 4, (
            f"Diz SIM, mas alto 'Não sei' em Acordos ({a_nao:.0%}) "
            f"e Ec. Fraterna ({ef_nao:.0%}) → ainda aprendendo as práticas"
        )
    if not active:
        return 4, "Diz SIM para a escola, mas sem papel ativo no momento"
    if many and overall >= 0.75 and ef_net >= 0.60:
        return 7, (
            f"{e_ativo} papéis ativos + concordância global ({overall:.0%}) "
            f"+ Ec. Fraterna forte ({ef_net:.0%})"
        )
    if multi and overall >= 0.70:
        return 6, f"{e_ativo} papéis ativos + concordância global ({overall:.0%})"
    if active and overall >= 0.60:
        return 5, f"{e_ativo} papel(éis) ativo(s) + concordância ({overall:.0%})"
    return 4, f"Papel ativo, mas concordância global ainda moderada ({overall:.0%})"

# ─── Alinhamento ──────────────────────────────────────────────────────────────

def alinhamento_label(q10, inf):
    if q10 is None: return "—"
    d = inf - q10
    if d == 0: return "Alinhado"
    if d == 1: return "Potencial +1 nível"
    if d >= 2: return "Potencial +2 ou mais níveis"
    if d == -1: return "Autoposicionamento +1 acima"
    return "Autoposicionamento +2 ou mais acima"

def parse_ano(v):
    if v is None or (isinstance(v, float) and pd.isna(v)): return None
    sv = str(v).strip()
    if "antes" in sv.lower(): return 2014
    try: return int(float(sv))
    except: return None

def parse_q10(v):
    if v is None or (isinstance(v, float) and pd.isna(v)): return None, "—"
    s = str(v).strip()
    for num, nome in ESTAGIO_NOME.items():
        if norm(nome) in norm(s): return num, nome
    return None, "—"

# ─── Estilos ──────────────────────────────────────────────────────────────────

COR_ESTAGIO = {
    "Novo": "F4CCCC", "Adaptando": "FCE5CD", "Inserindo": "FFF2CC",
    "Comprometido": "D9EAD3", "Ativo no Motivo": "C9DAF8",
    "Embaixador": "D9D2E9", "Cocriador": "E6B8A2",
}
COR_ALIGN = {
    "Alinhado": "D9EAD3",
    "Potencial +1 nível": "C9DAF8",
    "Potencial +2 ou mais níveis": "9FC5E8",
    "Autoposicionamento +1 acima": "FFE599",
    "Autoposicionamento +2 ou mais acima": "EA9999",
}

def hdr(ws, row=1):
    for cell in ws[row]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F3864")
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

def auto_w(ws):
    for col in ws.iter_cols():
        mx = max((len(str(c.value or "")) for c in col), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(mx + 3, 52)

def pf(v):
    return f"{v:.0%}" if v is not None else "—"

# ─── Main ──────────────────────────────────────────────────────────────────────

def main():
    df_resp, df_pd = carrega()
    resp_cols = list(df_resp.columns)

    esfera_cols = mapeia_esferas(df_pd, resp_cols)
    print("Esferas mapeadas:")
    for k, v in esfera_cols.items():
        print(f"  {k}: {len(v)} colunas")

    col_nome    = next((c for c in resp_cols if "nome completo" in norm(c)), None)
    col_ano     = next((c for c in resp_cols if "em que ano" in norm(c) and "arandu" in norm(c)), None)
    col_vinculo = next((c for c in resp_cols if "vinculo" in norm(c) and "escola" in norm(c)), None)
    col_q10     = next((c for c in resp_cols if "com base na imagem" in norm(c)), None)

    cols_c   = esfera_cols.get("Conhecimento institucional", [])
    cols_p   = esfera_cols.get("Pertencimento cultural", [])
    cols_a   = esfera_cols.get("Acordos sociais", [])
    cols_ef  = esfera_cols.get("Prática da Economia Fraterna", [])
    cols_eng = esfera_cols.get("Engajamento", [])

    print(f"Colunas: C={len(cols_c)}, P={len(cols_p)}, A={len(cols_a)}, EF={len(cols_ef)}, Eng={len(cols_eng)}")

    rows_out = []
    for _, row in df_resp.iterrows():
        nome = str(row[col_nome]).strip() if col_nome and pd.notna(row.get(col_nome)) else "—"
        entry_year = parse_ano(row.get(col_ano) if col_ano else None)
        years = (ANO_REF - entry_year) if entry_year else None

        vv = row.get(col_vinculo) if col_vinculo else None
        vinculo = str(vv).strip() if vv and pd.notna(vv) else "—"

        q10_num, q10_nome = parse_q10(row.get(col_q10) if col_q10 else None)

        st_c   = stats_sim_nao(row, cols_c)   if cols_c   else None
        st_p   = stats_concordancia(row, cols_p)  if cols_p  else None
        st_a   = stats_concordancia(row, cols_a)  if cols_a  else None
        st_ef  = stats_concordancia(row, cols_ef) if cols_ef else None
        st_eng = stats_engajamento(row, cols_eng) if cols_eng else None

        c_sim = st_c["sim"] if st_c else None
        inf_num, justif = inferred_stage(c_sim, st_p, st_a, st_ef, st_eng, entry_year)
        inf_nome = ESTAGIO_NOME[inf_num]

        next_nome = ESTAGIO_NOME[min(inf_num + 1, 7)]
        align = alinhamento_label(q10_num, inf_num)

        rows_out.append({
            "Nome": nome,
            "Ano de entrada": entry_year or "—",
            "Anos na escola": years if years is not None else "—",
            "Vínculo": vinculo,
            "Autoposicionamento Q10": q10_nome,
            "Nível inferido pelas respostas": inf_nome,
            "Alinhamento Q10 vs Inferido": align,
            "Justificativa da inferência": justif,
            "Próximo nível": next_nome,
            "Ação prioritária": PROXIMA_ACAO[inf_num],
            "Papel atual na comunidade": PAPEIS[inf_num],
            # Conhecimento
            "Conhecimento (% Sim)": pf(c_sim),
            # Pertencimento
            "Pertenc. % Concordo": pf(st_p["pleno"] if st_p else None),
            "Pertenc. % Concordo Parc.": pf(st_p["parc"] if st_p else None),
            "Pertenc. % Não sei": pf(st_p["nao"] if st_p else None),
            "Pertenc. % Discordo": pf(st_p["disc"] if st_p else None),
            # Acordos
            "Acordos % Concordo": pf(st_a["pleno"] if st_a else None),
            "Acordos % Concordo Parc.": pf(st_a["parc"] if st_a else None),
            "Acordos % Não sei": pf(st_a["nao"] if st_a else None),
            "Acordos % Discordo": pf(st_a["disc"] if st_a else None),
            # Ec. Fraterna
            "Ec. Fraterna % Concordo": pf(st_ef["pleno"] if st_ef else None),
            "Ec. Fraterna % Concordo Parc.": pf(st_ef["parc"] if st_ef else None),
            "Ec. Fraterna % Não sei": pf(st_ef["nao"] if st_ef else None),
            "Ec. Fraterna % Discordo": pf(st_ef["disc"] if st_ef else None),
            # Engajamento
            "Engajamento: papéis ativos": st_eng["ativo"] if st_eng else "—",
            "Engajamento: histórico (já foi)": st_eng["foi"] if st_eng else "—",
        })

    df_out = pd.DataFrame(rows_out)
    print(f"\nTotal: {len(df_out)}")
    print("\nNível Inferido:")
    for e in ESTAGIOS:
        n = (df_out["Nível inferido pelas respostas"] == e).sum()
        if n: print(f"  {e}: {n}")
    print("\nAlinhamento:")
    print(df_out["Alinhamento Q10 vs Inferido"].value_counts().to_string())

    # ── Export ────────────────────────────────────────────────────────────────
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    # Aba 1 — Classificação Individual
    ws1 = wb_out.create_sheet("Classificação Individual")
    cols = list(df_out.columns)
    ws1.append(cols)
    hdr(ws1, 1)
    col_inf_i  = cols.index("Nível inferido pelas respostas") + 1
    col_align_i = cols.index("Alinhamento Q10 vs Inferido") + 1

    for i, (_, row) in enumerate(df_out.iterrows(), start=2):
        ws1.append([row[c] for c in cols])
        inf_v = row["Nível inferido pelas respostas"]
        if inf_v in COR_ESTAGIO:
            ws1.cell(i, col_inf_i).fill = PatternFill("solid", fgColor=COR_ESTAGIO[inf_v])
        al_v = row["Alinhamento Q10 vs Inferido"]
        if al_v in COR_ALIGN:
            ws1.cell(i, col_align_i).fill = PatternFill("solid", fgColor=COR_ALIGN[al_v])

    ws1.freeze_panes = "F2"
    auto_w(ws1)

    # Aba 2 — Distribuição de níveis inferidos
    ws2 = wb_out.create_sheet("Distribuição Nível Inferido")
    ws2.append(["Nível Inferido", "N", "%", "Interpretação"])
    hdr(ws2, 1)
    interp = {
        "Novo": "Ainda em reconhecimento — não internalizou propósito/missão",
        "Adaptando": "Assimilando como a escola funciona — pode ter tensões ou ser recente",
        "Inserindo": "No ponto de virada — precisa decidir SIM ou NÃO",
        "Comprometido": "Disse SIM, mas ainda sem papel ativo estruturado",
        "Ativo no Motivo": "Engajado com ao menos 1 papel ativo na escola",
        "Embaixador": "2+ papéis ativos + alta concordância — referência comunitária",
        "Cocriador": "Co-cria o futuro da escola com papéis estratégicos e alta concordância",
    }
    total = len(df_out)
    for est in ESTAGIOS:
        n = (df_out["Nível inferido pelas respostas"] == est).sum()
        row_vals = [est, int(n), f"{100*n/total:.1f}%", interp.get(est, "")]
        ws2.append(row_vals)
        cor = COR_ESTAGIO.get(est)
        if cor:
            for ci in range(1, 5):
                ws2.cell(ws2.max_row, ci).fill = PatternFill("solid", fgColor=cor)
    ws2.append(["Total", total, "100%", ""])
    ws2.cell(ws2.max_row, 1).font = Font(bold=True)
    auto_w(ws2)

    # Aba 3 — Cruzamento Q10 × Inferido
    ws3 = wb_out.create_sheet("Q10 × Inferido")
    ws3.append(["Q10 ↓  /  Inferido →"] + ESTAGIOS + ["Total"])
    hdr(ws3, 1)
    df_cross = df_out[["Autoposicionamento Q10", "Nível inferido pelas respostas"]].copy()
    for q10_e in ESTAGIOS:
        sub = df_cross[df_cross["Autoposicionamento Q10"] == q10_e]
        if len(sub) == 0: continue
        row_v = [q10_e]
        for inf_e in ESTAGIOS:
            n = (sub["Nível inferido pelas respostas"] == inf_e).sum()
            row_v.append(int(n) if n else "")
        row_v.append(len(sub))
        ws3.append(row_v)
    # Total por coluna
    total_row = ["Total"]
    for inf_e in ESTAGIOS:
        n = (df_cross["Nível inferido pelas respostas"] == inf_e).sum()
        total_row.append(int(n) if n else "")
    total_row.append(len(df_cross))
    ws3.append(total_row)
    ws3.cell(ws3.max_row, 1).font = Font(bold=True)
    ws3.freeze_panes = "B2"
    auto_w(ws3)

    # Aba 4 — Alinhamento
    ws4 = wb_out.create_sheet("Alinhamento Q10 vs Inferido")
    ws4.append(["Categoria", "N", "%", "Interpretação"])
    hdr(ws4, 1)
    interp_al = {
        "Alinhado": "Percepção própria (Q10) coincide com o nível inferido pelas respostas",
        "Potencial +1 nível": "Respostas mostram 1 nível acima do autoposicionamento",
        "Potencial +2 ou mais níveis": "Respostas mostram 2+ níveis acima — subavaliação relevante",
        "Autoposicionamento +1 acima": "Declara 1 nível acima do que as respostas sustentam",
        "Autoposicionamento +2 ou mais acima": "Declara 2+ níveis acima — sobreposicionamento significativo",
        "—": "Sem resposta em Q10",
    }
    dist_al = df_out["Alinhamento Q10 vs Inferido"].value_counts()
    ordem_al = ["Alinhado", "Potencial +1 nível", "Potencial +2 ou mais níveis",
                "Autoposicionamento +1 acima", "Autoposicionamento +2 ou mais acima", "—"]
    for cat in ordem_al:
        n = dist_al.get(cat, 0)
        ws4.append([cat, int(n), f"{100*n/total:.1f}%", interp_al.get(cat, "")])
        cor = COR_ALIGN.get(cat)
        if cor: ws4.cell(ws4.max_row, 1).fill = PatternFill("solid", fgColor=cor)
    ws4.append(["Total", total, "100%", ""])
    ws4.cell(ws4.max_row, 1).font = Font(bold=True)
    auto_w(ws4)

    # Aba 5 — Legenda
    ws5 = wb_out.create_sheet("Legenda da Inferência")
    ws5.append(["Elemento", "Detalhe"])
    hdr(ws5, 1)
    leg = [
        ("THRESHOLDS PRINCIPAIS", ""),
        ("'Disse SIM'", f"Pertencimento net ≥ {THRESH_YES:.0%} E discordância Pertencimento < {THRESH_DISC:.0%}"),
        ("Net Signal", "Concordo×1.0 + Concordo parc×0.4 + Não sei×(−0.2) + Discordo×(−0.8) → normaliz. [0,1]"),
        ("Discordância relevante", f"≥ {THRESH_DISC:.0%} em qualquer esfera"),
        ("", ""),
        ("MODIFICADOR DE TEMPO", ""),
        ("≥ 5 anos + não disse SIM + disc. ≥ 25%", "→ Adaptando (tensão/sinal de possível saída)"),
        ("≥ 5 anos + não disse SIM", "→ Inserindo (no ponto de virada há muito tempo)"),
        ("", ""),
        ("LADO ESQUERDO — não disse SIM", ""),
        ("Discordância Pertencimento ≥ 40%", "→ Novo (resistência)"),
        ("Não sei Pertencimento ≥ 50%", "→ Novo (não reconhece a escola ainda)"),
        ("Conhecimento ≥ 75% + positivo Pertencimento ≥ 50%", "→ Inserindo"),
        ("Conhecimento ≥ 50% ou positivo Pertencimento ≥ 35%", "→ Adaptando"),
        ("Demais casos", "→ Novo"),
        ("", ""),
        ("LADO DIREITO — disse SIM", ""),
        ("Discordância Acordos ≥ 25% E Ec.Fraterna ≥ 25%", "→ Comprometido (tensões práticas)"),
        ("Discordância Acordos ≥ 25%", "→ Comprometido (tensão nos acordos)"),
        ("Discordância Ec. Fraterna ≥ 25%", "→ Comprometido (tensão econômica)"),
        ("Não sei Acordos ≥ 40% E Não sei EF ≥ 40%", "→ Comprometido (ainda aprendendo práticas)"),
        ("Sem papel ativo", "→ Comprometido"),
        ("3+ papéis ativos + overall ≥ 75% + EF ≥ 60%", "→ Cocriador"),
        ("2+ papéis ativos + overall ≥ 70%", "→ Embaixador"),
        ("1+ papel ativo + overall ≥ 60%", "→ Ativo no Motivo"),
        ("Papel ativo + overall < 60%", "→ Comprometido"),
        ("", ""),
        ("NET SIGNAL — exemplos", ""),
        ("Todos Concordo", f"net = (1,0+0,8)/1,8 = 1,00"),
        ("Todos Concordo parcialmente", f"net = (0,4+0,8)/1,8 = 0,67"),
        ("Todos Não sei opinar", f"net = (−0,2+0,8)/1,8 = 0,33"),
        ("Todos Discordo", f"net = (−0,8+0,8)/1,8 = 0,00"),
    ]
    for item in leg:
        ws5.append(list(item))
    auto_w(ws5)

    wb_out.save(SAIDA)
    print(f"\n✓ Salvo: {SAIDA}")


if __name__ == "__main__":
    main()
