#!/usr/bin/env python3
"""
Processa pesquisa de maturidade Arandu.
Dimensões: Conhecimento (20%), Atitude (35%), Prática (45%).
Normalização por pergunta: cada stage é reescalado para [1,7] com base no
min/max possível daquela pergunta na Lógica, tornando as escalas comparáveis
antes da ponderação.
"""

import re
import unicodedata
import argparse
from collections import Counter, defaultdict
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ─── Constantes ───────────────────────────────────────────────────────────────

ESTAGIOS = ["Novo", "Adaptando", "Inserindo", "Comprometido", "Ativo no Motivo", "Embaixador", "Cocriador"]
ESTAGIO_NUM  = {e: i+1 for i, e in enumerate(ESTAGIOS)}
ESTAGIO_NOME = {i+1: e for i, e in enumerate(ESTAGIOS)}

PESOS = {"Conhecimento": 0.20, "Atitude": 0.35, "Prática": 0.45}
DIMENSOES = list(PESOS.keys())

PERGUNTA_ANO = "1. Em que ano você entrou na Arandu?"

# Perguntas excluídas do cálculo de score (autoposicionamento subjetivo)
PERGUNTAS_EXCLUIR = [
    "com base na imagem anterior, marque a posicao que mais reflete",
]

# ─── Utilidades ───────────────────────────────────────────────────────────────

def normaliza(s):
    s = str(s) if s is not None else ""
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s.lower().strip())

def extrai_colchetes(s):
    m = re.search(r"\[([^\[\]]+)\]\s*$", str(s))
    return m.group(1).strip() if m else None

def normaliza_estagio(s):
    if s is None:
        return None
    n = normaliza(str(s))
    if not n:
        return None
    if "co" in n and "cri" in n:
        return 7
    if "embaixador" in n:
        return 6
    if "ativo" in n:
        return 5
    if "comprometido" in n:
        return 4
    if "inserindo" in n:
        return 3
    if "adaptando" in n:
        return 2
    if "novo" in n:
        return 1
    return None

def estagio_do_score(score):
    if score is None or (isinstance(score, float) and pd.isna(score)):
        return None
    if score >= 6.50: return "Cocriador"
    if score >= 5.50: return "Embaixador"
    if score >= 4.50: return "Ativo no Motivo"
    if score >= 3.50: return "Comprometido"
    if score >= 2.50: return "Inserindo"
    if score >= 1.50: return "Adaptando"
    return "Novo"

def casa_textos(a, b):
    na, nb = normaliza(a), normaliza(b)
    if na == nb:
        return True
    ca, cb = extrai_colchetes(a), extrai_colchetes(b)
    if ca and cb and normaliza(ca) == normaliza(cb):
        return True
    return False

def normaliza_score(stage, q_min, q_max):
    """Reescala stage bruto para [1,7] com base no range da pergunta."""
    if q_max == q_min:
        return float(stage)
    return 1.0 + 6.0 * (stage - q_min) / (q_max - q_min)

# ─── Carregamento ─────────────────────────────────────────────────────────────

def carrega_planilha(caminho):
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)

    def aba_df(nome):
        ws = wb[nome]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return pd.DataFrame()
        header = [str(c) if c is not None else f"_col{i}" for i, c in enumerate(rows[0])]
        data = [list(r) for r in rows[1:] if any(v is not None for v in r)]
        return pd.DataFrame(data, columns=header)

    return (
        aba_df("Respostas ao formulário 1"),
        aba_df("Lógica"),
        aba_df("Perguntas_Dimensoes"),
    )

# ─── Mapa de lógica + ranges por pergunta ────────────────────────────────────

def constroi_logica(df_logica):
    """
    Retorna:
      mapa:      {(norm_perg, norm_resp) → stage_num}
      mapa_chave:{norm_bracket_key → {norm_resp → stage_num}}
      q_ranges:  {norm_perg → (min_stage, max_stage)}
                 também indexado por norm_bracket_key para lookup rápido
    """
    mapa = {}
    mapa_chave = {}
    q_stages_raw = defaultdict(list)   # norm_perg → [stages]
    q_stages_ck  = defaultdict(list)   # norm_bracket_key → [stages]

    col_perg  = df_logica.columns[0]
    col_resp  = df_logica.columns[1]
    col_stage = df_logica.columns[2]

    for _, row in df_logica.iterrows():
        perg  = row[col_perg]
        resp  = row[col_resp]
        stage = row[col_stage]
        if pd.isna(perg) or perg is None:
            continue
        num = normaliza_estagio(stage)
        if num is None:
            continue
        np_ = normaliza(str(perg))
        nr_ = normaliza(str(resp)) if resp is not None else ""
        mapa[(np_, nr_)] = num
        q_stages_raw[np_].append(num)

        chave = extrai_colchetes(str(perg))
        if chave:
            ck = normaliza(chave)
            mapa_chave.setdefault(ck, {})[nr_] = num
            q_stages_ck[ck].append(num)

    # Ranges: por norm_perg e por bracket key
    q_ranges_np = {np_: (min(v), max(v)) for np_, v in q_stages_raw.items()}
    q_ranges_ck = {ck: (min(v), max(v)) for ck, v in q_stages_ck.items()}

    return mapa, mapa_chave, q_ranges_np, q_ranges_ck

# ─── Busca stage + range ─────────────────────────────────────────────────────

def busca_stage_e_range(col_name, resp_value, mapa, mapa_chave, q_ranges_np, q_ranges_ck):
    """
    Retorna (stage_num, q_min, q_max) ou (None, None, None).
    q_min/q_max definem o range teórico da pergunta para normalização.
    """
    if resp_value is None or (isinstance(resp_value, float) and pd.isna(resp_value)):
        return None, None, None
    s = str(resp_value).strip()
    if not s:
        return None, None, None

    nr = normaliza(s)
    nc = normaliza(col_name)
    chave_col = extrai_colchetes(col_name)
    ck = normaliza(chave_col) if chave_col else None

    stage = None

    # 1. Exato por (norm_perg, norm_resp)
    if (nc, nr) in mapa:
        stage = mapa[(nc, nr)]

    # 2. Por bracket key
    if stage is None and ck and ck in mapa_chave and nr in mapa_chave[ck]:
        stage = mapa_chave[ck][nr]

    # 3. Busca parcial por prefixo do texto da pergunta (sem colchetes)
    if stage is None:
        for (np_, nr_), num in mapa.items():
            if nr_ == nr and np_ == nc:
                stage = num
                break

    # 4. Resposta auto-contém o estágio (ex: "10.Com base..." / autoposicionamento)
    if stage is None:
        parte = s.split(" - ")[0].split(":")[0].strip()
        stage = normaliza_estagio(parte)

    if stage is None:
        return None, None, None

    # Range da pergunta
    q_min, q_max = 1, 7   # default: escala completa
    if nc in q_ranges_np:
        q_min, q_max = q_ranges_np[nc]
    elif ck and ck in q_ranges_ck:
        q_min, q_max = q_ranges_ck[ck]
    else:
        pass  # sem match: usa default [1,7]

    return stage, q_min, q_max

# ─── Mapeamento dimensões → colunas ──────────────────────────────────────────

def mapeia_dimensoes(df_pd, resp_cols):
    col_para_dim = {}
    sem_match = []

    col_perg = df_pd.columns[0]
    col_dim  = df_pd.columns[1]

    for _, row in df_pd.iterrows():
        perg = row[col_perg]
        dim  = row[col_dim]
        if pd.isna(perg) or perg is None:
            continue
        if pd.isna(dim) or dim is None:
            continue

        np_ = normaliza(str(perg))
        if any(excl in np_ for excl in PERGUNTAS_EXCLUIR):
            continue

        encontrado = False
        for rc in resp_cols:
            if casa_textos(str(perg), rc):
                col_para_dim[rc] = str(dim).strip()
                encontrado = True
                break

        if not encontrado:
            sem_match.append((str(perg), str(dim)))

    return col_para_dim, sem_match

# ─── Processamento ────────────────────────────────────────────────────────────

def processa(df_resp, df_logica, df_pd):
    mapa, mapa_chave, q_ranges_np, q_ranges_ck = constroi_logica(df_logica)

    resp_cols = list(df_resp.columns)
    col_para_dim, sem_match_pd = mapeia_dimensoes(df_pd, resp_cols)

    col_nome = next(
        (c for c in resp_cols if normaliza(c) == normaliza("1. Nome Completo")),
        next((c for c in resp_cols if "nome completo" in normaliza(c)), None)
    )
    col_ano = next(
        (c for c in resp_cols if normaliza(c) == normaliza(PERGUNTA_ANO)),
        next((c for c in resp_cols if "em que ano" in normaliza(c) and "arandu" in normaliza(c)), None)
    )
    col_vinculo = next(
        (c for c in resp_cols if "vinculo" in normaliza(c) and "escola" in normaliza(c)), None
    )

    dims_cols = defaultdict(list)
    for col, dim in col_para_dim.items():
        if dim in DIMENSOES:
            dims_cols[dim].append(col)

    sem_logica = set()
    respondentes_sem_respostas = []
    indicadores = []
    pontuacoes_detalhes = []

    for _, row in df_resp.iterrows():
        nome = str(row[col_nome]).strip() if col_nome and not pd.isna(row.get(col_nome)) else "—"

        ano_raw = row.get(col_ano) if col_ano else None
        if ano_raw is None or (isinstance(ano_raw, float) and pd.isna(ano_raw)):
            ano = None
        else:
            try:
                ano = int(float(str(ano_raw)))
            except Exception:
                ano = str(ano_raw)

        vinculo_raw = row.get(col_vinculo) if col_vinculo else None
        vinculo = (
            str(vinculo_raw).strip()
            if vinculo_raw and not (isinstance(vinculo_raw, float) and pd.isna(vinculo_raw))
            else "—"
        )

        scores_dim = {}
        stages_todos = []
        n_validas = 0

        for dim in DIMENSOES:
            scores_norm = []
            for col in dims_cols[dim]:
                val = row.get(col)
                stage, q_min, q_max = busca_stage_e_range(
                    col, val, mapa, mapa_chave, q_ranges_np, q_ranges_ck
                )
                if stage is not None:
                    score_norm = normaliza_score(stage, q_min, q_max)
                    scores_norm.append(score_norm)
                    stages_todos.append(stage)
                    n_validas += 1
                    pontuacoes_detalhes.append({
                        "Nome": nome,
                        "Dimensão": dim,
                        "Pergunta": col[:80],
                        "Resposta": str(val)[:60] if val is not None else "",
                        "Estágio bruto": stage,
                        "Estágio bruto (nome)": ESTAGIO_NOME.get(stage, "?"),
                        "Range pergunta (min)": q_min,
                        "Range pergunta (max)": q_max,
                        "Score normalizado [1-7]": round(score_norm, 2),
                    })
                elif val is not None and not (isinstance(val, float) and pd.isna(val)) and str(val).strip():
                    sem_logica.add(f"{col[:60]} → {str(val)[:40]}")

            if scores_norm:
                scores_dim[dim] = float(sum(scores_norm)) / len(scores_norm)

        if not scores_dim:
            respondentes_sem_respostas.append(nome)
            continue

        # Score Final ponderado (normaliza pesos se alguma dimensão faltar)
        peso_total = sum(PESOS[d] for d in scores_dim)
        score_final = sum(scores_dim[d] * PESOS[d] for d in scores_dim) / peso_total

        estagio_fin = estagio_do_score(score_final)

        # Estágio Predominante: stage bruto mais frequente entre as respostas
        if stages_todos:
            estagio_pred_num = Counter(stages_todos).most_common(1)[0][0]
            estagio_pred = ESTAGIO_NOME.get(estagio_pred_num, "?")
        else:
            estagio_pred = "—"

        indicadores.append({
            "Nome": nome,
            "Ano de entrada": ano,
            "Vínculo": vinculo,
            "Nº respostas válidas": n_validas,
            "Score Conhecimento": round(scores_dim.get("Conhecimento", float("nan")), 2),
            "Score Atitude": round(scores_dim.get("Atitude", float("nan")), 2),
            "Score Prática": round(scores_dim.get("Prática", float("nan")), 2),
            "Score Final": round(score_final, 2),
            "Estágio Final": estagio_fin,
            "Estágio Predominante": estagio_pred,
        })

    df_ind = pd.DataFrame(indicadores)
    df_det = pd.DataFrame(pontuacoes_detalhes)

    validacoes = {
        "sem_match_pd": sem_match_pd,
        "sem_logica": sorted(sem_logica),
        "respondentes_sem_respostas": respondentes_sem_respostas,
    }

    return df_ind, df_det, col_para_dim, validacoes

# ─── Resumo_Comunidade ────────────────────────────────────────────────────────

def resumo_comunidade(df_ind):
    rows = []
    rows.append({"Indicador": "Total de respondentes", "Valor": len(df_ind)})
    for dim in DIMENSOES:
        col = f"Score {dim}"
        val = df_ind[col].mean() if col in df_ind and not df_ind[col].isna().all() else float("nan")
        rows.append({"Indicador": f"Score {dim} médio (normalizado)", "Valor": round(val, 2) if not pd.isna(val) else "—"})
    rows.append({"Indicador": "Score Final médio", "Valor": round(df_ind["Score Final"].mean(), 2)})
    rows.append({"Indicador": "Estágio Final médio", "Valor": estagio_do_score(df_ind["Score Final"].mean())})
    rows.append({"Indicador": "", "Valor": ""})

    rows.append({"Indicador": "--- Distribuição por Estágio Final ---", "Valor": ""})
    dist = df_ind["Estágio Final"].value_counts()
    for est in ESTAGIOS:
        n = dist.get(est, 0)
        pct = round(100 * n / len(df_ind), 1) if len(df_ind) > 0 else 0
        rows.append({"Indicador": est, "Valor": f"{n} ({pct}%)"})

    rows.append({"Indicador": "", "Valor": ""})
    rows.append({"Indicador": "--- Distribuição por Estágio Predominante ---", "Valor": ""})
    dist_pred = df_ind["Estágio Predominante"].value_counts()
    for est in ESTAGIOS:
        n = dist_pred.get(est, 0)
        pct = round(100 * n / len(df_ind), 1) if len(df_ind) > 0 else 0
        rows.append({"Indicador": est, "Valor": f"{n} ({pct}%)"})

    rows.append({"Indicador": "", "Valor": ""})
    rows.append({"Indicador": "--- Score Final por Vínculo ---", "Valor": ""})
    if "Vínculo" in df_ind.columns:
        for vinc, grp in sorted(df_ind.groupby("Vínculo")):
            rows.append({
                "Indicador": str(vinc),
                "Valor": f"n={len(grp)}, score={round(grp['Score Final'].mean(), 2)}, estágio={estagio_do_score(grp['Score Final'].mean())}"
            })

    return pd.DataFrame(rows)

# ─── Tempo_Vinculo ────────────────────────────────────────────────────────────

def tempo_vinculo(df_ind):
    if "Ano de entrada" not in df_ind.columns or df_ind["Ano de entrada"].isna().all():
        return pd.DataFrame({"Nota": ["Coluna 'Ano de entrada' não encontrada."]})

    df = df_ind[["Ano de entrada", "Estágio Final", "Score Final"]].copy()
    df["Ano de entrada"] = pd.to_numeric(df["Ano de entrada"], errors="coerce")
    df = df.dropna(subset=["Ano de entrada"])
    df["Ano de entrada"] = df["Ano de entrada"].astype(int)

    rows = []
    for ano, grp in df.sort_values("Ano de entrada").groupby("Ano de entrada"):
        row = {
            "Ano de entrada": ano,
            "N": len(grp),
            "Score Final médio": round(grp["Score Final"].mean(), 2),
            "Estágio": estagio_do_score(grp["Score Final"].mean()),
        }
        for est in ESTAGIOS:
            row[est] = int((grp["Estágio Final"] == est).sum())
        rows.append(row)

    return pd.DataFrame(rows)

# ─── Auditoria Perguntas_Dimensoes ───────────────────────────────────────────

def auditoria_pd(col_para_dim, df_pd, q_ranges_np, q_ranges_ck):
    rows = []
    col_perg = df_pd.columns[0]
    col_dim  = df_pd.columns[1]

    for _, row in df_pd.iterrows():
        perg = row[col_perg]
        dim  = row[col_dim]
        if pd.isna(perg) or perg is None:
            continue
        matched_col = next((c for c in col_para_dim if casa_textos(str(perg), c)), None)
        matched = matched_col is not None

        # Range da pergunta
        nc = normaliza(str(perg))
        ck = normaliza(extrai_colchetes(str(perg))) if extrai_colchetes(str(perg)) else None
        rng = q_ranges_np.get(nc) or (q_ranges_ck.get(ck) if ck else None)
        # sem fallback de prefixo: bracket key já cobre os casos necessários

        rows.append({
            "Pergunta": str(perg)[:100],
            "Dimensão": str(dim) if not (isinstance(dim, float) and pd.isna(dim)) else "",
            "Coluna encontrada": "Sim" if matched else "NÃO",
            "Stage mín (Lógica)": rng[0] if rng else "—",
            "Stage máx (Lógica)": rng[1] if rng else "—",
        })

    return pd.DataFrame(rows)

# ─── Exportação XLSX ──────────────────────────────────────────────────────────

COR_HEADER = "D6E4F0"
COR_ALT    = "EBF3FB"

def estilo_cabecalho(cell):
    cell.font = Font(bold=True)
    cell.fill = PatternFill("solid", fgColor=COR_HEADER)
    cell.alignment = Alignment(horizontal="center", wrap_text=True)

def auto_largura(ws):
    for col_cells in ws.iter_cols():
        max_len = max((len(str(c.value or "")) for c in col_cells), default=0)
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(max_len + 4, 60)

def df_para_aba(wb, nome_aba, df):
    if nome_aba in wb.sheetnames:
        del wb[nome_aba]
    ws = wb.create_sheet(nome_aba)
    if df.empty:
        ws.append(["(sem dados)"])
        return ws
    ws.append(list(df.columns))
    for cell in ws[1]:
        estilo_cabecalho(cell)
    for i, (_, row) in enumerate(df.iterrows()):
        ws.append([v if not (isinstance(v, float) and pd.isna(v)) else "" for v in row])
        if i % 2 == 1:
            for cell in ws[i + 2]:
                cell.fill = PatternFill("solid", fgColor=COR_ALT)
    ws.freeze_panes = "A2"
    auto_largura(ws)
    return ws

def exporta(df_ind, df_det, col_para_dim, validacoes, df_pd,
            q_ranges_np, q_ranges_ck, caminho_saida):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    df_para_aba(wb, "Indicadores", df_ind)

    df_res = resumo_comunidade(df_ind)
    df_para_aba(wb, "Resumo_Comunidade", df_res)

    df_tv = tempo_vinculo(df_ind)
    df_para_aba(wb, "Tempo_Vinculo", df_tv)

    df_audit = auditoria_pd(col_para_dim, df_pd, q_ranges_np, q_ranges_ck)
    df_para_aba(wb, "Perguntas_Dimensoes_Auditoria", df_audit)

    if not df_det.empty:
        df_para_aba(wb, "Pontuacoes_Detalhadas", df_det)

    val_rows = []
    val_rows.append({"Categoria": "Perguntas sem correspondência", "Detalhe": ""})
    for perg, dim in validacoes["sem_match_pd"]:
        val_rows.append({"Categoria": f"[{dim}] sem coluna", "Detalhe": perg[:120]})
    val_rows.append({"Categoria": "", "Detalhe": ""})
    val_rows.append({"Categoria": "Respostas sem lógica mapeada", "Detalhe": ""})
    for item in validacoes["sem_logica"]:
        val_rows.append({"Categoria": "sem lógica", "Detalhe": item[:120]})
    val_rows.append({"Categoria": "", "Detalhe": ""})
    val_rows.append({"Categoria": "Respondentes sem respostas válidas", "Detalhe": ""})
    for nome in validacoes["respondentes_sem_respostas"]:
        val_rows.append({"Categoria": "sem respostas", "Detalhe": nome})
    df_para_aba(wb, "Validacoes", pd.DataFrame(val_rows))

    cortes_rows = [
        {"Estágio": e, "Score mínimo": mn, "Score máximo": mx}
        for e, mn, mx in zip(
            ESTAGIOS,
            [1.00, 1.50, 2.50, 3.50, 4.50, 5.50, 6.50],
            [1.49, 2.49, 3.49, 4.49, 5.49, 6.49, 7.00],
        )
    ]
    df_para_aba(wb, "Cortes_Utilizados", pd.DataFrame(cortes_rows))

    wb.save(caminho_saida)
    print(f"\n✓ Arquivo salvo: {caminho_saida}")

# ─── CLI ──────────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(description="Processa pesquisa de maturidade Arandu")
    ap.add_argument("arquivo", help="Arquivo .xlsx de entrada")
    ap.add_argument("-o", "--saida", default="pesquisa_processada.xlsx")
    args = ap.parse_args()

    print(f"Carregando {args.arquivo}...")
    df_resp, df_logica, df_pd = carrega_planilha(args.arquivo)
    print(f"  Respostas: {len(df_resp)} respondentes, {len(df_resp.columns)} colunas")
    print(f"  Lógica: {len(df_logica)} linhas")
    print(f"  Perguntas_Dimensoes: {len(df_pd)} linhas")

    print("Processando...")
    df_ind, df_det, col_para_dim, validacoes = processa(df_resp, df_logica, df_pd)

    print(f"\n=== Resultados ===")
    print(f"  Respondentes processados: {len(df_ind)}")
    print(f"  Sem respostas válidas: {len(validacoes['respondentes_sem_respostas'])}")

    if not df_ind.empty:
        print(f"\n  Score Conhecimento médio (normalizado): {df_ind['Score Conhecimento'].mean():.2f}")
        print(f"  Score Atitude médio (normalizado):      {df_ind['Score Atitude'].mean():.2f}")
        print(f"  Score Prática médio (normalizado):      {df_ind['Score Prática'].mean():.2f}")
        print(f"  Score Final médio:                      {df_ind['Score Final'].mean():.2f}")
        print(f"  Estágio Final médio:                    {estagio_do_score(df_ind['Score Final'].mean())}")
        print(f"\n  Distribuição por Estágio Final:")
        dist = df_ind["Estágio Final"].value_counts()
        for est in ESTAGIOS:
            n = dist.get(est, 0)
            pct = round(100 * n / len(df_ind), 1)
            print(f"    {est:20s}: {n:3d} ({pct}%)")

    print(f"\n=== Validações ===")
    print(f"  Perguntas sem correspondência: {len(validacoes['sem_match_pd'])}")
    for perg, dim in validacoes["sem_match_pd"]:
        print(f"    [{dim}] {perg[:90]}")
    print(f"  Respostas sem lógica: {len(validacoes['sem_logica'])}")
    for item in list(validacoes["sem_logica"])[:10]:
        print(f"    {item}")

    # Para exportar precisamos dos ranges — recomputar do df_logica
    _, _, q_ranges_np2, q_ranges_ck2 = constroi_logica(df_logica)
    exporta(df_ind, df_det, col_para_dim, validacoes, df_pd,
            q_ranges_np2, q_ranges_ck2, args.saida)

if __name__ == "__main__":
    main()
