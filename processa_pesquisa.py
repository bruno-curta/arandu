#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Processamento da Pesquisa de Maturidade da Comunidade Escolar (Trimembração Social)
====================================================================================

Lê o arquivo .ods (ou .xlsx) da pesquisa, aplica a lógica de pontuação e
atribui a cada respondente um estágio da curva dos 7 Processos Sociais:

    1. Novo (respiração)
    2. Adaptando (aquecimento)
    3. Inserindo (digestão)
    4. Comprometido (segregação)
    5. Ativo no Motivo (manutenção)
    6. Embaixador (crescimento)
    7. Cocriador (geração)

Metodologia de pontuação
-------------------------
Cada resposta é convertida no número do estágio atribuído na aba Lógica (1–7)
e depois transformada pela função exponencial: pontos = BASE_EXP ** (estágio - 1).

O score final é a média desses pontos. Os cortes para classificação na curva
são calculados automaticamente com base no range teórico real da pesquisa
(mínimo possível → máximo possível), dividido em 7 faixas iguais.

Isso resolve o efeito de teto: perguntas com resposta máxima em "Inserindo"
têm peso natural menor que perguntas que chegam até "Embaixador" ou "Cocriador",
e respostas de estágios avançados pesam exponencialmente mais.

Uso
---
    python processa_pesquisa.py "Base da pesquisa oara chatgpt.ods" -o resultado.xlsx

Dependências: pandas, odfpy (para .ods), openpyxl (para gravar .xlsx)
    pip install pandas odfpy openpyxl
"""

import argparse
import re
import sys
import unicodedata

import pandas as pd

# ---------------------------------------------------------------------------
# Configuração
# ---------------------------------------------------------------------------

ESTAGIOS = [
    "Novo",            # 1 – respiração
    "Adaptando",       # 2 – aquecimento
    "Inserindo",       # 3 – digestão
    "Comprometido",    # 4 – segregação
    "Ativo no Motivo", # 5 – manutenção
    "Embaixador",      # 6 – crescimento
    "Cocriador",       # 7 – geração
]

# Base da escala exponencial: pontos = BASE_EXP ** (estágio - 1)
# Base 1 = linear; base 2 = cada estágio vale o dobro do anterior.
BASE_EXP = 2.0

# CORTES são calculados automaticamente a partir do range teórico real;
# esta constante é ignorada quando BASE_EXP > 1.
CORTES = [1.5, 2.5, 3.5, 4.5, 5.5, 6.5]

# Palavras-chave para localizar as abas
ABAS = {
    "respostas": ["resposta", "formulario", "formulário"],
    "perguntas": ["pergunta"],
    "logica":    ["logica", "lógica"],
}

# Padrões para auto-detectar o pilar da trimembração a partir do nome da coluna
PILARES_PATTERN = [
    (re.compile(r"pertencimento cultural", re.I),     "Pertencimento Cultural"),
    (re.compile(r"clareza dos acordos|acordo", re.I), "Acordos Sociais"),
    (re.compile(r"economia fraterna|fraterna", re.I), "Economia Fraterna"),
    (re.compile(r"participa[çc]", re.I),              "Engajamento e Participação"),
]


# ---------------------------------------------------------------------------
# Utilitários
# ---------------------------------------------------------------------------

def normaliza(texto):
    """Remove acentos, pontuação e espaços extras; retorna em minúsculas."""
    if texto is None or (isinstance(texto, float) and pd.isna(texto)):
        return ""
    texto = str(texto)
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"[^\w\s]", " ", texto.lower())
    return re.sub(r"\s+", " ", texto).strip()


def extrai_colchetes(texto):
    """Retorna o conteúdo do PRIMEIRO par de colchetes, ou None."""
    m = re.search(r"\[(.+?)\]", str(texto))
    return m.group(1).strip() if m else None


def acha_aba(nomes, chaves):
    """Encontra a aba cujo nome normalizado contém alguma das palavras-chave."""
    for nome in nomes:
        for chave in chaves:
            if normaliza(chave) in normaliza(nome):
                return nome
    return None


def estagio_para_numero(texto):
    """Converte nome/número do estágio para 1–7, ou None."""
    if texto is None or (isinstance(texto, float) and pd.isna(texto)):
        return None
    t = normaliza(str(texto))
    m = re.match(r"^(\d)", t)
    if m and 1 <= int(m.group(1)) <= 7:
        return int(m.group(1))
    tabela = {
        "novo": 1, "respiracao": 1,
        "adaptando": 2, "aquecimento": 2,
        "inserindo": 3, "digestao": 3,
        "comprometido": 4, "segregacao": 4,
        "ativo no motivo": 5, "manutencao": 5,
        "embaixador": 6, "crescimento": 6,
        "cocriador": 7, "geracao": 7, "co criador": 7,
    }
    for chave, num in tabela.items():
        if chave in t:
            return num
    return None


def numero_para_estagio(score, cortes=CORTES):
    if pd.isna(score):
        return None
    for i, corte in enumerate(cortes):
        if score <= corte:
            return ESTAGIOS[i]
    return ESTAGIOS[-1]


def detecta_pilar(nome_coluna):
    for padrao, pilar in PILARES_PATTERN:
        if padrao.search(nome_coluna):
            return pilar
    return None


# ---------------------------------------------------------------------------
# Carregamento
# ---------------------------------------------------------------------------

def carrega_abas(caminho):
    engine = "odf" if caminho.lower().endswith(".ods") else None
    todas = pd.read_excel(caminho, sheet_name=None, engine=engine)
    nomes = list(todas.keys())
    print(f"Abas encontradas: {nomes}")

    abas = {}
    for papel, chaves in ABAS.items():
        nome = acha_aba(nomes, chaves)
        if nome is None and papel in ("respostas", "logica"):
            sys.exit(f"ERRO: não encontrei a aba '{papel}'.\n"
                     f"Procurei por: {chaves}\nAbas disponíveis: {nomes}")
        if nome is not None:
            abas[papel] = todas[nome].dropna(how="all").dropna(how="all", axis=1)
            print(f"  -> aba '{papel}': '{nome}' "
                  f"({len(abas[papel])} linhas, {len(abas[papel].columns)} colunas)")
    return abas


# ---------------------------------------------------------------------------
# Lógica de pontuação
# ---------------------------------------------------------------------------

def monta_logica(df_logica):
    """
    Constrói o mapa de pontuação a partir da aba Lógica.
    Retorna (mapa, usa_estagio, col_pergunta_nome).
    mapa[(perg_norm, opcao_norm)] = pontos
    mapa[("", opcao_norm)] = pontos  (fallback global)
    """
    def acha_col(palavras):
        for col in df_logica.columns:
            cn = normaliza(col)
            if any(normaliza(k) in cn for k in palavras):
                return col
        return None

    col_pergunta = acha_col(["pergunta", "questao", "afirmacao"])
    col_opcao    = acha_col(["opcao", "resposta", "categoria"])
    col_valor    = acha_col(["valor", "pontuacao", "pontos", "peso", "proposta"])
    col_estagio  = acha_col(["estagio", "processo", "curva", "logica"])

    if col_opcao is None:
        sys.exit(f"ERRO: coluna de opção não encontrada na aba Lógica.\n"
                 f"Colunas disponíveis: {list(df_logica.columns)}")

    usa_estagio = False
    if col_estagio is not None:
        nums = df_logica[col_estagio].map(estagio_para_numero)
        if nums.notna().mean() > 0.4:
            usa_estagio = True
            df_logica = df_logica.assign(_pontos=nums)
            print(f"Pontuação baseada na coluna de estágio '{col_estagio}' (escala 1–7).")

    if not usa_estagio:
        if col_valor is None:
            sys.exit("ERRO: aba Lógica sem coluna de estágio nem de valor numérico.")
        df_logica = df_logica.assign(
            _pontos=pd.to_numeric(df_logica[col_valor], errors="coerce"))
        print(f"Pontuação baseada na coluna de valor '{col_valor}'.")

    mapa = {}
    for _, linha in df_logica.iterrows():
        pontos = linha["_pontos"]
        if pd.isna(pontos):
            continue
        opcao = normaliza(linha[col_opcao])
        if not opcao:
            continue
        perg_norm = normaliza(linha[col_pergunta]) if col_pergunta else ""
        mapa[(perg_norm, opcao)] = float(pontos)
        mapa.setdefault(("", opcao), float(pontos))

    print(f"Lógica carregada: {len(mapa)} entradas (pergunta + opção).")
    return mapa, usa_estagio, col_pergunta


# ---------------------------------------------------------------------------
# Casamento lógica → colunas de resposta
# ---------------------------------------------------------------------------

def _casa_col_para_lp(resp_col, lp, lp_chave_col):
    """
    Tenta casar uma coluna de resposta com um texto de lógica (lp).

    Estratégia (em ordem de preferência):
    1. Texto exato (normalizado).
    2. Colchetes: extrai o texto entre [ ] de ambos os lados e compara.
    3. Texto completo: lp_n sem colchetes == col_n sem colchetes (para questões
       sem colchetes na lógica, e.g. "1. Em que ano...").

    NÃO usa casamento por prefixo, que causaria falsos positivos entre questões
    com o mesmo prefixo ("Marque o nível de concordância...").
    """
    col_n = normaliza(resp_col)
    lp_n  = normaliza(lp)

    # 1. Texto exato
    if col_n == lp_n:
        return True

    # 2. Colchetes
    col_chave = extrai_colchetes(resp_col)
    if col_chave and lp_chave_col:
        if normaliza(col_chave) == normaliza(lp_chave_col):
            return True
        # Usa os primeiros 40 chars do texto entre colchetes como chave única —
        # tolera variações ortográficas tardias (ex: "busca" vs "busco")
        cn = normaliza(col_chave)
        ln = normaliza(lp_chave_col)
        if len(cn) >= 40 and len(ln) >= 40 and cn[:40] == ln[:40]:
            return True

    # 3. Questões SEM colchetes (ex: "1. Em que ano entrou"): prefixo longo único
    if lp_chave_col is None and col_chave is None:
        # Só usa prefixo quando o texto começa com um número de questão (único)
        num_m = re.match(r"^\d+[\.\s]", lp_n)
        if num_m and lp_n[:40] == col_n[:40]:
            return True

    return False


def seleciona_perguntas(abas, df_respostas):
    """
    Para cada pergunta listada na aba Lógica, encontra a coluna correspondente
    nas respostas. Retorna (lista_colunas, {coluna: pilar}).
    """
    df_logica = abas["logica"]

    # Coluna de pergunta na lógica
    col_perg_logica = None
    for col in df_logica.columns:
        if normaliza("pergunta") in normaliza(col):
            col_perg_logica = col
            break
    col_perg_logica = col_perg_logica or df_logica.columns[0]

    logica_pergs = df_logica[col_perg_logica].dropna().unique().tolist()

    # Filtro opcional da aba "Perguntas" (lista de questões a incluir)
    filtro_lps = None
    if "perguntas" in abas:
        df_p = abas["perguntas"].dropna(how="all")
        lista = df_p.iloc[:, 0].dropna().astype(str).tolist()
        if len(lista) > 2:
            filtro_lps = set(normaliza(p) for p in lista)

    # Pré-computa chave de colchetes para cada lp
    lp_chaves = {lp: extrai_colchetes(lp) for lp in logica_pergs}

    relevantes = []
    pilares    = {}
    sem_match  = []

    for lp in logica_pergs:
        # Aplica filtro da aba "Perguntas" (se existir)
        if filtro_lps is not None:
            lp_n = normaliza(lp)
            if lp_n not in filtro_lps:
                lp_chave = lp_chaves[lp]
                # Tenta via chave de colchetes
                if lp_chave is None or not any(
                        normaliza(lp_chave) in f for f in filtro_lps):
                    continue

        lp_chave = lp_chaves[lp]
        achou = None
        pilar = None

        for col in df_respostas.columns:
            if _casa_col_para_lp(col, lp, lp_chave):
                achou = col
                pilar = detecta_pilar(col)
                break

        if achou is None:
            sem_match.append(lp[:80])
        elif achou not in relevantes:
            relevantes.append(achou)
            if pilar:
                pilares[achou] = pilar

    # Para colunas sem pilar detectado pelo nome, tenta pelo padrão
    for col in relevantes:
        if col not in pilares:
            p = detecta_pilar(col)
            if p:
                pilares[col] = p

    if sem_match:
        print(f"Aviso: {len(sem_match)} pergunta(s) da Lógica não localizadas "
              f"nas respostas (ignoradas):")
        for s in sem_match[:5]:
            print(f"   - {s}")
        if len(sem_match) > 5:
            print(f"   ... e mais {len(sem_match)-5}")

    nomes_pilares = sorted(set(pilares.values()))
    print(f"{len(relevantes)} perguntas mapeadas. "
          f"Pilares: {nomes_pilares if nomes_pilares else '(nenhum)'}")
    return relevantes, pilares


# ---------------------------------------------------------------------------
# Pontuação
# ---------------------------------------------------------------------------

def _busca_pontos(perg_norm, opcao_norm, mapa):
    """Busca a pontuação no mapa com fallback e correspondência parcial."""
    for chave in [(perg_norm, opcao_norm), ("", opcao_norm)]:
        if chave in mapa:
            return mapa[chave]
    # correspondência parcial da opção
    for (p, o), v in mapa.items():
        if (p == perg_norm or p == "") and (o in opcao_norm or opcao_norm in o):
            return v
    # drop palavras curtas (artigos/preposições) e tenta novamente
    # ex: "Antes de 2015" vs "Antes 2015"
    def _sem_curtas(t):
        return re.sub(r"\s+", " ", re.sub(r"\b\w{1,2}\b", " ", t)).strip()
    opcao_curta = _sem_curtas(opcao_norm)
    if opcao_curta and opcao_curta != opcao_norm:
        for (p, o), v in mapa.items():
            o_curta = _sem_curtas(o)
            if (p == perg_norm or p == "") and o_curta and o_curta == opcao_curta:
                return v
    return None


def pontua(df_respostas, perguntas, mapa_logica, df_logica, col_perg_logica):
    """Converte cada resposta em pontos. Retorna DataFrame de pontos."""
    # Mapa: chave_colchetes_da_col_resp (norm) → perg_norm_na_logica
    chave_para_perg_norm = {}
    for lp in df_logica[col_perg_logica].dropna().unique():
        lp_n = normaliza(lp)
        chave_para_perg_norm[lp_n] = lp_n  # match direto
        c = extrai_colchetes(lp)
        if c:
            chave_para_perg_norm[normaliza(c)] = lp_n

    pontos = pd.DataFrame(index=df_respostas.index)
    sem_logica = set()

    for col in perguntas:
        col_n = normaliza(col)
        # Determina a chave de pergunta para buscar no mapa_logica
        perg_norm = col_n
        c = extrai_colchetes(col)
        if c and normaliza(c) in chave_para_perg_norm:
            perg_norm = chave_para_perg_norm[normaliza(c)]
        elif col_n in chave_para_perg_norm:
            perg_norm = chave_para_perg_norm[col_n]

        def _val(resposta, pn=perg_norm):
            if pd.isna(resposta):
                return float("nan")
            rn = normaliza(str(resposta))
            if not rn:
                return float("nan")
            # Caso especial: resposta direta do estágio (questão de autoposicionamento)
            est = estagio_para_numero(resposta)
            if est is not None and ("posic" in pn or "nivel de envolvimento" in pn
                                    or "imagem anterior" in pn):
                estagio_raw = float(est)
            else:
                estagio_raw = _busca_pontos(pn, rn, mapa_logica)
                if estagio_raw is None:
                    sem_logica.add(f"{col[:50]} → {str(resposta)[:40]}")
                    return float("nan")
            # Aplica escala exponencial: BASE_EXP ** (estágio - 1)
            return BASE_EXP ** (estagio_raw - 1)

        pontos[col] = df_respostas[col].map(_val)

    if sem_logica:
        amostras = sorted(sem_logica)[:8]
        print(f"\nAviso: {len(sem_logica)} tipo(s) de resposta sem correspondência na "
              f"Lógica (ficaram como NaN). Amostras:")
        for ex in amostras:
            print(f"   - {ex}")
    return pontos


# ---------------------------------------------------------------------------
# Pipeline principal
# ---------------------------------------------------------------------------

def detecta_segmento(df_respostas, colunas_pontuadas):
    """Localiza a coluna que indica o segmento (família / colaborador)."""
    for col in df_respostas.columns:
        if col in colunas_pontuadas:
            continue
        cn = normaliza(col)
        if any(k in cn for k in ["vinculo", "segmento", "perfil"]):
            vals = " ".join(normaliza(str(v))
                            for v in df_respostas[col].dropna().unique()[:20])
            if "famil" in vals or "colabo" in vals or "mae" in vals or "pai" in vals:
                return col
    for col in df_respostas.columns:
        if col in colunas_pontuadas:
            continue
        vals = " ".join(normaliza(str(v))
                        for v in df_respostas[col].dropna().unique()[:20])
        if ("famil" in vals or "mae" in vals or "pai" in vals) and "colabo" in vals:
            return col
    return None


def processa(caminho_entrada, caminho_saida, cortes=CORTES):
    abas = carrega_abas(caminho_entrada)
    df_resp = abas["respostas"]

    mapa_logica, usa_estagio, col_perg_nome = monta_logica(abas["logica"])
    perguntas, pilares = seleciona_perguntas(abas, df_resp)

    if not perguntas:
        sys.exit("ERRO: nenhuma pergunta foi mapeada. Verifique a estrutura do arquivo.")

    pontos = pontua(df_resp, perguntas, mapa_logica, abas["logica"],
                    col_perg_nome or abas["logica"].columns[0])

    # Calcula cortes pelo range teórico real da pesquisa.
    # pontos já contém valores exponenciais: BASE_EXP^(estágio-1).
    # Min teórico = BASE_EXP^(1-1) = 1.0 para toda pergunta.
    # Max teórico por pergunta = máximo observado em pontos (proxy conservador;
    # com 143 respondentes é provável que toda pergunta tenha ao menos um respondente
    # no nível máximo).
    score_min_t = 1.0  # BASE_EXP^0
    score_max_t = float(pontos.max().mean())
    r = score_max_t - score_min_t
    cortes = [score_min_t + r / 7 * i for i in range(1, 7)]
    print(f"Escala exponencial base {BASE_EXP}: "
          f"range teórico [{score_min_t:.3f}, {score_max_t:.3f}]")
    print(f"Cortes: {[round(c, 3) for c in cortes]}")

    # ----- Scores individuais -----
    resultado = df_resp.copy()
    col_segmento = detecta_segmento(df_resp, set(perguntas))
    if col_segmento:
        print(f"Coluna de segmento: '{col_segmento}'")

    nomes_pilares = sorted(set(pilares.values()))
    for pilar in nomes_pilares:
        cols = [c for c, p in pilares.items() if p == pilar]
        resultado[f"Score – {pilar}"] = pontos[cols].mean(axis=1).round(2)

    resultado["Score Final"]          = pontos.mean(axis=1).round(2)
    resultado["Estágio na Curva"]     = resultado["Score Final"].map(
        lambda s: numero_para_estagio(s, cortes))
    resultado["Nº perguntas pontuadas"] = pontos.notna().sum(axis=1)
    resultado["Total de perguntas"]   = len(perguntas)

    # ----- Resumos -----
    resumo_estagio = (
        resultado["Estágio na Curva"]
        .value_counts()
        .reindex(ESTAGIOS).fillna(0).astype(int)
        .rename("Respondentes").to_frame()
    )
    resumo_estagio["%"] = (100 * resumo_estagio["Respondentes"]
                           / max(len(resultado), 1)).round(1)
    resumo_estagio.index.name = "Estágio"
    resumos = {"Resumo por Estágio": resumo_estagio.reset_index()}

    if nomes_pilares:
        resumo_pilar = pd.DataFrame({
            "Pilar":           nomes_pilares,
            "Nº perguntas":    [len([c for c, p in pilares.items() if p == pl])
                                for pl in nomes_pilares],
            "Score Médio":     [resultado[f"Score – {pl}"].mean().round(2)
                                for pl in nomes_pilares],
        })
        resumo_pilar["Estágio Médio"] = resumo_pilar["Score Médio"].map(
            lambda s: numero_para_estagio(s, cortes))
        resumos["Resumo por Pilar"] = resumo_pilar

    if col_segmento:
        grp = resultado.groupby(col_segmento)
        seg = grp["Score Final"].agg(["count", "mean"]).round(2).reset_index()
        seg.columns = ["Segmento", "Respondentes", "Score Médio"]
        seg["Estágio Médio"] = seg["Score Médio"].map(
            lambda s: numero_para_estagio(s, cortes))
        if nomes_pilares:
            for pl in nomes_pilares:
                seg[f"Score – {pl}"] = grp[f"Score – {pl}"].mean().round(2).values
        resumos["Resumo por Segmento"] = seg

    # Tabela de cortes usados
    limites = [score_min_t] + list(cortes) + [score_max_t]
    resumos["Cortes Utilizados"] = pd.DataFrame([
        {"Estágio": est,
         "De (score >)":   "mínimo" if i == 0 else round(limites[i], 2),
         "Até (score <=)": "máximo" if i == 6 else round(limites[i + 1], 2)}
        for i, est in enumerate(ESTAGIOS)
    ])

    # ----- Grava .xlsx -----
    with pd.ExcelWriter(caminho_saida, engine="openpyxl") as w:
        resultado.to_excel(w,  sheet_name="Scores Individuais",    index=False)
        pontos.to_excel(w,     sheet_name="Pontos por Pergunta",   index=False)
        for nome, df in resumos.items():
            df.to_excel(w, sheet_name=nome[:31], index=False)

    print(f"\nArquivo gerado: {caminho_saida}")
    print("\n=== Resumo por Estágio ===")
    print(resumo_estagio.to_string())
    if "Resumo por Pilar" in resumos:
        print("\n=== Resumo por Pilar (trimembração social) ===")
        print(resumos["Resumo por Pilar"].to_string(index=False))
    if "Resumo por Segmento" in resumos:
        print("\n=== Resumo por Segmento ===")
        print(resumos["Resumo por Segmento"]
              [["Segmento", "Respondentes", "Score Médio", "Estágio Médio"]]
              .to_string(index=False))
    return resultado


def main():
    parser = argparse.ArgumentParser(
        description="Processa a pesquisa e gera scores na curva dos 7 Processos Sociais.")
    parser.add_argument("entrada", help="Arquivo .ods ou .xlsx da pesquisa")
    parser.add_argument("-o", "--saida", default="pesquisa_processada.xlsx",
                        help="Arquivo de saída (padrão: pesquisa_processada.xlsx)")
    args = parser.parse_args()
    processa(args.entrada, args.saida)


if __name__ == "__main__":
    main()
