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
------------------------
Em vez de usar valores arbitrários, a recomendação é ancorar a pontuação
diretamente na curva: cada opção de resposta recebe (na aba "Lógica") o
estágio da curva correspondente, e esse estágio é convertido em um número
de 1 a 7. O score final do respondente é a média desses números — assim o
score já "vive" na mesma escala da curva e os intervalos de corte ficam
naturais (ex.: score 4,2 ≈ estágio 4, "Comprometido").

Se a aba "Lógica" tiver uma coluna de estágio, ela é usada como fonte da
pontuação (1–7). Caso contrário, usa-se a coluna de valor numérico da
própria aba, e os cortes são calculados proporcionalmente.

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
# Configuração — ajuste aqui se os nomes no seu arquivo forem diferentes
# ---------------------------------------------------------------------------

# Estágios da curva dos 7 Processos Sociais, em ordem evolutiva.
ESTAGIOS = [
    "Novo",            # respiração
    "Adaptando",       # aquecimento
    "Inserindo",       # digestão
    "Comprometido",    # segregação
    "Ativo no Motivo", # manutenção
    "Embaixador",      # crescimento
    "Cocriador",       # geração
]

# Intervalos de corte do score (escala 1–7) para cada estágio.
# O score S recebe o estágio i se S <= CORTES[i]. O padrão arredonda para o
# estágio mais próximo; ajuste livremente (ex.: exigir mais para "Cocriador").
CORTES = [1.5, 2.5, 3.5, 4.5, 5.5, 6.5]  # acima do último corte => Cocriador

# Palavras-chave para localizar as abas (a busca ignora acentos/maiúsculas).
ABAS = {
    "respostas": ["resposta formulario", "respostas"],
    "perguntas": ["pergunta"],
    "opcoes":    ["opcoes de reposta", "opcoes de resposta", "opcao"],
    "logica":    ["logica"],
}

# Palavras-chave para localizar colunas na aba "Lógica".
COLUNAS_LOGICA = {
    "pergunta": ["pergunta", "questao", "afirmacao", "afirmativa"],
    "opcao":    ["opcao", "resposta", "categoria"],
    "valor":    ["valor", "pontuacao", "pontos", "peso", "nota"],
    "estagio":  ["estagio", "processo", "curva", "posicao"],
    "nivel":    ["nivel", "qualificacao"],
}

# Palavras-chave para localizar colunas na aba "Perguntas".
COLUNAS_PERGUNTAS = {
    "pergunta": ["pergunta", "questao", "afirmacao", "afirmativa"],
    "pilar":    ["pilar", "dimensao", "trimembracao", "eixo", "tema"],
}

# Palavras-chave para identificar a coluna de segmento (família/colaborador)
# na aba de respostas.
SEGMENTO_KEYWORDS = ["famil", "colabor"]


# ---------------------------------------------------------------------------
# Utilitários
# ---------------------------------------------------------------------------

def normaliza(texto):
    """Remove acentos, pontuação e espaços repetidos; deixa em minúsculas."""
    if texto is None or (isinstance(texto, float) and pd.isna(texto)):
        return ""
    texto = str(texto)
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"[^\w\s]", " ", texto.lower())
    return re.sub(r"\s+", " ", texto).strip()


def acha_aba(nomes_abas, chaves):
    """Encontra a aba cujo nome contém alguma das palavras-chave."""
    for chave in chaves:
        chave_norm = normaliza(chave)
        for nome in nomes_abas:
            if chave_norm in normaliza(nome) or normaliza(nome) in chave_norm:
                return nome
    # fallback: interseção de palavras
    for chave in chaves:
        palavras = set(normaliza(chave).split())
        for nome in nomes_abas:
            if palavras & set(normaliza(nome).split()):
                return nome
    return None


def acha_coluna(df, chaves, obrigatoria=False, nome=""):
    """Encontra a coluna do DataFrame cujo nome contém alguma palavra-chave."""
    for chave in chaves:
        for col in df.columns:
            if normaliza(chave) in normaliza(col):
                return col
    if obrigatoria:
        sys.exit(
            f"ERRO: não encontrei a coluna '{nome}' (procurei por {chaves}). "
            f"Colunas disponíveis: {list(df.columns)}"
        )
    return None


def estagio_para_numero(texto):
    """Converte o nome de um estágio da curva no número 1–7 (ou None)."""
    t = normaliza(texto)
    if not t:
        return None
    # aceita "4", "4 - Comprometido", "Comprometido (segregação)" etc.
    m = re.match(r"^(\d+)", t)
    if m and 1 <= int(m.group(1)) <= 7:
        return int(m.group(1))
    sinonimos = {
        "novo": 1, "respiracao": 1,
        "adaptando": 2, "aquecimento": 2,
        "inserindo": 3, "digestao": 3,
        "comprometido": 4, "segregacao": 4,
        "ativo no motivo": 5, "manutencao": 5, "ativo": 5,
        "embaixador": 6, "crescimento": 6,
        "cocriador": 7, "geracao": 7, "co criador": 7,
    }
    for chave, num in sinonimos.items():
        if chave in t:
            return num
    return None


def numero_para_estagio(score, cortes=CORTES):
    """Aplica os intervalos de corte e devolve o nome do estágio."""
    if pd.isna(score):
        return None
    for i, corte in enumerate(cortes):
        if score <= corte:
            return ESTAGIOS[i]
    return ESTAGIOS[-1]


# ---------------------------------------------------------------------------
# Pipeline
# ---------------------------------------------------------------------------

def carrega_abas(caminho):
    """Lê todas as abas do arquivo e identifica cada uma pelo nome."""
    engine = "odf" if caminho.lower().endswith(".ods") else None
    todas = pd.read_excel(caminho, sheet_name=None, engine=engine)
    nomes = list(todas.keys())
    print(f"Abas encontradas: {nomes}")

    abas = {}
    for papel, chaves in ABAS.items():
        nome = acha_aba(nomes, chaves)
        if nome is None and papel in ("respostas", "logica"):
            sys.exit(f"ERRO: não encontrei a aba de '{papel}' (procurei por {chaves}).")
        if nome is not None:
            abas[papel] = todas[nome].dropna(how="all").dropna(how="all", axis=1)
            print(f"  -> aba de {papel}: '{nome}' ({len(abas[papel])} linhas)")
    return abas


def monta_logica(df_logica):
    """
    Constrói o dicionário de pontuação a partir da aba "Lógica".

    Retorna (mapa, usa_estagio):
      - mapa: {(pergunta_norm, opcao_norm): pontuacao} e também
              {("", opcao_norm): pontuacao} como fallback global.
      - usa_estagio: True se a pontuação veio da coluna de estágio (escala 1–7).
    """
    col_pergunta = acha_coluna(df_logica, COLUNAS_LOGICA["pergunta"])
    col_opcao = acha_coluna(df_logica, COLUNAS_LOGICA["opcao"], obrigatoria=True,
                            nome="opção de resposta (aba Lógica)")
    col_valor = acha_coluna(df_logica, COLUNAS_LOGICA["valor"])
    col_estagio = acha_coluna(df_logica, COLUNAS_LOGICA["estagio"])

    usa_estagio = False
    if col_estagio is not None:
        estagios_num = df_logica[col_estagio].map(estagio_para_numero)
        if estagios_num.notna().mean() > 0.5:  # a coluna é mesmo de estágios
            usa_estagio = True
            df_logica = df_logica.assign(_pontos=estagios_num)
            print(f"Pontuação ancorada na coluna de estágio '{col_estagio}' (escala 1–7).")
    if not usa_estagio:
        if col_valor is None:
            sys.exit("ERRO: a aba Lógica não tem coluna de estágio nem de valor utilizável.")
        df_logica = df_logica.assign(_pontos=pd.to_numeric(df_logica[col_valor], errors="coerce"))
        print(f"Pontuação baseada na coluna de valor '{col_valor}'.")

    mapa = {}
    for _, linha in df_logica.iterrows():
        pontos = linha["_pontos"]
        if pd.isna(pontos):
            continue
        opcao = normaliza(linha[col_opcao])
        if not opcao:
            continue
        if col_pergunta is not None:
            mapa[(normaliza(linha[col_pergunta]), opcao)] = float(pontos)
        # fallback global (vale para qualquer pergunta com essa opção)
        mapa.setdefault(("", opcao), float(pontos))

    print(f"Lógica carregada: {len(mapa)} combinações pergunta/opção pontuadas.")
    return mapa, usa_estagio


def seleciona_perguntas(abas, df_respostas):
    """
    Usa a aba "Perguntas" para filtrar as colunas relevantes das respostas e,
    se houver, associar cada pergunta a um pilar da trimembração.
    Retorna (lista de colunas relevantes, {coluna: pilar}).
    """
    colunas_resp = [c for c in df_respostas.columns]
    if "perguntas" not in abas:
        print("Aviso: aba 'Perguntas' não encontrada — usando todas as colunas de resposta.")
        return colunas_resp, {}

    df_perg = abas["perguntas"]
    col_texto = acha_coluna(df_perg, COLUNAS_PERGUNTAS["pergunta"])
    if col_texto is None:
        col_texto = df_perg.columns[0]
    col_pilar = acha_coluna(df_perg, COLUNAS_PERGUNTAS["pilar"])

    relevantes, pilares = [], {}
    nao_encontradas = []
    for _, linha in df_perg.iterrows():
        alvo = normaliza(linha[col_texto])
        if not alvo:
            continue
        achou = None
        for col in colunas_resp:
            cn = normaliza(col)
            if alvo == cn or alvo in cn or cn in alvo:
                achou = col
                break
        if achou is None:
            nao_encontradas.append(str(linha[col_texto])[:60])
            continue
        if achou not in relevantes:
            relevantes.append(achou)
        if col_pilar is not None and pd.notna(linha[col_pilar]):
            pilares[achou] = str(linha[col_pilar]).strip()

    if nao_encontradas:
        print(f"Aviso: {len(nao_encontradas)} pergunta(s) da aba 'Perguntas' não "
              f"localizadas nas respostas: {nao_encontradas}")
    print(f"{len(relevantes)} perguntas relevantes selecionadas"
          + (f", com pilares: {sorted(set(pilares.values()))}" if pilares else "."))
    return relevantes, pilares


def detecta_segmento(df_respostas, colunas_pontuadas):
    """Procura a coluna que identifica o segmento (família/colaborador)."""
    for col in df_respostas.columns:
        if col in colunas_pontuadas:
            continue
        valores = " ".join(normaliza(v) for v in df_respostas[col].dropna().unique()[:50])
        if all(kw in valores for kw in SEGMENTO_KEYWORDS):
            return col
    for col in df_respostas.columns:
        if any(k in normaliza(col) for k in ["vinculo", "segmento", "perfil", "voce e", "relacao com a escola"]):
            return col
    return None


def pontua(df_respostas, perguntas, mapa_logica):
    """Converte cada resposta em pontos. Retorna DataFrame de pontos."""
    pontos = pd.DataFrame(index=df_respostas.index)
    sem_logica = set()
    for col in perguntas:
        cn = normaliza(col)

        def valor(resposta, cn=cn):
            rn = normaliza(resposta)
            if not rn:
                return float("nan")
            # tenta casar (pergunta, opção); depois só a opção; depois por inclusão
            for chave in [(cn, rn), ("", rn)]:
                if chave in mapa_logica:
                    return mapa_logica[chave]
            for (p, o), v in mapa_logica.items():
                if (p == cn or p == "") and (o in rn or rn in o):
                    return v
            sem_logica.add(f"{col[:40]}... => {str(resposta)[:40]}")
            return float("nan")

        pontos[col] = df_respostas[col].map(valor)

    if sem_logica:
        print(f"Aviso: {len(sem_logica)} resposta(s) sem correspondência na Lógica "
              "(ficaram em branco). Exemplos:")
        for ex in sorted(sem_logica)[:10]:
            print(f"   - {ex}")
    return pontos


def processa(caminho_entrada, caminho_saida, cortes=CORTES):
    abas = carrega_abas(caminho_entrada)
    df_resp = abas["respostas"]
    mapa_logica, usa_estagio = monta_logica(abas["logica"])
    perguntas, pilares = seleciona_perguntas(abas, df_resp)

    pontos = pontua(df_resp, perguntas, mapa_logica)

    # Se a pontuação não veio da escala 1–7, reescala os cortes proporcionalmente.
    if not usa_estagio:
        vmin, vmax = pontos.min().min(), pontos.max().max()
        if pd.notna(vmin) and vmax > vmin:
            cortes = [vmin + (c - 1) / 6 * (vmax - vmin) for c in cortes]
            print(f"Cortes reescalados para a faixa de valores [{vmin:.2f}, {vmax:.2f}]: "
                  f"{[round(c, 2) for c in cortes]}")

    # --- Resultado individual ---------------------------------------------
    resultado = df_resp.copy()
    col_segmento = detecta_segmento(df_resp, set(perguntas))
    if col_segmento:
        print(f"Coluna de segmento detectada: '{col_segmento}'")

    # score por pilar da trimembração (se a aba Perguntas trouxe os pilares)
    nomes_pilares = sorted(set(pilares.values()))
    for pilar in nomes_pilares:
        cols = [c for c, p in pilares.items() if p == pilar]
        resultado[f"Score - {pilar}"] = pontos[cols].mean(axis=1).round(2)

    resultado["Score Final"] = pontos.mean(axis=1).round(2)
    resultado["Estágio na Curva"] = resultado["Score Final"].map(
        lambda s: numero_para_estagio(s, cortes))
    resultado["Nº de respostas pontuadas"] = pontos.notna().sum(axis=1)

    # --- Resumos ------------------------------------------------------------
    resumo_estagio = (
        resultado["Estágio na Curva"]
        .value_counts()
        .reindex(ESTAGIOS)
        .fillna(0).astype(int)
        .rename("Respondentes")
        .to_frame()
    )
    resumo_estagio["%"] = (100 * resumo_estagio["Respondentes"]
                           / max(len(resultado), 1)).round(1)
    resumo_estagio.index.name = "Estágio"

    resumos = {"Resumo por Estágio": resumo_estagio.reset_index()}

    if nomes_pilares:
        resumo_pilar = pd.DataFrame({
            "Pilar": nomes_pilares,
            "Score Médio": [resultado[f"Score - {p}"].mean().round(2) for p in nomes_pilares],
        })
        resumo_pilar["Estágio Médio"] = resumo_pilar["Score Médio"].map(
            lambda s: numero_para_estagio(s, cortes))
        resumos["Resumo por Pilar"] = resumo_pilar

    if col_segmento:
        grp = resultado.groupby(col_segmento)
        resumo_seg = grp["Score Final"].agg(["count", "mean"]).round(2).reset_index()
        resumo_seg.columns = [col_segmento, "Respondentes", "Score Médio"]
        resumo_seg["Estágio Médio"] = resumo_seg["Score Médio"].map(
            lambda s: numero_para_estagio(s, cortes))
        resumos["Resumo por Segmento"] = resumo_seg

    # tabela de referência dos cortes usados
    faixas = []
    limites = [pontos.min().min() if not usa_estagio else 1.0] + list(cortes) + \
              [pontos.max().max() if not usa_estagio else 7.0]
    for i, est in enumerate(ESTAGIOS):
        faixas.append({"Estágio": est,
                       "De (score >)": round(limites[i], 2) if i else "mínimo",
                       "Até (score <=)": round(limites[i + 1], 2) if i < 6 else "máximo"})
    resumos["Cortes Utilizados"] = pd.DataFrame(faixas)

    # --- Grava o arquivo de saída -------------------------------------------
    with pd.ExcelWriter(caminho_saida, engine="openpyxl") as escritor:
        resultado.to_excel(escritor, sheet_name="Scores Individuais", index=False)
        pontos.to_excel(escritor, sheet_name="Pontos por Pergunta", index=False)
        for nome, df in resumos.items():
            df.to_excel(escritor, sheet_name=nome[:31], index=False)

    print(f"\nArquivo processado gravado em: {caminho_saida}")
    print("\n=== Resumo por Estágio ===")
    print(resumo_estagio.to_string())
    if "Resumo por Pilar" in resumos:
        print("\n=== Resumo por Pilar ===")
        print(resumos["Resumo por Pilar"].to_string(index=False))
    if "Resumo por Segmento" in resumos:
        print("\n=== Resumo por Segmento ===")
        print(resumos["Resumo por Segmento"].to_string(index=False))
    return resultado


def main():
    parser = argparse.ArgumentParser(
        description="Processa a pesquisa e gera scores na curva dos 7 Processos Sociais.")
    parser.add_argument("entrada", help="Arquivo .ods ou .xlsx da pesquisa")
    parser.add_argument("-o", "--saida", default="pesquisa_processada.xlsx",
                        help="Arquivo .xlsx de saída (padrão: pesquisa_processada.xlsx)")
    args = parser.parse_args()
    processa(args.entrada, args.saida)


if __name__ == "__main__":
    main()
