#!/usr/bin/env python3
"""
Gera documento Word com a metodologia de inferência da classificação individual.
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

SAIDA = "/home/user/arandu/Metodologia_Classificacao_Individual.docx"

# ─── Helpers ──────────────────────────────────────────────────────────────────

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def set_cell_bold(cell, bold=True):
    for para in cell.paragraphs:
        for run in para.runs:
            run.bold = bold

def add_heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    return h

def add_paragraph(doc, text, bold_parts=None):
    """Adiciona parágrafo com texto. bold_parts = [(start, end)] (índices)."""
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    if bold_parts is None:
        run = p.add_run(text)
        run.font.size = Pt(11)
    else:
        run = p.add_run(text)
        run.font.size = Pt(11)
    return p

def add_body(doc, text):
    p = doc.add_paragraph(style="Normal")
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.space_before = Pt(2)
    run = p.add_run(text)
    run.font.size = Pt(11)
    return p

def add_bullet(doc, text, bold_prefix=None):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    if bold_prefix:
        r1 = p.add_run(bold_prefix)
        r1.bold = True
        r1.font.size = Pt(11)
        r2 = p.add_run(text)
        r2.font.size = Pt(11)
    else:
        r = p.add_run(text)
        r.font.size = Pt(11)
    return p

def add_note(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(1)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0x60, 0x60, 0x60)
    run.font.italic = True
    return p

def table_simple(doc, headers, rows, header_color="1F3864", col_widths=None):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.LEFT

    # Header row
    hrow = t.rows[0]
    for i, h in enumerate(headers):
        cell = hrow.cells[i]
        cell.text = h
        set_cell_bg(cell, header_color)
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                run.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.size = Pt(10)

    # Data rows
    alt_colors = ["FFFFFF", "EBF3FB"]
    for ri, row_data in enumerate(rows):
        row = t.rows[ri + 1]
        for ci, val in enumerate(row_data):
            cell = row.cells[ci]
            cell.text = str(val)
            set_cell_bg(cell, alt_colors[ri % 2])
            for para in cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for run in para.runs:
                    run.font.size = Pt(10)

    if col_widths:
        for i, w in enumerate(col_widths):
            for row in t.rows:
                row.cells[i].width = Cm(w)

    doc.add_paragraph()
    return t


# ─── Documento ────────────────────────────────────────────────────────────────

def main():
    doc = Document()

    # Configuração da página
    section = doc.sections[0]
    section.page_width  = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.5)
    section.top_margin    = Cm(2.5)
    section.bottom_margin = Cm(2.5)

    # ── Capa / Título ─────────────────────────────────────────────────────────
    title = doc.add_heading("Metodologia de Classificação Individual", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = subtitle.add_run("Inferência de Posição na Curva dos 7 Processos Sociais\nEscola Arandu · Pesquisa de Maturidade Comunitária · 2025")
    r.font.size = Pt(12)
    r.font.italic = True
    r.font.color.rgb = RGBColor(0x40, 0x40, 0x40)
    doc.add_paragraph()

    # ── 1. Objetivo ───────────────────────────────────────────────────────────
    add_heading(doc, "1. Objetivo", 1)
    add_body(doc,
        "Este documento descreve a metodologia utilizada para inferir, a partir das respostas "
        "individuais ao formulário de pesquisa, a posição de cada membro da comunidade escolar "
        "Arandu na Curva dos 7 Processos Sociais. A classificação é distinta do autoposicionamento "
        "declarado pelo respondente na pergunta Q10 e busca refletir o estágio real com base em "
        "evidências concretas: nível de conhecimento, padrões de concordância e discordância, "
        "presença de respostas de incerteza ('Não sei opinar') e engajamento ativo na escola."
    )
    add_body(doc,
        "A inferência também considera o tempo de vínculo com a escola como variável estrutural: "
        "membros recentes têm um teto máximo de estágio inferido, independentemente do nível de "
        "concordância, pois a curva pressupõe tempo de vivência antes da decisão de compromisso. "
        "Veteranos com muitos anos de escola e baixo comprometimento recebem tratamento específico."
    )

    # ── 2. A Curva dos 7 Processos Sociais ───────────────────────────────────
    add_heading(doc, "2. A Curva dos 7 Processos Sociais", 1)
    add_body(doc,
        "A curva é fundamentada na antroposofia e descreve o processo de maturidade de um membro "
        "dentro de uma comunidade de propósito. Ela possui uma estrutura em lemniscata (em forma de "
        "oito): um lado esquerdo de Capital Social Individual — onde o membro ainda se relaciona "
        "com a escola principalmente pelos filhos — e um lado direito de Capital Social Comunitário "
        "— onde o membro passou a se identificar com a escola como um todo."
    )
    add_body(doc,
        "O ponto de virada entre os dois lados é o estágio Comprometido: o momento em que a pessoa "
        "diz SIM para a escola e decide fazer parte. Antes desse ponto, o membro pode sair sem que "
        "isso seja visto como ruptura. Após essa virada, sair representa uma quebra de compromisso."
    )

    table_simple(doc,
        ["#", "Estágio", "Processo Social", "Característica central"],
        [
            ["1", "Novo",           "Respiração",  "Está chegando, reconhecendo a escola — ainda de fora"],
            ["2", "Adaptando",      "Aquecimento", "Entendendo como a escola funciona — assimilando"],
            ["3", "Inserindo",      "Digestão",    "Tem dúvidas e questionamentos — processando, decidindo"],
            ["4", "Comprometido",   "Segregação",  "Disse SIM para a escola — quer fazer parte → VIRADA"],
            ["5", "Ativo no Motivo","Manutenção",  "Tem papel de atuação dentro da escola"],
            ["6", "Embaixador",     "Crescimento", "É reconhecido como referência e representante"],
            ["7", "Cocriador",      "Geração",     "Atua com a escola para construir o que precisa surgir"],
        ],
        col_widths=[0.8, 3.2, 3.2, 7.5]
    )

    add_note(doc,
        "Nota: os estágios 1–3 compõem o lado esquerdo da curva (Capital Social Individual). "
        "Os estágios 4–7 compõem o lado direito (Capital Social Comunitário). "
        "O estágio 4 (Comprometido) é simultaneamente o ponto de chegada do lado esquerdo "
        "e o ponto de partida do lado direito."
    )

    # ── 3. Dados utilizados ───────────────────────────────────────────────────
    add_heading(doc, "3. Dados Utilizados na Inferência", 1)
    add_body(doc,
        "A inferência utiliza cinco conjuntos de dados extraídos das respostas individuais ao "
        "formulário, organizados por esfera temática:"
    )

    table_simple(doc,
        ["Esfera", "Nº de perguntas", "Tipo de resposta", "O que mede"],
        [
            ["Conhecimento institucional", "4",
             "Sim / Não",
             "Se o membro já conhecia propósito, missão, valores e imagem de ser humano da escola antes de entrar"],
            ["Pertencimento cultural", "8",
             "Escala de concordância",
             "Alinhamento com propósito, valores, confiança e atitudes de pertencimento à comunidade"],
            ["Acordos sociais", "9",
             "Escala de concordância",
             "Clareza e cumprimento dos acordos, relações, postura construtiva em crises"],
            ["Prática da Economia Fraterna", "8",
             "Escala de concordância",
             "Entendimento e prática da economia fraterna, participação ativa e confiança no processo"],
            ["Engajamento", "5",
             "Sou atualmente / Já fui / Nunca fui",
             "Participação efetiva em papéis formais: representante de turma, voluntário, grupos de trabalho, APG, mutirões"],
        ],
        col_widths=[3.8, 2.2, 3.2, 5.5]
    )

    add_body(doc,
        "Além dos dados de respostas, utiliza-se o ano de entrada na escola (pergunta P1) "
        "para calcular o tempo de vínculo, que modifica o teto e o piso do estágio inferido "
        "(ver Seção 6)."
    )

    # ── 4. Net Signal ─────────────────────────────────────────────────────────
    add_heading(doc, "4. Net Signal: Medida de Força do Sinal por Esfera", 1)
    add_body(doc,
        "Para as três esferas com escala de concordância (Pertencimento, Acordos e Economia Fraterna), "
        "calcula-se um indicador chamado Net Signal que resume a força do sinal de cada pessoa "
        "em uma escala de 0 a 1. Ele não é uma simples média — incorpora explicitamente os efeitos "
        "negativos da discordância e da incerteza."
    )

    add_heading(doc, "4.1 Pesos por tipo de resposta", 2)
    add_body(doc,
        "Cada resposta recebe um peso que reflete sua contribuição para o comprometimento com a escola:"
    )
    table_simple(doc,
        ["Resposta", "Peso", "Justificativa"],
        [
            ["Concordo",               "+1,0", "Sinal positivo forte — pleno alinhamento"],
            ["Concordo parcialmente",  "+0,4", "Sinal positivo moderado — há reservas"],
            ["Não sei opinar",         "−0,2", "Sinal negativo leve — não é neutro: indica desconexão com o tema"],
            ["Discordo",               "−0,8", "Sinal negativo forte — resistência ou desalinhamento"],
        ],
        col_widths=[4.5, 2.0, 8.2]
    )
    add_note(doc,
        "Nota: 'Não sei opinar' recebe peso negativo porque, na Lógica da pesquisa, essa resposta "
        "é mapeada para o estágio Novo — o mais inicial. Uma pessoa que não sabe opinar sobre "
        "os valores ou acordos da escola ainda não internalizou os fundamentos do vínculo comunitário."
    )

    add_heading(doc, "4.2 Cálculo do Net Signal", 2)
    add_body(doc,
        "O Net Signal é calculado em duas etapas:"
    )
    add_bullet(doc,
        "Calcula-se a média ponderada das respostas: "
        "raw = (% Concordo × 1,0) + (% Concordo parc. × 0,4) + (% Não sei × −0,2) + (% Discordo × −0,8)",
        bold_prefix="Passo 1 — Valor bruto: "
    )
    add_bullet(doc,
        "O valor bruto varia de −0,8 (todos Discordo) a +1,0 (todos Concordo). "
        "Para comparabilidade, normaliza-se para a escala [0, 1]: "
        "Net Signal = (raw + 0,8) / 1,8",
        bold_prefix="Passo 2 — Normalização: "
    )

    add_heading(doc, "4.3 Exemplos de Net Signal", 2)
    table_simple(doc,
        ["Situação", "Raw", "Net Signal"],
        [
            ["Respondeu 'Concordo' em todas as questões da esfera",               "+1,00", "1,00"],
            ["Respondeu 'Concordo parcialmente' em todas",                         "+0,40", "0,67"],
            ["Metade 'Concordo', metade 'Não sei opinar'",                         "+0,40", "0,67"],
            ["Metade 'Concordo', metade 'Discordo'",                               "+0,10", "0,50"],
            ["Respondeu 'Não sei opinar' em todas",                                "−0,20", "0,33"],
            ["Metade 'Concordo parcialmente', metade 'Discordo'",                  "−0,20", "0,33"],
            ["Respondeu 'Discordo' em todas",                                      "−0,80", "0,00"],
        ],
        col_widths=[9.5, 2.0, 3.2]
    )

    # ── 5. Lógica de Inferência ───────────────────────────────────────────────
    add_heading(doc, "5. Lógica de Inferência do Estágio", 1)
    add_body(doc,
        "A inferência segue uma árvore de decisão que replica a estrutura da curva lemniscata. "
        "O ponto central é determinar se a pessoa já fez a virada — se 'disse SIM' para a escola. "
        "Dependendo dessa decisão, aplica-se a lógica do lado esquerdo ou do lado direito da curva."
    )

    add_heading(doc, "5.1 Threshold 'Disse SIM'", 2)
    add_body(doc,
        "A virada para o lado direito (Comprometido ou acima) é inferida quando duas condições "
        "são simultaneamente atendidas:"
    )
    add_bullet(doc,
        "Net Signal de Pertencimento ≥ 0,60 — indica alinhamento suficiente com propósito, "
        "valores e confiança na escola",
        bold_prefix="Condição 1: "
    )
    add_bullet(doc,
        "Discordância em Pertencimento < 25% — ausência de resistência ativa ao vínculo",
        bold_prefix="Condição 2: "
    )
    add_note(doc,
        "Justificativa: o Pertencimento é a esfera que mais diretamente expressa a decisão "
        "de pertencer — ela inclui questões como 'Eu escolho estar na Arandu porque compartilho "
        "do seu propósito' e 'Eu confio na escola Arandu'. Um Net Signal ≥ 0,60 equivale a "
        "respostas majoritariamente positivas nessa esfera, o que configura o 'dizer SIM' que "
        "define o estágio Comprometido na curva."
    )

    add_heading(doc, "5.2 Lado Esquerdo — Não disse SIM", 2)
    add_body(doc,
        "Se a pessoa não atingiu o threshold de virada, aplica-se a seguinte sequência:"
    )
    table_simple(doc,
        ["Condição (verificada nesta ordem)", "Estágio inferido", "Interpretação"],
        [
            ["Discordância em Pertencimento ≥ 40%",
             "Novo",
             "Resistência ativa — não há vínculo afetivo com a escola"],
            ["'Não sei opinar' em Pertencimento ≥ 50%",
             "Novo",
             "Não reconhece a escola ainda — não internalizou os fundamentos"],
            ["Conhecimento ≥ 75% Sim E positivo Pertencimento ≥ 50%",
             "Inserindo",
             "Conhece a escola, há parcial identificação — no processo de absorção"],
            ["Conhecimento ≥ 50% OU positivo Pertencimento ≥ 35%",
             "Adaptando",
             "Está assimilando — algum nível de identificação com a escola"],
            ["Nenhuma das condições acima",
             "Novo",
             "Ainda em reconhecimento inicial"],
        ],
        col_widths=[6.5, 3.0, 5.2]
    )

    add_heading(doc, "5.3 Lado Direito — Disse SIM", 2)
    add_body(doc,
        "Confirmado o threshold de virada, verificam-se tensões e o nível de engajamento ativo:"
    )
    table_simple(doc,
        ["Condição (verificada nesta ordem)", "Estágio inferido", "Interpretação"],
        [
            ["Discord. Acordos ≥ 25% E Discord. Ec. Fraterna ≥ 25%",
             "Comprometido",
             "Diz SIM, mas há tensões práticas nas duas esferas de atuação"],
            ["Discordância em Acordos Sociais ≥ 25%",
             "Comprometido",
             "Tensão nos acordos — dificuldade nas relações comunitárias"],
            ["Discordância em Ec. Fraterna ≥ 25%",
             "Comprometido",
             "Tensão econômica — resistência à prática da economia fraterna"],
            ["'Não sei' em Acordos ≥ 40% E 'Não sei' em Ec. Fraterna ≥ 40%",
             "Comprometido",
             "Diz SIM, mas ainda não conhece as práticas — em absorção"],
            ["Sem papel ativo no momento (nenhum 'Sou atualmente')",
             "Comprometido",
             "Disse SIM mas ainda não assumiu responsabilidade formal"],
            ["3+ papéis ativos E overall ≥ 75% E Ec. Fraterna ≥ 60%",
             "Cocriador",
             "Alta dedicação, múltiplos papéis e forte prática econômica"],
            ["2+ papéis ativos E overall ≥ 70%",
             "Embaixador",
             "Múltiplos papéis e alta concordância — referência comunitária"],
            ["1+ papel ativo E overall ≥ 60%",
             "Ativo no Motivo",
             "Engajado com pelo menos um papel formal na escola"],
            ["Papel ativo E overall < 60%",
             "Comprometido",
             "Participa, mas concordância geral ainda moderada"],
        ],
        col_widths=[6.5, 3.0, 5.2]
    )
    add_note(doc,
        "Overall = média simples dos Net Signals de Pertencimento, Acordos e Economia Fraterna."
    )

    # ── 6. Modificador de Tempo ───────────────────────────────────────────────
    add_heading(doc, "6. Modificador de Tempo de Escola", 1)
    add_body(doc,
        "O tempo de vínculo com a escola é uma variável estrutural da curva: a maturidade "
        "comunitária não é alcançada apenas pela concordância com valores, mas pelo tempo de "
        "vivência, absorção de processos e experiências reais. Por isso, aplicam-se dois "
        "modificadores que sobrepõem a lógica de concordância:"
    )

    add_heading(doc, "6.1 Teto para novatos (≤ 1 ano na escola)", 2)
    add_body(doc,
        "Membros que entraram em 2024 ou 2025 têm o estágio inferido limitado a Inserindo (3), "
        "independentemente de qualquer nível de concordância."
    )
    add_bullet(doc,
        "Alta concordância é registrada na justificativa como sinal de potencial rápido.")
    add_bullet(doc,
        "O campo 'Próximo nível sugerido' aponta para Comprometido — indicando a conversão esperada.")
    add_bullet(doc,
        "A lógica: concordar com os valores da escola em um questionário não equivale a ter "
        "vivenciado os processos, feito parte de situações de crise, construído relações "
        "comunitárias e decidido permanecer. O tempo é condição necessária para a virada.")

    add_heading(doc, "6.2 Piso para veteranos (≥ 5 anos sem virada)", 2)
    add_body(doc,
        "Membros com 5 ou mais anos na escola que ainda não apresentam o sinal de virada "
        "(threshold 'Disse SIM' não atingido) são tratados como anomalias da curva:"
    )
    table_simple(doc,
        ["Situação do veterano (≥ 5 anos)", "Estágio inferido", "Interpretação"],
        [
            ["Discordância em Pertencimento ≥ 25% OU Discordância em Acordos ≥ 30%",
             "Adaptando",
             "Tensão ativa com a escola — possível sinal de saída iminente"],
            ["Sem discordância significativa",
             "Inserindo",
             "No ponto de virada há muito tempo — decisão pendente"],
        ],
        col_widths=[7.0, 3.0, 4.7]
    )
    add_note(doc,
        "Justificativa: pela lógica da curva, um membro que está em Inserindo por 5+ anos "
        "deveria ter tomado a decisão de comprometer-se ou de sair. A permanência prolongada "
        "nesse estágio é ambígua — pode indicar que a decisão foi tomada informalmente mas "
        "não se reflete nas práticas, ou que há uma tensão não resolvida com a escola."
    )

    # ── 7. Engajamento ────────────────────────────────────────────────────────
    add_heading(doc, "7. Engajamento: Papéis Formais na Escola", 1)
    add_body(doc,
        "O engajamento é medido por 5 papéis formais da escola, cada um com três opções de resposta:"
    )
    table_simple(doc,
        ["Papel", "Resposta", "Sinal na inferência"],
        [
            ["Representante ou tesoureiro de turma", "Sou atualmente", "Papel ativo — conta para 'active'"],
            ["Voluntário em eventos",                "Sou atualmente", "Papel ativo — conta para 'active'"],
            ["Membro de grupos de trabalho ou comissões", "Sou atualmente", "Papel ativo — conta para 'active'"],
            ["Membro da APG",                        "Sou atualmente", "Papel ativo — conta para 'active'"],
            ["Mutirões da escola",                   "Sou atualmente", "Papel ativo — conta para 'active'"],
            ["Qualquer papel",                       "Já fui",         "Engajamento histórico — registrado mas não pesa na inferência do lado direito"],
            ["Qualquer papel",                       "Nunca fui",      "Ausência de engajamento"],
        ],
        col_widths=[5.2, 3.0, 6.5]
    )
    add_body(doc,
        "Para a inferência, o que importa é o engajamento ATUAL ('Sou atualmente'). "
        "Os thresholds são:"
    )
    add_bullet(doc, "1 papel ativo: condição mínima para Ativo no Motivo", bold_prefix="active (≥1): ")
    add_bullet(doc, "qualifica para Embaixador (combinado com overall ≥ 70%)", bold_prefix="multi (≥2): ")
    add_bullet(doc, "qualifica para Cocriador (combinado com overall ≥ 75% e Ec. Fraterna ≥ 60%)", bold_prefix="many (≥3): ")

    # ── 8. Alinhamento Q10 vs Inferido ───────────────────────────────────────
    add_heading(doc, "8. Alinhamento entre Autoposicionamento (Q10) e Inferido", 1)
    add_body(doc,
        "Cada respondente declarou sua posição na curva na pergunta Q10. O alinhamento compara "
        "essa autodeclaração com o estágio inferido pelas respostas:"
    )
    table_simple(doc,
        ["Diferença (Inferido − Q10)", "Categoria de Alinhamento", "Interpretação"],
        [
            ["= 0", "Alinhado",                           "A percepção própria coincide com as evidências"],
            ["+1",  "Potencial +1 nível",                 "As respostas sustentam 1 nível acima do autodeclarado"],
            ["≥+2", "Potencial +2 ou mais níveis",        "Subavaliação relevante — respostas mostram nível muito acima"],
            ["−1",  "Autoposicionamento +1 acima",        "Declara 1 nível acima do que as respostas sustentam"],
            ["≤−2", "Autoposicionamento +2 ou mais acima","Sobreposicionamento significativo — autopercepção elevada"],
        ],
        col_widths=[3.5, 5.5, 5.7]
    )
    add_note(doc,
        "Importante: o alinhamento não é um juízo de valor. Uma pessoa 'Autoposicionamento +2 acima' "
        "pode estar correta em sua percepção e as respostas ao formulário podem não capturar "
        "toda a sua contribuição. O dado é um ponto de partida para conversa, não um veredito."
    )

    # ── 9. Fluxo resumido ─────────────────────────────────────────────────────
    add_heading(doc, "9. Fluxo Resumido da Inferência", 1)
    add_body(doc, "A sequência completa de decisões para cada respondente:")

    passos = [
        ("Passo 1", "Calcular Net Signal para Pertencimento, Acordos e Economia Fraterna."),
        ("Passo 2", "Calcular % Sim no Conhecimento institucional."),
        ("Passo 3", "Contar papéis com resposta 'Sou atualmente' no Engajamento."),
        ("Passo 4", "Verificar tempo de escola (anos = 2025 − ano de entrada)."),
        ("Passo 5",
         "Aplicar modificador de TEMPO:\n"
         "  • Se ≤ 1 ano → marcar como novato (teto = Inserindo após toda a lógica)\n"
         "  • Se ≥ 5 anos E não disse SIM → retornar Adaptando (se discord ≥ 25%) ou Inserindo"),
        ("Passo 6",
         "Verificar se 'Disse SIM': Net Pertencimento ≥ 0,60 E discordância Pertencimento < 25%\n"
         "  • Não → aplicar lógica do lado esquerdo (Seção 5.2)\n"
         "  • Sim → aplicar lógica do lado direito (Seção 5.3)"),
        ("Passo 7",
         "Aplicar teto de novato: se anos ≤ 1 e estágio inferido > Inserindo → cap em Inserindo."),
        ("Passo 8", "Registrar estágio inferido, justificativa textual, próximo nível e papel na comunidade."),
    ]

    for num, desc in passos:
        p = doc.add_paragraph(style="List Number")
        p.paragraph_format.space_after = Pt(4)
        r1 = p.add_run(num + ": ")
        r1.bold = True
        r1.font.size = Pt(11)
        r2 = p.add_run(desc)
        r2.font.size = Pt(11)

    doc.add_paragraph()

    # ── 10. Limitações ────────────────────────────────────────────────────────
    add_heading(doc, "10. Limitações e Considerações", 1)
    add_body(doc,
        "A metodologia de inferência é uma aproximação baseada em respostas a um formulário. "
        "Algumas limitações devem ser consideradas ao interpretar os resultados:"
    )
    limitacoes = [
        ("Desejabilidade social",
         "Respondentes podem ter respondido de forma mais positiva do que sua vivência real "
         "reflete. Concordância alta em um formulário não é equivalente à prática comunitária."),
        ("Momento da resposta",
         "O formulário captura um momento. Pessoas em transição entre estágios podem oscilar "
         "entre classificações dependendo de eventos recentes."),
        ("Engajamento histórico",
         "A variável 'Já fui' (engajamento passado) não pesa na inferência do lado direito. "
         "Uma pessoa que foi muito ativa e se afastou recentemente pode estar subestimada."),
        ("Threshold de novato",
         "O teto de 1 ano pode ser conservador para pessoas que já tinham conexão prévia com a "
         "escola antes de entrar formalmente (ex: irmãos ou amigos já na comunidade)."),
        ("Perguntas não respondidas",
         "Respostas em branco são ignoradas. Em esferas com poucas respostas válidas, "
         "o Net Signal pode ser menos representativo."),
        ("Net Signal não é score de motivação",
         "Mede padrões de resposta em um questionário — não é uma avaliação de comprometimento "
         "pessoal. Deve ser usado como ponto de partida para conversa, não como veredito."),
    ]
    for titulo, descricao in limitacoes:
        add_bullet(doc, f" {descricao}", bold_prefix=titulo + ":")

    # ── Rodapé ────────────────────────────────────────────────────────────────
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(
        "Documento gerado a partir da pesquisa de maturidade comunitária Arandu · 2025 · "
        "143 respondentes · 34 perguntas pontuadas"
    )
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
    run.font.italic = True

    doc.save(SAIDA)
    print(f"✓ Salvo: {SAIDA}")


if __name__ == "__main__":
    main()
