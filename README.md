# Pesquisa de Maturidade da Comunidade Escolar — Arandu

Script de processamento da pesquisa sobre o estágio de maturidade da comunidade
escolar (famílias e colaboradores) em relação à proposta antroposófica, com base
nos três pilares da **trimembração social** (Pertencimento Cultural, Acordos
Sociais e Economia Fraterna) e na curva dos **7 Processos Sociais**:

1. Novo (respiração)
2. Adaptando (aquecimento)
3. Inserindo (digestão)
4. Comprometido (segregação)
5. Ativo no Motivo (manutenção)
6. Embaixador (crescimento)
7. Cocriador (geração)

## Como usar

```bash
pip install pandas odfpy openpyxl

python processa_pesquisa.py "Base da pesquisa oara chatgpt.ods" -o pesquisa_processada.xlsx
```

O script espera um arquivo `.ods` (ou `.xlsx`) com as abas:

| Aba | Conteúdo |
|---|---|
| Respostas o formulario | uma linha por respondente, uma coluna por pergunta |
| Perguntas | perguntas relevantes para análise (opcionalmente com a coluna **Pilar**) |
| Opções de reposta | categorias de resposta (apenas referência) |
| Lógica | para cada pergunta + opção de resposta: nível, valor e **estágio da curva** |

Os nomes de abas e colunas são localizados de forma flexível (ignora acentos,
maiúsculas e pequenas variações). Se algo não for encontrado, o script avisa no
terminal indicando o que ajustar.

## Como o score é calculado

Sobre a dúvida de "não estar certa de que os valores levam ao resultado
esperado": a recomendação implementada é **ancorar a pontuação diretamente na
curva**, em vez de usar valores arbitrários.

1. Na aba **Lógica**, cada opção de resposta já tem um estágio da curva
   atribuído. O script converte esse estágio em um número de 1 (Novo) a 7
   (Cocriador) e usa esse número como pontuação da resposta.
2. O **score final** de cada respondente é a média desses números — ou seja, o
   score já vive na escala da própria curva (1 a 7).
3. Os **intervalos de corte** padrão arredondam para o estágio mais próximo
   (1,5 / 2,5 / 3,5 / 4,5 / 5,5 / 6,5), mas são configuráveis na constante
   `CORTES` no topo do script — por exemplo, para exigir um score mais alto
   para "Cocriador".
4. Se a aba Lógica não tiver coluna de estágio, o script usa a coluna de
   **Valor** e reescala os cortes proporcionalmente à faixa de valores.

## Saída (`pesquisa_processada.xlsx`)

| Aba | Conteúdo |
|---|---|
| Scores Individuais | respostas originais + score por pilar, score final e estágio na curva de cada respondente |
| Pontos por Pergunta | a pontuação atribuída a cada resposta (para auditoria) |
| Resumo por Estágio | quantos respondentes em cada estágio da curva (visão geral) |
| Resumo por Pilar | score médio por pilar da trimembração |
| Resumo por Segmento | comparação Famílias × Colaboradores (se a coluna existir) |
| Cortes Utilizados | faixas de score usadas para cada estágio |
