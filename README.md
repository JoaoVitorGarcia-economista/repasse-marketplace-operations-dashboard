Ferramenta automatizada para diagnóstico operacional de conciliação em marketplaces. A partir de apenas 3 dados por canal, esta ferramenta gera um relatório Excel completo com 7 análises, KPIs interativos, gráficos e ranking de priorização para automação.

O Problema
Empresas que vendem em múltiplos marketplaces enfrentam um desafio crescente: cada plataforma possui seu próprio formato de repasse financeiro, com regras, layouts e complexidades distintas. Sem medição precisa, há sobrecarga oculta da equipe, escalonamento reativo (surpresas ao crescer volume), vulnerabilidade operacional e ausência de critério para automação. Esta ferramenta resolve isso.

Como Funciona
Você alimenta uma planilha simples com os dados dos seus canais, executa o script Python, e recebe um Excel completo com dashboard executivo, 7 análises detalhadas, gráficos automáticos e ranking de prioridade.

Arquivo de Entrada
Crie um arquivo dados_conciliacao.xlsx com a seguinte estrutura: na primeira linha, os nomes dos marketplaces como colunas (Amazon, Mercado Livre, Shopee, Magalu, etc.). Na linha "quantidade mês", coloque o número de repasses por mês de cada canal. Na linha "hora gasta para conclilar", coloque o tempo em horas gasto para conciliar cada repasse. O script detecta automaticamente novos marketplaces adicionados - basta incluir a nova coluna e executar novamente.

Premissas Configuráveis
No início do arquivo conciliacao_marketplace.py você pode ajustar: HORAS_MES_PESSOA = 176 (jornada mensal por pessoa, padrão 22 dias x 8h), CUSTO_HORA = 20.00 (custo por hora do conciliador em reais), VARIACAO_TEMPO_LENTO = 0.30 (simulação de piora de 30% no tempo), VARIACAO_TEMPO_RAPIDO = 0.30 (simulação de melhoria de 30% no tempo), CUSTO_AUTOMACAO = 15000.00 (investimento de referência para automação em reais).

O Que Você Recebe
Dashboard Executivo (aba principal): 8 KPIs principais (totais, horas, pessoas, custos), tabela consolidada de todos os canais e ranking de prioridade com medalhas (ouro/prata/bronze).
Sete análises detalhadas: Análise Base (quanto trabalho cada canal gera por mês), Complexidade (quão mais lento é cada canal vs o mais rápido), Concentração de Esforço (se o esforço da equipe é proporcional ao volume), Custo de Oportunidade (quantas horas e reais estamos perdendo), Sensibilidade de Tempo (impacto se o tempo de repasse mudar ±30%), Payback de Automação (em quantos meses o investimento se paga por canal), Ranking de Prioridade (qual canal atacar primeiro).
Gráficos automáticos incluídos: barras de horas por canal, barras de índice de complexidade, barras agrupadas de percentual volume vs esforço, barras de horas poupadas, barras comparativas dos três cenários (atual/lento/rápido), barras de payback por canal, barras de score de prioridade e pizza de distribuição de horas.

Principais Fórmulas
Horas Totais = Repasses x Tempo por Repasse. Pessoas Necessárias = Horas Totais dividido por 176h. Índice de Complexidade = Tempo do canal dividido pelo tempo do canal mais rápido. Delta de Concentração = Percentual de Horas menos Percentual de Repasses. Horas Poupadas = Horas Atuais menos Horas no cenário ideal. Custo Mensal = Horas Totais vezes Custo por Hora. Payback = Custo da Automação dividido pelo Custo Mensal. Score de Prioridade = Horas Totais vezes (Tempo por Repasse vezes Custo por Hora).

Executando
No terminal, execute: python conciliacao_marketplace.py
Saída esperada: "✅ Excel gerado: analise_conciliacao_marketplace.xlsx"

Atualizando a Análise
Basta substituir o arquivo dados_conciliacao.xlsx com novos dados e executar o script novamente. Todo o relatório é regenerado automaticamente.

Casos de Uso
Para reuniões de planejamento, use o dashboard executivo com todos os KPIs. Para justificar investimento em automação, use a análise de payback e ranking que mostram o ROI por canal. Para dimensionar equipe, veja quantas pessoas são necessárias por canal e no total. Para análise de risco, use a sensibilidade de tempo para ver o impacto de mudanças nos processos.

Tecnologias
Python, Pandas, OpenPyXL
Licença
MIT

