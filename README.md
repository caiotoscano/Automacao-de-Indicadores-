# Automação de Indicadores 
 Projeto feito para automatizar processos dentro de um cenário empresarial
# Objetivo
 Treinar e criar um Projeto Completo que envolva a automatização de um processo feito no computador
# Explicação do Projeto
 Imagine que você trabalha em uma grande rede de lojas de roupa com 25 lojas espalhadas por todo o Brasil.
 Todo dia, pela manhã, a equipe de análise de dados calcula os chamados One Pages e envia para o gerente de cada loja o OnePage da sua loja, bem como todas as informações usadas no cálculo dos indicadores.
# O que é um OnePage? 
 Um One Page é um resumo muito simples e direto ao ponto, usado pela equipe de gerência de loja para saber os principais indicadores de cada loja e permitir em 1 página (daí o nome OnePage) tanto a comparação entre diferentes lojas, quanto quais indicadores aquela    loja conseguiu cumprir naquele dia ou não.
 # O que se espera do projeto?
 Conseguir criar um processo da forma mais automática possível para calcular o OnePage de cada loja e enviar um email para o gerente de cada loja com o seu OnePage no corpo do e-mail e também o arquivo completo com os dados da sua respectiva loja em anexo.
 Ao final, envia ainda um e-mail para a diretoria (informações também estão no arquivo Emails.xlsx) com 2 rankings das lojas em anexo, 1 ranking do dia e outro ranking anual. Além disso, no corpo do e-mail, ressalta qual foi a melhor e a pior loja do dia e também a   melhor e pior loja do ano. O ranking de uma loja é dado pelo faturamento da loja.
# Estrutura do código
  Pode se observar no arquivo "Cod.Projeto.py" todas as linhas de código que detalharei neste arquivo. Para começar organizei o que deveria ser feito em partes fazendo um comentário de o que se trataria cada célula de código (projeto realizado no Jupyter Notebook). Feito o brainstorm inicial, comecei importando as bibliotecas que usaria nesse primeiro momento e importando a base de dados (pandas). Após importar toda a base de dados fiz um tratamento para aglutinar o ID de cada loja com o dia de cada venda efetuada naquela loja em específico e criei uma planilha com esses dados. Calculei todos os dados para cada loja (automático) como faturamento_dia/faturamento_ano/ticket_medio etc... Após ter todos esses valores criei o corpo do e-mail (HTML) e configurei a mensagem do e-mail. Dando continuidade enviei todos os indicadores para todas as lojas de forma automática utilizando o processo de integração Python com Gmail. Consequentemente enviei também os indicadores para os gerentes e um relatório geral de cada loja para cada gerente explicando qual loja obteve melhor/pior desempenho. 
