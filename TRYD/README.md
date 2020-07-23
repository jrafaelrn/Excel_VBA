## Controle Operacional de Ações no Excel ##
(integrado com Tryd&reg;)


------------



📜	**Requisitos** 📜
-  Instalar o Tryd e registrar a DLL 
- Ter uma versão do Microsoft Excel&reg; e habilita-lá para executar macros.
- Ter uma conexão com a internet para receber as cotações em tempo real do Tryd para o Excel


------------


🤔 **Como funciona?** 🤔

Ao abrir a planilha, ela será direcionada para a tela abaixo.
Aqui, a jornada foi pensada no seguinte fluxo (cada item será detalhado posteriormente):
1. Antes de mais nada, é feita a ANÁLISE DE RISCO da operação;
1. Após isso, podem ser cadastrados ALERTAS para que a planilha "fale" quando o valor da cotação atingir o gatilho configurado;
1. Quando um gatilho é acionado e de fato você realiza a operação, então ela pode ser cadastrada em ORDENS.


Na barra **superior** existem 2 opções:
- **Visão Mercado:** é possível acompanhar a situação geral do mercado e da sua carteira em uma mesma tela;
- **Desempenho:** Resumo das suas operações com resultados finais processados, para analisar o desempenho de cada operação.


Na barra **lateral** existem as seguintes configurações:
- **Contas:** é possível cadastrar as corretoras e as contas com as quais opera
- **Book:** onde é cadastrado os papéis que vocês gostaria de acompanhar
- **Cotas:** opção disponível para quem opera em fundos de investimento privado e precisa gerir as cotas
- **Mais configurações:** ajustes e tabelas de suporte para o funcionamento da planilha
- Ainda é possível encontrar os botões de zoom e o *Sobre.*



![Página Inicial](https://github.com/jrafaelrn/excel_VBA/blob/master/TRYD/HOW_TO/Tela%20Inicial.PNG?raw=true "Página Inicial")

------------


![Análise de Risco](https://github.com/jrafaelrn/excel_VBA/blob/master/TRYD/HOW_TO/Analise%20de%20Risco_icon.PNG?raw=true "Análise de Risco")

Abaixo se encontram os campos que podem ser ajustados conforme a estratégia de cada investidor, alguns estão comentados dentro da própria planilha .

Após definido os valores de entrada e saída da operação e analisado se a relação **RISCO/RETORNO** está dentro de sua estratégia, você pode:
	 ⏰ **ADICIONAR UM ALERTA** e quando o papel atingir o **VALOR DISPARO**, uma mensagem de voz será gerada pela planilha.*
	 (Uma tela de "envio para automatizador" irá aparecer, para que seja possível enviar a ordem para o Tryd lançar no sistema da Bolsa, mas essa função ainda não está funcional, aguardem as próximas versões)*

![Análise de Risco Detalhada](https://github.com/jrafaelrn/excel_VBA/blob/master/TRYD/HOW_TO/Analise%20de%20Risco.PNG?raw=true "Análise de Risco Detalhada")
