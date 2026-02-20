 Sistema de Controle de Amostras, Padrões e Gabaritos (VBA)

Este é um sistema desenvolvido em VBA para Microsoft Excel projetado para gerenciar o inventário, a movimentação (saídas e retornos) e a geração de relatórios de Amostras de Referência, Padrões de Clientes e Gabaritos (GBs).

 Funcionalidades Principais

Sincronização em Rede: Macros dedicadas para buscar e atualizar automaticamente a lista mestre de Amostras e Gabaritos a partir de planilhas matrizes localizadas em diretórios de rede (`\\s01\...`).
Gestão de Movimentações: Criação dinâmica de botões de "Registrar Saída" para cada item do inventário.
   Registro automático de data e hora de saída transferindo os dados para uma aba de histórico ("Movimentações").
    Botões dinâmicos de "Registrar Retorno" que calculam e fecham o ciclo do item, destacando a linha visualmente.
Automação de Relatórios: Geração automatizada de relatórios mensais em PDF (salvos em subpastas específicas), filtrando apenas as movimentações do mês corrente.
  Limpeza e Arquivamento: Rotina segura para arquivar e limpar dados antigos após a geração dos relatórios.
Segurança: Proteção automatizada das planilhas por senha (`1234`), permitindo que as macros executem suas funções sem deixar o código ou a estrutura expostos a edições acidentais de usuários.

## 📁 Estrutura do Código

O projeto está dividido em três componentes principais dentro do VBE (Visual Basic Editor) do Excel:

1.  `EstaPastaDeTrabalho` (Workbook): Contém os eventos de inicialização, garantindo que as planilhas sejam protegidas corretamente ao abrir e verificando se é o último dia do mês para acionar o relatório PDF.
2.  `Planilha1_Amostras` (Worksheet): Contém os eventos locais da planilha de inventário. Identifica quando um novo código (CI) é digitado manualmente e gera instantaneamente o botão de saída correspondente.
3.  `ModuloPrincipal` (Module): O "motor" do sistema. Contém todas as Sub-rotinas executáveis (`AtualizarListaMestra`, `AtualizarListaGBs`, `RegistrarSaida`, `RegistrarRetornoBotao`, `ExportarMovimentacoesPDF`).

🛠️ Como Instalar e Configurar

1.  Abra seu arquivo Excel habilitado para macros (`.xlsm`).
2.  Pressione `ALT + F11` para abrir o Editor VBA.
3.  No painel à esquerda (Project Explorer):
    * Dê um duplo-clique em **EstaPastaDeTrabalho** e cole o código correspondente.
    * Dê um duplo-clique na aba **Planilha1 (Amostra Referência e Padrão)** e cole o código correspondente.
    * Vá em **Inserir > Módulo** e cole todo o código do `ModuloPrincipal`.
4.  Salve o arquivo e reinicie o Excel.

 ⚠️ Requisitos e Configurações de Ambiente

* O sistema assume a existência de uma aba chamada `Amostra Referência e Padrão` e outra chamada `Movimentações`.
* Caminhos de Rede: As macros de atualização (`AtualizarListaMestra` e `AtualizarListaGBs`) contêm caminhos de rede hardcoded (`\\s01\...`). É necessário ajustar essas strings no código caso os caminhos dos arquivos matrizes mudem.
* O acesso aos caminhos de rede deve estar liberado pelo Firewall/Antivírus da máquina local.

---
*Projeto desenvolvido para otimização de fluxos de qualidade e calibração de instrumentos.*
