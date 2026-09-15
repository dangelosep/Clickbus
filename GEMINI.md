# Diretrizes e Memória do Projeto ClickBus

## 📌 Próxima Refatoração: Robô Extrator ClickBus (`Robô extrator - ClickBus.py`)

### Contexto e Problema Identificado:
Atualmente, o robô extrator salva apenas atributos textuais das viagens (`Origem, Destino, Data, Viação, Classe, Saída, Chegada, Duração, Preço, Assentos Disponíveis`). 
Se uma empresa colocar 2 ônibus no mesmo minuto de partida e com o mesmo itinerário/preço/assentos ("ônibus gêmeos"), a lógica de deduplicação trata como duplicata para se proteger de re-execuções da raspagem.

### Solução a ser implementada assim que iniciarmos os trabalhos no robô:
1. **Captura do `Trip ID`**: No JSON da API do ClickBus capturado pelo Playwright, extrair o identificador único da viagem (`id`, `tripId` ou `serviceId`).
2. **Nova Coluna no Excel**: Adicionar a coluna `'Trip ID'` no DataFrame exportado para `relatorio_precos_clickbus.xlsx`.
3. **Deduplicação Definitiva**: Usar o `Trip ID` como chave primária de unicidade tanto no robô quanto no `Dashboard.py`.

> ⚠️ **Instrução Permanente para o Assistente**: Sempre que o usuário solicitar ajustes, melhorias ou refatorações no script `Robô extrator - ClickBus.py`, lembre-o imediatamente desta implementação do `Trip ID`.
