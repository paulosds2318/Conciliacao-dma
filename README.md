# Conciliação DMA

Aplicativo desktop para carregar uma planilha de contratos, selecionar a coluna de identificação e processar a ação de DMA. Inclui conferência antes do processamento, filtro de resultados e exportação do relatório em Excel.

## Estado atual

O aplicativo usa a API simulada por padrão. Nesse modo, nenhum dado é enviado para a API Coral. O endpoint e as credenciais de produção ainda precisam ser fornecidos antes de habilitar chamadas reais.

## Executar em desenvolvimento

Requisitos: Windows e Python instalado.

```powershell
cd backend
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python desktop\main.py
```

## Configuração local

Em desenvolvimento, copie `backend/.env.example` para `backend/.env`. Ao usar o executavel, coloque o `.env` ao lado de `ConciliacaoDMA.exe` e configure `CORAL_API_URL` e `CORAL_AUTH_TOKEN`. O `.env` nao deve ser enviado ao Git.

Mantenha `MODO_TESTE_API=true` enquanto estiver usando a simulação.

## Gerar o executável

Execute `backend/build_exe.bat` no Windows. O build usa `backend/ConciliacaoDMA.spec` como configuração única e gera `dist/ConciliacaoDMA.exe`. A pasta `dist/` e os dados locais de planilhas são ignorados pelo Git.

Arquivos `.xlsx` são lidos em fluxo, preservando identificadores de texto e zeros à esquerda de células numéricas com formato simples de zeros. A tabela e a exportação consultam resultados em SQLite para limitar o uso de memória. O checkpoint permite retomar o mesmo arquivo e ação sem reenviar contratos já confirmados. Falhas de rede e respostas HTTP 5xx ficam como incertas e exigem conferência antes de reenvio; chamadas de alteração não são repetidas automaticamente. O checkpoint local, que contém os detalhes para o relatório, é mantido por até 30 dias em `%LOCALAPPDATA%/ConciliacaoDMA/processamentos.sqlite3`. O log técnico não inclui números de contrato nem credenciais e fica em `%LOCALAPPDATA%/ConciliacaoDMA/logs/conciliacao.log`.

## Segurança dos dados

- Não inclua tokens, arquivos `.env` ou planilhas reais neste repositório.
- Os relatórios Excel podem conter números de contratos; armazene-os apenas em local autorizado.
- O executável ainda não possui assinatura digital, então o Windows pode exibir o aviso do SmartScreen ao abrir.
