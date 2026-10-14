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

Copie `backend/.env.example` para `backend/.env` e preencha os valores somente quando o endpoint oficial estiver disponível. O arquivo `.env` não deve ser enviado ao Git.

Mantenha `MODO_TESTE_API=true` enquanto estiver usando a simulação.

## Gerar o executável

Execute `backend/build_exe.bat` no Windows. O executável será gerado em `dist/ConciliacaoDMA.exe`. A pasta `dist/` e os dados locais de planilhas são ignorados pelo Git.

## Segurança dos dados

- Não inclua tokens, arquivos `.env` ou planilhas reais neste repositório.
- Os relatórios Excel podem conter números de contratos; armazene-os apenas em local autorizado.
- O executável ainda não possui assinatura digital, então o Windows pode exibir o aviso do SmartScreen ao abrir.
