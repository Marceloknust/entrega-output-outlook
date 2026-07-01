# output_outlook

Automacao de validacao e consolidacao do cronograma de Demanda Primaria com integracao ao Outlook.

## Estrutura esperada

- `input/DE_PARA.xlsx` (aba `DE_PARA`)
- `input/Base_Dados.xlsx` (aba `Base Dados`)
- `output/old/` para backup do DE_PARA
- `scripts/script1_gerar_validacao.py`
- `scripts/script2_aplicar_ajustes.py`

## Dependencias

```powershell
.venv\Scripts\pip install -r src\Entregas\output_outlook\requirements.txt
```

## Script 1 - gerar validacao

Objetivo:
- Selecionar ano/mes via interface.
- Expandir cronograma com regra 26->25 (referencia do `expandir_cronoweb.py`).
- Aplicar DE_PARA (`Bacia Ajustada` + `Bloco Ajustado`) para preencher `Atendimento`, `Contrato`, `Modalidade`.
- Criar colunas `_ajustado` e `Validado` (inicial = `Nao`).
- Gerar arquivo `cronograma_outlook_validacao_aaaamm.xlsx` com abas `cronograma_base`, `Base Dados` e `TD`.
- Perguntar/permitir editar destinatarios antes de enviar e-mail de validacao.

Execucao:

```powershell
.venv\Scripts\python.exe src\Entregas\output_outlook\scripts\script1_gerar_validacao.py
```

## Script 2 - aplicar ajustes

Objetivo:
- Ler arquivo validado (`cronograma_outlook_validacao_aaaamm.xlsx`).
- Identificar mudancas nas colunas `_ajustado`.
- Fazer backup de `DE_PARA.xlsx` em `output/old/DE_PARA_AAAAMM_HHMMSS.xlsx`.
- Atualizar `Atendimento`, `Contrato`, `Modalidade` no DE_PARA.
- Gerar `cronograma_outlook_final_aaaamm.xlsx` sem colunas `_ajustado` e `Validado`.
- Perguntar/permitir editar destinatario antes de enviar e-mail final.

Execucao:

```powershell
.venv\Scripts\python.exe src\Entregas\output_outlook\scripts\script2_aplicar_ajustes.py
```
