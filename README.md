# Contagem automática de relatórios do Auvo

Automatiza a conferência manual dos relatórios de execução do Auvo (higienização /
troca de cortinas). O script busca as tarefas pela API, lê o questionário preenchido
em campo e compara a quantidade **esperada** com a quantidade **realizada**,
apontando as divergências.

## Como a contagem funciona

A quantidade esperada está no nome da pergunta e a realizada na resposta:

| Pergunta (questionário)                      | Resposta em campo      | Resultado        |
|----------------------------------------------|------------------------|------------------|
| `CENTRO CIRÚRGICO 10º - RPA  I  3 Cortinas`  | `3`                    | OK               |
| `CENTRO CIRÚRGICO 4º - RPA  I  12 Cortinas`  | `11 (1 não enviada)`   | FALTA (-1)       |

Regras aplicadas:

- A quantidade esperada é lida após o separador ` I ` no nome da pergunta.
- A quantidade realizada é o número no início da resposta.
- Texto entre parênteses é capturado como observação da divergência.
- Respostas que são URL (assinaturas e fotos) são ignoradas na contagem.
- Perguntas sem quantidade esperada (ex: `Manguito`, `Observações`) entram como
  informativas, sem checagem de divergência.

## Configuração

As credenciais saem do painel do Auvo em **Menu > Integrações > Chaves de integração**.

Opção 1 — variáveis de ambiente:

```powershell
$env:AUVO_API_KEY   = "sua-app-key"
$env:AUVO_API_TOKEN = "seu-token"
```

Opção 2 — arquivo local: copie `config.exemplo.json` para `config.local.json` e
preencha. Esse arquivo é ignorado pelo Git e não vai para o repositório.

## Uso

```powershell
# Tarefas de hoje
.\auvo-contagem.ps1

# Um dia específico
.\auvo-contagem.ps1 -StartDate 2026-09-15 -EndDate 2026-09-15

# Um mês inteiro, filtrando por cliente
.\auvo-contagem.ps1 -StartDate 2026-09-01 -EndDate 2026-09-30 -CustomerId 24201924
```

Saída no terminal com o resumo de cada ficha e as divergências em destaque, mais
dois CSVs (separados por `;`, prontos para o Excel) na pasta `saida/`:

- `itens-<data>.csv` — uma linha por setor/item conferido
- `resumo-<data>.csv` — uma linha por ficha, com totais e status

## Agendamento

Para rodar sozinho todo dia, use o Agendador de Tarefas do Windows apontando para:

```
powershell.exe -ExecutionPolicy Bypass -File "C:\caminho\auvo-contagem.ps1"
```

## Referência da API

- Base: `https://api.auvo.com.br/v2`
- Login: `GET /login?apiKey=...&apiToken=...` (token vale 30 min, o script renova sozinho)
- Tarefas: `GET /tasks?paramFilter={json}&page=1&pageSize=100`
  (filtros vão dentro de `paramFilter` como JSON; datas no formato `yyyy-MM-dd`)
- Detalhe: `GET /tasks/{id}` — traz `questionnaires[].answers[]` com as respostas
- Documentação oficial: `developer.auvo.com.br`
