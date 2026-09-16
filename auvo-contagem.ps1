<#
    Contagem automatica dos relatorios de execucao do Auvo.

    Le as tarefas via API, extrai o questionario de cada uma e compara a
    quantidade esperada (que fica no nome da pergunta, ex: "... I 12 Cortinas")
    com a quantidade informada em campo (ex: "11 (1 nao enviada)").

    Uso:
        .\auvo-contagem.ps1
        .\auvo-contagem.ps1 -StartDate 2026-09-15 -EndDate 2026-09-15
        .\auvo-contagem.ps1 -StartDate 2026-09-01 -EndDate 2026-09-30 -CustomerId 24201924
#>
[CmdletBinding()]
param(
    [string]$StartDate  = (Get-Date).ToString("yyyy-MM-dd"),
    [string]$EndDate    = (Get-Date).ToString("yyyy-MM-dd"),
    [int]   $CustomerId = 0,
    [string]$OutputDir  = (Join-Path $PSScriptRoot "saida")
)

$ErrorActionPreference = "Stop"
$BaseUrl = "https://api.auvo.com.br/v2"

$script:Headers     = $null
$script:TokenExpira = [datetime]::MinValue

function Get-AuvoCredencial {
    $key   = $env:AUVO_API_KEY
    $token = $env:AUVO_API_TOKEN

    if (-not $key -or -not $token) {
        $cfg = Join-Path $PSScriptRoot "config.local.json"
        if (Test-Path $cfg) {
            $c = Get-Content $cfg -Raw -Encoding UTF8 | ConvertFrom-Json
            if (-not $key)   { $key   = $c.apiKey }
            if (-not $token) { $token = $c.apiToken }
        }
    }

    if (-not $key -or -not $token) {
        throw "Credenciais nao encontradas. Defina AUVO_API_KEY e AUVO_API_TOKEN, ou crie o arquivo config.local.json (veja config.exemplo.json)."
    }

    return @{ Key = $key; Token = $token }
}

# O token do Auvo vale 30 minutos; renova sozinho quando necessario.
function Get-AuvoHeaders {
    if ($script:Headers -and (Get-Date) -lt $script:TokenExpira) { return $script:Headers }

    $c   = Get-AuvoCredencial
    $url = "$BaseUrl/login?apiKey=$([uri]::EscapeDataString($c.Key))&apiToken=$([uri]::EscapeDataString($c.Token))"
    $r   = Invoke-RestMethod -Uri $url -Method Get

    $script:Headers     = @{ Authorization = "Bearer $($r.result.accessToken)" }
    $script:TokenExpira = (Get-Date).AddMinutes(25)
    return $script:Headers
}

function Get-AuvoTarefas {
    param([string]$StartDate, [string]$EndDate, [int]$CustomerId)

    $filtro = [ordered]@{ startDate = $StartDate; endDate = $EndDate; status = 4 }
    if ($CustomerId -gt 0) { $filtro.customerId = $CustomerId }
    $pf = [uri]::EscapeDataString(($filtro | ConvertTo-Json -Compress))

    $tarefas = @()
    $pagina  = 1
    $total   = 0

    do {
        $url  = "$BaseUrl/tasks?paramFilter=$pf&page=$pagina&pageSize=100"
        $r    = Invoke-RestMethod -Uri $url -Headers (Get-AuvoHeaders) -Method Get
        $lote = @($r.result.entityList)
        $total = $r.result.pagedSearchReturnData.totalItems
        if ($lote.Count -gt 0) { $tarefas += $lote }
        $pagina++
    } while ($lote.Count -gt 0 -and $tarefas.Count -lt $total)

    return $tarefas
}

function Get-AuvoTarefaDetalhe {
    param([int]$TaskId)
    (Invoke-RestMethod -Uri "$BaseUrl/tasks/$TaskId" -Headers (Get-AuvoHeaders) -Method Get).result
}

function ConvertTo-Itens {
    param($Tarefa)

    $itens = @()

    foreach ($q in $Tarefa.questionnaires) {
        foreach ($a in $q.answers) {
            $desc  = [string]$a.questionDescription
            $reply = [string]$a.reply

            # Respostas de assinatura/foto sao URLs - nao entram na contagem.
            if ($reply -match '^\s*https?://') { continue }

            # Quantidade esperada: vem depois do separador " I " no nome da pergunta.
            $esperado = $null
            $unidade  = ""
            if ($desc -match '\sI\s+(\d+)\s*(\S*)\s*$') {
                $esperado = [int]$Matches[1]
                $unidade  = $Matches[2]
            }

            # Quantidade realizada: numero no inicio da resposta.
            $realizado = $null
            if ($reply -match '^\s*(\d+)') { $realizado = [int]$Matches[1] }

            # Texto entre parenteses costuma explicar a divergencia.
            $observacao = ""
            if ($reply -match '\(([^)]+)\)') { $observacao = $Matches[1] }

            $situacao  = "INFO"
            $diferenca = $null

            if ($null -ne $esperado) {
                if ($null -eq $realizado) {
                    $situacao = "SEM RESPOSTA"
                } else {
                    $diferenca = $realizado - $esperado
                    if     ($diferenca -lt 0) { $situacao = "FALTA" }
                    elseif ($diferenca -gt 0) { $situacao = "EXCEDENTE" }
                    else                      { $situacao = "OK" }
                }
            }

            $itens += [pscustomobject]@{
                Ficha        = $Tarefa.externalId
                TaskID       = $Tarefa.taskID
                Data         = $Tarefa.taskDate
                Cliente      = $Tarefa.customerDescription
                Responsavel  = $Tarefa.userToName
                Servico      = $Tarefa.taskTypeDescription
                Setor        = $desc
                Unidade      = $unidade
                Esperado     = $esperado
                Realizado    = $realizado
                Diferenca    = $diferenca
                Situacao     = $situacao
                Observacao   = $observacao
                RespostaBruta= $reply
            }
        }
    }

    return $itens
}

# ---------------------------------------------------------------- execucao ---

Write-Host "Buscando tarefas de $StartDate ate $EndDate..." -ForegroundColor Cyan
$tarefas = Get-AuvoTarefas -StartDate $StartDate -EndDate $EndDate -CustomerId $CustomerId
Write-Host "$($tarefas.Count) tarefa(s) encontrada(s)." -ForegroundColor Cyan

$todosItens = @()
$resumo     = @()

foreach ($t in $tarefas) {
    $detalhe = Get-AuvoTarefaDetalhe -TaskId $t.taskID
    $itens   = ConvertTo-Itens -Tarefa $detalhe
    if ($itens.Count -eq 0) { continue }

    $todosItens += $itens

    $comEsperado  = @($itens | Where-Object { $null -ne $_.Esperado })
    $divergencias = @($comEsperado | Where-Object { $_.Situacao -ne "OK" })

    $totalEsperado  = ($comEsperado | Measure-Object -Property Esperado  -Sum).Sum
    $totalRealizado = ($comEsperado | Measure-Object -Property Realizado -Sum).Sum
    if (-not $totalEsperado)  { $totalEsperado  = 0 }
    if (-not $totalRealizado) { $totalRealizado = 0 }

    $resumo += [pscustomobject]@{
        Ficha          = $detalhe.externalId
        TaskID         = $detalhe.taskID
        Data           = $detalhe.taskDate
        Cliente        = $detalhe.customerDescription
        Responsavel    = $detalhe.userToName
        Servico        = $detalhe.taskTypeDescription
        Itens          = $comEsperado.Count
        TotalEsperado  = $totalEsperado
        TotalRealizado = $totalRealizado
        Diferenca      = $totalRealizado - $totalEsperado
        Divergencias   = $divergencias.Count
        Status         = if ($divergencias.Count -eq 0) { "CONFERIDO" } else { "VERIFICAR" }
    }

    $cor = if ($divergencias.Count -eq 0) { "Green" } else { "Yellow" }
    Write-Host ""
    Write-Host ("Ficha {0} | {1} | {2:yyyy-MM-dd}" -f $detalhe.externalId, $detalhe.customerDescription, [datetime]$detalhe.taskDate) -ForegroundColor $cor
    Write-Host ("  Esperado: {0}  Realizado: {1}  Diferenca: {2}" -f $totalEsperado, $totalRealizado, ($totalRealizado - $totalEsperado))

    foreach ($d in $divergencias) {
        Write-Host ("  [{0}] {1} -> esperado {2}, informado '{3}'" -f $d.Situacao, $d.Setor, $d.Esperado, $d.RespostaBruta) -ForegroundColor Red
    }
}

if ($todosItens.Count -eq 0) {
    Write-Host "`nNenhum item de questionario encontrado no periodo." -ForegroundColor Yellow
    return
}

if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }
$carimbo     = Get-Date -Format "yyyyMMdd-HHmmss"
$arquivoItens = Join-Path $OutputDir "itens-$carimbo.csv"
$arquivoResumo= Join-Path $OutputDir "resumo-$carimbo.csv"

$todosItens | Export-Csv -Path $arquivoItens  -NoTypeInformation -Encoding UTF8 -Delimiter ";"
$resumo     | Export-Csv -Path $arquivoResumo -NoTypeInformation -Encoding UTF8 -Delimiter ";"

$geralEsperado  = ($resumo | Measure-Object -Property TotalEsperado  -Sum).Sum
$geralRealizado = ($resumo | Measure-Object -Property TotalRealizado -Sum).Sum
$fichasComProblema = @($resumo | Where-Object { $_.Divergencias -gt 0 })

Write-Host ""
Write-Host "=============================== TOTAL ===============================" -ForegroundColor Cyan
Write-Host ("Fichas conferidas : {0}" -f $resumo.Count)
Write-Host ("Total esperado    : {0}" -f $geralEsperado)
Write-Host ("Total realizado   : {0}" -f $geralRealizado)
Write-Host ("Diferenca         : {0}" -f ($geralRealizado - $geralEsperado))
Write-Host ("Fichas a verificar: {0}" -f $fichasComProblema.Count) -ForegroundColor $(if ($fichasComProblema.Count -eq 0) { "Green" } else { "Yellow" })
Write-Host ""
Write-Host "CSV por item  : $arquivoItens"
Write-Host "CSV por ficha : $arquivoResumo"
