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
    [string]$OutputDir  = (Join-Path $PSScriptRoot "saida"),
    [switch]$NoBrowser
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
        throw "Credenciais não encontradas. Defina AUVO_API_KEY e AUVO_API_TOKEN, ou crie o arquivo config.local.json (veja config.exemplo.json)."
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

function Protect-Html {
    param([string]$Texto)
    if ([string]::IsNullOrEmpty($Texto)) { return "" }
    return ($Texto -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;')
}

function Format-DataBr {
    param($Valor)
    try { return ([datetime]$Valor).ToString('dd/MM/yyyy') } catch { return [string]$Valor }
}

function New-RelatorioHtml {
    param($Resumo, $Itens, [string]$StartDate, [string]$EndDate, [string]$Caminho)

    $gEsperado  = ($Resumo | Measure-Object -Property TotalEsperado  -Sum).Sum
    $gRealizado = ($Resumo | Measure-Object -Property TotalRealizado -Sum).Sum
    if (-not $gEsperado)  { $gEsperado  = 0 }
    if (-not $gRealizado) { $gRealizado = 0 }
    $gDif       = $gRealizado - $gEsperado
    $aVerificar = @($Resumo | Where-Object { $_.Divergencias -gt 0 }).Count

    $classeDif = if ($gDif -eq 0) { "ok" } else { "falta" }
    $classeVer = if ($aVerificar -eq 0) { "ok" } else { "falta" }

    $css = @"
:root{--bg:#f4f5f7;--card:#fff;--linha:#e5e7eb;--txt:#1f2430;--sec:#6b7280;--ok:#12805c;--falta:#c0392b;--exc:#b45309}
*{box-sizing:border-box}
body{margin:0;padding:32px 16px;background:var(--bg);color:var(--txt);font-family:-apple-system,'Segoe UI',Roboto,Arial,sans-serif}
.wrap{max-width:1080px;margin:0 auto}
h1{font-size:22px;margin:0 0 4px}
.sub{color:var(--sec);font-size:14px;margin-bottom:24px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:24px}
.kpi{background:var(--card);border:1px solid var(--linha);border-radius:10px;padding:14px 16px}
.kpi .rot{font-size:12px;color:var(--sec);text-transform:uppercase;letter-spacing:.04em}
.kpi .val{font-size:26px;font-weight:600;margin-top:4px}
.kpi .val.ok{color:var(--ok)}
.kpi .val.falta{color:var(--falta)}
.filtro{display:block;margin-bottom:16px;font-size:14px;color:var(--sec);cursor:pointer}
.card{background:var(--card);border:1px solid var(--linha);border-radius:10px;margin-bottom:16px;overflow:hidden}
.card>header{display:flex;justify-content:space-between;align-items:center;gap:12px;padding:14px 16px;border-bottom:1px solid var(--linha);flex-wrap:wrap}
.tit{font-weight:600}
.meta{color:var(--sec);font-size:13px;margin-top:2px}
.badge{font-size:12px;font-weight:600;padding:4px 10px;border-radius:999px;white-space:nowrap}
.badge.ok{background:#e7f5ef;color:var(--ok)}
.badge.alerta{background:#fdecea;color:var(--falta)}
table{width:100%;border-collapse:collapse;font-size:14px}
th{text-align:left;font-size:12px;text-transform:uppercase;letter-spacing:.04em;color:var(--sec);padding:10px 16px;border-bottom:1px solid var(--linha);font-weight:600}
td{padding:10px 16px;border-bottom:1px solid #f1f2f4}
tr:last-child td{border-bottom:none}
th.num,td.num{text-align:right;width:90px;font-variant-numeric:tabular-nums}
tr.falta{background:#fdf3f2}
tr.exc{background:#fdf6ec}
tr.sem{background:#f7f7f8}
.sit{font-weight:600;font-size:13px}
.sit.ok{color:var(--ok)}.sit.falta{color:var(--falta)}.sit.exc{color:var(--exc)}.sit.sem{color:var(--sec)}
.obs{color:var(--sec);font-size:13px}
.info{padding:12px 16px;border-top:1px solid var(--linha);background:#fafafa;font-size:13px;color:var(--sec)}
.info b{color:var(--txt);font-weight:600}
.vazio{background:var(--card);border:1px solid var(--linha);border-radius:10px;padding:32px;text-align:center;color:var(--sec)}
@media print{body{background:#fff;padding:0}.filtro{display:none}.card{break-inside:avoid}}
"@

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<!DOCTYPE html>')
    [void]$sb.AppendLine('<html lang="pt-BR"><head><meta charset="utf-8">')
    [void]$sb.AppendLine('<meta name="viewport" content="width=device-width,initial-scale=1">')
    [void]$sb.AppendLine('<title>Contagem Auvo</title>')
    [void]$sb.AppendLine("<style>$css</style></head><body><div class=""wrap"">")

    $periodo = if ($StartDate -eq $EndDate) { Format-DataBr $StartDate } else { "$(Format-DataBr $StartDate) a $(Format-DataBr $EndDate)" }
    [void]$sb.AppendLine("<h1>Contagem de execução &middot; Auvo</h1>")
    [void]$sb.AppendLine("<div class=""sub"">Período: $periodo &middot; gerado em $(Get-Date -Format 'dd/MM/yyyy HH:mm')</div>")

    [void]$sb.AppendLine('<div class="kpis">')
    [void]$sb.AppendLine("<div class=""kpi""><div class=""rot"">Fichas</div><div class=""val"">$($Resumo.Count)</div></div>")
    [void]$sb.AppendLine("<div class=""kpi""><div class=""rot"">Esperado</div><div class=""val"">$gEsperado</div></div>")
    [void]$sb.AppendLine("<div class=""kpi""><div class=""rot"">Realizado</div><div class=""val"">$gRealizado</div></div>")
    [void]$sb.AppendLine("<div class=""kpi""><div class=""rot"">Diferença</div><div class=""val $classeDif"">$gDif</div></div>")
    [void]$sb.AppendLine("<div class=""kpi""><div class=""rot"">A verificar</div><div class=""val $classeVer"">$aVerificar</div></div>")
    [void]$sb.AppendLine('</div>')

    [void]$sb.AppendLine('<label class="filtro"><input type="checkbox" id="soDiv"> Mostrar apenas as divergências</label>')

    foreach ($f in $Resumo) {
        $daFicha  = @($Itens | Where-Object { $_.TaskID -eq $f.TaskID })
        $contados = @($daFicha | Where-Object { $null -ne $_.Esperado })
        $infos    = @($daFicha | Where-Object { $null -eq $_.Esperado })

        $badge = if ($f.Divergencias -eq 0) { '<span class="badge ok">Conferido</span>' }
                 else { "<span class=""badge alerta"">$($f.Divergencias) divergência(s)</span>" }

        [void]$sb.AppendLine("<section class=""card"" data-div=""$($f.Divergencias)"">")
        [void]$sb.AppendLine('<header><div>')
        [void]$sb.AppendLine("<div class=""tit"">Ficha $(Protect-Html $f.Ficha) &middot; $(Protect-Html $f.Cliente)</div>")
        [void]$sb.AppendLine("<div class=""meta"">$(Format-DataBr $f.Data) &middot; $(Protect-Html $f.Responsavel) &middot; $(Protect-Html $f.Servico) &middot; esperado $($f.TotalEsperado) / realizado $($f.TotalRealizado)</div>")
        [void]$sb.AppendLine("</div>$badge</header>")

        if ($contados.Count -gt 0) {
            [void]$sb.AppendLine('<table><thead><tr><th>Setor</th><th class="num">Esperado</th><th class="num">Realizado</th><th class="num">Dif.</th><th>Situação</th><th>Observação</th></tr></thead><tbody>')
            foreach ($i in $contados) {
                $cls = switch ($i.Situacao) {
                    "FALTA"        { "falta" }
                    "EXCEDENTE"    { "exc" }
                    "SEM RESPOSTA" { "sem" }
                    default        { "ok" }
                }
                $real = if ($null -eq $i.Realizado) { "-" } else { $i.Realizado }
                $dif  = if ($null -eq $i.Diferenca) { "-" } else { $i.Diferenca }
                [void]$sb.AppendLine("<tr class=""$cls"" data-sit=""$($i.Situacao)"">")
                [void]$sb.AppendLine("<td>$(Protect-Html $i.Setor)</td><td class=""num"">$($i.Esperado)</td><td class=""num"">$real</td><td class=""num"">$dif</td>")
                [void]$sb.AppendLine("<td><span class=""sit $cls"">$($i.Situacao)</span></td><td class=""obs"">$(Protect-Html $i.Observacao)</td></tr>")
            }
            [void]$sb.AppendLine('</tbody></table>')
        }

        $infosUteis = @($infos | Where-Object { $_.RespostaBruta -and $_.RespostaBruta.Trim() -notmatch '^\.*$' })
        if ($infosUteis.Count -gt 0) {
            $partes = $infosUteis | ForEach-Object { "<b>$(Protect-Html $_.Setor):</b> $(Protect-Html $_.RespostaBruta)" }
            [void]$sb.AppendLine("<div class=""info"">$($partes -join ' &nbsp;&middot;&nbsp; ')</div>")
        }

        [void]$sb.AppendLine('</section>')
    }

    if ($Resumo.Count -eq 0) {
        [void]$sb.AppendLine('<div class="vazio">Nenhuma ficha encontrada no período.</div>')
    }

    [void]$sb.AppendLine('</div><script>')
    [void]$sb.AppendLine('document.getElementById("soDiv").addEventListener("change",function(e){var s=e.target.checked;')
    [void]$sb.AppendLine('document.querySelectorAll("tr[data-sit]").forEach(function(t){t.style.display=(s&&t.dataset.sit==="OK")?"none":"";});')
    [void]$sb.AppendLine('document.querySelectorAll(".card").forEach(function(c){c.style.display=(s&&c.dataset.div==="0")?"none":"";});});')
    [void]$sb.AppendLine('</script></body></html>')

    [System.IO.File]::WriteAllText($Caminho, $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))
}

# ---------------------------------------------------------------- execucao ---

Write-Host "Buscando tarefas de $StartDate até $EndDate..." -ForegroundColor Cyan
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

    $cor     = if ($divergencias.Count -eq 0) { "DarkGray" } else { "Yellow" }
    $marca   = if ($divergencias.Count -eq 0) { "  " } else { "! " }
    $cliente = [string]$detalhe.customerDescription
    if ($cliente.Length -gt 38) { $cliente = $cliente.Substring(0, 37) + "." }

    Write-Host ("{0}Ficha {1,-7} {2,-38} esp {3,4}  real {4,4}  dif {5,4}" -f `
        $marca, $detalhe.externalId, $cliente, $totalEsperado, $totalRealizado, ($totalRealizado - $totalEsperado)) -ForegroundColor $cor
}

if ($todosItens.Count -eq 0) {
    Write-Host "`nNenhum item de questionário encontrado no período." -ForegroundColor Yellow
    return
}

if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }
# .NET resolve caminho relativo pelo diretorio do processo, que pode diferir do $PWD do PowerShell.
$OutputDir = (Resolve-Path -LiteralPath $OutputDir).Path
$carimbo       = Get-Date -Format "yyyyMMdd-HHmmss"
$arquivoItens  = Join-Path $OutputDir "itens-$carimbo.csv"
$arquivoResumo = Join-Path $OutputDir "resumo-$carimbo.csv"
$arquivoHtml   = Join-Path $OutputDir "relatorio-$carimbo.html"

$todosItens | Export-Csv -Path $arquivoItens  -NoTypeInformation -Encoding UTF8 -Delimiter ";"
$resumo     | Export-Csv -Path $arquivoResumo -NoTypeInformation -Encoding UTF8 -Delimiter ";"
New-RelatorioHtml -Resumo $resumo -Itens $todosItens -StartDate $StartDate -EndDate $EndDate -Caminho $arquivoHtml

$geralEsperado  = ($resumo | Measure-Object -Property TotalEsperado  -Sum).Sum
$geralRealizado = ($resumo | Measure-Object -Property TotalRealizado -Sum).Sum
$fichasComProblema = @($resumo | Where-Object { $_.Divergencias -gt 0 })

Write-Host ""
Write-Host ("TOTAL: {0} ficha(s) | esperado {1} | realizado {2} | diferença {3} | a verificar {4}" -f `
    $resumo.Count, $geralEsperado, $geralRealizado, ($geralRealizado - $geralEsperado), $fichasComProblema.Count) `
    -ForegroundColor $(if ($fichasComProblema.Count -eq 0) { "Green" } else { "Yellow" })
Write-Host ""
Write-Host "Relatório: $arquivoHtml" -ForegroundColor Cyan
Write-Host "CSVs     : $arquivoItens / $arquivoResumo" -ForegroundColor DarkGray

if (-not $NoBrowser) { Invoke-Item $arquivoHtml }
