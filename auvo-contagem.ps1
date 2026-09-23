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
    [string]$OutputDir  = "",
    [switch]$NoBrowser
)

$ErrorActionPreference = "Stop"
$BaseUrl = "https://api.auvo.com.br/v2"

# No PowerShell 5.1 o $PSScriptRoot ainda esta vazio quando o bloco param e avaliado
# (acontece ao rodar via "powershell -File"), entao a pasta e resolvida aqui.
$RaizScript = $PSScriptRoot
if (-not $RaizScript) { $RaizScript = Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not $RaizScript) { $RaizScript = (Get-Location).Path }
if (-not $OutputDir)  { $OutputDir  = Join-Path $RaizScript "saida" }

$script:Headers     = $null
$script:TokenExpira = [datetime]::MinValue

function Get-AuvoCredencial {
    $key   = $env:AUVO_API_KEY
    $token = $env:AUVO_API_TOKEN

    if (-not $key -or -not $token) {
        $cfg = Join-Path $RaizScript "config.local.json"
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
            $setor    = $desc
            if ($desc -match '\sI\s+(\d+)\s*(\S*)\s*$') {
                $esperado = [int]$Matches[1]
                $unidade  = $Matches[2]
                $setor    = ($desc -replace '\sI\s+\d+\s*\S*\s*$', '').Trim()
            }

            # Quantidade realizada: numero no inicio da resposta.
            $realizado = $null
            if ($reply -match '^\s*(\d+)') { $realizado = [int]$Matches[1] }

            # Tudo que o tecnico escreveu depois do numero e a justificativa.
            $observacao = ""
            if ($reply -match '^\s*\d+\b\s*(.+)$') { $observacao = $Matches[1] }
            $observacao = $observacao.Trim("() -:.".ToCharArray())
            if ($unidade -and $observacao) {
                $u = [regex]::Escape($unidade.TrimEnd('s', 'S'))
                $observacao = [regex]::Replace($observacao, "^$u" + "s?\b[\s,;:.-]*", "", 'IgnoreCase').Trim()
            }

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
                Setor        = $setor
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

    # No relatorio, quem precisa de atencao vem primeiro; o resto segue por data.
    $Resumo = @($Resumo | Sort-Object -Property @{Expression={ [int]($_.Divergencias -eq 0) }}, Data, Ficha)

    $gEsperado  = ($Resumo | Measure-Object -Property TotalEsperado  -Sum).Sum
    $gRealizado = ($Resumo | Measure-Object -Property TotalRealizado -Sum).Sum
    if (-not $gEsperado)  { $gEsperado  = 0 }
    if (-not $gRealizado) { $gRealizado = 0 }
    $gDif       = $gRealizado - $gEsperado
    $aVerificar = @($Resumo | Where-Object { $_.Divergencias -gt 0 }).Count

    $faltaQtd = 0
    $semQtd   = 0
    foreach ($i in $Itens) {
        if     ($i.Situacao -eq "FALTA")        { $faltaQtd += ($i.Esperado - $i.Realizado) }
        elseif ($i.Situacao -eq "SEM RESPOSTA") { $semQtd   += $i.Esperado }
    }
    $detalhe = @()
    if ($faltaQtd -gt 0) { $detalhe += "$faltaQtd não executada$(if ($faltaQtd -gt 1) { 's' })" }
    if ($semQtd   -gt 0) { $detalhe += "$semQtd não registrada$(if ($semQtd -gt 1) { 's' })" }

    $classeDif = if ($gDif -eq 0) { "ok" } else { "falta" }
    $classeVer = if ($aVerificar -eq 0) { "ok" } else { "falta" }

    $css = @"
:root{
--plane:#f7f7f5;--surface:#ffffff;--ink:#0b0b0b;--ink2:#52514e;--muted:#898781;
--hair:#e1e0d9;--hair2:#f0efe9;
--good:#006300;--good-mark:#0ca30c;--good-bg:#e9f5e9;
--crit:#b3261e;--crit-mark:#d03b3b;--crit-bg:#fdecea;
--warn:#8a5a00;--warn-mark:#fab219;--warn-bg:#fdf3e2;
--neutro-bg:#f2f2ef;--r:12px}
*{box-sizing:border-box}
body{margin:0;padding:36px 18px 56px;background:var(--plane);color:var(--ink);font-size:14px;line-height:1.45;
-webkit-font-smoothing:antialiased;font-family:ui-sans-serif,-apple-system,'Segoe UI',Roboto,'Helvetica Neue',Arial,sans-serif}
.wrap{max-width:1100px;margin:0 auto}
.top{display:flex;align-items:flex-end;justify-content:space-between;gap:16px;flex-wrap:wrap;margin-bottom:22px}
h1{font-size:21px;font-weight:650;margin:0;display:flex;align-items:center;gap:10px;letter-spacing:-.015em}
h1::before{content:"";width:4px;height:20px;border-radius:2px;background:linear-gradient(180deg,#0ca30c,#006300)}
.sub{color:var(--muted);font-size:12.5px;margin-top:5px}
.gerado{color:var(--muted);font-size:12px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(158px,1fr));gap:12px;margin-bottom:20px}
.kpi{background:var(--surface);border:1px solid var(--hair);border-radius:var(--r);padding:14px 16px;position:relative;overflow:hidden}
.kpi .rot{font-size:10.5px;color:var(--muted);text-transform:uppercase;letter-spacing:.07em;font-weight:600}
.kpi .val{font-size:27px;font-weight:650;margin-top:5px;letter-spacing:-.02em;font-variant-numeric:tabular-nums}
.kpi .val.ok{color:var(--good)}
.kpi .val.falta{color:var(--crit)}
.kpi .det{font-size:11px;color:var(--muted);margin-top:5px;line-height:1.35}
.kpi.alerta::before{content:"";position:absolute;inset:0 0 auto 0;height:3px;background:var(--crit-mark)}
.filtro{display:inline-flex;align-items:center;gap:7px;margin-bottom:14px;font-size:13px;color:var(--ink2);cursor:pointer;
background:var(--surface);border:1px solid var(--hair);border-radius:999px;padding:7px 13px}
.filtro:hover{border-color:var(--muted)}
.filtro input{margin:0;cursor:pointer}
.card{background:var(--surface);border:1px solid var(--hair);border-radius:var(--r);margin-bottom:14px;overflow:hidden;
box-shadow:0 1px 2px rgba(16,24,40,.03),0 4px 10px -6px rgba(16,24,40,.08)}
.card>header{display:flex;justify-content:space-between;align-items:flex-start;gap:14px;padding:15px 17px;border-bottom:1px solid var(--hair2);flex-wrap:wrap}
.tit{font-weight:650;font-size:14.5px;letter-spacing:-.01em}
.tit a{color:inherit;text-decoration:none;border-bottom:1px solid var(--hair)}
.tit a:hover{border-bottom-color:var(--muted)}
.meta{color:var(--muted);font-size:12.5px;margin-top:3px}
.meter{display:flex;align-items:center;gap:9px;margin-top:10px;max-width:330px}
.meter-track{flex:1;height:6px;border-radius:999px;background:#f2cecb;overflow:hidden}
.meter-track.cheio{background:var(--hair)}
.meter-fill{height:100%;border-radius:999px;background:var(--good-mark)}
.meter-val{font-size:11.5px;color:var(--muted);font-variant-numeric:tabular-nums;min-width:36px}
.badge{font-size:11.5px;font-weight:650;padding:5px 11px;border-radius:999px;white-space:nowrap}
.badge.ok{background:var(--good-bg);color:var(--good)}
.badge.alerta{background:var(--crit-bg);color:var(--crit)}
table{width:100%;border-collapse:collapse;font-size:13.5px}
th{text-align:left;font-size:10.5px;text-transform:uppercase;letter-spacing:.07em;color:var(--muted);font-weight:600;padding:10px 17px;border-bottom:1px solid var(--hair)}
td{padding:9px 17px;border-bottom:1px solid var(--hair2);vertical-align:middle}
tbody tr:last-child td{border-bottom:none}
th.num,td.num{text-align:right;width:92px;font-variant-numeric:tabular-nums}
tbody tr{transition:background .12s ease}
tbody tr:hover{background:#fbfbf9}
tr.falta{background:#fdf6f5}
tr.falta:hover{background:#fbeeeb}
tr.falta td:first-child{box-shadow:inset 3px 0 0 var(--crit-mark)}
tr.exc{background:#fdf7ec}
tr.exc:hover{background:#fbf2e2}
tr.sem{background:#f8f8f6}
tr.sem:hover{background:#f2f2ef}
tr.sem td:first-child{box-shadow:inset 3px 0 0 #c9c9c1}
.pill{display:inline-flex;align-items:center;gap:5px;font-size:11.5px;font-weight:650;padding:3px 9px;border-radius:999px;white-space:nowrap}
.pill.ok{background:var(--good-bg);color:var(--good)}
.pill.falta{background:var(--crit-bg);color:var(--crit)}
.pill.exc{background:var(--warn-bg);color:var(--warn)}
.pill.sem{background:var(--neutro-bg);color:var(--ink2)}
.pill .gl{font-size:9px;line-height:1}
.ok-mini{color:#9a9a92;font-size:13px}
.obs{color:var(--ink2);font-size:12.5px}
.semjust{color:var(--warn);font-style:italic}
tfoot td{font-weight:650;background:#fafaf8;border-top:1px solid var(--hair);border-bottom:none;padding-top:11px;padding-bottom:11px}
tfoot td.rot{color:var(--muted);font-size:10.5px;text-transform:uppercase;letter-spacing:.07em}
tfoot td.neg{color:var(--crit)}
td.neg{color:var(--crit)}
details.motivos>summary{display:flex;align-items:center;gap:10px;padding:14px 17px;cursor:pointer;list-style:none;-webkit-user-select:none;user-select:none}
details.motivos>summary::-webkit-details-marker{display:none}
details.motivos>summary::before{content:"";width:0;height:0;border-left:5px solid var(--muted);border-top:4px solid transparent;border-bottom:4px solid transparent;transition:transform .18s ease}
details.motivos[open]>summary::before{transform:rotate(90deg)}
details.motivos>summary:hover{background:#fbfbf9}
details.motivos[open]>summary{border-bottom:1px solid var(--hair2)}
details.motivos .cont{background:var(--crit-bg);color:var(--crit);font-size:11px;font-weight:650;padding:2px 9px;border-radius:999px}
details.motivos .dica{margin-left:auto;font-size:11.5px;color:var(--muted);font-weight:400}
details.motivos .dica::after{content:"clique para abrir"}
details.motivos[open] .dica::after{content:"clique para fechar"}
.motivos td.fic{white-space:nowrap;color:var(--muted);font-size:12.5px}
.info{padding:11px 17px;border-top:1px solid var(--hair2);background:#fafaf8;font-size:12.5px;color:var(--ink2)}
.info b{color:var(--ink);font-weight:650}
.vazio{background:var(--surface);border:1px solid var(--hair);border-radius:var(--r);padding:38px;text-align:center;color:var(--muted)}
@media print{body{background:#fff;padding:0}.filtro{display:none}.card{break-inside:avoid;box-shadow:none}}
@media(max-width:640px){th,td{padding-left:12px;padding-right:12px}}
"@

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<!DOCTYPE html>')
    [void]$sb.AppendLine('<html lang="pt-BR"><head><meta charset="utf-8">')
    [void]$sb.AppendLine('<meta name="viewport" content="width=device-width,initial-scale=1">')
    [void]$sb.AppendLine('<title>Contagem Auvo</title>')
    [void]$sb.AppendLine("<style>$css</style></head><body><div class=""wrap"">")

    $periodo = if ($StartDate -eq $EndDate) { Format-DataBr $StartDate } else { "$(Format-DataBr $StartDate) a $(Format-DataBr $EndDate)" }
    [void]$sb.AppendLine('<div class="top"><div>')
    [void]$sb.AppendLine("<h1>Contagem de execução</h1>")
    [void]$sb.AppendLine("<div class=""sub"">Período: $periodo</div>")
    [void]$sb.AppendLine("</div><div class=""gerado"">gerado em $(Get-Date -Format 'dd/MM/yyyy HH:mm')</div></div>")

    [void]$sb.AppendLine('<div class="kpis">')
    [void]$sb.AppendLine("<div class=""kpi""><div class=""rot"">Fichas</div><div class=""val"">$($Resumo.Count)</div></div>")
    [void]$sb.AppendLine("<div class=""kpi""><div class=""rot"">Esperado</div><div class=""val"">$gEsperado</div></div>")
    [void]$sb.AppendLine("<div class=""kpi""><div class=""rot"">Realizado</div><div class=""val"">$gRealizado</div></div>")
    $linhaDet = if ($detalhe.Count -gt 0) { "<div class=""det"">$($detalhe -join '<br>')</div>" } else { "" }
    [void]$sb.AppendLine("<div class=""kpi $(if ($gDif -ne 0) { ""alerta"" })""><div class=""rot"">Diferença</div><div class=""val $classeDif"">$gDif</div>$linhaDet</div>")
    [void]$sb.AppendLine("<div class=""kpi $(if ($aVerificar -gt 0) { ""alerta"" })""><div class=""rot"">A verificar</div><div class=""val $classeVer"">$aVerificar</div></div>")
    [void]$sb.AppendLine('</div>')

    $porCliente = @($Resumo | Group-Object Cliente | ForEach-Object {
        $g = $_.Group
        [pscustomobject]@{
            Cliente      = $_.Name
            Fichas       = $_.Count
            Esperado     = ($g | Measure-Object -Property TotalEsperado  -Sum).Sum
            Realizado    = ($g | Measure-Object -Property TotalRealizado -Sum).Sum
            Divergencias = ($g | Measure-Object -Property Divergencias   -Sum).Sum
        }
    } | Sort-Object @{Expression={ $_.Realizado - $_.Esperado }}, Cliente)

    if ($porCliente.Count -gt 1) {
        [void]$sb.AppendLine('<section class="card"><header><div class="tit">Por cliente</div></header>')
        [void]$sb.AppendLine('<table><thead><tr><th>Cliente</th><th class="num">Fichas</th><th class="num">Esperado</th><th class="num">Realizado</th><th class="num">Dif.</th><th class="num">Diverg.</th></tr></thead><tbody>')
        foreach ($c in $porCliente) {
            $dif = $c.Realizado - $c.Esperado
            $clsDif = if ($dif -lt 0) { "num neg" } else { "num" }
            [void]$sb.AppendLine("<tr><td>$(Protect-Html $c.Cliente)</td><td class=""num"">$($c.Fichas)</td><td class=""num"">$($c.Esperado)</td><td class=""num"">$($c.Realizado)</td><td class=""$clsDif"">$(if ($dif -ne 0) { $dif })</td><td class=""num"">$(if ($c.Divergencias -gt 0) { $c.Divergencias })</td></tr>")
        }
        [void]$sb.AppendLine('</tbody></table></section>')
    }

    [void]$sb.AppendLine('<label class="filtro"><input type="checkbox" id="soDiv"> Mostrar apenas as divergências</label>')

    $divs = @($Itens | Where-Object { $null -ne $_.Esperado -and $_.Situacao -ne "OK" })
    if ($divs.Count -gt 0) {
        [void]$sb.AppendLine("<details class=""card motivos""><summary><span class=""tit"">Por que houve diferença</span><span class=""cont"">$($divs.Count)</span><span class=""dica""></span></summary><table><tbody>")
        foreach ($d in $divs) {
            $motivo = if ($d.Observacao) { Protect-Html $d.Observacao } else { '<span class="semjust">sem justificativa informada</span>' }
            [void]$sb.AppendLine("<tr><td class=""fic"">Ficha $(Protect-Html $d.Ficha)</td><td>$(Protect-Html $d.Setor)</td><td class=""num"">$($d.Diferenca)</td><td>$motivo</td></tr>")
        }
        [void]$sb.AppendLine('</tbody></table></details>')
    }

    foreach ($f in $Resumo) {
        $daFicha  = @($Itens | Where-Object { $_.TaskID -eq $f.TaskID })
        $contados = @($daFicha | Where-Object { $null -ne $_.Esperado })
        $infos    = @($daFicha | Where-Object { $null -eq $_.Esperado })

        $badge = if ($f.Divergencias -eq 0) { '<span class="badge ok">Conferido</span>' }
                 elseif ($f.Divergencias -eq 1) { '<span class="badge alerta">1 divergência</span>' }
                 else { "<span class=""badge alerta"">$($f.Divergencias) divergências</span>" }

        [void]$sb.AppendLine("<section class=""card"" data-div=""$($f.Divergencias)"">")
        [void]$sb.AppendLine('<header><div>')
        $titulo = "Ficha $(Protect-Html $f.Ficha) &middot; $(Protect-Html $f.Cliente)"
        if ($f.TaskUrl) { $titulo = "<a href=""$(Protect-Html $f.TaskUrl)"" target=""_blank"" rel=""noopener"" title=""Abrir a tarefa no Auvo"">$titulo</a>" }
        [void]$sb.AppendLine("<div class=""tit"">$titulo</div>")
        [void]$sb.AppendLine("<div class=""meta"">$(Format-DataBr $f.Data) &middot; $(Protect-Html $f.Responsavel) &middot; $(Protect-Html $f.Servico)</div>")
        $pct = 0
        if ($f.TotalEsperado -gt 0) { $pct = [math]::Round(100 * $f.TotalRealizado / $f.TotalEsperado) }
        $larg = [math]::Min($pct, 100)
        $trilha = if ($larg -ge 100) { "cheio" } else { "" }
        [void]$sb.AppendLine("<div class=""meter""><div class=""meter-track $trilha""><div class=""meter-fill"" style=""width:$larg%""></div></div><span class=""meter-val"">$pct%</span></div>")
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
                $dif  = if ($null -eq $i.Diferenca) { "-" } elseif ($i.Diferenca -eq 0) { "" } else { $i.Diferenca }
                [void]$sb.AppendLine("<tr class=""$cls"" data-sit=""$($i.Situacao)"">")
                [void]$sb.AppendLine("<td>$(Protect-Html $i.Setor)</td><td class=""num"">$($i.Esperado)</td><td class=""num"">$real</td><td class=""num"">$dif</td>")
                $obsCel = if ($i.Observacao)        { Protect-Html $i.Observacao }
                          elseif ($i.Situacao -ne "OK") { '<span class="semjust">sem justificativa</span>' }
                          else                          { "" }
                $glifo = switch ($i.Situacao) {
                    "FALTA"        { "&#9660;" }
                    "EXCEDENTE"    { "&#9650;" }
                    "SEM RESPOSTA" { "&#8211;" }
                    default        { "&#10003;" }
                }
                $celSit = if ($i.Situacao -eq "OK") { '<span class="ok-mini">&#10003;</span>' }
                          else { "<span class=""pill $cls""><span class=""gl"">$glifo</span>$($i.Situacao)</span>" }
                [void]$sb.AppendLine("<td>$celSit</td><td class=""obs"">$obsCel</td></tr>")
            }
            [void]$sb.AppendLine('</tbody><tfoot><tr>')
            [void]$sb.AppendLine("<td class=""rot"">Total da ficha &middot; $($contados.Count) setores</td>")
            [void]$sb.AppendLine("<td class=""num"">$($f.TotalEsperado)</td><td class=""num"">$($f.TotalRealizado)</td><td class=""num $(if ($f.Diferenca -ne 0) { ""neg"" })"">$($f.Diferenca)</td>")
            [void]$sb.AppendLine('<td colspan="2"></td></tr></tfoot></table>')
        }

        $partes = @()
        foreach ($i in @($infos | Where-Object { $_.RespostaBruta -and $_.RespostaBruta.Trim() -notmatch '^\.*$' })) {
            $partes += "<b>$(Protect-Html $i.Setor):</b> $(Protect-Html $i.RespostaBruta)"
        }
        if ($f.Pendencia) { $partes += "<b>Pendência:</b> $(Protect-Html $f.Pendencia)" }
        if ($f.Relatorio) { $partes += "<b>Relato do técnico:</b> $(Protect-Html $f.Relatorio)" }
        if ($partes.Count -gt 0) {
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
        TaskUrl        = $detalhe.taskUrl
        Pendencia      = $detalhe.pendency
        Relatorio      = $detalhe.report
        Duracao        = $detalhe.duration
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
