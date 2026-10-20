# Testes do consulta-ad.ps1 — usa o stub de ActiveDirectory em .\stub
$ErrorActionPreference = 'Stop'

$testDir = $PSScriptRoot
$stubRoot = Join-Path $testDir 'stub'
$scriptPath = Join-Path (Split-Path -Path $PSScriptRoot -Parent) 'consulta-ad.ps1'
$entradaBom = Join-Path $testDir 'entrada_utf8bom.csv'
$entradaAnsi = Join-Path $testDir 'entrada_ansi.csv'
$entradaUtf8 = Join-Path $testDir 'entrada_utf8.csv'

$env:PSModulePath = "$stubRoot;$env:PSModulePath"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$pwshExe = (Get-Process -Id $PID).Path
$ps51Exe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

$script:passou = 0
$script:falhou = 0

function Ok {
    param([string]$Msg)
    $script:passou++
    Write-Host "  [OK]   $Msg" -ForegroundColor Green
}

function Falha {
    param([string]$Msg)
    $script:falhou++
    Write-Host "  [FALHA] $Msg" -ForegroundColor Red
}

function Assert-True {
    param([bool]$Cond, [string]$Msg)
    if ($Cond) { Ok $Msg } else { Falha $Msg }
}

function Assert-Eq {
    param($Actual, $Expected, [string]$Msg)
    if ("$Actual" -eq "$Expected") { Ok $Msg }
    else { Falha "$Msg (esperado: '$Expected'; obtido: '$Actual')" }
}

function Invoke-Child {
    param([string]$Exe, [string]$Comando)
    $saida = & $Exe -NoProfile -Command "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8; $Comando" 2>&1
    $codigo = $LASTEXITCODE
    $texto = ($saida | ForEach-Object { "$_" }) -join "`n"
    return [PSCustomObject]@{ Codigo = $codigo; Saida = $texto }
}

function Ler-JSON {
    param([string]$Dir)
    $f = Get-ChildItem -LiteralPath $Dir -Filter '*.json' | Select-Object -First 1
    if (-not $f) { Falha "JSON nao gerado em $Dir"; return $null }
    return (Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8 | ConvertFrom-Json)
}

# ---- limpar saidas antigas ----
Get-ChildItem -LiteralPath $testDir -Filter 'out_*' -ErrorAction SilentlyContinue |
    Remove-Item -Recurse -Force

# ---- gerar CSVs nos 3 encodings ----
$texto = @(
    'Nome;E-mail;SamAccountName'
    'Jose Silva;jose.silva@corp.local;jsilva'
    "Marco D'Ávila;;"
    'Maria Duplicada;;'
    'João Desativado;joao@corp.local;'
    'Nao Existe;;naoexiste'
    ';;;'
) -join "`r`n"

[System.IO.File]::WriteAllText($entradaBom, $texto, (New-Object System.Text.UTF8Encoding($true)))
[System.IO.File]::WriteAllText($entradaUtf8, $texto, (New-Object System.Text.UTF8Encoding($false)))
[System.IO.File]::WriteAllText($entradaAnsi, $texto, [System.Text.Encoding]::GetEncoding(1252))

# ---- S0: stub vence o AD real ----
Write-Host ""
Write-Host '== S0: modulo AD carregado == ' -ForegroundColor Cyan
$r = Invoke-Child $pwshExe "Import-Module ActiveDirectory; (Get-Module ActiveDirectory).Path"
Assert-True ($r.Saida -like '*stub*') "modulo AD vem do stub (nao do sistema)"

# ---- validador de execucao basica ----
function Test-Basico {
    param([string]$Label, [string]$Exe, [string]$Entrada, [string]$Tag, [string]$EncEsperado)

    Write-Host ""
    Write-Host "== $Label ==" -ForegroundColor Cyan

    $saida = Join-Path $testDir "out_$Tag"
    $r = Invoke-Child $Exe "& '$scriptPath' -Arquivo '$Entrada' -PastaSaida '$saida'"

    Assert-Eq $r.Codigo 0 "$Tag : exit 0"
    Assert-True ($r.Saida -like '*PROCESSAMENTO CONCLUÍDO*') "$Tag : resumo final exibido"
    Assert-True ($r.Saida -like "*$EncEsperado*") "$Tag : encoding detectado ($EncEsperado)"
    Assert-True ($r.Saida -like '*Linhas ignoradas sem Nome, E-mail e SamAccountName: 1*') "$Tag : linha vazia ignorada"
    Assert-True ($r.Saida -like '*NÃO ENCONTRADO - Nao Existe*') "$Tag : acentos no console (NAO ENCONTRADO)"

    $j = Ler-JSON $saida
    if (-not $j) { return }

    Assert-Eq $j.Total 5 "$Tag : Total = 5"
    Assert-Eq (@($j.Usuarios).Count) 5 "$Tag : Usuarios = 5"

    $u = @($j.Usuarios)
    Assert-Eq $u[0].Status 'Encontrado' "$Tag : u0 status"
    Assert-Eq $u[0].EncontradoPor 'SamAccountName' "$Tag : u0 por sam"
    Assert-Eq $u[0].Divergencia 'Nome' "$Tag : u0 divergencia Nome"
    Assert-Eq $u[0].Grupos 'TI' "$Tag : u0 grupos diretos"
    Assert-Eq $u[1].Status 'Encontrado' "$Tag : u1 status"
    Assert-Eq $u[1].EncontradoPor 'Nome' "$Tag : u1 por nome (apostrofo)"
    Assert-Eq $u[2].Status 'Duplicado' "$Tag : u2 status"
    Assert-Eq $u[2].Duplicados 'mdup1; mdup2' "$Tag : u2 candidatos"
    Assert-Eq $u[3].Status 'Encontrado' "$Tag : u3 status"
    Assert-Eq $u[3].EncontradoPor 'E-mail' "$Tag : u3 por email"
    Assert-Eq $u[3].Ativo $false "$Tag : u3 inativo"
    Assert-Eq $u[3].UltimoLogon 'Nunca' "$Tag : u3 nunca logou"
    Assert-Eq $u[3].OU_Caminho 'Arquivo' "$Tag : u3 caminho OU"
    Assert-Eq $u[4].Status 'Não Encontrado' "$Tag : u4 nao encontrado"

    $c = $null
    $cf = Get-ChildItem -LiteralPath $saida -Filter '*.csv' | Select-Object -First 1
    if ($cf) { $c = @(Import-Csv -LiteralPath $cf.FullName -Delimiter ';') } else { Falha "$Tag : CSV nao gerado" }
    if ($c) {
        Assert-Eq $c.Count 5 "$Tag : CSV = 5 linhas"
        Assert-Eq $c[0].Grupos 'TI' "$Tag : CSV grupos"
    }

    $h = $null
    $hf = Get-ChildItem -LiteralPath $saida -Filter '*.html' | Select-Object -First 1
    if ($hf) { $h = Get-Content -LiteralPath $hf.FullName -Raw -Encoding UTF8 } else { Falha "$Tag : HTML nao gerado" }
    if ($h) {
        Assert-True ($h.Contains('&#8226; Users')) "$Tag : arvore no sintetico CN=Users"
        Assert-Eq ([regex]::Matches($h, '&#8226;').Count) 6 "$Tag : arvore 6 nos"
        Assert-True ($h.Contains('Candidatos: mdup1; mdup2')) "$Tag : HTML mostra duplicados"
        Assert-Eq ([regex]::Matches($h, '<tr[\s>]').Count) 6 "$Tag : HTML 6 linhas de tabela"
        Assert-True ($h.Contains('id="ouTree"')) "$Tag : arvore dentro do container ouTree"
        Assert-True ($h.Contains("data-dn='OU=TI,OU=Funcionarios,DC=corp,DC=local'")) "$Tag : no com data-dn"
        Assert-True ($h.Contains('data-ou="OU=TI,OU=Funcionarios,DC=corp,DC=local"')) "$Tag : linha com data-ou"
        Assert-True ($h.Contains('tree-toggle')) "$Tag : botoes recolher/expandir"
        Assert-True ($h.Contains('selecionarOU')) "$Tag : JS de filtro por OU"
    }
}

# ---- S1-S3: pwsh nos 3 encodings ----
Test-Basico -Label 'S1: pwsh UTF-8 com BOM' -Exe $pwshExe -Entrada $entradaBom -Tag 'pwsh_bom' -EncEsperado 'UTF-8'
Test-Basico -Label 'S2: pwsh ANSI cp1252' -Exe $pwshExe -Entrada $entradaAnsi -Tag 'pwsh_ansi' -EncEsperado 'ANSI (cp 1252)'
Test-Basico -Label 'S3: pwsh UTF-8 sem BOM' -Exe $pwshExe -Entrada $entradaUtf8 -Tag 'pwsh_utf8' -EncEsperado 'UTF-8'

# ---- S7: PowerShell 5.1 ----
Test-Basico -Label 'S7a: PS5.1 UTF-8 com BOM' -Exe $ps51Exe -Entrada $entradaBom -Tag 'ps51_bom' -EncEsperado 'UTF-8'
Test-Basico -Label 'S7b: PS5.1 ANSI cp1252' -Exe $ps51Exe -Entrada $entradaAnsi -Tag 'ps51_ansi' -EncEsperado 'ANSI (cp 1252)'

# ---- funcao auxiliar para os cenarios S4 ----
function Teste-S4 {
    param([string]$Tag, [string]$Extras, [hashtable]$Esperados)

    $saida = Join-Path $testDir "out_$Tag"
    $r = Invoke-Child $pwshExe "& '$scriptPath' -Arquivo '$entradaBom' -PastaSaida '$saida' $Extras"
    Assert-Eq $r.Codigo 0 "$Tag : exit 0"

    $j = Ler-JSON $saida
    if (-not $j) { return }

    $u = @($j.Usuarios)

    if ($Esperados.ContainsKey('u0_OU_OK')) { Assert-Eq $u[0].OU_OK $Esperados['u0_OU_OK'] "$Tag : u0 OU_OK" }
    if ($Esperados.ContainsKey('u1_OU_OK')) { Assert-Eq $u[1].OU_OK $Esperados['u1_OU_OK'] "$Tag : u1 OU_OK" }
    if ($Esperados.ContainsKey('u0_Grupos')) { Assert-Eq $u[0].Grupos $Esperados['u0_Grupos'] "$Tag : u0 Grupos" }
    if ($Esperados.ContainsKey('u0_Faltando')) { Assert-Eq $u[0].GruposFaltando $Esperados['u0_Faltando'] "$Tag : u0 GruposFaltando" }
    if ($Esperados.ContainsKey('u3_OU_OK')) { Assert-Eq $u[3].OU_OK $Esperados['u3_OU_OK'] "$Tag : u3 OU_OK" }
}

Write-Host ""
Write-Host '== S4: -OUEsperada / -Grupo / -Aninhados ==' -ForegroundColor Cyan

Teste-S4 -Tag 's4a_nome' `
    "-OUEsperada 'Funcionarios' -Grupo 'VPN-Remoto','Inexistente'" `
    @{ u0_OU_OK = 'Sim'; u1_OU_OK = 'Não'; u0_Grupos = 'TI'; u0_Faltando = 'VPN-Remoto; Inexistente' }

Teste-S4 -Tag 's4b_substring' `
    "-OUEsperada 'ncional'" `
    @{ u0_OU_OK = 'Não' }

Teste-S4 -Tag 's4c_dn' `
    "-OUEsperada 'OU=Funcionarios,DC=corp,DC=local'" `
    @{ u0_OU_OK = 'Sim'; u1_OU_OK = 'Não' }

Teste-S4 -Tag 's4d_aninhados' `
    "-Aninhados -Grupo 'Inexistente'" `
    @{ u0_Grupos = 'TI; Todos; VPN-Remoto'; u0_Faltando = 'Inexistente' }

# ---- S5: -EnviarEmail sem parametros nao pode abortar ----
Write-Host ""
Write-Host '== S5: -EnviarEmail sem -SmtpServer ==' -ForegroundColor Cyan

$saida5 = Join-Path $testDir 'out_s5'
$r = Invoke-Child $pwshExe "& '$scriptPath' -Arquivo '$entradaBom' -PastaSaida '$saida5' -EnviarEmail"
Assert-True ($r.Saida -like '*Para enviar e-mail, informe -SmtpServer*') "S5 : mensagem de parametrizacao exibida"
Assert-True ($r.Saida -like '*PROCESSAMENTO CONCLUÍDO*') "S5 : script nao abortou (resumo exibido)"
$j5 = Ler-JSON $saida5
if ($j5) { Assert-Eq $j5.Total 5 "S5 : JSON ainda gerado" }

# ---- S6: arquivo inexistente => exit 1 ----
Write-Host ""
Write-Host '== S6: arquivo inexistente ==' -ForegroundColor Cyan

$r = Invoke-Child $pwshExe "& '$scriptPath' -Arquivo '$testDir\inexistente.csv' -PastaSaida '$testDir\out_s6'"
Assert-Eq $r.Codigo 1 "S6 : exit 1"
Assert-True ($r.Saida -like '*Arquivo não encontrado*') "S6 : mensagem de erro"

# ---- S8: -Inventario (usuarios e computadores de todo o AD) ----
Write-Host ""
Write-Host '== S8: -Inventario ==' -ForegroundColor Cyan

$saida8 = Join-Path $testDir 'out_s8'
$r = Invoke-Child $pwshExe "& '$scriptPath' -Arquivo '$entradaBom' -PastaSaida '$saida8' -Inventario"
Assert-Eq $r.Codigo 0 "S8 : exit 0"
Assert-True ($r.Saida -like '*Usuários no AD*5*') "S8 : total de usuarios no console"
Assert-True ($r.Saida -like '*Computadores no AD*3*') "S8 : total de computadores no console"

$j8 = Ler-JSON $saida8
if ($j8) {
    Assert-Eq $j8.Inventario.TotalUsuarios 5 "S8 : JSON usuarios = 5"
    Assert-Eq $j8.Inventario.TotalComputadores 3 "S8 : JSON computadores = 3"

    $comp = @($j8.Inventario.Computadores)
    Assert-Eq $comp.Count 3 "S8 : JSON lista de computadores"
    Assert-Eq $comp[0].Nome 'DESK-JSILVA' "S8 : computadores ordenados por nome"
    Assert-Eq $comp[0].IP '10.0.1.55' "S8 : IP do DESK-JSILVA"
    Assert-Eq $comp[1].IP '' "S8 : DESK-OLD sem IP"
    Assert-Eq $comp[1].Ativo 'Não' "S8 : DESK-OLD inativo"
    Assert-Eq $comp[2].IP '10.0.0.10' "S8 : IP do SRV-DC01"
    Assert-True ($comp[2].SO -like 'Windows Server*') "S8 : SO do servidor"

    $u8 = @($j8.Inventario.Usuarios)
    Assert-Eq $u8.Count 5 "S8 : JSON lista de usuarios"
    Assert-Eq $u8[0].SamAccountName 'jdesat' "S8 : usuarios ordenados por SAM"
}

$csvInv = @(Get-ChildItem -LiteralPath $saida8 -Filter 'inventario_usuarios_*.csv')
if ($csvInv.Count -gt 0) {
    $linhasInv = @(Import-Csv -LiteralPath $csvInv[0].FullName -Delimiter ';')
    Assert-Eq $linhasInv.Count 5 "S8 : CSV inventario usuarios = 5"
} else { Falha "S8 : CSV inventario usuarios nao gerado" }

$csvComp = @(Get-ChildItem -LiteralPath $saida8 -Filter 'inventario_computadores_*.csv')
if ($csvComp.Count -gt 0) {
    $linhasComp = @(Import-Csv -LiteralPath $csvComp[0].FullName -Delimiter ';')
    Assert-Eq $linhasComp.Count 3 "S8 : CSV inventario computadores = 3"
    Assert-Eq $linhasComp[2].Hostname 'srv-dc01.corp.local' "S8 : CSV hostname"
} else { Falha "S8 : CSV inventario computadores nao gerado" }

$hf8 = Get-ChildItem -LiteralPath $saida8 -Filter '*.html' | Select-Object -First 1
if ($hf8) {
    $h8 = Get-Content -LiteralPath $hf8.FullName -Raw -Encoding UTF8
    Assert-True ($h8.Contains('id="invUserTable"')) "S8 : HTML tabela de usuarios"
    Assert-True ($h8.Contains('id="invCompTable"')) "S8 : HTML tabela de computadores"
    Assert-True ($h8.Contains('10.0.0.10')) "S8 : HTML exibe IP"
    Assert-True ($h8.Contains('COMPUTADORES NO AD')) "S8 : HTML cartao de metrica"
    Assert-True ($h8.Contains('filtrarTabela')) "S8 : HTML busca do inventario"
    Assert-True (-not $h8.Contains('<th>Online</th>')) "S8 : sem coluna Online sem -VerificarOnline"
} else { Falha "S8 : HTML nao gerado" }

# ---- S9: sem -Inventario continua igual (sem secoes extras) ----
Write-Host ""
Write-Host '== S9: sem -Inventario ==' -ForegroundColor Cyan

$hf9 = Get-ChildItem -LiteralPath (Join-Path $testDir 'out_pwsh_bom') -Filter '*.html' | Select-Object -First 1
$h9 = Get-Content -LiteralPath $hf9.FullName -Raw -Encoding UTF8
Assert-True (-not $h9.Contains('id="invUserTable"')) "S9 : HTML sem inventario por padrao"
$j9 = Ler-JSON (Join-Path $testDir 'out_pwsh_bom')
if ($j9) { Assert-True ($null -eq $j9.Inventario) "S9 : JSON sem Inventario por padrao" }

# ---- S10: sem emoji / compativel com PowerShell antigo ----
Write-Host ""
Write-Host '== S10: sem emoji e compativel com PS 3.0+ ==' -ForegroundColor Cyan

$scriptText = [System.IO.File]::ReadAllText($scriptPath, [System.Text.Encoding]::UTF8)
Assert-True (-not [regex]::IsMatch($scriptText, '[\u0100-\uFFFF]')) "S10 : script so com Latin-1 (sem emoji)"
Assert-True (-not [regex]::IsMatch($h9, '[\u0100-\uFFFF]')) "S10 : HTML sem emoji"
Assert-True (-not [regex]::IsMatch($scriptText, '::new\(')) "S10 : sem ::new() (PS3)"
Assert-True ($scriptText.Contains('#Requires -Version 3.0')) "S10 : versao minima declarada"
Assert-True ($scriptText.Contains('#9662')) "S10 : setas como entidades HTML"

# ---- resumo ----
Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "Passou: $script:passou   Falhou: $script:falhou" -ForegroundColor $(if ($script:falhou -gt 0) { 'Red' } else { 'Green' })
Write-Host '==================================================' -ForegroundColor Cyan

if ($script:falhou -gt 0) { exit 1 } else { exit 0 }
