<#
.SYNOPSIS
    Consulta usuários no AD DS, gera mapeamento completo de Grupos e OUs,
    constrói a árvore hierárquica do domínio e exporta em HTML, JSON e CSV.
    XLSX e envio por SMTP são opcionais.

.DESCRIPTION
    Entrada:
        CSV ou TSV contendo Nome e/ou E-mail.

    Pesquisa:
        1. SamAccountName, se informado
        2. E-mail
        3. Nome completo

    Saídas:
        CSV
        JSON
        HTML
        XLSX opcional

    Com -Inventario:
        Também enumera todos os usuários e computadores do domínio,
        trazendo hostname e IP das máquinas (atributo IPv4Address do AD).
        Com -VerificarOnline, testa ping/DNS de cada computador para
        confirmar IP e status online (mais lento em domínios grandes).

.NOTES
    LastLogonDate usa lastLogonTimestamp replicado (atraso de até ~14 dias).
    Linhas sem Nome, E-mail e SamAccountName são ignoradas.
    -OUEsperada aceita nome exato de OU (ex.: "Users") ou trecho de DN
    (ex.: "OU=Users,OU=Corp").
#>

#Requires -Version 3.0

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Arquivo,

    [string]$PastaSaida,

    [string]$Server,

    [string[]]$Grupo,

    [string]$OUEsperada,

    [switch]$Aninhados,

    [switch]$Inventario,

    [switch]$VerificarOnline,

    [switch]$ExportarExcel,

    [switch]$EnviarEmail,

    [string]$SmtpServer,

    [int]$SmtpPort = 587,

    [string]$Remetente,

    [string[]]$Destinatarios,

    [switch]$UseSSL,

    [PSCredential]$Credenciais
)

# ============================================================
# CONFIGURAÇÃO INICIAL
# ============================================================

$ErrorActionPreference = "Stop"

Set-StrictMode -Version Latest

try {

    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {

    Write-Error -ErrorAction Continue "Módulo ActiveDirectory indisponível: $($_.Exception.Message)"
    exit 1
}

# ============================================================
# VALIDAR ARQUIVO
# ============================================================

if (-not (Test-Path -LiteralPath $Arquivo)) {

    Write-Error -ErrorAction Continue "Arquivo não encontrado: $Arquivo"
    exit 1
}

$ArquivoResolvido = (Resolve-Path -LiteralPath $Arquivo).Path

if (-not $PastaSaida) {

    $PastaSaida = Join-Path `
        (Split-Path -Path $ArquivoResolvido -Parent) `
        "Resultado_AD"
}

# ============================================================
# CONFIGURAR SERVER AD
# ============================================================

$srv = @{}

if ($Server) {

    $srv.Server = $Server
}

# ============================================================
# CONECTAR AO DOMÍNIO
# ============================================================

try {

    if ($Server) {

        $domainObj = Get-ADDomain -Server $Server
    }
    else {

        $domainObj = Get-ADDomain
    }

    $domainDN   = $domainObj.DistinguishedName
    $domainName = $domainObj.DNSRoot
}
catch {

    Write-Error -ErrorAction Continue `
        "Não foi possível conectar ao Active Directory. Verifique a conexão de rede, credenciais e parâmetros."

    Write-Error -ErrorAction Continue $_.Exception.Message

    exit 1
}

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " CONSULTA / MAPEAMENTO ACTIVE DIRECTORY" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Domínio : $domainName" -ForegroundColor Green
Write-Host "DN      : $domainDN" -ForegroundColor Green
Write-Host ""

# ============================================================
# DETECTAR ENCODING DO ARQUIVO
# ============================================================

try {

    $bytes = [System.IO.File]::ReadAllBytes($ArquivoResolvido)
}
catch {

    Write-Error -ErrorAction Continue "Não foi possível ler o arquivo: $($_.Exception.Message)"
    exit 1
}

$encObj = $null

if (
    $bytes.Length -ge 3 -and
    $bytes[0] -eq 0xEF -and
    $bytes[1] -eq 0xBB -and
    $bytes[2] -eq 0xBF
) {

    $encObj = [System.Text.Encoding]::UTF8
}
elseif (
    $bytes.Length -ge 4 -and
    $bytes[0] -eq 0xFF -and
    $bytes[1] -eq 0xFE -and
    $bytes[2] -eq 0 -and
    $bytes[3] -eq 0
) {

    $encObj = [System.Text.Encoding]::UTF32
}
elseif (
    $bytes.Length -ge 2 -and
    $bytes[0] -eq 0xFF -and
    $bytes[1] -eq 0xFE
) {

    $encObj = [System.Text.Encoding]::Unicode
}
elseif (
    $bytes.Length -ge 2 -and
    $bytes[0] -eq 0xFE -and
    $bytes[1] -eq 0xFF
) {

    $encObj = [System.Text.Encoding]::BigEndianUnicode
}
else {

    $utf8Estrito = New-Object System.Text.UTF8Encoding($false, $true)

    try {

        $null = $utf8Estrito.GetString($bytes)

        $encObj = [System.Text.Encoding]::UTF8
    }
    catch {

        $encObj = [System.Text.Encoding]::GetEncoding(
            [System.Globalization.CultureInfo]::CurrentCulture.TextInfo.ANSICodePage
        )
    }
}

$encLabel = switch ($encObj.CodePage) {

    65001  { "UTF-8" }
    1200   { "UTF-16 LE" }
    1201   { "UTF-16 BE" }
    12000  { "UTF-32" }
    default { "ANSI (cp $($encObj.CodePage))" }
}

if ($PSVersionTable.PSVersion.Major -ge 6) {

    $encParam = $encObj
}
else {

    $encParam = switch ($encObj.CodePage) {

        65001  { "UTF8" }
        1200   { "Unicode" }
        1201   { "BigEndianUnicode" }
        12000  { "UTF32" }
        default { "Default" }
    }
}

# ============================================================
# DETECTAR DELIMITADOR
# ============================================================

try {

    $cabecalho = Get-Content `
        -LiteralPath $ArquivoResolvido `
        -TotalCount 1 `
        -ErrorAction Stop

    if ($ArquivoResolvido -match '\.tsv$') {

        $delim = "`t"
    }
    elseif ($cabecalho -match "`t") {

        $delim = "`t"
    }
    elseif ($cabecalho -match ';') {

        $delim = ";"
    }
    else {

        $delim = ","
    }
}
catch {

    Write-Error -ErrorAction Continue "Não foi possível ler o cabeçalho do arquivo."
    Write-Error -ErrorAction Continue $_.Exception.Message
    exit 1
}

Write-Host "Arquivo de entrada : $ArquivoResolvido"
Write-Host "Delimitador        : " -NoNewline

switch ($delim) {

    "`t" {
        Write-Host "TAB" -ForegroundColor Yellow
    }

    ";" {
        Write-Host "Ponto e vírgula" -ForegroundColor Yellow
    }

    "," {
        Write-Host "Vírgula" -ForegroundColor Yellow
    }
}

Write-Host "Encoding         : $encLabel" -ForegroundColor Yellow

# ============================================================
# IMPORTAR LISTA
# ============================================================

try {

    $lista = @(
        Import-Csv `
            -LiteralPath $ArquivoResolvido `
            -Delimiter $delim `
            -Encoding $encParam
    )
}
catch {

    Write-Error -ErrorAction Continue "Erro ao importar o arquivo."
    Write-Error -ErrorAction Continue $_.Exception.Message
    exit 1
}

if ($lista.Count -eq 0) {

    Write-Error -ErrorAction Continue "O arquivo não possui registros."
    exit 1
}

# ============================================================
# IDENTIFICAR COLUNAS
# ============================================================

$colunas = @(
    $lista[0].PSObject.Properties.Name
)

Write-Host ""
Write-Host "Colunas identificadas:" -ForegroundColor Cyan

foreach ($coluna in $colunas) {

    Write-Host "  - [$coluna]"
}

# ============================================================
# FUNÇÃO PARA LOCALIZAR COLUNA
# ============================================================

function Get-Coluna {

    param(
        [object]$Row,
        [string[]]$PossiveisNomes
    )

    foreach ($nome in $PossiveisNomes) {

        $prop = $Row.PSObject.Properties |
            Where-Object {
                $_.Name.Trim().TrimStart([char]0xFEFF) -ieq $nome
            } |
            Select-Object -First 1

        if ($prop) {

            return $prop.Name
        }
    }

    return $null
}

$colunaNome = Get-Coluna `
    -Row $lista[0] `
    -PossiveisNomes @(
        "Nome",
        "Nome Completo",
        "Name",
        "DisplayName"
    )

$colunaEmail = Get-Coluna `
    -Row $lista[0] `
    -PossiveisNomes @(
        "E-mail",
        "Email",
        "E-Mail",
        "mail"
    )

$colunaSam = Get-Coluna `
    -Row $lista[0] `
    -PossiveisNomes @(
        "SamAccountName",
        "SAM",
        "Login",
        "Usuário",
        "Usuario"
    )

if (-not $colunaNome -and -not $colunaEmail -and -not $colunaSam) {

    Write-Error -ErrorAction Continue "Nenhuma coluna compatível foi encontrada."
    Write-Error -ErrorAction Continue "O arquivo precisa possuir pelo menos uma destas informações:"
    Write-Error -ErrorAction Continue "  Nome"
    Write-Error -ErrorAction Continue "  E-mail"
    Write-Error -ErrorAction Continue "  SamAccountName"

    exit 1
}

Write-Host ""
Write-Host "Mapeamento das colunas:" -ForegroundColor Cyan
Write-Host "  Nome  : $colunaNome"
Write-Host "  E-mail: $colunaEmail"
Write-Host "  SAM   : $colunaSam"

# ============================================================
# FUNÇÃO PARA LER VALOR DA COLUNA
# ============================================================

function Get-Valor {

    param(
        [object]$Row,
        [string]$Coluna
    )

    if (-not $Coluna) {

        return ""
    }

    $prop = $Row.PSObject.Properties[$Coluna]

    if ($null -eq $prop) {

        return ""
    }

    $valor = $prop.Value

    if ($null -eq $valor) {

        return ""
    }

    return "$valor".Trim()
}

# ============================================================
# FUNÇÃO ESCAPE HTML
# ============================================================

function HtmlEncode {

    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {

        return ""
    }

    return [System.Net.WebUtility]::HtmlEncode(
        [string]$Value
    )
}

# ============================================================
# FUNÇÃO SEPARAR PARTES DO DN
# ============================================================

function Split-DN {

    param(
        [string]$DN
    )

    $partes = New-Object System.Collections.ArrayList

    $atual = New-Object System.Text.StringBuilder

    $escape = $false

    foreach ($c in $DN.ToCharArray()) {

        if ($escape) {

            [void]$atual.Append('\')
            [void]$atual.Append($c)
            $escape = $false
        }
        elseif ($c -eq '\') {

            $escape = $true
        }
        elseif ($c -eq ',') {

            [void]$partes.Add($atual.ToString())
            [void]$atual.Clear()
        }
        else {

            [void]$atual.Append($c)
        }
    }

    if ($escape) {

        [void]$atual.Append('\')
    }

    [void]$partes.Add($atual.ToString())

    return $partes
}

# ============================================================
# FUNÇÃO EXTRAIR RDN
# ============================================================

function NomeRDN {

    param(
        [string]$DN
    )

    if (-not $DN) {

        return ""
    }

    $primeiro = @(Split-DN $DN)[0]

    return (
        $primeiro -replace '^[A-Za-z]+=', ''
    )
}

# ============================================================
# FUNÇÃO PAI DO DN
# ============================================================

function Get-PaiDN {

    param(
        [string]$DN
    )

    if (-not $DN) {

        return ""
    }

    $partes = @(Split-DN $DN)

    if ($partes.Count -le 1) {

        return ""
    }

    return (
        $partes[1..($partes.Count - 1)] -join ','
    )
}

# ============================================================
# FUNÇÃO ESCAPE LDAP
# ============================================================

function Escape-LDAP {

    param(
        [string]$Valor
    )

    if (-not $Valor) {

        return ""
    }

    return ($Valor.Replace('\', '\5c').Replace('*', '\2a').Replace('(', '\28').Replace(')', '\29').Replace("`0", '\00'))
}

# ============================================================
# FUNÇÃO CAMINHO DA OU
# ============================================================

function CaminhoOU {

    param(
        [string]$OUDN
    )

    if (-not $OUDN) {

        return ""
    }

    $partes = @(
        @(Split-DN $OUDN) |
        Where-Object {
            $_ -like "OU=*"
        } |
        ForEach-Object {
            $_.Substring(3)
        }
    )

    if ($partes.Count -eq 0) {

        return $OUDN
    }

    [array]::Reverse($partes)

    return ($partes -join " / ")
}

# ============================================================
# MAPEAR OUS
# ============================================================

Write-Host ""
Write-Host "-> Mapeando estrutura de OUs do domínio..." -ForegroundColor Cyan

try {

    $allOUs = @(
        Get-ADOrganizationalUnit `
            @srv `
            -Filter * `
            -Properties `
                DistinguishedName,
                Name,
                CanonicalName |
        Sort-Object CanonicalName
    )
}
catch {

    Write-Warning "Não foi possível obter todas as OUs."
    $allOUs = @()
}

Write-Host "OUs encontradas: $($allOUs.Count)" -ForegroundColor Green

# ============================================================
# CONSTRUIR MAPA DE OUS
# ============================================================

$mapaNome = @{}

$mapaFilhas = @{}

function Add-MapaOU {

    param(
        [string]$DN,
        [string]$Nome
    )

    if (-not $DN -or $DN -eq $domainDN) {

        return
    }

    if (-not $mapaNome.ContainsKey($DN)) {

        $mapaNome[$DN] = $Nome

        $pai = Get-PaiDN $DN

        if ($pai) {

            if (-not $mapaFilhas.ContainsKey($pai)) {

                $mapaFilhas[$pai] = New-Object System.Collections.ArrayList
            }

            [void]$mapaFilhas[$pai].Add($DN)
        }
    }

    $paiDireto = Get-PaiDN $DN

    if (
        $paiDireto -and
        $paiDireto -ne $domainDN -and
        -not $mapaNome.ContainsKey($paiDireto)
    ) {

        Add-MapaOU `
            -DN $paiDireto `
            -Nome (NomeRDN $paiDireto)
    }
}

foreach ($ou in $allOUs) {

    Add-MapaOU `
        -DN $ou.DistinguishedName `
        -Nome $ou.Name
}

# ============================================================
# CONSTRUIR ÁRVORE DE OUS
# ============================================================

function Build-OUHtmlTree {

    param(
        [string]$ParentDN
    )

    if (-not $mapaFilhas.ContainsKey($ParentDN)) {

        return ""
    }

    $html = "<ul>"

    foreach ($filho in $mapaFilhas[$ParentDN]) {

        $nome = HtmlEncode $mapaNome[$filho]

        $html += "<li>"

        if ($mapaFilhas.ContainsKey($filho)) {

            $html += "<button type='button' class='tree-toggle' aria-label='Recolher'>&#9662;</button>"
        }

        $html += "<span class='tree-node' data-dn='$(HtmlEncode $filho)'>&#8226; $nome</span>"

        $html += Build-OUHtmlTree `
            -ParentDN $filho

        $html += "</li>"
    }

    $html += "</ul>"

    return $html
}

$treeHtml = @"
<ul class="tree-root">
    <li>
        <span class="tree-node root" data-dn="">
            &#9679; $(HtmlEncode $domainName)
        </span>
        $(Build-OUHtmlTree -ParentDN $domainDN)
    </li>
</ul>
"@

# ============================================================
# PROPRIEDADES DO AD
# ============================================================

$adProps = @{}

foreach ($item in $srv.GetEnumerator()) {

    $adProps[$item.Key] = $item.Value
}

$adProps["Properties"] = @(
    "DisplayName",
    "Mail",
    "UserPrincipalName",
    "Enabled",
    "LastLogonDate",
    "MemberOf",
    "DistinguishedName",
    "SamAccountName",
    "Department",
    "Title"
)

# ============================================================
# PREPARAR ENTRADAS
# ============================================================

Write-Host ""
Write-Host "-> Preparando lista de consulta..." -ForegroundColor Cyan

$entradas = New-Object System.Collections.ArrayList

foreach ($l in $lista) {

    $nomeL = Get-Valor `
        -Row $l `
        -Coluna $colunaNome

    $mailL = (
        Get-Valor `
            -Row $l `
            -Coluna $colunaEmail
    ).ToLower()

    $samL = Get-Valor `
        -Row $l `
        -Coluna $colunaSam

    if ($nomeL -or $mailL -or $samL) {

        [void]$entradas.Add(
            [PSCustomObject]@{
                Nome = $nomeL

                Mail = $mailL

                Sam = $samL
            }
        )
    }
}

$linhasIgnoradas = $lista.Count - $entradas.Count

if ($linhasIgnoradas -gt 0) {

    Write-Warning `
        "Linhas ignoradas sem Nome, E-mail e SamAccountName: $linhasIgnoradas"
}

if ($entradas.Count -eq 0) {

    Write-Error -ErrorAction Continue "Nenhuma linha válida para consulta."
    exit 1
}

# ============================================================
# FUNÇÕES DE ÍNDICE E CONSULTA EM LOTE
# ============================================================

function Get-Chave {

    param(
        [string]$Valor
    )

    if (-not $Valor) {

        return ""
    }

    return $Valor.Trim().ToLowerInvariant()
}

function Add-Indice {

    param(
        [hashtable]$Indice,
        [string]$Chave,
        [object]$Usuario
    )

    if (-not $Chave) {

        return
    }

    if (-not $Indice.ContainsKey($Chave)) {

        $Indice[$Chave] = New-Object System.Collections.ArrayList
    }

    $existente = $Indice[$Chave]

    if (-not $existente.Contains($Usuario)) {

        [void]$existente.Add($Usuario)
    }
}

function Get-UsuariosLote {

    param(
        [string[]]$Valores,
        [string]$Modelo
    )

    $resultado = New-Object System.Collections.ArrayList

    if (-not $Valores) {

        return $resultado
    }

    $tamanho = 400

    for ($ini = 0; $ini -lt $Valores.Count; $ini += $tamanho) {

        $fim = [Math]::Min(
            $ini + $tamanho - 1,
            $Valores.Count - 1
        )

        $pedaco = @($Valores[$ini..$fim])

        $condicoes = (
            $pedaco |
            ForEach-Object {
                $Modelo -f (Escape-LDAP $_)
            }
        ) -join ""

        try {

            $lote = @(
                Get-ADUser `
                    @adProps `
                    -LDAPFilter "(|$condicoes)"
            )

            foreach ($u in $lote) {

                [void]$resultado.Add($u)
            }
        }
        catch {

            Write-Warning `
                "Erro na consulta em lote: $($_.Exception.Message)"
        }
    }

    return $resultado
}

function Resolve-Entrada {

    param(
        [object]$Entrada
    )

    if ($Entrada.Sam) {

        $k = Get-Chave $Entrada.Sam

        if ($idxSam.ContainsKey($k)) {

            return @{
                Por = "SamAccountName"

                Lista = $idxSam[$k]
            }
        }
    }

    if ($Entrada.Mail) {

        $k = Get-Chave $Entrada.Mail

        if ($idxMail.ContainsKey($k)) {

            return @{
                Por = "E-mail"

                Lista = $idxMail[$k]
            }
        }
    }

    if ($Entrada.Nome) {

        $k = Get-Chave $Entrada.Nome

        if ($idxNome.ContainsKey($k)) {

            return @{
                Por = "Nome"

                Lista = $idxNome[$k]
            }
        }
    }

    return @{
        Por = ""

        Lista = @()
    }
}

# ============================================================
# CONSULTA EM LOTE
# ============================================================

Write-Host ""
Write-Host "-> Consultando usuários no AD DS..." -ForegroundColor Cyan
Write-Host ""

$idxSam = @{}

$idxMail = @{}

$idxNome = @{}

$hsSam = New-Object 'System.Collections.Generic.HashSet[string]' -ArgumentList (
    [System.StringComparer]::OrdinalIgnoreCase
)

$hsMail = New-Object 'System.Collections.Generic.HashSet[string]' -ArgumentList (
    [System.StringComparer]::OrdinalIgnoreCase
)

$hsNome = New-Object 'System.Collections.Generic.HashSet[string]' -ArgumentList (
    [System.StringComparer]::OrdinalIgnoreCase
)

foreach ($e in $entradas) {

    if ($e.Sam) {

        $null = $hsSam.Add($e.Sam)
    }

    if ($e.Mail) {

        $null = $hsMail.Add($e.Mail)
    }

    if ($e.Nome) {

        $null = $hsNome.Add($e.Nome)
    }
}

$usuariosSam = @(
    Get-UsuariosLote `
        -Valores @($hsSam) `
        -Modelo '(sAMAccountName={0})'
)

foreach ($u in $usuariosSam) {

    Add-Indice `
        -Indice $idxSam `
        -Chave (Get-Chave $u.SamAccountName) `
        -Usuario $u
}

$usuariosMail = @(
    Get-UsuariosLote `
        -Valores @($hsMail) `
        -Modelo '(mail={0})(userPrincipalName={0})'
)

foreach ($u in $usuariosMail) {

    $km = Get-Chave $u.Mail

    $ku = Get-Chave $u.UserPrincipalName

    Add-Indice `
        -Indice $idxMail `
        -Chave $km `
        -Usuario $u

    if ($ku -ne $km) {

        Add-Indice `
            -Indice $idxMail `
            -Chave $ku `
            -Usuario $u
    }
}

$usuariosNome = @(
    Get-UsuariosLote `
        -Valores @($hsNome) `
        -Modelo '(displayName={0})(name={0})'
)

foreach ($u in $usuariosNome) {

    $kn = Get-Chave $u.DisplayName

    $kr = Get-Chave $u.Name

    Add-Indice `
        -Indice $idxNome `
        -Chave $kn `
        -Usuario $u

    if ($kr -ne $kn) {

        Add-Indice `
            -Indice $idxNome `
            -Chave $kr `
            -Usuario $u
    }
}

# ============================================================
# PROCESSAMENTO
# ============================================================

$i = 0

$res = foreach ($entrada in $entradas) {

    $i++

    $percentual = [math]::Round(
        (100 * $i / $entradas.Count),
        0
    )

    Write-Progress `
        -Activity "Consultando AD DS" `
        -Status "$i de $($entradas.Count)" `
        -PercentComplete $percentual

    # --------------------------------------------------------
    # DADOS DA LISTA
    # --------------------------------------------------------

    $nome = $entrada.Nome

    $mail = $entrada.Mail

    $sam = $entrada.Sam

    # --------------------------------------------------------
    # RESOLVER ENTRADA
    # --------------------------------------------------------

    $resolvido = Resolve-Entrada -Entrada $entrada

    $usuarios = @($resolvido.Lista)

    $por = $resolvido.Por

    # --------------------------------------------------------
    # STATUS
    # --------------------------------------------------------

    $n = $usuarios.Count

    if ($n -eq 0) {

        $status = "Não Encontrado"
    }
    elseif ($n -gt 1) {

        $status = "Duplicado"
    }
    else {

        $status = "Encontrado"
    }

    # --------------------------------------------------------
    # VARIÁVEIS
    # --------------------------------------------------------

    $div = @()
    $grupos = @()
    $falta = @()

    $ouDn = ""
    $ouCam = ""
    $ouOk = ""

    $ativo = $null
    $ultimo = ""
    $qtd = 0

    $usuario = $null

    # --------------------------------------------------------
    # PROCESSAR USUÁRIO ÚNICO
    # --------------------------------------------------------

    if ($n -eq 1) {

        $usuario = $usuarios[0]

        $ativo = $usuario.Enabled

        if ($usuario.LastLogonDate) {

            $ultimo = $usuario.LastLogonDate.ToString(
                "dd/MM/yyyy HH:mm"
            )
        }
        else {

            $ultimo = "Nunca"
        }

        # ----------------------------------------------------
        # COMPARAÇÕES
        # ----------------------------------------------------

        if (
            $sam -and
            $usuario.SamAccountName -ne $sam
        ) {

            $div += "Sam"
        }

        if (
            $mail -and
            $usuario.Mail -and
            $usuario.Mail.ToLower() -ne $mail
        ) {

            $div += "E-mail"
        }

        if (
            $nome -and
            $usuario.DisplayName -and
            $usuario.DisplayName -ne $nome
        ) {

            $div += "Nome"
        }

        # ----------------------------------------------------
        # GRUPOS
        # ----------------------------------------------------

        if ($Aninhados) {

            $escapedDN = $usuario.DistinguishedName

            $escapedDN = $escapedDN.Replace(
                '\',
                '\5c'
            )

            $escapedDN = $escapedDN.Replace(
                '*',
                '\2a'
            )

            $escapedDN = $escapedDN.Replace(
                '(',
                '\28'
            )

            $escapedDN = $escapedDN.Replace(
                ')',
                '\29'
            )

            try {

                $grupos = @(
                    Get-ADGroup `
                        @srv `
                        -LDAPFilter `
                        "(member:1.2.840.113556.1.4.1941:=$escapedDN)" |
                    Select-Object -ExpandProperty Name
                )
            }
            catch {

                Write-Warning `
                    "Erro ao consultar grupos aninhados de $($usuario.SamAccountName)"

                $grupos = @()
            }
        }
        else {

            $grupos = @(
                $usuario.MemberOf |
                ForEach-Object {
                    NomeRDN $_
                }
            )
        }

        $grupos = @(
            $grupos |
            Where-Object {
                $_
            } |
            Sort-Object -Unique
        )

        $qtd = $grupos.Count

        # ----------------------------------------------------
        # GRUPOS FALTANTES
        # ----------------------------------------------------

        if ($Grupo) {

            $falta = @(
                $Grupo |
                Where-Object {
                    $grupos -notcontains $_
                }
            )
        }

        # ----------------------------------------------------
        # OU
        # ----------------------------------------------------

        $ouDn = Get-PaiDN $usuario.DistinguishedName

        $ouCam = CaminhoOU $ouDn

        if ($OUEsperada) {

            $esperado = $OUEsperada.Trim()

            $ouOk = "Não"

            if ($esperado.Contains('=')) {

                if (
                    $ouDn.IndexOf(
                        $esperado,
                        [StringComparison]::OrdinalIgnoreCase
                    ) -ge 0
                ) {

                    $ouOk = "Sim"
                }
            }
            else {

                $nomesOU = @(
                    @(Split-DN $ouDn) |
                    Where-Object {
                        $_ -like "OU=*"
                    } |
                    ForEach-Object {
                        $_.Substring(3).Trim()
                    }
                )

                if ($nomesOU -contains $esperado) {

                    $ouOk = "Sim"
                }
            }
        }
    }

    # --------------------------------------------------------
    # RESULTADO
    # --------------------------------------------------------

    [PSCustomObject]@{

        Nome_Lista = $nome

        Sam_Lista = $sam

        Email_Lista = $mail

        Status = $status

        EncontradoPor = $por

        Duplicados = if ($n -gt 1) {

            (
                $usuarios |
                ForEach-Object {
                    $_.SamAccountName
                }
            ) -join "; "
        }
        else {

            ""
        }

        Sam_AD = if ($usuario) {
            $usuario.SamAccountName
        }
        else {
            ""
        }

        Nome_AD = if ($usuario) {
            $usuario.DisplayName
        }
        else {
            ""
        }

        Email_AD = if ($usuario) {
            $usuario.Mail
        }
        else {
            ""
        }

        UserPrincipalName = if ($usuario) {
            $usuario.UserPrincipalName
        }
        else {
            ""
        }

        Ativo = $ativo

        Departamento = if ($usuario) {
            $usuario.Department
        }
        else {
            ""
        }

        Cargo = if ($usuario) {
            $usuario.Title
        }
        else {
            ""
        }

        UltimoLogon = $ultimo

        Divergencia = (
            $div -join ", "
        )

        OU_Caminho = $ouCam

        OU_DN = $ouDn

        OU_OK = $ouOk

        QtdGrupos = $qtd

        Grupos = (
            $grupos -join "; "
        )

        GruposFaltando = (
            $falta -join "; "
        )
    }

    # --------------------------------------------------------
    # STATUS NO CONSOLE
    # --------------------------------------------------------

    switch ($status) {

        "Encontrado" {

            Write-Host `
                "[$i/$($entradas.Count)] OK - $nome" `
                -ForegroundColor Green
        }

        "Duplicado" {

            Write-Host `
                "[$i/$($entradas.Count)] DUPLICADO - $nome" `
                -ForegroundColor Yellow
        }

        "Não Encontrado" {

            Write-Host `
                "[$i/$($entradas.Count)] NÃO ENCONTRADO - $nome" `
                -ForegroundColor Red
        }
    }
}

Write-Progress `
    -Activity "Consultando AD DS" `
    -Completed

# ============================================================
# CRIAR PASTA DE SAÍDA
# ============================================================

New-Item `
    -ItemType Directory `
    -Path $PastaSaida `
    -Force |
    Out-Null

# ============================================================
# NOMES DOS ARQUIVOS
# ============================================================

$timeStamp = Get-Date -Format "yyyyMMdd_HHmmss"

$fileCsv = Join-Path `
    $PastaSaida `
    "resultado_AD_$timeStamp.csv"

$fileJson = Join-Path `
    $PastaSaida `
    "resultado_AD_$timeStamp.json"

$fileHtml = Join-Path `
    $PastaSaida `
    "resultado_AD_$timeStamp.html"

$fileXlsx = Join-Path `
    $PastaSaida `
    "resultado_AD_$timeStamp.xlsx"

$fileInvUserCsv = Join-Path `
    $PastaSaida `
    "inventario_usuarios_$timeStamp.csv"

$fileInvCompCsv = Join-Path `
    $PastaSaida `
    "inventario_computadores_$timeStamp.csv"

# ============================================================
# INVENTÁRIO DO AD (USUÁRIOS E COMPUTADORES)
# ============================================================

$inventarioUsuarios = @()

$inventarioComputadores = @()

$invUsuariosRows = ""

$invComputadoresRows = ""

$inventarioHtml = ""

$metricasInventario = ""

$jsonInventario = $null

function Formatar-DataAD {

    param(
        [object]$Valor
    )

    if (-not $Valor) {

        return ""
    }

    return ([datetime]$Valor).ToString(
        "dd/MM/yyyy HH:mm"
    )
}

function Obter-IPComputador {

    param(
        [string]$DNSHostName
    )

    if (-not $DNSHostName) {

        return ""
    }

    try {

        $ips = @(
            [System.Net.Dns]::GetHostAddresses(
                $DNSHostName
            ) |
            Where-Object {
                $_.AddressFamily -eq `
                    [System.Net.Sockets.AddressFamily]::InterNetwork
            }
        )

        if ($ips.Count -gt 0) {

            return $ips[0].IPAddressToString
        }
    }
    catch {
    }

    return ""
}

if ($Inventario) {

    Write-Host ""
    Write-Host "-> Inventariando todo o AD (usuários e computadores)..." -ForegroundColor Cyan

    # --------------------------------------------------------
    # USUÁRIOS DO AD
    # --------------------------------------------------------

    try {

        $adUsuarios = @(
            Get-ADUser `
                @srv `
                -Filter * `
                -Properties `
                DisplayName, Mail, Enabled, Department, `
                Title, Description, Created, LastLogonDate, `
                DistinguishedName
        )
    }
    catch {

        Write-Warning `
            "Falha ao listar usuários do AD: $($_.Exception.Message)"

        $adUsuarios = @()
    }

    $inventarioUsuarios = @(
        $adUsuarios |
        Sort-Object SamAccountName |
        ForEach-Object {

            [PSCustomObject]@{
                SamAccountName = $_.SamAccountName

                Nome = $_.DisplayName

                Email = $_.Mail

                Ativo = if ($_.Enabled) {
                    "Sim"
                }
                else {
                    "Não"
                }

                Departamento = $_.Department

                Cargo = $_.Title

                Criado = Formatar-DataAD $_.Created

                UltimoLogon = if ($_.LastLogonDate) {
                    (Formatar-DataAD $_.LastLogonDate)
                }
                else {
                    "Nunca"
                }

                OU = CaminhoOU (Get-PaiDN $_.DistinguishedName)

                Descricao = $_.Description
            }
        }
    )

    # --------------------------------------------------------
    # COMPUTADORES DO AD
    # --------------------------------------------------------

    try {

        $adComputadores = @(
            Get-ADComputer `
                @srv `
                -Filter * `
                -Properties `
                DNSHostName, IPv4Address, OperatingSystem, `
                OperatingSystemVersion, Description, Enabled, `
                Created, LastLogonDate, DistinguishedName
        )
    }
    catch {

        Write-Warning `
            "Falha ao listar computadores do AD: $($_.Exception.Message)"

        $adComputadores = @()
    }

    $inventarioComputadores = @(
        $adComputadores |
        Sort-Object Name |
        ForEach-Object {

            $ip = $_.IPv4Address

            $online = ""

            if ($VerificarOnline -and $_.DNSHostName) {

                try {

                    if (
                        Test-Connection `
                            -ComputerName $_.DNSHostName `
                            -Count 1 `
                            -Quiet `
                            -ErrorAction Stop
                    ) {

                        $online = "Sim"
                    }
                    else {

                        $online = "Não"
                    }
                }
                catch {

                    $online = "Não"
                }

                if (-not $ip) {

                    $ip = Obter-IPComputador $_.DNSHostName
                }
            }

            [PSCustomObject]@{
                Nome = $_.Name

                Hostname = $_.DNSHostName

                IP = $ip

                Ativo = if ($_.Enabled) {
                    "Sim"
                }
                else {
                    "Não"
                }

                Online = $online

                SO = $_.OperatingSystem

                VersaoSO = $_.OperatingSystemVersion

                Criado = Formatar-DataAD $_.Created

                UltimoLogon = if ($_.LastLogonDate) {
                    (Formatar-DataAD $_.LastLogonDate)
                }
                else {
                    "Nunca"
                }

                OU = CaminhoOU (Get-PaiDN $_.DistinguishedName)

                Descricao = $_.Description
            }
        }
    )

    Write-Host `
        "Usuários no AD     : $($inventarioUsuarios.Count)" `
        -ForegroundColor Green

    Write-Host `
        "Computadores no AD : $($inventarioComputadores.Count)" `
        -ForegroundColor Green

    # --------------------------------------------------------
    # CSV DO INVENTÁRIO
    # --------------------------------------------------------

    if ($inventarioUsuarios.Count -gt 0) {

        $inventarioUsuarios |
            Export-Csv `
                -Path $fileInvUserCsv `
                -NoTypeInformation `
                -Delimiter ";" `
                -Encoding UTF8
    }

    if ($inventarioComputadores.Count -gt 0) {

        $inventarioComputadores |
            Export-Csv `
                -Path $fileInvCompCsv `
                -NoTypeInformation `
                -Delimiter ";" `
                -Encoding UTF8
    }

    # --------------------------------------------------------
    # JSON DO INVENTÁRIO
    # --------------------------------------------------------

    $jsonInventario = [PSCustomObject]@{
        TotalUsuarios = $inventarioUsuarios.Count

        TotalComputadores = $inventarioComputadores.Count

        Usuarios = @($inventarioUsuarios)

        Computadores = @($inventarioComputadores)
    }

    # --------------------------------------------------------
    # LINHAS HTML DO INVENTÁRIO
    # --------------------------------------------------------

    $invUsuariosRows = @(
        $inventarioUsuarios |
        ForEach-Object {

            $tdAtivo = if ($_.Ativo -eq "Sim") {
                "<span class='badge badge-success'>Sim</span>"
            }
            else {
                "<span class='badge badge-danger'>Não</span>"
            }

            @"
<tr>
    <td><b>$(HtmlEncode $_.SamAccountName)</b></td>
    <td>$(HtmlEncode $_.Nome)</td>
    <td>$(HtmlEncode $_.Email)</td>
    <td>$tdAtivo</td>
    <td>$(HtmlEncode $_.Departamento)</td>
    <td>$(HtmlEncode $_.Cargo)</td>
    <td>$(HtmlEncode $_.Criado)</td>
    <td>$(HtmlEncode $_.UltimoLogon)</td>
    <td><small>$(HtmlEncode $_.OU)</small></td>
</tr>
"@
        }
    ) -join "`n"

    $invComputadoresRows = @(
        $inventarioComputadores |
        ForEach-Object {

            $tdAtivo = if ($_.Ativo -eq "Sim") {
                "<span class='badge badge-success'>Sim</span>"
            }
            else {
                "<span class='badge badge-danger'>Não</span>"
            }

            $tdOnline = ""

            if ($VerificarOnline) {

                if ($_.Online -eq "Sim") {
                    $tdOnline = "<td><span class='badge badge-success'>Sim</span></td>"
                }
                elseif ($_.Online -eq "Não") {
                    $tdOnline = "<td><span class='badge badge-danger'>Não</span></td>"
                }
                else {
                    $tdOnline = "<td>-</td>"
                }
            }

            @"
<tr>
    <td><b>$(HtmlEncode $_.Nome)</b></td>
    <td>$(HtmlEncode $_.Hostname)</td>
    <td>$(HtmlEncode $_.IP)</td>
    <td>$tdAtivo</td>
    $tdOnline
    <td>$(HtmlEncode $_.SO)</td>
    <td>$(HtmlEncode $_.Criado)</td>
    <td>$(HtmlEncode $_.UltimoLogon)</td>
    <td><small>$(HtmlEncode $_.OU)</small></td>
    <td><small>$(HtmlEncode $_.Descricao)</small></td>
</tr>
"@
        }
    ) -join "`n"

    $thOnline = ""

    if ($VerificarOnline) {

        $thOnline = "<th>Online</th>"
    }

    # --------------------------------------------------------
    # PAINÉIS HTML DO INVENTÁRIO
    # --------------------------------------------------------

    $inventarioHtml = @"
<div class="inventario">

    <div class="panel">

        <h2>
            Inventário de Usuários - todo o AD
            ($($inventarioUsuarios.Count))
        </h2>

        <input
            type="text"
            id="invUserSearch"
            class="search-box"
            onkeyup="filtrarTabela('invUserTable', this.value)"
            placeholder="Buscar por Nome, SAM, E-mail, Departamento ou OU..."
        >

        <table id="invUserTable">

            <thead>

                <tr>

                    <th>SAM</th>

                    <th>Nome</th>

                    <th>E-mail</th>

                    <th>Ativo</th>

                    <th>Departamento</th>

                    <th>Cargo</th>

                    <th>Criado</th>

                    <th>Último Logon</th>

                    <th>OU</th>

                </tr>

            </thead>

            <tbody>

                $invUsuariosRows

            </tbody>

        </table>

    </div>

    <div class="panel">

        <h2>
            Inventário de Computadores - todo o AD
            ($($inventarioComputadores.Count))
        </h2>

        <input
            type="text"
            id="invCompSearch"
            class="search-box"
            onkeyup="filtrarTabela('invCompTable', this.value)"
            placeholder="Buscar por Nome, Hostname, IP, SO ou OU..."
        >

        <table id="invCompTable">

            <thead>

                <tr>

                    <th>Nome</th>

                    <th>Hostname</th>

                    <th>IP</th>

                    <th>Ativo</th>

                    $thOnline

                    <th>Sistema Operacional</th>

                    <th>Criado</th>

                    <th>Último Logon</th>

                    <th>OU</th>

                    <th>Descrição</th>

                </tr>

            </thead>

            <tbody>

                $invComputadoresRows

            </tbody>

        </table>

    </div>

</div>
"@

    # --------------------------------------------------------
    # CARTÕES DE MÉTRICA DO INVENTÁRIO
    # --------------------------------------------------------

    $metricasInventario = @"
    <div class="card">

        <h3>
            USUÁRIOS NO AD
        </h3>

        <p style="color:var(--accent-color);">
            $($inventarioUsuarios.Count)
        </p>

    </div>

    <div class="card">

        <h3>
            COMPUTADORES NO AD
        </h3>

        <p style="color:var(--accent-color);">
            $($inventarioComputadores.Count)
        </p>

    </div>
"@
}

# ============================================================
# EXPORTAR CSV
# ============================================================

$res |
    Export-Csv `
        -Path $fileCsv `
        -NoTypeInformation `
        -Delimiter ";" `
        -Encoding UTF8

# ============================================================
# EXPORTAR JSON
# ============================================================

$json = @{
    GeradoEm = (
        Get-Date
    ).ToString("o")

    Dominio = $domainName

    Total = $res.Count

    Usuarios = @($res)

    ArvoreOUs = @(
        $allOUs |
        Select-Object `
            Name,
            DistinguishedName,
            CanonicalName
    )
}

if ($Inventario) {

    $json['Inventario'] = $jsonInventario
}

$json |
    ConvertTo-Json `
        -Depth 6 |
    Set-Content `
        -Path $fileJson `
        -Encoding UTF8

# ============================================================
# EXPORTAR EXCEL
# ============================================================

$excelOk = $false

if ($ExportarExcel) {

    Write-Host ""
    Write-Host "-> Gerando planilha Excel..." -ForegroundColor Cyan

    try {

        if (-not (
            Get-Module `
                -ListAvailable `
                -Name ImportExcel
        )) {

            Write-Host `
                "Módulo ImportExcel não encontrado." `
                -ForegroundColor Yellow

            Write-Host `
                "Instalando no escopo CurrentUser..." `
                -ForegroundColor Yellow

            Install-Module `
                -Name ImportExcel `
                -Scope CurrentUser `
                -Force `
                -AllowClobber `
                -ErrorAction Stop
        }

        Import-Module ImportExcel -ErrorAction Stop

        $res |
            Export-Excel `
                -Path $fileXlsx `
                -WorksheetName "Usuários AD" `
                -TableStyle Medium2 `
                -AutoSize `
                -AutoFilter `
                -FreezeTopRow `
                -BoldTopRow `
                -ErrorAction Stop

        if (
            $Inventario -and
            $inventarioUsuarios.Count -gt 0
        ) {

            $inventarioUsuarios |
                Export-Excel `
                    -Path $fileXlsx `
                    -WorksheetName "Inventário Usuários" `
                    -TableStyle Medium4 `
                    -AutoSize `
                    -AutoFilter `
                    -FreezeTopRow `
                    -BoldTopRow `
                    -Append `
                    -ErrorAction Stop
        }

        if (
            $Inventario -and
            $inventarioComputadores.Count -gt 0
        ) {

            $inventarioComputadores |
                Export-Excel `
                    -Path $fileXlsx `
                    -WorksheetName "Inventário Computadores" `
                    -TableStyle Medium6 `
                    -AutoSize `
                    -AutoFilter `
                    -FreezeTopRow `
                    -BoldTopRow `
                    -Append `
                    -ErrorAction Stop
        }

        $excelOk = $true
    }
    catch {

        Write-Warning `
            "Não foi possível gerar o XLSX: $($_.Exception.Message)"

        $excelOk = $false
    }
}

# ============================================================
# ESTATÍSTICAS
# ============================================================

$totEncontrados = @(
    $res |
    Where-Object {
        $_.Status -eq "Encontrado"
    }
).Count

$totNaoEncontrados = @(
    $res |
    Where-Object {
        $_.Status -eq "Não Encontrado"
    }
).Count

$totDuplicados = @(
    $res |
    Where-Object {
        $_.Status -eq "Duplicado"
    }
).Count

# ============================================================
# CONSTRUIR LINHAS HTML
# ============================================================

$rowsHtml = foreach ($r in $res) {

    $statusClass = switch ($r.Status) {

        "Encontrado" {
            "badge-success"
        }

        "Não Encontrado" {
            "badge-danger"
        }

        default {
            "badge-warning"
        }
    }

    $ativoBadge = if ($r.Ativo -eq $true) {

        "<span class='badge badge-success'>Sim</span>"
    }
    elseif ($r.Ativo -eq $false) {

        "<span class='badge badge-danger'>Não</span>"
    }
    else {

        "-"
    }

    if ($r.Grupos) {

        $gruposFormatted = (
            $r.Grupos -split "; " |
            ForEach-Object {

                "<span class='tag'>$(HtmlEncode $_)</span>"
            }
        ) -join " "
    }
    else {

        $gruposFormatted =
            "<i style='color:#888;'>Nenhum</i>"
    }

    if ($r.GruposFaltando) {

        $faltandoFormatted = (
            $r.GruposFaltando -split "; " |
            ForEach-Object {

                "<span class='tag tag-alert'>$(HtmlEncode $_)</span>"
            }
        ) -join " "
    }
    else {

        $faltandoFormatted = "-"
    }

    @"
<tr data-ou="$(HtmlEncode $r.OU_DN)">
    <td>
        <span class="badge $statusClass">
            $(HtmlEncode $r.Status)
        </span>
    </td>

    <td>
        <b>$(HtmlEncode $r.Sam_AD)</b>
        <br>
        <small>
            $(HtmlEncode $r.Nome_AD)
        </small>
        $(
            if ($r.Duplicados) {

                "<br>
                <small style='color:#eab308;'>
                    Candidatos: $(HtmlEncode $r.Duplicados)
                </small>"
            }
        )
    </td>

    <td>
        $(HtmlEncode $r.Email_AD)
    </td>

    <td>
        $ativoBadge
    </td>

    <td>
        <small>
            $(HtmlEncode $r.OU_Caminho)
        </small>
    </td>

    <td>
        $gruposFormatted
    </td>

    <td>
        $faltandoFormatted
    </td>

    <td>
        <small>
            $(HtmlEncode $r.UltimoLogon)
        </small>
    </td>
</tr>
"@
}

# ============================================================
# TEMPLATE HTML
# ============================================================

$htmlTemplate = @"
<!DOCTYPE html>

<html lang="pt-BR">

<head>

<meta charset="UTF-8">

<meta
    name="viewport"
    content="width=device-width, initial-scale=1.0"
>

<title>
Relatório AD DS - $(HtmlEncode $domainName)
</title>

<style>

:root {
    --bg-color: #0f172a;
    --card-bg: #1e293b;
    --text-color: #f8fafc;
    --border-color: #334155;
    --accent-color: #38bdf8;
    --success-color: #22c55e;
    --warning-color: #eab308;
    --danger-color: #ef4444;
}

body {
    font-family:
        'Segoe UI',
        Tahoma,
        Geneva,
        Verdana,
        sans-serif;

    background-color: var(--bg-color);

    color: var(--text-color);

    margin: 0;

    padding: 24px;
}

.header {
    display: flex;

    justify-content: space-between;

    align-items: center;

    border-bottom:
        1px solid var(--border-color);

    padding-bottom: 16px;

    margin-bottom: 24px;
}

.header h1 {
    margin: 0;

    font-size: 1.8rem;

    color: var(--accent-color);
}

.metrics {
    display: grid;

    grid-template-columns:
        repeat(
            auto-fit,
            minmax(200px, 1fr)
        );

    gap: 16px;

    margin-bottom: 24px;
}

.card {
    background:
        var(--card-bg);

    border:
        1px solid var(--border-color);

    border-radius: 8px;

    padding: 16px;

    text-align: center;
}

.card h3 {
    margin: 0;

    font-size: 0.9rem;

    color: #94a3b8;
}

.card p {
    margin: 8px 0 0 0;

    font-size: 1.8rem;

    font-weight: bold;
}

.main-layout {
    display: grid;

    grid-template-columns:
        320px 1fr;

    gap: 24px;
}

.inventario {
    display: grid;

    grid-template-columns:
        1fr;

    gap: 24px;

    margin-top: 24px;
}

@media (max-width: 1024px) {

    .main-layout {
        grid-template-columns: 1fr;
    }
}

.panel {
    background:
        var(--card-bg);

    border:
        1px solid var(--border-color);

    border-radius: 8px;

    padding: 16px;

    overflow-x: auto;
}

.panel h2 {
    margin-top: 0;

    font-size: 1.2rem;

    color: var(--accent-color);
}

.tree-root,
.tree-root ul {
    list-style: none;

    padding-left: 16px;

    margin: 0;
}

.tree-node {
    display: inline-block;

    padding: 4px 8px;

    margin: 2px 0;

    border-radius: 4px;

    font-size: 0.85rem;

    background: #0f172a;

    border:
        1px solid var(--border-color);

    cursor: pointer;

    transition:
        background 0.15s ease,
        color 0.15s ease;
}

.tree-node:hover {
    background: #1e3a5f;

    color: #fff;
}

.tree-node.active {
    background: var(--accent-color);

    color: #fff;

    font-weight: bold;
}

.tree-toggle {
    background: none;

    border: none;

    color: #94a3b8;

    cursor: pointer;

    font-size: 0.75rem;

    padding: 0 2px 0 0;

    vertical-align: middle;
}

.tree-toggle:hover {
    color: #fff;
}

li.collapsed > ul {
    display: none;
}

.tree-hint {
    color: #94a3b8;

    font-size: 0.75rem;

    margin: 0 0 10px 0;
}

.tree-node.root {
    background: #0284c7;

    color: #fff;

    font-weight: bold;
}

.search-box {
    width: 100%;

    padding: 10px;

    border-radius: 6px;

    border:
        1px solid var(--border-color);

    background: #0f172a;

    color: #fff;

    margin-bottom: 16px;

    box-sizing: border-box;
}

table {
    width: 100%;

    border-collapse: collapse;

    font-size: 0.88rem;
}

th,
td {
    padding: 10px;

    text-align: left;

    border-bottom:
        1px solid var(--border-color);

    vertical-align: top;
}

th {
    background-color: #0f172a;

    color: var(--accent-color);

    position: sticky;

    top: 0;
}

.badge {
    padding: 3px 8px;

    border-radius: 12px;

    font-size: 0.75rem;

    font-weight: bold;

    white-space: nowrap;
}

.badge-success {
    background:
        rgba(34, 197, 94, 0.2);

    color: var(--success-color);

    border:
        1px solid var(--success-color);
}

.badge-danger {
    background:
        rgba(239, 68, 68, 0.2);

    color: var(--danger-color);

    border:
        1px solid var(--danger-color);
}

.badge-warning {
    background:
        rgba(234, 179, 8, 0.2);

    color: var(--warning-color);

    border:
        1px solid var(--warning-color);
}

.tag {
    display: inline-block;

    background: #334155;

    color: #e2e8f0;

    padding: 2px 6px;

    border-radius: 4px;

    font-size: 0.75rem;

    margin: 1px;
}

.tag-alert {
    background:
        rgba(239, 68, 68, 0.3);

    color: #fca5a5;

    border:
        1px solid var(--danger-color);
}

</style>

</head>

<body>

<div class="header">

    <div>

        <h1>
            Dashboard Active Directory
        </h1>

        <small style="color:#94a3b8;">

            Domínio:
            <b>$(HtmlEncode $domainName)</b>

            |

            Gerado em:
            $(
                (Get-Date).ToString(
                    "dd/MM/yyyy HH:mm:ss"
                )
            )

        </small>

    </div>

</div>

<div class="metrics">

    <div class="card">

        <h3>
            TOTAL CONSULTADO
        </h3>

        <p>
            $($res.Count)
        </p>

    </div>

    <div class="card">

        <h3>
            ENCONTRADOS
        </h3>

        <p style="color:var(--success-color);">

            $totEncontrados

        </p>

    </div>

    <div class="card">

        <h3>
            NÃO ENCONTRADOS
        </h3>

        <p style="color:var(--danger-color);">

            $totNaoEncontrados

        </p>

    </div>

    <div class="card">

        <h3>
            DUPLICADOS
        </h3>

        <p style="color:var(--warning-color);">

            $totDuplicados

        </p>

    </div>

    $metricasInventario

</div>

<div class="main-layout">

    <div class="panel">

        <h2>
            Estrutura de OUs
            ($(HtmlEncode $domainName))
        </h2>

        <p class="tree-hint">
            Clique em uma OU para filtrar os usuários.
            Use as setas &#9662;/&#9656; para recolher. Clique no domínio para limpar.
        </p>

        <div id="ouTree">

            $treeHtml

        </div>

    </div>

    <div class="panel">

        <h2>
            Usuários Mapeados
        </h2>

        <input
            type="text"
            id="searchInput"
            class="search-box"
            onkeyup="filterTable()"
            placeholder="Buscar por Nome, SAM, E-mail, OU ou Grupo..."
        >

        <table id="userTable">

            <thead>

                <tr>

                    <th>Status</th>

                    <th>
                        Usuário / Nome
                    </th>

                    <th>
                        E-mail
                    </th>

                    <th>
                        Ativo
                    </th>

                    <th>
                        Caminho OU
                    </th>

                    <th>
                        Grupos
                    </th>

                    <th>
                        Faltando
                    </th>

                    <th>
                        Último Logon
                    </th>

                </tr>

            </thead>

            <tbody>

                $($rowsHtml -join "`n")

            </tbody>

        </table>

    </div>

</div>

$inventarioHtml

<script>

var ouFiltro = null;

function aplicarFiltros() {

    var texto =
        document.getElementById("searchInput")
            .value
            .toLowerCase();

    var linhas =
        document.getElementById("userTable")
            .tBodies[0]
            .rows;

    for (
        var i = 0;
        i < linhas.length;
        i++
    ) {

        var linha = linhas[i];

        var visivel = true;

        if (ouFiltro) {

            var dn =
                (linha.getAttribute("data-ou") || "")
                    .toLowerCase();

            var alvo =
                ouFiltro.toLowerCase();

            visivel =
                dn === alvo ||
                dn.slice(-(alvo.length + 1)) === "," + alvo;
        }

        if (visivel && texto) {

            visivel =
                linha.textContent
                    .toLowerCase()
                    .indexOf(texto) > -1;
        }

        linha.style.display =
            visivel ? "" : "none";
    }
}

function filterTable() {

    aplicarFiltros();
}

function filtrarTabela(idTabela, texto) {

    var linhas =
        document.getElementById(idTabela)
            .tBodies[0]
            .rows;

    texto = (texto || "").toLowerCase();

    for (
        var i = 0;
        i < linhas.length;
        i++
    ) {

        var visivel =
            !texto ||
            linhas[i].textContent
                .toLowerCase()
                .indexOf(texto) > -1;

        linhas[i].style.display =
            visivel ? "" : "none";
    }
}

function selecionarOU(node) {

    var dn =
        node.getAttribute("data-dn") || "";

    var jaAtiva =
        node.classList.contains("active");

    var ativos =
        document.querySelectorAll(
            ".tree-node.active"
        );

    for (
        var i = 0;
        i < ativos.length;
        i++
    ) {

        ativos[i].classList.remove("active");
    }

    if (dn && !jaAtiva) {

        ouFiltro = dn;

        node.classList.add("active");
    }
    else {

        ouFiltro = null;
    }

    aplicarFiltros();
}

document
    .getElementById("ouTree")
    .addEventListener(
        "click",
        function (ev) {

            var alvo = ev.target;

            if (
                alvo.classList.contains(
                    "tree-toggle"
                )
            ) {

                var li = alvo.closest("li");

                li.classList.toggle("collapsed");

                alvo.textContent =
                    li.classList.contains(
                        "collapsed"
                    )
                        ? "\u25B8"
                        : "\u25BE";

                return;
            }

            var node = alvo.closest(".tree-node");

            if (node) {

                selecionarOU(node);
            }
        }
    );

</script>

</body>

</html>
"@

# ============================================================
# SALVAR HTML
# ============================================================

$htmlTemplate |
    Set-Content `
        -Path $fileHtml `
        -Encoding UTF8

# ============================================================
# ENVIO DE E-MAIL
# ============================================================

if ($EnviarEmail) {

    if (
        -not $SmtpServer -or
        -not $Remetente -or
        -not $Destinatarios
    ) {

        Write-Error -ErrorAction Continue `
            "Para enviar e-mail, informe -SmtpServer, -Remetente e -Destinatarios."
    }
    else {

        Write-Host ""
        Write-Host `
            "-> Enviando relatório por SMTP..." `
            -ForegroundColor Cyan

        $mailBody = @"
<html>

<body style="font-family: Segoe UI, Arial, sans-serif; color: #1e293b;">

<h2 style="color: #0284c7;">
    Relatório AD DS - $(HtmlEncode $domainName)
</h2>

<p>
    Gerado em:
    $((Get-Date).ToString("dd/MM/yyyy HH:mm:ss"))
</p>

<table cellpadding="6" style="border-collapse: collapse;">

    <tr>
        <td>Total consultado</td>
        <td><b>$($res.Count)</b></td>
    </tr>

    <tr>
        <td>Encontrados</td>
        <td><b>$totEncontrados</b></td>
    </tr>

    <tr>
        <td>Não encontrados</td>
        <td><b>$totNaoEncontrados</b></td>
    </tr>

    <tr>
        <td>Duplicados</td>
        <td><b>$totDuplicados</b></td>
    </tr>

</table>

<p>
    Os arquivos completos estão anexos (HTML, CSV e JSON).
</p>

</body>

</html>
"@

        $anexos = @(
            $fileHtml,
            $fileCsv,
            $fileJson
        )

        if (
            $excelOk -and
            (Test-Path -LiteralPath $fileXlsx)
        ) {

            $anexos += $fileXlsx
        }

        if ($Inventario) {

            foreach ($invFile in @(
                $fileInvUserCsv,
                $fileInvCompCsv
            )) {

                if (Test-Path -LiteralPath $invFile) {

                    $anexos += $invFile
                }
            }
        }

        $mailProps = @{
            SmtpServer  = $SmtpServer
            Port        = $SmtpPort
            From        = $Remetente
            To          = $Destinatarios
            Subject     = "[Relatório AD DS] Mapeamento de Usuários e OUs - $domainName"
            Body        = $mailBody
            BodyAsHtml  = $true
            Attachments = $anexos
        }

        if ($UseSSL) {

            $mailProps['UseSsl'] = $true
        }

        if ($Credenciais) {

            $mailProps['Credential'] = $Credenciais
        }

        try {

            Send-MailMessage `
                @mailProps `
                -ErrorAction Stop

            Write-Host `
                "E-mail enviado com sucesso." `
                -ForegroundColor Green
        }
        catch {

            Write-Error -ErrorAction Continue `
                "Falha ao enviar e-mail: $($_.Exception.Message)"
        }
    }
}

# ============================================================
# RESUMO FINAL
# ============================================================

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " PROCESSAMENTO CONCLUÍDO" -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "Domínio        : $domainName"
Write-Host "Total           : $($res.Count)"
Write-Host "Encontrados     : $totEncontrados" -ForegroundColor Green
Write-Host "Não encontrados : $totNaoEncontrados" -ForegroundColor Red
Write-Host "Duplicados      : $totDuplicados" -ForegroundColor Yellow
Write-Host ""

if ($Inventario) {

    Write-Host "Usuários no AD  : $($inventarioUsuarios.Count)"
    Write-Host "Computadores    : $($inventarioComputadores.Count)"

    if ($VerificarOnline) {

        $totOnline = @(
            $inventarioComputadores |
            Where-Object {
                $_.Online -eq "Sim"
            }
        ).Count

        Write-Host "Online          : $totOnline" -ForegroundColor Green
    }

    Write-Host ""
}

Write-Host "Arquivos gerados:" -ForegroundColor Cyan
Write-Host ""
Write-Host "[CSV]  $fileCsv"
Write-Host "[JSON] $fileJson"
Write-Host "[HTML] $fileHtml"

if ($excelOk) {

    Write-Host "[XLSX] $fileXlsx"
}

if ($Inventario) {

    if (Test-Path -LiteralPath $fileInvUserCsv) {

        Write-Host "[CSV]  $fileInvUserCsv"
    }

    if (Test-Path -LiteralPath $fileInvCompCsv) {

        Write-Host "[CSV]  $fileInvCompCsv"
    }
}

Write-Host ""