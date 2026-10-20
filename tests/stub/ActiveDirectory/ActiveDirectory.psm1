# Módulo ActiveDirectory FALSO — apenas para testes do consulta-ad.ps1

$script:Dominio = [PSCustomObject]@{
    DistinguishedName = 'DC=corp,DC=local'
    DNSRoot = 'corp.local'
}

$script:OUs = @(
    [PSCustomObject]@{
        Name = 'Funcionarios'
        DistinguishedName = 'OU=Funcionarios,DC=corp,DC=local'
        CanonicalName = 'corp.local/Funcionarios'
    }
    [PSCustomObject]@{
        Name = 'TI'
        DistinguishedName = 'OU=TI,OU=Funcionarios,DC=corp,DC=local'
        CanonicalName = 'corp.local/Funcionarios/TI'
    }
    [PSCustomObject]@{
        Name = 'Grupos'
        DistinguishedName = 'OU=Grupos,DC=corp,DC=local'
        CanonicalName = 'corp.local/Grupos'
    }
    [PSCustomObject]@{
        Name = 'RH'
        DistinguishedName = 'OU=RH,DC=corp,DC=local'
        CanonicalName = 'corp.local/RH'
    }
    [PSCustomObject]@{
        Name = 'Arquivo'
        DistinguishedName = 'OU=Arquivo,CN=Users,DC=corp,DC=local'
        CanonicalName = 'corp.local/Users/Arquivo'
    }
)

$script:Usuarios = @(
    [PSCustomObject]@{
        SamAccountName = 'jsilva'
        DisplayName = 'José Silva'
        Name = 'José Silva'
        Mail = 'jose.silva@corp.local'
        UserPrincipalName = 'jsilva@corp.local'
        Enabled = $true
        LastLogonDate = (Get-Date).AddDays(-3)
        MemberOf = @('CN=TI,OU=Grupos,DC=corp,DC=local')
        DistinguishedName = 'CN=José Silva,OU=TI,OU=Funcionarios,DC=corp,DC=local'
        Department = 'TI'
        Title = 'Analista de Sistemas'
    }
    [PSCustomObject]@{
        SamAccountName = 'mdavila'
        DisplayName = "Marco D'Ávila"
        Name = "Marco D'Ávila"
        Mail = ''
        UserPrincipalName = 'mdavila@corp.local'
        Enabled = $true
        LastLogonDate = $null
        MemberOf = @()
        DistinguishedName = "CN=Marco D'Ávila,OU=RH,DC=corp,DC=local"
        Department = 'RH'
        Title = 'Coordenador'
    }
    [PSCustomObject]@{
        SamAccountName = 'mdup1'
        DisplayName = 'Maria Duplicada'
        Name = 'Maria Duplicada'
        Mail = 'maria1@corp.local'
        UserPrincipalName = 'mdup1@corp.local'
        Enabled = $true
        LastLogonDate = (Get-Date).AddDays(-30)
        MemberOf = @()
        DistinguishedName = 'CN=Maria Duplicada I,OU=RH,DC=corp,DC=local'
        Department = 'RH'
        Title = 'Assistente'
    }
    [PSCustomObject]@{
        SamAccountName = 'mdup2'
        DisplayName = 'Maria Duplicada'
        Name = 'Maria Duplicada'
        Mail = 'maria2@corp.local'
        UserPrincipalName = 'mdup2@corp.local'
        Enabled = $true
        LastLogonDate = $null
        MemberOf = @()
        DistinguishedName = 'CN=Maria Duplicada II,OU=RH,DC=corp,DC=local'
        Department = 'RH'
        Title = 'Assistente'
    }
    [PSCustomObject]@{
        SamAccountName = 'jdesat'
        DisplayName = 'João Desativado'
        Name = 'João Desativado'
        Mail = 'joao@corp.local'
        UserPrincipalName = 'jdesat@corp.local'
        Enabled = $false
        LastLogonDate = $null
        MemberOf = @()
        DistinguishedName = 'CN=João Desativado,OU=Arquivo,CN=Users,DC=corp,DC=local'
        Department = 'TI'
        Title = 'Estagiário'
    }
)

foreach ($u in $script:Usuarios) {

    $u | Add-Member `
        -NotePropertyName Created `
        -NotePropertyValue (Get-Date).AddYears(-2) `
        -Force

    $u | Add-Member `
        -NotePropertyName Description `
        -NotePropertyValue '' `
        -Force
}

$script:Computadores = @(
    [PSCustomObject]@{
        Name = 'SRV-DC01'
        DNSHostName = 'srv-dc01.corp.local'
        SamAccountName = 'SRV-DC01$'
        Enabled = $true
        OperatingSystem = 'Windows Server 2022 Standard'
        OperatingSystemVersion = '10.0 (20348)'
        Description = 'Controlador de dominio'
        LastLogonDate = (Get-Date).AddHours(-1)
        Created = (Get-Date).AddYears(-3)
        IPv4Address = '10.0.0.10'
        DistinguishedName = 'CN=SRV-DC01,OU=TI,OU=Funcionarios,DC=corp,DC=local'
    }
    [PSCustomObject]@{
        Name = 'DESK-JSILVA'
        DNSHostName = 'desk-jsilva.corp.local'
        SamAccountName = 'DESK-JSILVA$'
        Enabled = $true
        OperatingSystem = 'Windows 11 Pro'
        OperatingSystemVersion = '10.0 (22631)'
        Description = 'Estacao do Jose'
        LastLogonDate = (Get-Date).AddDays(-1)
        Created = (Get-Date).AddMonths(-8)
        IPv4Address = '10.0.1.55'
        DistinguishedName = 'CN=DESK-JSILVA,OU=Funcionarios,DC=corp,DC=local'
    }
    [PSCustomObject]@{
        Name = 'DESK-OLD'
        DNSHostName = 'desk-old.corp.local'
        SamAccountName = 'DESK-OLD$'
        Enabled = $false
        OperatingSystem = 'Windows 7 Professional'
        OperatingSystemVersion = '6.1 (7601)'
        Description = 'Maquina antiga'
        LastLogonDate = $null
        Created = (Get-Date).AddYears(-6)
        IPv4Address = ''
        DistinguishedName = 'CN=DESK-OLD,OU=Arquivo,CN=Users,DC=corp,DC=local'
    }
)

$script:Grupos = @(
    [PSCustomObject]@{
        Name = 'TI'
        DistinguishedName = 'CN=TI,OU=Grupos,DC=corp,DC=local'
        Members = @('CN=José Silva,OU=TI,OU=Funcionarios,DC=corp,DC=local')
    }
    [PSCustomObject]@{
        Name = 'VPN-Remoto'
        DistinguishedName = 'CN=VPN-Remoto,OU=Grupos,DC=corp,DC=local'
        Members = @('CN=TI,OU=Grupos,DC=corp,DC=local')
    }
    [PSCustomObject]@{
        Name = 'Todos'
        DistinguishedName = 'CN=Todos,OU=Grupos,DC=corp,DC=local'
        Members = @('CN=VPN-Remoto,OU=Grupos,DC=corp,DC=local')
    }
)

function Unescape-LDAPValue {
    param([string]$Valor)

    $sb = New-Object System.Text.StringBuilder
    $i = 0

    while ($i -lt $Valor.Length) {
        $c = $Valor[$i]

        if ($c -eq '\' -and ($i + 2) -lt $Valor.Length) {
            $hex = $Valor.Substring($i + 1, 2)

            switch ($hex) {
                '5c' { [void]$sb.Append('\'); $i += 3; continue }
                '2a' { [void]$sb.Append('*'); $i += 3; continue }
                '28' { [void]$sb.Append('('); $i += 3; continue }
                '29' { [void]$sb.Append(')'); $i += 3; continue }
                '00' { [void]$sb.Append([char]0); $i += 3; continue }
            }
        }

        [void]$sb.Append($c)
        $i++
    }

    return $sb.ToString()
}

function Testar-FiltroUsuarios {
    param([string]$Filtro, [object]$Usuario)

    $atributos = @{
        'sAMAccountName' = 'SamAccountName'
        'mail' = 'Mail'
        'userPrincipalName' = 'UserPrincipalName'
        'displayName' = 'DisplayName'
        'name' = 'Name'
    }

    $clausulas = [regex]::Matches(
        $Filtro,
        '\((?<attr>[A-Za-z0-9.]+)=(?<val>(?:\\.|[^()])*)\)'
    )

    foreach ($cl in $clausulas) {
        $prop = $atributos[$cl.Groups['attr'].Value]

        if (-not $prop) {
            continue
        }

        $esperado = Unescape-LDAPValue $cl.Groups['val'].Value
        $atual = "$($Usuario.$prop)"

        if ($atual -and [string]::Equals(
                $atual,
                $esperado,
                [System.StringComparison]::OrdinalIgnoreCase
            )) {
            return $true
        }
    }

    return $false
}

function Get-GruposCadeia {
    param([string]$DN)

    $resultado = New-Object 'System.Collections.Generic.HashSet[string]'
    $visitado = New-Object 'System.Collections.Generic.HashSet[string]'
    $fila = New-Object 'System.Collections.Generic.Queue[string]'

    $fila.Enqueue($DN)

    while ($fila.Count -gt 0) {
        $atual = $fila.Dequeue()

        if (-not $visitado.Add($atual)) {
            continue
        }

        foreach ($g in $script:Grupos) {
            if ($g.Members -contains $atual) {
                if ($resultado.Add($g.DistinguishedName)) {
                    $fila.Enqueue($g.DistinguishedName)
                }
            }
        }
    }

    return $resultado
}

function Get-ADDomain {
    [CmdletBinding()]
    param([string]$Server)

    return $script:Dominio
}

function Get-ADOrganizationalUnit {
    [CmdletBinding()]
    param(
        [string]$Server,
        $Filter,
        [string[]]$Properties
    )

    return $script:OUs
}

function Get-ADUser {
    [CmdletBinding()]
    param(
        [string]$Server,
        [string[]]$Properties,
        $Filter,
        [string]$LDAPFilter
    )

    if ($LDAPFilter) {
        return @(
            $script:Usuarios |
            Where-Object {
                Testar-FiltroUsuarios -Filtro $LDAPFilter -Usuario $_
            }
        )
    }

    return @($script:Usuarios)
}

function Get-ADComputer {
    [CmdletBinding()]
    param(
        [string]$Server,
        [string[]]$Properties,
        $Filter,
        [string]$LDAPFilter
    )

    if ($LDAPFilter) {
        return @(
            $script:Computadores |
            Where-Object {
                Testar-FiltroUsuarios -Filtro $LDAPFilter -Usuario $_
            }
        )
    }

    return @($script:Computadores)
}

function Get-ADGroup {
    [CmdletBinding()]
    param(
        [string]$Server,
        $Filter,
        [string]$LDAPFilter,
        [string[]]$Properties
    )

    if ($LDAPFilter -match '^\(member:1\.2\.840\.113556\.1\.4\.1941:=(?<dn>.+)\)$') {
        $dn = Unescape-LDAPValue $Matches['dn']
        $cadeia = @(Get-GruposCadeia $dn)

        return @(
            $script:Grupos |
            Where-Object {
                $cadeia -contains $_.DistinguishedName
            }
        )
    }

    return @()
}
