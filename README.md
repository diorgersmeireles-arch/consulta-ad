# consulta-ad

Consulta usuários no Active Directory (AD DS) a partir de uma lista em CSV/TSV
e gera relatório em HTML, CSV e JSON — opcionalmente XLSX e envio por e-mail.
Com `-Inventario`, enumera também **todos os usuários e computadores do
domínio**, trazendo hostname e IP das máquinas.

## Estrutura

```
consulta-ad/
├─ consulta-ad.ps1                      # script principal
├─ exemplos/
│  └─ lista_exemplo.csv                 # arquivo de entrada de exemplo
├─ tests/
│  ├─ run_tests.ps1                     # suíte de testes (usa stub do módulo AD)
│  └─ stub/ActiveDirectory/             # módulo ActiveDirectory falso
├─ README.md
├─ .gitignore
└─ .gitattributes
```

## Requisitos

- Windows PowerShell 3.0+ ou PowerShell 7+
  - Server 2008 SP2 com WMF 3.0, 2008 R2/2012/2012 R2/2016+/2025 com o
    PowerShell nativo ou WMF atualizado
  - PowerShell 2.0 não é suportado (`ConvertTo-Json` e `[PSCustomObject]`
    exigem 3.0; o script declara `#Requires -Version 3.0`)
- Módulo `ActiveDirectory` (RSAT) — o script valida no início
- Opcional: módulo `ImportExcel` para `-ExportarExcel`
  (sem ele, o script avisa e gera apenas HTML/CSV/JSON)

## Uso

```powershell
# Consulta básica (saída em .\Resultado_AD ao lado do arquivo de entrada)
.\consulta-ad.ps1 -Arquivo .\exemplos\lista_exemplo.csv

# Pasta de saída, OU esperada e grupos obrigatórios (inclusive aninhados)
.\consulta-ad.ps1 -Arquivo lista.csv -PastaSaida .\saida `
    -OUEsperada Funcionarios -Grupo 'VPN-Remoto' -Aninhados

# Domínio específico e exportação Excel
.\consulta-ad.ps1 -Arquivo lista.csv -Server dc01.empresa.local -ExportarExcel

# Envio por e-mail
.\consulta-ad.ps1 -Arquivo lista.csv -EnviarEmail `
    -SmtpServer smtp.empresa.local -Remetente ad@empresa.local `
    -Destinatarios ti@empresa.local -UseSSL

# Inventário de todo o AD (usuários + computadores, hostname e IP)
.\consulta-ad.ps1 -Arquivo lista.csv -Inventario

# Inventário + ping/DNS das máquinas (confirma IP e status online)
.\consulta-ad.ps1 -Arquivo lista.csv -Inventario -VerificarOnline
```

## Arquivo de entrada

- Delimitador detectado automaticamente: `;`, `,` ou TAB (`.tsv`)
- Encoding detectado automaticamente: UTF-8 (com ou sem BOM) ou ANSI
- Colunas reconhecidas (nomes flexíveis, sem diferenciar maiúsculas):
  - Nome: `Nome`, `Nome Completo`, `Name`, `DisplayName`
  - E-mail: `E-mail`, `Email`, `E-Mail`, `mail`
  - Login: `SamAccountName`, `SAM`, `Login`, `Usuário`, `Usuario`
- Linhas sem nenhuma das três informações são ignoradas (com aviso)

A busca segue a precedência **SamAccountName → E-mail → Nome**, em lote
(`-LDAPFilter`), o que mantém o script rápido mesmo com listas grandes.

## Saídas

- `resultado_AD_<data>.html` — relatório navegável (árvore de OUs clicável,
  duplicados, inventário do AD com busca própria)
- `resultado_AD_<data>.csv` — delimitado por `;`
- `resultado_AD_<data>.json` — dados completos (inclui `Inventario` com
  `-Inventario`)
- `resultado_AD_<data>.xlsx` — apenas com `-ExportarExcel` (uma aba por
  seção: usuários da lista, inventário de usuários, inventário de
  computadores)
- `inventario_usuarios_<data>.csv` e `inventario_computadores_<data>.csv` —
  apenas com `-Inventario`

## Observações

- `UltimoLogon` usa `LastLogonDate`, derivado de `LastLogonTimestamp`
  (replicado, precisão de ~14 dias; pode vir vazio em contas que nunca logaram)
- `-OUEsperada` aceita o nome exato de uma OU (ex.: `Funcionarios`) ou o DN
  completo/início do DN (ex.: `OU=Funcionarios,DC=empresa,DC=local`)
- `Duplicados` lista os `SamAccountName` dos candidatos quando o nome retorna
  mais de um usuário
- Na árvore de OUs do HTML, clicar em uma OU filtra a tabela (incluindo
  sub-OUs); clicar de novo ou no domínio limpa o filtro; ▾/▸ recolhe e expande
- `-Inventario` traz o IP do atributo `IPv4Address` do computador (populado
  por DHCP/DNS integrado); `-VerificarOnline` complementa com ping e resolução
  DNS — mais lento em domínios grandes e sujeito a bloqueio de ICMP/firewall
- Computadores sem `IPv4Address` nem DNS resolvível aparecem com IP vazio

## Desenvolvimento e testes

Os scripts `.ps1`/`.psm1` do repositório são gravados em **UTF-8 com BOM** —
obrigatório para que o Windows PowerShell 5.1 leia acentos e caracteres
especiais corretamente. Mantenha o BOM ao salvar.

O conteúdo gerado é **100% ASCII na estrutura**: o script não usa emojis; os
ícones da árvore e as setas de recolher/expandir são entidades HTML
(`&#8226;`, `&#9679;`, `&#9662;`/`&#9656;`), que renderizam em qualquer
navegador. Textos em português (acentos) continuam em UTF-8.

A suíte usa um módulo `ActiveDirectory` falso (`tests/stub`), portanto **não
consulta o AD real**:

```powershell
pwsh -NoProfile -File .\tests\run_tests.ps1
```

Ela executa o script nos dois hosts (pwsh e Windows PowerShell 5.1), nos três
encodings de entrada, e valida HTML/CSV/JSON, `-OUEsperada`, `-Grupo`,
`-Aninhados`, validação de e-mail, `-Inventario` (usuários, computadores, IP)
e códigos de erro.
