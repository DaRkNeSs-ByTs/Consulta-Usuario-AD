Import-Module ActiveDirectory

# Define caminhos para arquivos de saída e entrada
$desktopPath = [Environment]::GetFolderPath("Desktop")
$outputFileUsuarios = "$desktopPath\UsuariosInativos_AD.txt"
$outputFileComputadores = "$desktopPath\ComputadoresInativos_TodasMaquinas.txt"
$maquinasInputFile = "$desktopPath\MaquinasParaBusca.txt"  # Arquivo para nomes de máquinas

# Obtém data atual e informações da máquina
$dataAtual = Get-Date
$nomeMaquinaAtual = $env:COMPUTERNAME
$dataAtualInicio = $dataAtual.Date

# Solicita e valida a data limite
$dataLimiteInput = Read-Host "Digite a data limite para busca (dd/MM/yyyy)"
while ($true) {
    try {
        $dataLimite = [datetime]::ParseExact($dataLimiteInput, "dd/MM/yyyy", $null)
        if ($dataLimite -gt $dataAtual) {
            Write-Host "A data não pode ser futura. Digite uma data válida." -ForegroundColor Red
            $dataLimiteInput = Read-Host "Digite a data limite para busca (dd/MM/yyyy)"
        } else {
            break
        }
    } catch {
        Write-Host "Data inválida. Digite no formato dd/MM/yyyy (ex.: 01/03/2025)." -ForegroundColor Red
        $dataLimiteInput = Read-Host "Digite a data limite para busca (dd/MM/yyyy)"
    }
}

$dataLimiteFormatada = $dataLimite.ToString('dd/MM/yyyy')
$dataAtualFormatada = $dataAtual.ToString('dd/MM/yyyy')

# Busca usuários inativos
$usuariosInativos = Get-ADUser -Filter {Enabled -eq $true} -Properties LastLogonDate, Name, SamAccountName |
    Where-Object { $_.LastLogonDate -ne $null -and $_.LastLogonDate.Date -lt $dataLimite } |
    Select-Object Name, SamAccountName, @{Name='LastLogonDate';Expression={$_.LastLogonDate.ToString('dd/MM/yyyy')}} |
    Sort-Object Name

# Gera relatório de usuários
if ($usuariosInativos) {
    $conteudoUsuarios = "Relatório de Usuários Inativos no AD `n"
    $conteudoUsuarios += "Data Atual do Sistema: $dataAtualFormatada`n"
    $conteudoUsuarios += "Data Limite da Busca: $dataLimiteFormatada`n"
    $conteudoUsuarios += "Período de Busca: de $dataLimiteFormatada até $dataAtualFormatada`n"
    $conteudoUsuarios += "Máquina que gerou o relatório: $nomeMaquinaAtual`n"
    $conteudoUsuarios += "Total de Usuários Encontrados: $($usuariosInativos.Count)`n"
    $conteudoUsuarios += "Nota: Máquina do último logon não disponível.`n"
    $conteudoUsuarios += "--------------------------------------------`n`n"
    
    foreach ($usuario in $usuariosInativos) {
        $conteudoUsuarios += "Nome: $($usuario.Name)`n"
        $conteudoUsuarios += "Login: $($usuario.SamAccountName)`n"
        $conteudoUsuarios += "Último Login (LastLogonDate): $($usuario.LastLogonDate)`n"
        $conteudoUsuarios += "Máquina do Último Logon: Não disponível`n"
        $conteudoUsuarios += "--------------------------------------------`n"
    }
    
    $conteudoUsuarios | Out-File -FilePath $outputFileUsuarios -Encoding UTF8
    Write-Host "Relatório de usuários gerado com sucesso!" -ForegroundColor Green
} else {
    "Nenhum usuário inativo encontrado no período (de $dataLimiteFormatada até $dataAtualFormatada)" | Out-File -FilePath $outputFileUsuarios -Encoding UTF8
    Write-Host "Nenhum usuário inativo encontrado" -ForegroundColor Yellow
}

# Verifica se o arquivo MaquinasParaBusca.txt existe e contém nomes
$caracterBusca = $null
$caracterBuscaOriginal = $null
$maquinasFromFile = $null
if (Test-Path $maquinasInputFile) {
    $maquinasFromFile = Get-Content $maquinasInputFile | Where-Object { $_ -match '\S' } | ForEach-Object { $_.Trim().ToLower() }
    if ($maquinasFromFile) {
        Write-Host "Usando nomes de máquinas do arquivo $maquinasInputFile (case-insensitive)" -ForegroundColor Cyan
    }
}

if (-not $maquinasFromFile) {
    $caracterBuscaOriginal = Read-Host "Digite o nome da máquina para buscar ou deixe em branco para listar todas as máquinas"
    $caracterBusca = $caracterBuscaOriginal.ToLower()
}

# Busca computadores
if ($maquinasFromFile) {
    $resultadosInativos = @()
    $resultadosAtivos = @()
    $resultadosHoje = @()
    $nomesOriginais = Get-Content $maquinasInputFile | Where-Object { $_ -match '\S' } | ForEach-Object { $_.Trim() }
    
    foreach ($nomeMaquina in $maquinasFromFile) {
        $resultadosInativos += Get-ADComputer -Filter "Name -like '*$nomeMaquina*'" -Property Name, LastLogonDate, OperatingSystem, DistinguishedName | 
            Where-Object { ($_.LastLogonDate -eq $null) -or ($_.LastLogonDate -lt $dataLimite) }
        $resultadosAtivos += Get-ADComputer -Filter "Name -like '*$nomeMaquina*'" -Property Name, LastLogonDate, OperatingSystem, DistinguishedName | 
            Where-Object { $_.LastLogonDate -ge $dataLimite -and $_.LastLogonDate -le $dataAtual }
        $resultadosHoje += Get-ADComputer -Filter "Name -like '*$nomeMaquina*'" -Property Name, LastLogonDate, OperatingSystem, DistinguishedName | 
            Where-Object { $_.LastLogonDate -ge $dataAtualInicio -and $_.LastLogonDate -le $dataAtual }
    }
    
    $resultadosInativos = $resultadosInativos | Sort-Object Name -Unique | 
        Select-Object Name, OperatingSystem, DistinguishedName, @{Name='LastLogonDate';Expression={if ($_.LastLogonDate) {$_.LastLogonDate.ToString('dd/MM/yyyy HH:mm')} else {"Nunca"}}}
    $resultadosAtivos = $resultadosAtivos | Sort-Object Name -Unique | 
        Select-Object Name, OperatingSystem, DistinguishedName, @{Name='LastLogonDate';Expression={if ($_.LastLogonDate) {$_.LastLogonDate.ToString('dd/MM/yyyy HH:mm')} else {"Nunca"}}}
    $resultadosHoje = $resultadosHoje | Sort-Object Name -Unique | 
        Select-Object Name, OperatingSystem, DistinguishedName, @{Name='LastLogonDate';Expression={if ($_.LastLogonDate) {$_.LastLogonDate.ToString('dd/MM/yyyy HH:mm')} else {"Nunca"}}}
    
    $tituloInativos = "Relatório de Máquinas Inativas com Nomes do Arquivo (Sem Logon de $dataLimiteFormatada até $dataAtualFormatada)"
    $tituloAtivos = "Relatório de Máquinas Ativas com Nomes do Arquivo (Logon entre $dataLimiteFormatada e $dataAtualFormatada)"
    $tituloHoje = "Relatório de Máquinas Logadas Hoje com Nomes do Arquivo ($dataAtualFormatada)"
} else {
    if ([string]::IsNullOrWhiteSpace($caracterBusca)) {
        $resultadosInativos = Get-ADComputer -Filter * -Property Name, LastLogonDate, OperatingSystem, DistinguishedName | 
            Where-Object { ($_.LastLogonDate -eq $null) -or ($_.LastLogonDate -lt $dataLimite) } | 
            Sort-Object Name | 
            Select-Object Name, OperatingSystem, DistinguishedName, @{Name='LastLogonDate';Expression={if ($_.LastLogonDate) {$_.LastLogonDate.ToString('dd/MM/yyyy HH:mm')} else {"Nunca"}}}
        
        $resultadosAtivos = Get-ADComputer -Filter * -Property Name, LastLogonDate, OperatingSystem, DistinguishedName | 
            Where-Object { $_.LastLogonDate -ge $dataLimite -and $_.LastLogonDate -le $dataAtual } | 
            Sort-Object Name | 
            Select-Object Name, OperatingSystem, DistinguishedName, @{Name='LastLogonDate';Expression={if ($_.LastLogonDate) {$_.LastLogonDate.ToString('dd/MM/yyyy HH:mm')} else {"Nunca"}}}
        
        $resultadosHoje = Get-ADComputer -Filter * -Property Name, LastLogonDate, OperatingSystem, DistinguishedName | 
            Where-Object { $_.LastLogonDate -ge $dataAtualInicio -and $_.LastLogonDate -le $dataAtual } | 
            Sort-Object Name | 
            Select-Object Name, OperatingSystem, DistinguishedName, @{Name='LastLogonDate';Expression={if ($_.LastLogonDate) {$_.LastLogonDate.ToString('dd/MM/yyyy HH:mm')} else {"Nunca"}}}
        
        $tituloInativos = "Relatório de Todas as Máquinas Inativas (Sem Logon de $dataLimiteFormatada até $dataAtualFormatada)"
        $tituloAtivos = "Relatório de Todas as Máquinas Ativas (Logon entre $dataLimiteFormatada e $dataAtualFormatada)"
        $tituloHoje = "Relatório de Todas as Máquinas Logadas Hoje ($dataAtualFormatada)"
    } else {
        $resultadosInativos = Get-ADComputer -Filter "Name -like '*$caracterBusca*'" -Property Name, LastLogonDate, OperatingSystem, DistinguishedName | 
            Where-Object { ($_.LastLogonDate -eq $null) -or ($_.LastLogonDate -lt $dataLimite) } | 
            Sort-Object Name | 
            Select-Object Name, OperatingSystem, DistinguishedName, @{Name='LastLogonDate';Expression={if ($_.LastLogonDate) {$_.LastLogonDate.ToString('dd/MM/yyyy HH:mm')} else {"Nunca"}}}
        
        $resultadosAtivos = Get-ADComputer -Filter "Name -like '*$caracterBusca*'" -Property Name, LastLogonDate, OperatingSystem, DistinguishedName | 
            Where-Object { $_.LastLogonDate -ge $dataLimite -and $_.LastLogonDate -le $dataAtual } | 
            Sort-Object Name | 
            Select-Object Name, OperatingSystem, DistinguishedName, @{Name='LastLogonDate';Expression={if ($_.LastLogonDate) {$_.LastLogonDate.ToString('dd/MM/yyyy HH:mm')} else {"Nunca"}}}
        
        $resultadosHoje = Get-ADComputer -Filter "Name -like '*$caracterBusca*'" -Property Name, LastLogonDate, OperatingSystem, DistinguishedName | 
            Where-Object { $_.LastLogonDate -ge $dataAtualInicio -and $_.LastLogonDate -le $dataAtual } | 
            Sort-Object Name | 
            Select-Object Name, OperatingSystem, DistinguishedName, @{Name='LastLogonDate';Expression={if ($_.LastLogonDate) {$_.LastLogonDate.ToString('dd/MM/yyyy HH:mm')} else {"Nunca"}}}
        
        $tituloInativos = "Relatório de Máquinas Inativas com '$caracterBuscaOriginal' no Nome (Sem Logon de $dataLimiteFormatada até $dataAtualFormatada)"
        $tituloAtivos = "Relatório de Máquinas Ativas com '$caracterBuscaOriginal' no Nome (Logon entre $dataLimiteFormatada e $dataAtualFormatada)"
        $tituloHoje = "Relatório de Máquinas Logadas Hoje com '$caracterBuscaOriginal' no Nome ($dataAtualFormatada)"
    }
}

# Gera relatório de computadores
$conteudoComputadores = "Relatório de Computadores no AD `n"
$conteudoComputadores += "Data Atual do Sistema: $dataAtualFormatada`n"
$conteudoComputadores += "Data Limite da Busca: $dataLimiteFormatada`n"
$conteudoComputadores += "Período de Busca: de $dataLimiteFormatada até $dataAtualFormatada`n"
$conteudoComputadores += "Máquina que gerou o relatório: $nomeMaquinaAtual`n"
if ($maquinasFromFile) {
    $conteudoComputadores += "Nomes de Máquinas Buscados (do arquivo): $($nomesOriginais -join ', ')`n"
} elseif ($caracterBuscaOriginal) {
    $conteudoComputadores += "Nome de Máquina Buscado: $caracterBuscaOriginal`n"
} else {
    $conteudoComputadores += "Busca: Todas as máquinas`n"
}
$conteudoComputadores += "--------------------------------------------`n`n"

$conteudoComputadores += "$tituloInativos `n"
$conteudoComputadores += "Total de Computadores Inativos Encontrados: $($resultadosInativos.Count)`n"
$conteudoComputadores += "--------------------------------------------`n"
if ($resultadosInativos) {
    foreach ($computador in $resultadosInativos) {
        $conteudoComputadores += "Nome: $($computador.Name)`n"
        $conteudoComputadores += "Sistema Operacional: $($computador.OperatingSystem)`n"
        $conteudoComputadores += "Último Login: $($computador.LastLogonDate)`n"
        $conteudoComputadores += "DistinguishedName: $($computador.DistinguishedName)`n"
        $conteudoComputadores += "--------------------------------------------`n"
    }
} else {
    $conteudoComputadores += "Nenhum computador inativo encontrado.`n"
    $conteudoComputadores += "--------------------------------------------`n"
}

$conteudoComputadores += "`n$tituloAtivos `n"
$conteudoComputadores += "Total de Computadores Ativos Encontrados: $($resultadosAtivos.Count)`n"
$conteudoComputadores += "--------------------------------------------`n"
if ($resultadosAtivos) {
    foreach ($computador in $resultadosAtivos) {
        $conteudoComputadores += "Nome: $($computador.Name)`n"
        $conteudoComputadores += "Sistema Operacional: $($computador.OperatingSystem)`n"
        $conteudoComputadores += "Último Login: $($computador.LastLogonDate)`n"
        $conteudoComputadores += "DistinguishedName: $($computador.DistinguishedName)`n"
        $conteudoComputadores += "--------------------------------------------`n"
    }
} else {
    $conteudoComputadores += "Nenhum computador ativo encontrado.`n"
    $conteudoComputadores += "--------------------------------------------`n"
}

$conteudoComputadores += "`n$tituloHoje `n"
$conteudoComputadores += "Total de Computadores Logados Hoje Encontrados: $($resultadosHoje.Count)`n"
$conteudoComputadores += "--------------------------------------------`n"
if ($resultadosHoje) {
    foreach ($computador in $resultadosHoje) {
        $conteudoComputadores += "Nome: $($computador.Name)`n"
        $conteudoComputadores += "Sistema Operacional: $($computador.OperatingSystem)`n"
        $conteudoComputadores += "Último Login: $($computador.LastLogonDate)`n"
        $conteudoComputadores += "DistinguishedName: $($computador.DistinguishedName)`n"
        $conteudoComputadores += "--------------------------------------------`n"
    }
} else {
    $conteudoComputadores += "Nenhum computador logado hoje encontrado.`n"
    $conteudoComputadores += "--------------------------------------------`n"
}

$conteudoComputadores | Out-File -FilePath $outputFileComputadores -Encoding UTF8
Write-Host "Relatório de computadores gerado com sucesso!" -ForegroundColor Green

# Exibe resultados no console
Write-Host "====================================================================================" -ForegroundColor Yellow
Write-Host "`n============================== MÁQUINAS INATIVAS ====================================" -ForegroundColor Yellow
Write-Host "$tituloInativos" -ForegroundColor Yellow
Write-Host "Data Atual do Sistema: $dataAtualFormatada"
Write-Host "Data Limite da Busca: $dataLimiteFormatada"
Write-Host "Período de Busca: de $dataLimiteFormatada até $dataAtualFormatada"
if ($resultadosInativos) {
    Write-Host "Total de Computadores Inativos Encontrados: $($resultadosInativos.Count)" -ForegroundColor Yellow
    Write-Host "Máquinas Inativas:" -ForegroundColor Yellow
    Write-Host ("{0,-20} {1,-30} {2,-20}" -f "Nome", "Sistema Operacional", "Último Login")
    Write-Host ("-" * 20 + " " + "-" * 30 + " " + "-" * 20)
    foreach ($computador in $resultadosInativos) {
        $nome = $computador.Name
        $sistema = if ($computador.OperatingSystem) { $computador.OperatingSystem } else { "Não informado" }
        $ultimoLogin = $computador.LastLogonDate
        Write-Host ("{0,-20} {1,-30} {2,-20}" -f $nome, $sistema, $ultimoLogin)
    }
} else {
    $suffix = if ($maquinasFromFile) { " com nomes do arquivo" } elseif ($caracterBuscaOriginal) { " com '$caracterBuscaOriginal' no nome" } else { "" }
    $msgNenhumInativo = "Nenhum computador inativo encontrado (de $dataLimiteFormatada até $dataAtualFormatada)$suffix"
    Write-Host $msgNenhumInativo -ForegroundColor Red
}
Write-Host "====================================================================================" -ForegroundColor Yellow

Write-Host "`n============================== MÁQUINAS ATIVAS ====================================" -ForegroundColor Green
Write-Host "$tituloAtivos" -ForegroundColor Green
Write-Host "Data Atual do Sistema: $dataAtualFormatada"
Write-Host "Data Limite da Busca: $dataLimiteFormatada"
Write-Host "Período de Busca: de $dataLimiteFormatada até $dataAtualFormatada"
if ($resultadosAtivos) {
    Write-Host "Total de Computadores Ativos Encontrados: $($resultadosAtivos.Count)" -ForegroundColor Green
    Write-Host "Máquinas Ativas:" -ForegroundColor Green
    Write-Host ("{0,-20} {1,-30} {2,-20}" -f "Nome", "Sistema Operacional", "Último Login")
    Write-Host ("-" * 20 + " " + "-" * 30 + " " + "-" * 20)
    foreach ($computador in $resultadosAtivos) {
        $nome = $computador.Name
        $sistema = if ($computador.OperatingSystem) { $computador.OperatingSystem } else { "Não informado" }
        $ultimoLogin = $computador.LastLogonDate
        Write-Host ("{0,-20} {1,-30} {2,-20}" -f $nome, $sistema, $ultimoLogin)
    }
} else {
    $suffix = if ($maquinasFromFile) { " com nomes do arquivo" } elseif ($caracterBuscaOriginal) { " com '$caracterBuscaOriginal' no nome" } else { "" }
    $msgNenhumAtivo = "Nenhum computador ativo encontrado (de $dataLimiteFormatada até $dataAtualFormatada)$suffix"
    Write-Host $msgNenhumAtivo -ForegroundColor Red
}
Write-Host "====================================================================================" -ForegroundColor Green

Write-Host "`n============================== MÁQUINAS LOGADAS HOJE ====================================" -ForegroundColor Cyan
Write-Host "$tituloHoje" -ForegroundColor Cyan
Write-Host "Data Atual do Sistema: $dataAtualFormatada"
Write-Host "Busca de Logons Hoje: $dataAtualFormatada"
if ($resultadosHoje) {
    Write-Host "Total de Computadores Logados Hoje Encontrados: $($resultadosHoje.Count)" -ForegroundColor Cyan
    Write-Host "Máquinas Logadas Hoje:" -ForegroundColor Cyan
    Write-Host ("{0,-20} {1,-30} {2,-20}" -f "Nome", "Sistema Operacional", "Último Login")
    Write-Host ("-" * 20 + " " + "-" * 30 + " " + "-" * 20)
    foreach ($computador in $resultadosHoje) {
        $nome = $computador.Name
        $sistema = if ($computador.OperatingSystem) { $computador.OperatingSystem } else { "Não informado" }
        $ultimoLogin = $computador.LastLogonDate
        Write-Host ("{0,-20} {1,-30} {2,-20}" -f $nome, $sistema, $ultimoLogin)
    }
} else {
    $suffix = if ($maquinasFromFile) { " com nomes do arquivo" } elseif ($caracterBuscaOriginal) { " com '$caracterBuscaOriginal' no nome" } else { "" }
    $msgNenhumHoje = "Nenhum computador logado hoje encontrado ($dataAtualFormatada)$suffix"
    Write-Host $msgNenhumHoje -ForegroundColor Red
}
Write-Host "====================================================================================" -ForegroundColor Cyan
