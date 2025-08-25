# --- Pré-requisitos ---
# 1. Instale o módulo Microsoft Graph PowerShell SDK, se ainda não o tiver:
#    Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force

# --- Conexão e Permissões ---
# 2. Conecte-se ao Microsoft Graph com as permissões necessárias.
#    Você precisará de uma conta com privilégios de Administrador Global ou um papel equivalente
#    para conceder consentimento a essas permissões na primeira execução.

# Permissões (Scopes) necessárias para o script:
# User.Read.All: Para ler todos os perfis de usuário.
# Directory.Read.All: Para ler informações de diretório, incluindo funções administrativas e SKUs de licença.
# Organization.Read.All: Para obter detalhes dos SKUs de licença (e.g., transformar IDs em nomes legíveis).
# NOTA: AuditLog.Read.All e a propriedade signInActivity foram removidas devido à ausência de licença Premium.
$RequiredScopes = @(
    "User.Read.All",
    "Directory.Read.All",
    "Organization.Read.All"
)

Write-Host "Iniciando conexão ao Microsoft Graph com as permissões necessárias..."
Try {
    # Tenta desconectar para garantir uma conexão limpa
    If (Get-MgContext -ErrorAction SilentlyContinue) {
        Disconnect-MgGraph
        Write-Host "Contexto MgGraph anterior desconectado."
    }

    Connect-MgGraph -Scopes $RequiredScopes -ErrorAction Stop
    Write-Host "Conexão estabelecida com sucesso."

    Write-Host "Coletando SKUs de licença para mapeamento (ID para Nome legível)..."
    $allSkus = Get-MgSubscribedSku -All -ErrorAction Stop | Select-Object SkuId, SkuPartNumber
    $skuMap = @{}
    ForEach ($sku in $allSkus) {
        $skuMap[$sku.SkuId] = $sku.SkuPartNumber
    }
    Write-Host "SKUs de licença mapeados."

} Catch {
    Write-Error "ERRO: Falha ao conectar ao Microsoft Graph ou obter SKUs."
    Write-Error "Por favor, verifique suas permissões no Azure AD e a conectividade de rede."
    Write-Error "Detalhes do Erro: $($_.Exception.Message)"
    Exit 1 # Sai do script com erro
}

# --- Coleta de Dados: Inventário de Usuários e Licenças ---
Write-Host "Iniciando coleta de informações de usuários e licenças."
Write-Host "Isso pode levar alguns minutos dependendo do número de usuários e do desempenho da API."
Write-Host "NOTA: Dados de 'Último Login' e 'Inatividade' serão N/A ou baseados apenas em 'Conta Desabilitada' devido à ausência de licença Azure AD Premium para acesso a 'signInActivity'."

$usersData = @()

Try {
    # Obtém todos os usuários com as propriedades necessárias.
    # A propriedade 'signInActivity' FOI REMOVIDA daqui.
    $allUsers = Get-MgUser -All -Property "id,displayName,userPrincipalName,accountEnabled,assignedLicenses,department,jobTitle" -ErrorAction Stop

    If (-not $allUsers) {
        Write-Warning "AVISO: Nenhum usuário foi retornado pelo comando Get-MgUser. O arquivo CSV pode ficar vazio."
        Write-Warning "Verifique se o seu tenant possui usuários ou se a conta de autenticação tem permissão para lê-los."
    } Else {
        Write-Host "Total de usuários encontrados: $($allUsers.Count)"
    }

    ForEach ($user in $allUsers) {
        # Progresso visual para tenants grandes
        Write-Progress -Activity "Processando Usuários" -Status "Usuário: $($user.UserPrincipalName)" -PercentComplete (($usersData.Count / $allUsers.Count) * 100)

        # 1. Processar Licenças Atribuídas
        $userLicenses = @()
        If ($user.AssignedLicenses) {
            ForEach ($license in $user.AssignedLicenses) {
                $licenseName = If ($skuMap.ContainsKey($license.SkuId)) { $skuMap[$license.SkuId] } Else { "SKU Desconhecido ($($license.SkuId))" }

                $disabledServicePlans = @()
                If ($license.DisabledPlans) {
                    $disabledServicePlans = ($license.DisabledPlans.ServicePlanName | Sort-Object) -join ", "
                }

                $userLicenses += [PSCustomObject]@{
                    LicenseName        = $licenseName
                    DisabledServices   = $disabledServicePlans
                    AssignedSKUId      = $license.SkuId
                }
            }
        }

        # 2. Processar Funções Administrativas
        $userAdminRoles = @()
        Try {
            # Obtém a lista de grupos e funções de diretório que o usuário é membro
            $memberOf = Get-MgUserMemberOf -UserId $user.Id -All -ErrorAction SilentlyContinue

            ForEach ($member in $memberOf) {
                # Verifica se o membro é uma função de diretório (directoryRole) e não um grupo
                If ($member.AdditionalProperties -and $member.AdditionalProperties.ContainsKey('odata.type') -and $member.AdditionalProperties['odata.type'] -eq '#microsoft.graph.directoryRole') {
                    # Obtém os detalhes da função de diretório para o nome de exibição
                    If ($member.Id) {
                        $roleDetails = Get-MgDirectoryRole -DirectoryRoleId $member.Id -ErrorAction SilentlyContinue
                        If ($roleDetails) {
                            $userAdminRoles += $roleDetails.DisplayName
                        }
                    }
                }
            }
            $userAdminRoles = ($userAdminRoles | Sort-Object | Select-Object -Unique) -join ", " # Junta em uma string separada por vírgulas, remove duplicatas
        } Catch {
            Write-Warning "AVISO: Não foi possível obter funções administrativas para o usuário $($user.UserPrincipalName). Erro: $($_.Exception.Message)"
            $userAdminRoles = "ERRO ao obter funções"
        }

        # 3. Processar Último Login e Status de Inatividade/Órfão (Agora sem signInActivity)
        # Estas colunas serão N/A ou baseadas apenas na conta estar desabilitada
        $lastSignInUtc = "N/A (Requer AAD Premium P1/P2)"
        $lastSignInLocal = "N/A (Requer AAD Premium P1/P2)"
        $isInactive90Days = $false
        $isPotentiallyOrphaned = $false

        If ($user.AccountEnabled -eq $false) {
            $isInactive90Days = $true
            $isPotentiallyOrphaned = $true
            $lastSignInUtc = "Conta Desabilitada"
            $lastSignInLocal = "Conta Desabilitada"
        }

        # Adiciona os dados do usuário ao array principal
        $usersData += [PSCustomObject]@{
            DisplayName           = $user.DisplayName
            UserPrincipalName     = $user.UserPrincipalName
            Department            = $user.Department
            JobTitle              = $user.JobTitle
            AccountEnabled        = $user.AccountEnabled
            LastSignInDateTimeUTC = $lastSignInUtc
            LastSignInDateTimeLocal = $lastSignInLocal
            AssignedLicenses      = ($userLicenses | ConvertTo-Json -Compress) # Armazena como string JSON para manter a estrutura
            AdministrativeRoles   = $userAdminRoles
            IsInactive90Days      = $isInactive90Days
            IsPotentiallyOrphaned = $isPotentiallyOrphaned
        }
    }

} Catch {
    Write-Error "ERRO durante a coleta de dados de usuários."
    Write-Error "Detalhes do Erro: $($_.Exception.Message)"
    Exit 1
}

# --- Exportar para CSV ---
Write-Host "Coleta de dados concluída. Exportando para CSV..."

# Define o caminho de saída para o arquivo CSV
$outputPath = Join-Path (Get-Location) "M365_Tenant_Users_Licenses_Report.csv"

Try {
    # Exporta os dados para um arquivo CSV
    If ($usersData.Count -gt 0) {
        $usersData | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8 -Force
        Write-Host "Relatório de Inventário de Usuários e Licenças salvo com sucesso em: $outputPath"
    } Else {
        Write-Warning "AVISO: Não há dados para exportar. O arquivo CSV pode estar vazio ou não foi criado."
    }

} Catch {
    Write-Error "ERRO ao exportar dados para CSV."
    Write-Error "Detalhes do Erro: $($_.Exception.Message)"
    Exit 1
}

Write-Host "--- Análise de Inventário de Usuários e Licenças Concluída ---"

# --- Desconectar Microsoft Graph ---
Try {
    Disconnect-MgGraph
    Write-Host "Desconectado do Microsoft Graph."
} Catch {
    Write-Warning "AVISO: Não foi possível desconectar do Microsoft Graph. Pode não haver uma conexão ativa."
}