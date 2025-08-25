# --- Pré-requisitos ---
# 1. Instale o módulo Microsoft Graph PowerShell SDK, se ainda não o tiver:
#    Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force

# --- Conexão e Permissões ---
# 2. Conecte-se ao Microsoft Graph com as permissões necessárias.
#    Você precisará de uma conta com privilégios de Administrador Global ou um papel equivalente
#    para conceder consentimento a essas permissões na primeira execução.

# Permissões (Scopes) necessárias para o script:
# Group.Read.All: Para ler todas as propriedades de grupos.
# GroupMember.Read.All: Para ler os membros de todos os grupos.
# User.Read.All: Para obter os nomes de exibição (DisplayName) e UserPrincipalName de proprietários e membros.
$RequiredScopes = @(
    "Group.Read.All",
    "GroupMember.Read.All",
    "User.Read.All"
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

} Catch {
    Write-Error "ERRO: Falha ao conectar ao Microsoft Graph."
    Write-Error "Por favor, verifique suas permissões no Azure AD e a conectividade de rede."
    Write-Error "Detalhes do Erro: $($_.Exception.Message)"
    Exit 1 # Sai do script com erro
}

# --- Coleta de Dados: Análise de Grupos e Membros ---
Write-Host "Iniciando coleta de informações de grupos e seus membros."
Write-Host "Isso pode levar alguns minutos, especialmente para tenants com muitos grupos ou grupos grandes."

$groupsData = @()

Try {
    # Obtém todos os grupos com as propriedades necessárias
    $allGroups = Get-MgGroup -All -Property "id,displayName,description,groupTypes,mailEnabled,securityEnabled,visibility,createdDateTime" -ErrorAction Stop

    If (-not $allGroups) {
        Write-Warning "AVISO: Nenhum grupo foi retornado pelo comando Get-MgGroup. O arquivo CSV pode ficar vazio."
        Write-Warning "Verifique se o seu tenant possui grupos ou se a conta de autenticação tem permissão para lê-los."
    } Else {
        Write-Host "Total de grupos encontrados: $($allGroups.Count)"
    }

    $processedGroupsCount = 0
    ForEach ($group in $allGroups) {
        $processedGroupsCount++
        # Progresso visual para tenants grandes
        Write-Progress -Activity "Processando Grupos" -Status "Grupo: $($group.DisplayName) ($processedGroupsCount de $($allGroups.Count))" -PercentComplete (($processedGroupsCount / $allGroups.Count) * 100)

        # Classificação do Tipo de Grupo
        $groupType = "Outro"
        If ($group.GroupTypes -contains "Unified") {
            $groupType = "Microsoft 365 Group"
        } ElseIf ($group.MailEnabled -eq $true -and $group.SecurityEnabled -eq $false) {
            $groupType = "Distribution Group"
        } ElseIf ($group.MailEnabled -eq $false -and $group.SecurityEnabled -eq $true) {
            $groupType = "Security Group (Non-Mail Enabled)"
        } ElseIf ($group.MailEnabled -eq $true -and $group.SecurityEnabled -eq $true) {
            $groupType = "Security Group (Mail Enabled)" # Um grupo de segurança pode ser mail-enabled
        }

        # 1. Obter Proprietários do Grupo
        $ownerUPNs = @()
        $ownerCount = 0
        Try {
            # Seleciona apenas o UserPrincipalName dos proprietários
            $owners = Get-MgGroupOwner -GroupId $group.Id -All -ErrorAction SilentlyContinue | Select-Object UserPrincipalName
            If ($owners) {
                $ownerUPNs = ($owners.UserPrincipalName | Sort-Object) -join "; "
                $ownerCount = $owners.Count
            }
        } Catch {
            Write-Warning "AVISO: Não foi possível obter proprietários para o grupo '$($group.DisplayName)'. Erro: $($_.Exception.Message)"
            $ownerUPNs = "ERRO ao obter proprietários"
        }

        # 2. Obter Membros do Grupo
        $memberUPNs = @()
        $memberCount = 0
        Try {
            # Seleciona apenas o UserPrincipalName dos membros
            $members = Get-MgGroupMember -GroupId $group.Id -All -ErrorAction SilentlyContinue | Select-Object UserPrincipalName
            If ($members) {
                $memberUPNs = ($members.UserPrincipalName | Sort-Object) -join "; "
                $memberCount = $members.Count
            }
        } Catch {
            Write-Warning "AVISO: Não foi possível obter membros para o grupo '$($group.DisplayName)'. Erro: $($_.Exception.Message)"
            $memberUPNs = "ERRO ao obter membros"
        }

        # Adiciona os dados do grupo ao array principal
        $groupsData += [PSCustomObject]@{
            GroupName           = $group.DisplayName
            GroupId             = $group.Id
            GroupType           = $groupType
            MailEnabled         = $group.MailEnabled
            SecurityEnabled     = $group.SecurityEnabled
            Visibility          = $group.Visibility # Public/Private para M365 Groups
            Description         = $group.Description
            CreatedDateTime     = $group.CreatedDateTime
            OwnerCount          = $ownerCount
            Owners              = $ownerUPNs # Agora uma string de UPNs separados por ;
            MemberCount         = $memberCount
            Members             = $memberUPNs # Agora uma string de UPNs separados por ;
        }
    }

} Catch {
    Write-Error "ERRO durante a coleta de dados de grupos."
    Write-Error "Detalhes do Erro: $($_.Exception.Message)"
    Exit 1
}

# --- Exportar para CSV ---
Write-Host "Coleta de dados de grupos concluída. Exportando para CSV..."

# Define o caminho de saída para o arquivo CSV
$outputPath = Join-Path (Get-Location) "M365_Tenant_Groups_Members_Report.csv"

Try {
    # Exporta os dados para um arquivo CSV
    If ($groupsData.Count -gt 0) {
        $groupsData | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8 -Force
        Write-Host "Relatório de Análise de Grupos e Membros salvo com sucesso em: $outputPath"
    } Else {
        Write-Warning "AVISO: Não há dados de grupos para exportar. O arquivo CSV pode estar vazio ou não foi criado."
    }

} Catch {
    Write-Error "ERRO ao exportar dados de grupos para CSV."
    Write-Error "Detalhes do Erro: $($_.Exception.Message)"
    Exit 1
}

Write-Host "--- Análise de Grupos e Membros Concluída ---"

# --- Desconectar Microsoft Graph ---
Try {
    Disconnect-MgGraph
    Write-Host "Desconectado do Microsoft Graph."
} Catch {
    Write-Warning "AVISO: Não foi possível desconectar do Microsoft Graph. Pode não haver uma conexão ativa."
}