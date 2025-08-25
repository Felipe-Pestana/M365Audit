# --- Pré-requisitos ---
# 1. Instale o módulo Microsoft Graph PowerShell SDK, se ainda não o tiver:
#    Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force

# --- Conexão e Permissões ---
# 2. Conecte-se ao Microsoft Graph com as permissões necessárias.
#    Você precisará de uma conta com privilégios de Administrador Global ou um papel equivalente
#    para conceder consentimento a essas permissões na primeira execução.

# Permissões (Scopes) necessárias para o script:
# Directory.Read.All: Para configurações de diretório (ex: políticas de senha, guests).
# Policy.Read.All: Para políticas como Acesso Condicional e Security Defaults.
# SecurityEvents.Read.All: Para Secure Score.
# Organization.Read.All: Informações gerais da organização.
# AuditLog.Read.All: Embora não usemos diretamente, essa permissão é para contexto de segurança.
$RequiredScopes = @(
    "Directory.Read.All",
    "Policy.Read.All",
    "SecurityEvents.Read.All",
    "Organization.Read.All",
    "AuditLog.Read.All"
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

    # Captura o TenantId para uso em chamadas específicas (ex: Get-MgOrganization)
    $tenantId = (Get-MgContext).TenantId

} Catch {
    Write-Error "ERRO CRÍTICO: Falha ao conectar ao Microsoft Graph."
    Write-Error "Por favor, verifique suas permissões no Azure AD e a conectividade de rede."
    Write-Error "Detalhes do Erro: $($_.Exception.Message)"
    Exit 1 # Sai do script com erro
}

# Define o diretório de saída para os relatórios
$outputDirectory = Join-Path (Get-Location) "M365_Tenant_Security_Reports"
If (-not (Test-Path $outputDirectory)) {
    Try {
        New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
        Write-Host "Diretório de saída '$outputDirectory' criado."
    } Catch {
        Write-Error "ERRO CRÍTICO: Não foi possível criar o diretório de saída '$outputDirectory'. Verifique as permissões de escrita. Detalhes: $($_.Exception.Message)"
        Exit 1
    }
} Else {
    Write-Host "Diretório de saída '$outputDirectory' já existe."
}

# --- 1. Coleta de Dados: Políticas de Senha do Azure AD e Security Defaults ---
Write-Host "`n--- Iniciando coleta de Políticas de Senha e Security Defaults ---"
$passwordSecurityData = @()

Try {
    # Obter política de proteção de senha (Smart Lockout, Custom Banned Passwords)
    $authMethodPolicy = Get-MgPolicyAuthenticationMethodPolicy -ErrorAction SilentlyContinue
    $passwordProtection = $authMethodPolicy.PasswordProtection
    
    If ($passwordProtection) {
        $passwordSecurityData += [PSCustomObject]@{
            Setting = "Smart Lockout Enabled"
            Value = If ($passwordProtection.ProtectingSmartlockout -eq "Enabled") { "Sim" } Else { "Não" }
        }
        $passwordSecurityData += [PSCustomObject]@{
            Setting = "Smart Lockout Threshold"
            Value = $passwordProtection.BlockedPasswordCountThreshold
        }
        $passwordSecurityData += [PSCustomObject]@{
            Setting = "Custom Banned Passwords Enabled"
            Value = If ($passwordProtection.ProtectingCustomBannedPasswords -eq "Enabled") { "Sim" } Else { "Não" }
        }
        $passwordSecurityData += [PSCustomObject]@{
            Setting = "Banned Password List State"
            Value = $passwordProtection.PasswordHashSynchronizationType # Pode ser "None" ou "Cloud"
        }
    } Else {
        Write-Warning "AVISO: Não foi possível obter políticas de proteção de senha ou elas não estão configuradas."
    }

    # Obter status dos Padrões de Segurança (Security Defaults)
    $securityDefaults = Get-MgPolicySecurityDefault -ErrorAction SilentlyContinue
    If ($securityDefaults) {
        $passwordSecurityData += [PSCustomObject]@{
            Setting = "Security Defaults Enabled"
            Value = If ($securityDefaults.IsEnabled -eq $true) { "Sim" } Else { "Não" }
        }
    } Else {
        Write-Warning "AVISO: Não foi possível obter o status dos Security Defaults (pode não estar configurado ou sem permissão)."
    }
    
    # Exportar dados de políticas de senha e security defaults
    $outputPath = Join-Path $outputDirectory "M365_Security_Password_And_Defaults.csv"
    If ($passwordSecurityData.Count -gt 0) {
        $passwordSecurityData | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8 -Force
        Write-Host "Relatório 'M365_Security_Password_And_Defaults.csv' salvo com sucesso."
    } Else {
        Write-Warning "AVISO: Nenhum dado de Políticas de Senha/Security Defaults para exportar. Arquivo CSV não gerado."
    }

} Catch {
    Write-Error "ERRO ao coletar Políticas de Senha e Security Defaults: $($_.Exception.Message)"
}
Write-Host "--- Coleta de Políticas de Senha e Security Defaults Concluída ---`n"


# --- 2. Coleta de Dados: Políticas de Acesso Condicional ---
Write-Host "--- Iniciando coleta de Políticas de Acesso Condicional ---"
$conditionalAccessData = @()

Try {
    $caPolicies = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction SilentlyContinue

    If ($caPolicies -and $caPolicies.Count -gt 0) {
        ForEach ($policy in $caPolicies) {
            $policyState = $policy.State # Ex: enabled, disabled, enabledForReportingButNotEnforced
            $grantControls = @()
            $sessionControls = @()
            
            # Processar Grant Controls
            If ($policy.GrantControls) {
                If ($policy.GrantControls.BuiltInControls) {
                    $grantControls += ($policy.GrantControls.BuiltInControls | ForEach-Object { "$($_)" }) -join ", "
                }
                If ($policy.GrantControls.CustomAuthenticationFactors) {
                    $grantControls += ($policy.GrantControls.CustomAuthenticationFactors | ForEach-Object { "Custom Factor: $($_.DisplayName)" }) -join ", "
                }
                If ($policy.GrantControls.TermsOfUse) {
                    $grantControls += ($policy.GrantControls.TermsOfUse | ForEach-Object { "ToU: $($_.DisplayName)" }) -join ", "
                }
                $grantControls = ($grantControls | Select-Object -Unique) -join "; "
            }

            # Processar Session Controls
            If ($policy.SessionControls) {
                If ($policy.SessionControls.ApplicationEnforcedRestrictions) {
                    $sessionControls += "App Enforced Restrictions"
                }
                If ($policy.SessionControls.CloudAppSecurity) {
                    $sessionControls += "Cloud App Security: $($policy.SessionControls.CloudAppSecurity.CloudAppSecurityType)"
                }
                If ($policy.SessionControls.PersistentBrowser) {
                    $sessionControls += "Persistent Browser: $($policy.SessionControls.PersistentBrowser.Mode)"
                }
                If ($policy.SessionControls.SignInFrequency) {
                    $sessionControls += "Sign-in Frequency: $($policy.SessionControls.SignInFrequency.Value) $($policy.SessionControls.SignInFrequency.Unit)"
                }
                If ($policy.SessionControls.DisableResilienceDefaults) {
                    $sessionControls += "Resilience Defaults Disabled"
                }
                If ($policy.SessionControls.DisableStrongAuthenticationFactorRemembrance) {
                    $sessionControls += "Strong Auth Factor Remembrance Disabled"
                }
                $sessionControls = ($sessionControls | Select-Object -Unique) -join "; "
            }
            
            # Incluir/Excluir Usuários/Grupos
            $includedUsersGroups = ($policy.Conditions.Users.IncludeUsers | Where-Object { $_ -ne "None" }) + ($policy.Conditions.Users.IncludeGroups | Where-Object { $_ -ne "None" })
            If ($includedUsersGroups -contains "All") { $includedUsersGroups = @("All Users/Groups") } # Simplifica se for 'All'
            
            $excludedUsersGroups = ($policy.Conditions.Users.ExcludeUsers | Where-Object { $_ -ne "None" }) + ($policy.Conditions.Users.ExcludeGroups | Where-Object { $_ -ne "None" })
            
            # Incluir/Excluir Aplicações
            $includedApplications = $policy.Conditions.Applications.IncludeApplications | Where-Object { $_ -ne "None" }
            If ($includedApplications -contains "All") { $includedApplications = @("All Cloud Apps") } # Simplifica se for 'All'
            
            $excludedApplications = $policy.Conditions.Applications.ExcludeApplications | Where-Object { $_ -ne "None" }

            $conditionalAccessData += [PSCustomObject]@{
                PolicyName          = $policy.DisplayName
                PolicyState         = $policyState
                GrantControls       = $grantControls
                SessionControls     = $sessionControls
                IncludedUsersGroups = ($includedUsersGroups -join "; ")
                ExcludedUsersGroups = ($excludedUsersGroups -join "; ")
                IncludedApplications = ($includedApplications -join "; ")
                ExcludedApplications = ($excludedApplications -join "; ")
            }
        }
    } Else {
        Write-Warning "AVISO: Nenhuma política de Acesso Condicional encontrada ou sem permissão para ler. Arquivo CSV não gerado."
    }

    # Exportar dados de políticas de acesso condicional
    $outputPath = Join-Path $outputDirectory "M365_Security_ConditionalAccess_Policies.csv"
    If ($conditionalAccessData.Count -gt 0) {
        $conditionalAccessData | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8 -Force
        Write-Host "Relatório 'M365_Security_ConditionalAccess_Policies.csv' salvo com sucesso."
    } Else {
        Write-Warning "AVISO: Nenhum dado de Políticas de Acesso Condicional para exportar. Arquivo CSV não gerado."
    }

} Catch {
    Write-Error "ERRO ao coletar Políticas de Acesso Condicional: $($_.Exception.Message)"
}
Write-Host "--- Coleta de Políticas de Acesso Condicional Concluída ---`n"


# --- 3. Coleta de Dados: Configurações de Colaboração Externa (Azure AD) ---
Write-Host "--- Iniciando coleta de Configurações de Colaboração Externa ---"
$externalCollaborationData = @()

Try {
    $orgSettings = Get-MgOrganization -OrganizationId $tenantId -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Settings
    If ($orgSettings) {
        $externalCollaborationData += [PSCustomObject]@{
            Setting = "Guest User Access Restriction"
            Value = If ($orgSettings.GuestUserRoleId) { $orgSettings.GuestUserRoleId } Else { "N/A" } # Can be Guest, Admin, etc. Or 'None' for no restriction
        }
        $externalCollaborationData += [PSCustomObject]@{
            Setting = "Guest Invite Restrictions"
            Value = If ($orgSettings.GuestUserInviteSettings) { $orgSettings.GuestUserInviteSettings.GuestUserInviteRestrictions } Else { "N/A" } # e.g., "AnyUser", "AdminsAndGuestInviters", "AdminsOnly"
        }
        $externalCollaborationData += [PSCustomObject]@{
            Setting = "Allow Create Tenancy"
            Value = If ($orgSettings.TenantAdminSettings) { $orgSettings.TenantAdminSettings.AllowCreateTenancy } Else { "N/A" }
        }
    } Else {
        Write-Warning "AVISO: Não foi possível obter as configurações da organização para colaboração externa ou sem permissão."
    }

    # Exportar dados de colaboração externa
    $outputPath = Join-Path $outputDirectory "M365_Security_ExternalCollaboration_Settings.csv"
    If ($externalCollaborationData.Count -gt 0) {
        $externalCollaborationData | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8 -Force
        Write-Host "Relatório 'M365_Security_ExternalCollaboration_Settings.csv' salvo com sucesso."
    } Else {
        Write-Warning "AVISO: Nenhum dado de Colaboração Externa para exportar. Arquivo CSV não gerado."
    }

} Catch {
    Write-Error "ERRO ao coletar Configurações de Colaboração Externa: $($_.Exception.Message)"
}
Write-Host "--- Coleta de Configurações de Colaboração Externa Concluída ---`n"


# --- 4. Coleta de Dados: Secure Score ---
Write-Host "--- Iniciando coleta de Secure Score ---"
$secureScoreData = @()

Try {
    # Pega a pontuação mais recente (assumindo que a API retorna por data decrescente)
    $secureScore = Get-MgSecuritySecureScore -ErrorAction SilentlyContinue | Sort-Object -Property LastUpdatedDateTime -Descending | Select-Object -First 1

    If ($secureScore) {
        $secureScoreData += [PSCustomObject]@{
            Metric              = "Current Score"
            Value               = $secureScore.CurrentScore
        }
        $secureScoreData += [PSCustomObject]@{
            Metric              = "Max Score"
            Value               = $secureScore.MaxScore
        }
        $secureScoreData += [PSCustomObject]@{
            Metric              = "Azure AD Score"
            Value               = $secureScore.AzureAdScore
        }
        $secureScoreData += [PSCustomObject]@{
            Metric              = "Exchange Score"
            Value               = $secureScore.ExchangeScore
        }
        $secureScoreData += [PSCustomObject]@{
            Metric              = "Endpoint Score"
            Value               = $secureScore.EndpointScore
        }
        $secureScoreData += [PSCustomObject]@{
            Metric              = "Applications Score"
            Value               = $secureScore.ApplicationsScore
        }
        $secureScoreData += [PSCustomObject]@{
            Metric              = "SharePoint Score"
            Value               = $secureScore.SharepointScore
        }
        $secureScoreData += [PSCustomObject]@{
            Metric              = "Last Updated"
            Value               = $secureScore.LastUpdatedDateTime
        }
    } Else {
        Write-Warning "AVISO: Não foi possível obter o Secure Score ou sem permissão. Certifique-se de ter a licença apropriada e a permissão SecurityEvents.Read.All. Arquivo CSV não gerado."
    }

    # Exportar dados do Secure Score
    $outputPath = Join-Path $outputDirectory "M365_Security_SecureScore.csv"
    If ($secureScoreData.Count -gt 0) {
        $secureScoreData | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8 -Force
        Write-Host "Relatório 'M365_Security_SecureScore.csv' salvo com sucesso."
    } Else {
        Write-Warning "AVISO: Nenhum dado de Secure Score para exportar. Arquivo CSV não gerado."
    }

} Catch {
    Write-Error "ERRO ao coletar Secure Score: $($_.Exception.Message)"
}
Write-Host "--- Coleta de Secure Score Concluída ---`n"


# --- 5. Considerações sobre Detecção de Ameaças (Defender for O365) e Logs de Auditoria ---
Write-Host "Notas Adicionais:"
Write-Host "  - Políticas detalhadas do Microsoft Defender for Office 365 (Anti-Phishing, Safe Links, Safe Attachments) geralmente são gerenciadas via o módulo 'ExchangeOnlineManagement' ou APIs específicas de Segurança, e não são diretamente recuperáveis de forma granular via 'Microsoft.Graph' PowerShell SDK."
Write-Host "  - O Unified Audit Log (log de auditoria unificado) é habilitado por padrão em novos tenants e é ativado automaticamente. A permissão 'AuditLog.Read.All' neste script permitiria a recuperação de logs, indicando que o recurso está ativo."
Write-Host "  - A verificação de conformidade com as políticas de segurança da Microsoft e melhores práticas é em grande parte avaliada através do Secure Score (coletado acima) e da análise manual das políticas coletadas."

Write-Host "--- Análise de Configurações de Segurança e Conformidade Concluída ---"

# --- Desconectar Microsoft Graph ---
Try {
    Disconnect-MgGraph
    Write-Host "Desconectado do Microsoft Graph."
} Catch {
    Write-Warning "AVISO: Não foi possível desconectar do Microsoft Graph. Pode não haver uma conexão ativa."
}