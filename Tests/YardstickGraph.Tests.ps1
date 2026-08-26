BeforeAll {
    Import-Module "$PSScriptRoot\..\Modules\YardstickGraph.psm1" -Force
}

AfterAll {
    Remove-Variable TenantID, ClientID, ClientSecret, Token, AuthenticationHeader -Scope Global -ErrorAction SilentlyContinue
}

Describe 'Connect-YardstickGraph' {
    BeforeEach {
        Remove-Variable Token, AuthenticationHeader -Scope Global -ErrorAction SilentlyContinue
        InModuleScope YardstickGraph { $script:AccessToken = $null }
        Mock Invoke-RestMethod -ModuleName YardstickGraph {
            [pscustomobject]@{ access_token = 'test-token'; token_type = 'Bearer'; expires_in = 3600 }
        }
    }

    It 'uses client credentials and publishes the compatibility header' {
        Connect-YardstickGraph -TenantID 'tenant' -ClientID 'client' -ClientSecret 'secret' | Out-Null
        $Global:AuthenticationHeader.Authorization | Should -Be 'Bearer test-token'
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGraph -Times 1 -Exactly -ParameterFilter {
            $Uri -eq 'https://login.microsoftonline.com/tenant/oauth2/v2.0/token' -and
            $Body.grant_type -eq 'client_credentials' -and
            $Body.scope -eq 'https://graph.microsoft.com/.default'
        }
    }

    It 'reuses a token that is safely inside its lifetime' {
        Connect-YardstickGraph -TenantID 'tenant' -ClientID 'client' -ClientSecret 'secret' | Out-Null
        Connect-YardstickGraph -TenantID 'tenant' -ClientID 'client' -ClientSecret 'secret' | Out-Null
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGraph -Times 1 -Exactly
    }

    It 'reports missing credentials clearly in a fresh process' {
        Remove-Variable TenantID, ClientID, ClientSecret -Scope Global -ErrorAction SilentlyContinue
        { Connect-YardstickGraph } | Should -Throw '*TenantID, ClientID, and ClientSecret*'
    }

    It 'does not emit the token unless PassThru is requested' {
        $output = Connect-YardstickGraph -TenantID 'tenant' -ClientID 'client' -ClientSecret 'secret'
        $output | Should -BeNullOrEmpty
        $cached = Connect-YardstickGraph -TenantID 'tenant' -ClientID 'client' -ClientSecret 'secret'
        $cached | Should -BeNullOrEmpty
        (Connect-YardstickGraph -TenantID 'tenant' -ClientID 'client' -ClientSecret 'secret' -PassThru).AccessToken | Should -Be 'test-token'
    }
}

Describe 'Invoke-YardstickGraphRequest' {
    BeforeEach {
        $Global:Token = [pscustomobject]@{ ExpiresOn = [datetimeoffset]::UtcNow.AddHours(1) }
        $Global:AuthenticationHeader = @{ Authorization = 'Bearer cached'; ExpiresOn = [datetimeoffset]::UtcNow.AddHours(1) }
    }

    It 'unwraps and follows OData collection pages' {
        Mock Invoke-RestMethod -ModuleName YardstickGraph {
            if ($Uri -eq 'https://graph.microsoft.com/beta/next') {
                return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'two' }) }
            }
            [pscustomobject]@{ value = @([pscustomobject]@{ id = 'one' }); '@odata.nextLink' = 'https://graph.microsoft.com/beta/next' }
        }
        $result = @(Invoke-YardstickGraphRequest -Resource 'deviceAppManagement/mobileApps')
        $result.id | Should -Be @('one', 'two')
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGraph -Times 2 -Exactly
    }

    It 'returns an empty collection without manufacturing an item' {
        Mock Invoke-RestMethod -ModuleName YardstickGraph { [pscustomobject]@{ value = @() } }
        @(Invoke-YardstickGraphRequest -Resource 'deviceAppManagement/mobileApps').Count | Should -Be 0
    }

    It 'serializes request bodies as JSON' {
        Mock Invoke-RestMethod -ModuleName YardstickGraph { [pscustomobject]@{ id = 'created' } }
        $result = Invoke-YardstickGraphRequest -Method Post -ApiVersion v1.0 -Resource 'deviceAppManagement/mobileApps' -Body @{ name = 'App' }
        $result.id | Should -Be 'created'
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGraph -Times 1 -Exactly -ParameterFilter {
            $Method -eq 'Post' -and $ContentType -eq 'application/json' -and $Body -match '"name":"App"'
        }
    }
}
