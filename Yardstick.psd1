#
# Module manifest for module 'Yardstick'
#
# Import-Module .\Yardstick.psd1 from a Yardstick folder. Publish-Yardstick.ps1
# also uses this file to recognise a Yardstick folder.
#

@{

RootModule = 'Yardstick.psm1'

ModuleVersion = '1.1.0'

GUID = '975bd557-3e81-4a23-b56d-3eb1c6563a95'

Author = 'Jeffrey Pearson'

CompanyName = 'Jeffrey Pearson'

Copyright = '(c) Jeffrey Pearson. All rights reserved.'

Description = 'Automates adding and updating Win32 applications in Microsoft Intune: downloads each app, packages it and publishes it with assignments, supersedence and retention.'

PowerShellVersion = '7.0'

# powershell-yaml is deliberately not a RequiredModule: Install-Yardstick has to
# load this module in order to install it. The runner and the health check report
# it clearly when it is missing.

FunctionsToExport = @('Invoke-Yardstick', 'Install-Yardstick', 'Test-YardstickSetup')

CmdletsToExport = @()

VariablesToExport = @()

AliasesToExport = @()

PrivateData = @{
    PSData = @{
        Tags = @('Intune', 'Win32', 'Packaging', 'Windows', 'MicrosoftGraph')
        LicenseUri = 'https://github.com/jeffpearson1/Yardstick/blob/main/LICENSE.md'
        ProjectUri = 'https://github.com/jeffpearson1/Yardstick'
        ReleaseNotes = 'https://github.com/jeffpearson1/Yardstick/blob/main/CHANGELOG.md'
    }
}

}
