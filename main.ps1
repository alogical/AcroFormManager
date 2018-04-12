$Global:AppPath = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition)

Import-Module "$Global:AppPath\modules\ManagerConsole\ManagerConsole.psm1"

Show-ManagerConsole