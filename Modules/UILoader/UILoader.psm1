# UILoader.psm1 - UI functions for Winget Discovery Tool

function Show-DiscoveryWindow {
    Add-Type -AssemblyName PresentationFramework
    
    $xamlPath = Join-Path $PSScriptRoot "..\..\UI\DiscoveryWindow.xaml"
    if (-not (Test-Path $xamlPath)) {
        throw "XAML file not found: $xamlPath"
    }
    
    $reader = (New-Object System.Xml.XmlNodeReader ([xml](Get-Content $xamlPath)))
    $window = [Windows.Markup.XamlReader]::Load($reader)
    
    return $window
}

Export-ModuleMember -Function Show-DiscoveryWindow
