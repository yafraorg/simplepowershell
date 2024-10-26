# Load presentation framework assembly
Add-Type -AssemblyName PresentationFramework

# Define XAML
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="PowerShell WPF" Height="350" Width="525">
    <StackPanel>
        <Button Name="OpenCalculator">Open Calculator</Button>
        <Button Name="OpenWord">Open Word</Button>
        <Button Name="OpenPowerPoint">Open PowerPoint</Button>
        <Button Name="OpenTeams">Open Teams</Button>
        <Button Name="OpenOneNote">Open OneNote</Button>
        <Button Name="OpenBrowser">Open Browser</Button>
    </StackPanel>
</Window>
"@

# Load XAML
#$reader = New-Object System.Xml.XmlNodeReader $xaml
#$window = [Windows.Markup.XamlReader]::Load($reader)
# Load XAML
$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
$window = [Windows.Markup.XamlReader]::Load($reader)

# Define functions
function Open-Calculator {
    Start-Process calc.exe
}

function Open-Word {
    $fileName = Read-Host "Enter the name for the new Word document:"
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $true
    $wordDoc = $wordApp.Documents.Add()
    $wordDoc.SaveAs("$fileName.docx")
}

function Open-PowerPoint {
    $fileName = Read-Host "Enter the name for the new PowerPoint presentation:"
    $pptApp = New-Object -ComObject PowerPoint.Application
    $pptApp.Visible = $true
    $pptPresentation = $pptApp.Presentations.Add()
    $pptPresentation.SaveAs("$fileName.pptx")
}

function Open-Teams {
    Start-Process "C:\Users\nelo07\AppData\Local\Microsoft\Teams\x"
}

function Open-Onenote {
    Start-Process OneNote
}

function Open-Browser {
    Start-Process microsoft-edge:https://www.youtube.com/watch?v=xm3YgoEiEDc
}

# Add event handlers
$window.OpenCalculator.Add_Click({Open-Calculator})
$window.OpenWord.Add_Click({Open-Word})
$window.OpenPowerPoint.Add_Click({Open-PowerPoint})
$window.OpenTeams.Add_Click({Open-Teams})
$window.OpenOneNote.Add_Click({Open-Onenote})
$window.OpenBrowser.Add_Click({Open-Browser})

# Show window
$window.ShowDialog() | Out-Null