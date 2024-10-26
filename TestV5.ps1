# ============================================================================================================================
#Programm: SkriptVersuch2.ps1
#Aufruf:   SkriptVersuch2.ps1 ohne parameter 
#Autor:     Nelo Nissle
#Beschreibung: es öffnet Verschiedene Applikationen und man kann optionen davon sehen 
#Version: 0.2
#Datum: 06.05.2024
# ==============================================================================================================================

# GUI laden
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Das ist die FUnktion um en Taschenrechner zu öffnen 
function Open-Calculator {
    Start-Process calc.exe
}

# Das ist die Funktion um Microsoft Words zu öffnen und eine Presentation zu erstellen mit einem Namen 
function Open-Word {
    $fileName = Read-Host "Enter the name for the new Word document:"
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $true
    $wordDoc = $wordApp.Documents.Add()
    $wordDoc.SaveAs("$fileName.docx")
}

# Das ist die Funktion um Microsoft PowerPoint zu öffnen und eine Presentation zu erstellen mit einem Namen 
function Open-PowerPoint {
    $fileName = Read-Host "Enter the name for the new PowerPoint presentation:"
    $pptApp = New-Object -ComObject PowerPoint.Application
    $pptApp.Visible = $true
    $pptPresentation = $pptApp.Presentations.Add()
    $pptPresentation.SaveAs("$fileName.pptx")
}

# Das ist die Funktion um Teams zu öfnnen das geht aber nicht 
function Open-Teams {
    Start-Process "C:\Users\nelo07\AppData\Local\Microsoft\Teams\x"
}

# Das ist die Funktion um OneNote zu öfnnen 
function Open-Onenote {
    Start-Process OneNote
}

# Diese FUnctionene öffnet einen link im Microsoft edge Browser 
function Open-Browser {
    Start-Process microsoft-edge:https://www.youtube.com/watch?v=xm3YgoEiEDc
}

# Dasist die FUnction um die Laufenden Events zu sehen 
function Show-Events {
    $form2 = New-Object System.Windows.Forms.Form
    $form2.Text = 'Events'
    $form2.Size = New-Object System.Drawing.Size(1000, 1000)
    
    $okButton2 = New-Object System.Windows.Forms.Button
    $okButton2.Location = New-Object System.Drawing.Point(10, 10)
    $okButton2.Size = New-Object System.Drawing.Size(75, 23)
    $okButton2.Text = 'OK'
    $okButton2.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form2.AcceptButton = $okButton2
    $form2.Controls.Add($okButton2)
    
    #Get-EventLog -LogName Application -Newest 10 | Format-Table -Property TimeGenerated, Source, EventID, Message -AutoSize
    # put eventlog output in a array
    $events = Get-EventLog -LogName Application -Newest 10
    # sort eventlog output with EventID
    $sortedEvents = $events | Sort-Object EventID 
    # loop for each event and print
    foreach ($event in $sortedEvents) {
        # create a new string with event information
        # append to string
        $myEvents += " Source: $($event.Source) Datum: $($event.TimeGenerated) ID: $($event.EventID)`n"
    }
    $label2 = New-Object System.Windows.Forms.Label
    $label2.Location = New-Object System.Drawing.Point(10, 40)
    $label2.Size = New-Object System.Drawing.Size(900, 900)
    $label2.Text = $myEvents
    $form2.Controls.Add($label2)
    $result2 = $form2.ShowDialog()
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Nelo TestV5 Menu'
$form.Size = New-Object System.Drawing.Size(300, 300)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75, 120)
$okButton.Size = New-Object System.Drawing.Size(75, 23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)


$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10, 20)
$label.Size = New-Object System.Drawing.Size(280, 20)
$label.Text = 'Select App or Events:'
$form.Controls.Add($label)

# Add a button to open Browser
$buttonBrowser = New-Object System.Windows.Forms.Button
$buttonBrowser.Location = New-Object System.Drawing.Point(10, 40)
$buttonBrowser.Size = New-Object System.Drawing.Size(100, 23)
$buttonBrowser.Text = 'Open Browser'
$buttonBrowser.Add_Click({ Open-Browser })
$form.Controls.Add($buttonBrowser)

# Add a button to open Calculator
$buttonCalculator = New-Object System.Windows.Forms.Button
$buttonCalculator.Location = New-Object System.Drawing.Point(140, 40)
$buttonCalculator.Size = New-Object System.Drawing.Size(100, 23)
$buttonCalculator.Text = 'Open Calculator'
$buttonCalculator.Add_Click({ Open-Calculator })
$form.Controls.Add($buttonCalculator)

# Add a button to open a windows with the events
$buttonEvents = New-Object System.Windows.Forms.Button
$buttonEvents.Location = New-Object System.Drawing.Point(10, 80)
$buttonEvents.Size = New-Object System.Drawing.Size(140, 23)
$buttonEvents.Text = 'Show Events'
$buttonEvents.Add_Click({ Show-Events })
$form.Controls.Add($buttonEvents)

$form.Topmost = $true
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $userInput = $textBox.Text
    Write-Host "User entered: $userInput"
}
