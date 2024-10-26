# ============================================================================================================================
#Programm: TestV4.ps1
#Aufruf:   TestV4.ps1 ohne parameter 
#Autor:     Nelo Nissle
#Beschreibung: Das ist mein programm, Man kann mit ihm Appikationen öffnen oder events sehen Auch kann man mit meinem programm zwischen 6 verschiedenen aps entscheiden 
#              Auch hatt man die option Infos über dieses Programm Heraus Zu finden Diese sind infos die die app Kurz und knackig Beschreibt. 
#Version: 0.2
#Datum: 06.05.2024
# ==============================================================================================================================

# Das ist die FUnktion um en Taschenrechner zu öffnen 
function Open-Calculator {
    Start-Process calc.exe
}

# Das ist die Funktion um Microsoft Word zu öffnen und ein Dokument zu erstellen mit einem Namen 
function Open-Word {
    $fileName = Read-Host "Enter the name for the new Word document:"
    #erstellt das Word 
    $wordApp = New-Object -ComObject Word.Application
    #Mach die Herstellung sichtbar 
    $wordApp.Visible = $true
    # erstellt das Document 
    $wordDoc = $wordApp.Documents.Add()
    #Speichert es als 
    $wordDoc.SaveAs("$fileName.docx")
}

# Das ist die Funktion um Microsoft PowerPoint zu öffnen und eine Presentation zu erstellen mit einem Namen 
function Open-PowerPoint {
    $fileName = Read-Host "Enter the name for the new PowerPoint presentation:"
    # erstellt die PowerPoint 
    $pptApp = New-Object -ComObject PowerPoint.Application
    # Erstellt die Neue Presentation
    $pptPresentation = $pptApp.Presentations.Add()
    #Speichert es mit dem file Namen 
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

# Diese Funktionen sind da um dan schluss endlich informationen zu den Jeweiligen Programmen zu Bekommen das sind nur wörtliche infromatoenen 
function App-Info {
    param ([string]$appName)
    #$appName = Read-Host "Enter the application name for which you want to see information:"
    #Diese Informationen sind vorgelegt und sind zusammen gefasste infos aus dem Internet 
    switch ($appName) {
        "Calculator" {
            Write-Host "Calculator is a built-in application for performing mathematical calculations."
        }
        "Word" {
            Write-Host "Microsoft Word is a word processing application used for creating documents."
        }
        "PowerPoint" {
            Write-Host "Microsoft PowerPoint is a presentation software used for creating slideshows."
        }
        "Teams" {
            Write-Host "Microsoft Teams is a collaboration platform for chat, meetings, and file sharing."
        }
        "OneNote" {
            Write-Host "Microsoft OneNote is a digital notebook application for organizing notes and information."
        }
        "Browser" {
            Write-Host "Microsoft Edge is a web browser developed by Microsoft."
        }
        default {
            Write-Host "Information about the selected application is not available."
        }
    }
}

# Dasist die FUnction um die Laufenden Events zu sehen 
function Show-Events {
    # Setzt den eventlog Output in eine array
    $events = Get-EventLog -LogName Application -Newest 10
    
    $frage = Read-Host "do you wanna see just the failure (f) or sorted by Eventid (id)? "
    if ($frage -eq "f") {
        # wählt nur die error events
        $errorEvents = $events | Where-Object { $_.EntryType -eq "Error" }
        # Schreibe eventlog in datei
        $errorEvents | Out-File -FilePath Eventlog.txt
        # das ist die Loop für die events und schreibt sie aus
        foreach ($event in $errorEvents) {
            Write-Host "Time Generated: $($event.TimeGenerated)"
            Write-Host "Event ID: $($event.EventID)"
            Write-Host "-------------------------------------------------"
        }
    }
    elseif ($frage -eq "id") {
        # Sortiert eventlog output mit der Eventid
        $sortedEvents = $events | Sort-Object EventID 
        # Schreibe eventlog in datei
        $sortedEvents | Out-File -FilePath Eventlog.txt
        #Loop für events und Schreibt sie dan
        foreach ($event in $sortedEvents) {
            Write-Host "Time Generated: $($event.TimeGenerated)"
            Write-Host "Event ID: $($event.EventID)"
            Write-Host "-------------------------------------------------"
        }
    }
    $frage = Read-Host "view Output Datei (j/n)"
    if ($frage -eq "j") {
        # open file Eventlog.txt with notepad
        Start-Process notepad.exe -ArgumentList "Eventlog.txt"
    }
}

Write-Host "-------------------------------------------------"
Write-Host "Nelo Praxisarbeit Powershell"
Write-Host "-------------------------------------------------"

#======================================================================================================================================
# Ab Hier beginen die Loop und somit das HauptProgramm 
#======================================================================================================================================

# Das ist die Mainloop wo man sich entscheided ob man die apps oder Die events zu sheen und das script zu beenden 
:MainMenu while ($true) {
    # Da muss man sich entscheiden ob man apps oder events sehen will
    $option = Read-Host "Type 'events' to see current events, 'apps' to open applications, or 'q' to quit"

    # checkt was man ausgewählt hatt
    switch ($option) {
        "events" {
            Show-Events
        }

        "apps" {
            # Loop um die applikationen zu öffnen 
            # Das ist die Zweite Loop Und ich habe diese Loop AppsMenu genannt aus dem einfachen grund das ich mit einem Break nur aus Dieser bestimmten loop raus kann
            :AppsMenu while ($true) {
                # Frägt den Benutzer welche Applikation er aussführen will 
                $input = Read-Host "Type 'App1' for Calculator, 'App2' for Word, 'App3' for PowerPoint, 'App4' for Teams, 'App5' for OneNote, 'App6' for Browser, 'back' to return to the main menu"

                # Performed dan die aktion die man ausgewählt hatt 
                switch ($input) {
                
                    # App1 ist der Taschenrechner 
                    "App1" {
                        $action = Read-Host "Do you want to open Calculator or see information about it? Type 'o' to open or 'infos' for information:"
                        if ($action -eq "o") {
                            Open-Calculator
                        }
                        elseif ($action -eq "infos") {
                            App-Info "Calculator"
                        }
                        else {
                            Write-Host "Invalid input. Please enter 'o' to open or 'infos' for information."
                        }
                    }
                    # App2 ist Word 
                    "App2" {
                        $action = Read-Host "Do you want to open Word or see information about it? Type 'o' to open or 'infos' for information:"
                        if ($action -eq "o") {
                            Open-Word
                        }
                        elseif ($action -eq "infos") {
                            App-Info "Word"
                        }
                        else {
                            Write-Host "Invalid input. Please enter 'o' to open or 'infos' for information."
                        }
                    }
                    #App3 ist Powerpoint 
                    "App3" {
                        $action = Read-Host "Do you want to open PowerPoint or see information about it? Type 'o' to open or 'infos' for information:"
                        if ($action -eq "o") {
                            Open-PowerPoint
                        }
                        elseif ($action -eq "infos") {
                            App-Info "PowerPoint"
                        }
                        else {
                            Write-Host "Invalid input. Please enter 'o' to open or 'infos' for information."
                        }
                    }
                    #App4 ist Teams Da kann man schon etwas Spezielles Sehen 
                    "App4" {
                        $action = Read-Host "Do you want to open Teams or see information about it? Type 'o' to open or 'infos' for information:"
                        try {
                            if ($action -eq "o") {
                                Open-Teams
                            }
                            elseif ($action -eq "infos") {
                                App-Info "Teams"
                            }
                            else {
                                Write-Host "Invalid input. Please enter 'o' to open or 'infos' for information."
                            }
                        }
                        catch {
                            # Fehler Behandlung 
                            Write-host -f red "There has been an error Pleas Retry or start something else:"$_.Exception.Message
                            #Ausgabe in .Txt Datei 
                            $Error[0] | Out-File -FilePath ErrorNelo.txt -Append
                            # erstellt eventlog fehrler in Appklikations log 
                            New-EventLog –LogName Application –Source "Nelo Script”
                            Write-EventLog -LogName Application -Source "Nelo Script" -EntryType Error -EventID 1 -Message "Error in Teams function: $($_.Exception.Message)"
                        }
                    }
                    #App5 ist OneNote
                    "App5" {
                        $action = Read-Host "Do you want to open OneNote or see information about it? Type 'o' to open or 'infos' for information:"
                        if ($action -eq "o") {
                            Open-Onenote
                        }
                        elseif ($action -eq "infos") {
                            App-Info "OneNote"
                        }
                        else {
                            Write-Host "Invalid input. Please enter 'o' to open or 'infos' for information."
                        }
                    }
                    #App6 ist Der Link in den Microsoft edge Browser 
                    "App6" {
                        $action = Read-Host "Do you want to open Browser or see information about it? Type 'o' to open or 'infos' for information:"
                        if ($action -eq "o") {
                            Open-Browser
                        }
                        elseif ($action -eq "infos") {
                            App-Info "Browser"
                        }
                        else {
                            Write-Host "Invalid input. Please enter 'o' to open or 'infos' for information."
                        }
                    }
                    # mit dem Back geht man dan einfachwider zur MainLoop zurück wo es wider fagrt ob man die apps oder events sheen will oder ob man das skript beednen will
                    # Würde ich dieses Break nur als break schreiebn würde es das ganze Script beenden 
                    "back" {
                        break AppsMenu
                    }
                    #das ist eine Fehler meldung und heist das man eiene Richtige option asuführen solte 
                    default {
                        Write-Host "Invalid input. Please enter a valid option"
                    }
                }
            }
        }
        # Mit dem q tut man dan das skript beenden 
        "q" {
            exit 
        }
        # das Default heist einfach das es ein falsche input war und ob man dan events oder apps sehen will oder das programm beenden will 
        default {
            Write-Host "Invalid input. Type 'events' to see current events, 'apps' to open applications, or 'q' to quit"
        }
    }
}