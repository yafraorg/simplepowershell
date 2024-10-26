# Function to open the calculator
function Open-Calculator {
   Start-Process calc.exe
}
 
# Function to open Microsoft Word and create a new document
function Open-Word {
   $fileName = Read-Host "Enter the name for the new Word document:"
   $wordApp = New-Object -ComObject Word.Application
   $wordApp.Visible = $true
   $wordDoc = $wordApp.Documents.Add()
   $wordDoc.SaveAs("$fileName.docx")
}
 
# Function to open Microsoft PowerPoint and create a new presentation
function Open-PowerPoint {
   $fileName = Read-Host "Enter the name for the new PowerPoint presentation:"
   $pptApp = New-Object -ComObject PowerPoint.Application
   $pptApp.Visible = $true
   $pptPresentation = $pptApp.Presentations.Add()
   $pptPresentation.SaveAs("$fileName.pptx")
}
 
# Function to open Microsoft Teams
function Open-Teams {
   Start-Process "C:\Users\nelo07\AppData\Local\Microsoft\Teams\app.ico"
}
 
# Function to open OneNote
function Open-Onenote {
   Start-Process OneNote
}
 
# Function to open Microsoft Edge Browser
function Open-Browser {
   Start-Process microsoft-edge:https://www.youtube.com/watch?v=xm3YgoEiEDc
}
 
# Function to display information about the selected application
function App-Info {
   $appName = Read-Host "Enter the application name for which you want to see information:"
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

# Function to display current events
function Show-Events {
   Get-EventLog -LogName Application -Newest 10 | Format-Table -Property TimeGenerated, Source, EventID, Message -AutoSize
}

# Main loop to select between events and applications
:MainMenu while ($true) {
   # Prompt user to choose between events and applications
   $option = Read-Host "Type 'events' to see current events, 'apps' to open applications, or 'q' to quit:"

   # Check user's choice
   switch ($option) {
      "events" {
         Show-Events
      }

      "apps" {
         # Loop for opening applications
         :AppsMenu while ($true) {
            # Prompt user to choose an application
            $input = Read-Host "Type 'App1' for Calculator, 'App2' for Word, 'App3' for PowerPoint, 'App4' for Teams, 'App5' for OneNote, 'App6' for Browser, 'back' to return to the main menu:"

            # Perform action based on user input
            switch ($input) {
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
               "App4" {
                  $action = Read-Host "Do you want to open Teams or see information about it? Type 'o' to open or 'infos' for information:"
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
               "back" {
                  break AppsMenu
               }
               default {
                  Write-Host "Invalid input. Please enter a valid option."
               }
            }
         }
      }

      "q" {
         exit 
      }
      default {
         Write-Host "Invalid input. Type 'events' to see current events, 'apps' to open applications, or 'q' to quit."
      }
   }
}