#Script to create a word document based on variables from a computer running the deploy task sequence.
 
#Set the path where the CSV files are stored
$csvpath = 'C:\CSV'
 
#Get CSV files in the specified path
$csvs = Get-ChildItem "$csvpath\*.csv"
 
#Run a loop to import CSV files and data
ForEach ($csv in $csvs) {
 
    $csvdata = Import-Csv "$csv"
    ForEach ($data in $csvdata)
        {
            $SN = $data.sn
            $Asset = $data.asset
            $Model = $data.model
        }
 
    #-----------------------------------------------
    #-----------------------------------------------
 
    #Create a new Word document
    $word = New-Object -ComObject Word.Application
    $Document = $Word.Documents.Add()
    $Selection = $Word.Selection
    
    #Uncomment to make the Word Document visible
    #$Word.Visible = $True 
 
    #-----------------------------------------------
    #-----------------------------------------------
 
    #Format the Word document
    #Set margins
    $Selection.PageSetup.LeftMargin = 36
    $Selection.PageSetup.RightMargin = 36
    $Selection.PageSetup.TopMargin = 36
    $Selection.PageSetup.BottomMargin = 36
 
    #Title
    $Selection.Style = 'Title'
    $Selection.ParagraphFormat.Alignment = 1
    $Selection.Borders.OutsideLineStyle = 1
    $Selection.TypeText("New Computer Information")
    $Selection.TypeParagraph()
    $Selection.TypeParagraph()
  
    #Numbered List
    $Selection.Style = 'List Number'
    $Selection.Font.Size = 12
    $Selection.TypeText("The computer model is: ")
    $Selection.Font.Bold = 1
    $Selection.TypeText("$model")
    $Selection.Font.Bold = 0
    $Selection.TypeParagraph()
    $Selection.TypeText("The computer serial number is: ")
    $Selection.Font.Bold = 1
    $Selection.TypeText("$SN")
    $Selection.Font.Bold = 0
    $Selection.TypeParagraph()
    $Selection.TypeText("The computer asset tag is: ")
    $Selection.Font.Bold = 1
    $Selection.TypeText("$asset")
    $Selection.Font.Bold = 0
    $Selection.TypeParagraph()
    $Selection.Style = 'No Spacing'
 
    #Insert and center image
    $Selection.ParagraphFormat.Alignment = 1
    $Selection.InlineShapes.AddPicture("C:\CSV\Images\Image.png")
    $Selection.TypeParagraph()
    $Selection.TypeParagraph()
 
    #Inser date imaged and add a line for the technician to sign
    $Selection.Font.Size = 12
    $Selection.ParagraphFormat.Alignment = 0
    $date = Get-Date -UFormat "%m / %d / %Y"
    $Selection.TypeText("Imaged Date: $date  -  ")
    $Selection.TypeText("Imaging Technician:_________________")
    $Selection.TypeParagraph()
 
    #-----------------------------------------------
    #-----------------------------------------------
 
    #Save the Word document
    $SaveLoc = "$csvpath\$SN.docx"
    $Document.SaveAs([ref]$SaveLoc,[ref]$SaveFormat::wdFormatDocument)
    $word.Quit()
 
    #-----------------------------------------------
    #-----------------------------------------------
 
    #Print the Word document
    Start-Process -Filepath "$csvpath\$SN.docx" -verb print
    Start-Sleep -Seconds 10
 
    #-----------------------------------------------
    #-----------------------------------------------
 
    #Move the Word document and delete the CSV file
    Move-Item -Path "$csvpath\$SN.docx" -Destination "$csvpath\Completed Documents\$SN.docx" -Force
    Remove-Item "$csv"
 
    #-----------------------------------------------
    #-----------------------------------------------
 
}
 
#Cleanup the Word COM Object
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$word)
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
Remove-Variable word