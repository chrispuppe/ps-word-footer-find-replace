[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

$Word = New-Object -ComObject Word.Application

# Search for all Word document types (.doc, .docx, .doct, etc.)
$WordFiles = New-Object System.Windows.Forms.OpenFileDialog
$WordFiles.Multiselect = $True
$WordFiles.Filter = "All files (*.*)| *.*"
$WordFiles.showHelp = $true
$WordFiles.ShowDialog() | Out-Null


$FindText = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter text to find", "Find") # <= Find this text
$ReplaceText = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter new text", "Replace With") # <= Replace it with this text

$MatchCase = $false
$MatchWholeWorld = $true
$MatchWildcards = $false
$MatchSoundsLike = $false
$MatchAllWordForms = $false
$Forward = $false
$Wrap = 1
$Format = $false
$Replace = 2


foreach($WordFile in $WordFiles.FileNames) {
	# Open the document
    # Write-Output($WordFile.FileName)
    $Document = $Word.Documents.Open($WordFile)
    
    # Find and replace the text using the variables we just setup
    $Document.Sections.Item(1).Footers.Item(1).Range.Find.Execute($FindText, $MatchCase, $MatchWholeWorld, $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap, $Format, $ReplaceText, $Replace)

    # Save and close the document
    $Document.Close(-1) # The -1 corresponds to https://docs.microsoft.com/en-us/office/vba/api/word.wdsaveoptions
}

$Word.Quit()