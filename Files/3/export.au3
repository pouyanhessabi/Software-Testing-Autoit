; Include the Word UDF (User Defined Function) for automation
#include <Word.au3>

; Initialize Word application
$oWordApp = _Word_Create()

; Check if Word was successfully launched
If @error Then
    MsgBox(16, "Error", "Unable to start Microsoft Word.")
    Exit
EndIf


; Create a new Word document
$oDoc = _Word_DocAdd($oWordApp)

; Check if the document was created successfully
If @error Then
    MsgBox(16, "Error", "Unable to create a new document.")
    _Word_Quit($oWordApp)
    Exit
EndIf

; Measure the time taken to perform the conversion
Local $iStartTime = TimerInit()

Local $oRange = _Word_DocRangeSet($oDoc, -1)

; Type some text into the document
$oRange.Text = "3: Export:       "

; Convert the Word document to PDF
Local $sPDFName = @DesktopDir & "\export.pdf"
_Word_DocExport($oDoc, $sPDFName)

; Calculate and display the conversion time
Local $iEndTime = TimerDiff($iStartTime)
$vRange = $oDoc.Range
$vRange.InsertAfter($iEndTime)
MsgBox(64, "Conversion Time", "Conversion completed in " & $iEndTime & " milliseconds.")
; Save the document with a unique filename 
Local $sSavePath = @DesktopDir & "\export.docx"
Local $oResult = _Word_DocSaveAs($oDoc, $sSavePath)
If $oResult = 0 Then
    MsgBox(16, "Error", "Unable to save the document.")
    _Word_DocClose($oDoc)
    _Word_Quit($oWordApp)
    Exit
EndIf
; Close the document
_Word_DocClose($oDoc)

; Quit Microsoft Word
_Word_Quit($oWordApp)

Exit
