Sub HighlightTargets()

'   ____         __    __    __ ___      __   ___      __   __ 
'  / _(_)__  ___/ / __/ /_  / // (_)__ _/ /  / (_)__ _/ /  / /_
' / _/ / _ \/ _  / /_  __/ / _  / / _ `/ _ \/ / / _ `/ _ \/ __/
'/_//_/_//_/\_,_/   /_/   /_//_/_/\_, /_//_/_/_/\_, /_//_/\__/ 
'                                /___/         /___/           
'
'Script highlights cells which contain any of the following codes


Dim range As range
Dim i As Long
Dim DxArray(1 To 11) As String

Dim RedCount As Long
Dim RedCountDuplicates As Long
Dim RedCountNonUnique As Long

RedCount = 0
RedCountDuplicates = 0
RedCountNonUnique = 0

DxArray(1) = "A40"
DxArray(2) = "A40.0"
DxArray(3) = "A40.1"
DxArray(4) = "A40.3"
DxArray(5) = "A40.8"
DxArray(6) = "A40.9"
DxArray(7) = "A41"
DxArray(8) = "A41.0"
DxArray(9) = "A41.01"
DxArray(10) = "A41.02"
DxArray(11) = "TESTTESTTEST"

For i = 1 To UBound(DxArray)

    Set range = ActiveDocument.range
    
    With range.Find
        .Text = DxArray(i)
        .Format = True
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

                
        Do While .Execute(Forward:=True) = True
        
            range.HighlightColorIndex = wdRed
                
                RedCount = RedCount + 1
        
        Loop
    
    End With

Next

MsgBox RedCount & " High Cost ICD codes were found."

End Sub
