Option Explicit

Public Sub FiveHundredKnNotice()

'===================================================================
'                        _ ________           __           
'  ___  ____ ___  ____ _(_) / ____/___ ______/ /____  _____
' / _ \/ __ `__ \/ __ `/ / / /_  / __ `/ ___/ __/ _ \/ ___/
'/  __/ / / / / / /_/ / / / __/ / /_/ (__  ) /_/  __/ /    
'\___/_/ /_/ /_/\__,_/_/_/_/    \__,_/____/\__/\___/_/     
'                                                          
'script notifies of $300k+ claimants.
'you paste in the directory of the notice, such as
'\\\Shared\\Notice
'===================================================================

Dim OApp As Object
Dim OMail As Object
Dim Signature As String

Set OApp = CreateObject("Outlook.Application")

Set OMail = OApp.CreateItem(0)

    With OMail
        .Display
        .ReadReceiptRequested = True 'request read receipt.
    End With

    Signature = OMail.Body

Dim DirPath As String
Dim CheckForSubDirPath As String
Dim missingVariable As String 'fix this if you need it
Dim MidDirPath As String
Dim DirPathArray() As String 'array for splitting
Dim Product As String
Dim Group As String
Dim Claims As String
Dim Year As String

Dim DateArray() As String
Dim NoticeDate As String

Dim person_a As String
Dim person_b As String

person_a = "person_a@website.com"
Joann = "person_b@website.com"

On Error GoTo Errormsg:

    'take the folder DirPath, paste it into place.
    DirPath = InputBox("Paste the folder Path", "Directory Path")

    'parse it apart. let's get out everything after the 60th character
    CheckForSubDirPath = (Mid(DirPath, 60, 999))

    CheckForSubDirPath = Left(CheckForSubDirPath, 17)

        If CheckForSubDirPath = "something_special" Then
            Product = "something_special"

        Else: Product = "something_else"

        End If

    Select Case Product

        Case "something_special"

            'Parse path
            MidDirPath = (Mid(DirPath, 76, 999))
            DirPathArray = Split(MidDirPath, "\")

            'give array elements some clearer names
            Group = DirPathArray(0)
            Subgroup = DirPathArray(1)
            Year = DirPathArray(2)

            'remove Notice_, and keep just the date on righthand
            DateArray() = Split(DirPathArray(3), "_")
            NoticeDate = DateArray(1)

        Case "something_else"

            'Parse path
            MidDirPath = (Mid(DirPath, 47, 999))
            DirPathArray = Split(MidDirPath, "\")

            'give array elements some clearer names
            Group = DirPathArray(0)
            Subgroup = DirPathArray(1)
            Year = DirPathArray(2)

            'remove Notice_, and keep just the date on righthand
            DateArray() = Split(DirPathArray(3), "_")
            NoticeDate = DateArray(1)

        Case Else

        MsgBox "You probably didn't paste in the entire directory path, or there is something wrong with the directory, such as a change in root folder name."

    End Select

    With OMail

        .To = person_a
        .CC = person_b
        .Subject = Product & "  -  " & Group & "  -  PY " & Year & "  -  " & NoticeDate

        .Body = "Hello person_a and person_b," & vbNewLine & vbNewLine & _
        "This is your sample message." & vbNewLine & vbNewLine & _
        Product & " group: " & Group & vbNewLine & _
        "Year: " & Year & vbNewLine & _
        "Receive date: " & NoticeDate & vbNewLine & _
        "Requested $: " & vbNewLine & _
        "Directory: " & DirPath & vbNewLine & _
        Signature

        .Categories = Product & "Notice"
        '.BodyFormat = olFormatPlain 'format as plain text
        .Display
        '.Send 'we can do this.

    End With


Set OMail = Nothing

Set OApp = Nothing

Errormsg:
    Debug.Print Err.Number

End Sub
