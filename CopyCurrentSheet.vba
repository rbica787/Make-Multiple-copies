Sub CopyCurrentSheet()
    Dim sheetName As String
    Dim numCopies As Long
    Dim i As Long
    Dim originalSheet As Worksheet
    Dim resp As VbMsgBoxResult
    Dim baseName As String
    Dim newName As String
    Dim pos As Long
    Dim suffix As String
    Dim origNumber As Long
    Dim j As Long
    Dim nextNum As Long
    Dim targetSheet As Worksheet
    Dim userInput As String

    ' Get input safely
    userInput = InputBox("How many copies would you like to make?", "Copy Sheet")
    If Trim(userInput) = "" Or Not IsNumeric(userInput) Then
        MsgBox "Please enter a valid number of copies.", vbExclamation
        Exit Sub
    End If

    numCopies = CLng(userInput)
    If numCopies <= 0 Then
        MsgBox "Please enter a valid number of copies.", vbExclamation
        Exit Sub
    End If

    Set originalSheet = ActiveSheet
    sheetName = originalSheet.Name

    resp = MsgBox("Would you like to increment the numbers in the copied worksheets' titles?", vbYesNo + vbQuestion, "Increment Titles?")

    ' Find trailing numeric part
    pos = 0
    For j = Len(sheetName) To 1 Step -1
        If Mid(sheetName, j, 1) Like "[0-9]" Then
            pos = j
        Else
            If pos <> 0 Then Exit For
        End If
    Next j

    If pos <> 0 Then
        suffix = Mid(sheetName, pos)
        If IsNumeric(suffix) Then
            baseName = Left(sheetName, pos - 1)
            origNumber = CLng(suffix)
        Else
            baseName = sheetName
            origNumber = 0
            pos = 0
        End If
    Else
        baseName = sheetName
        origNumber = 0
    End If

    nextNum = origNumber + 1
    Set targetSheet = originalSheet

    For i = 1 To numCopies

        ' Find next available name (skip collisions)
        Do
            If resp = vbYes And pos <> 0 Then
                newName = baseName & CStr(nextNum)
            ElseIf resp = vbYes And pos = 0 Then
                newName = sheetName & " " & CStr(nextNum)
            Else
                newName = sheetName & " Copy " & CStr(i)
                Exit Do
            End If

            newName = Left(newName, 31)

            If Not SheetNameExists(newName) Then
                Exit Do
            Else
                ' Move anchor to existing sheet
                Set targetSheet = Sheets(newName)
                nextNum = nextNum + 1
            End If
        Loop

        ' Copy sheet to the right of current anchor
        originalSheet.Copy After:=targetSheet
        Set targetSheet = targetSheet.Next

        ' Assign name safely
        On Error Resume Next
        targetSheet.Name = newName
        If Err.Number <> 0 Then
            Err.Clear
            j = 1
            Do
                newName = Left(newName & " (" & j & ")", 31)
                j = j + 1
            Loop While SheetNameExists(newName)
            targetSheet.Name = newName
        End If
        On Error GoTo 0

        nextNum = nextNum + 1
    Next i
End Sub

Private Function SheetNameExists(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    SheetNameExists = False
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then
            SheetNameExists = True
            Exit Function
        End If
    Next ws
End Function
