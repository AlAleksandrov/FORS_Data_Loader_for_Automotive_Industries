Attribute VB_Name = "FORS_Data_Loader_Transaction"
Option Compare Database

Public Function FORS_Data_Loader(Transaction As String, Delay As Variant)
Dim a As Long, y As Long, last_row As Long, x As Integer, z As Long, last_column As Integer, MyTraName As String, MyDelay As Variant
Dim report_path As String, file_name As String, OpenAt As Variant, MyFORSApp, Response, MyLinkDest As String, DataArray(1000000, 10) As Variant, IndexArray(1000000) As Variant

MyTraName = Left(Transaction, 4)
MyDelay = Delay

OpenAt = "\\This PC\" 'SVBG1FILE01\Groups\AO\01-Change managment\01-C Class Project\03 Masterdata_LEPS\02-FORS Data Loader\FORS Transactions File\"

report_path = BrowseForFolder(OpenAt) & "\"

If Right(Dir(report_path & "*.xlsx", vbDirectory), 4) = "xlsx" Then
    file_name = Dir(report_path & "*.xlsx", vbDirectory)
End If

MyLinkDest = (report_path & file_name)
'MsgBox (MyLinkDest)
Workbooks.OpenText FileName:=MyLinkDest _
    , Origin:=xlWindows, StartRow:=1, DataType:=xlDelimited, TextQualifier _
    :=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:= _
    False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1) _
    , TrailingMinusNumbers:=True
If MyTraName = "APFW" Then
    Worksheets(MyTraName).Activate
Else
    Worksheets(Transaction).Activate
End If
last_row = ActiveWorkbook.ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
last_column = ActiveWorkbook.ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column

z = 0

x = 1
y = 2
For y = 1 To last_row
    For x = 1 To last_column
        DataArray(y, x) = Cells(y, x)
        If x = 1 Then
            If z = o Then
                If IndexArray(z) <> Cells(y, x) Then
                    IndexArray(z) = Cells(y, x)
                    z = z + 1
                End If
            Else
                If IndexArray(z - 1) <> Cells(y, x) Then
                    IndexArray(z) = Cells(y, x)
                    z = z + 1
                End If
            End If
        End If
    Next
Next

ActiveWindow.Close

Label1:
Responce = MsgBox("Do you want to run FORS?", vbQuestion + vbYesNoCancel, "Run FORS")
  
If Responce = vbYes Then
    Call Run_MyFORSApp(MyFORSApp)
    Response = MsgBox("Please only Log In (don't touch any other buttons!) and click 'OK' when is done", vbInformation, "FORS Login")
ElseIf Responce = vbNo Then
    Responce = MsgBox("You already open FORS?", vbQuestion + vbYesNo, "Run FORS")
    If Responce = vbYes Then
        Response = MsgBox("Please open new FORS window", vbInformation, "Run FORS")
        GoTo Label1
    Else
        Response = MsgBox("Please open new FORS window", vbInformation, "Run FORS")
        GoTo Label1
    End If
Else
    GoTo Label100
End If

' Insert Data in FORS
AppActivate MyFORSApp ' Activate FORS window

SendKeys "{F2}", True
If MyTraName = "FFSC" Then
    SendKeys "" & Transaction & "", True
Else
    SendKeys "" & MyTraName & "", True
End If
SendKeys "{ENTER}", True

If MyTraName = "FFSC" Then
    Response = MsgBox("Your choice is FORS Data Loader to start with '" & Transaction & "'", vbInformation, "Start '" & Transaction & "' in FORS")
Else
    Response = MsgBox("Your choice is FORS Data Loader to start with '" & MyTraName & "'", vbInformation, "Start '" & MyTraName & "' in FORS")
End If
AppActivate MyFORSApp ' Activate FORS window

' Call diferent Transactions
If MyTraName = "APAB" Then
    Call APAB(DataArray(), MyTraName, MyDelay, MyFORSApp, last_row, last_column)
ElseIf MyTraName = "APAG" Then
    Call APAG(DataArray(), MyTraName, MyDelay, MyFORSApp, last_row, last_column)
ElseIf MyTraName = "APAZ" Then
    Call APAZ(DataArray(), MyTraName, MyDelay, MyFORSApp, last_row, last_column)
ElseIf MyTraName = "MAKK" Then
    Call MAKK(DataArray(), MyTraName, MyDelay, MyFORSApp, last_row, last_column)
ElseIf MyTraName = "APAR" Then
    'Call APAR(DataArray(), MyTraName, MyDelay, MyFORSApp, last_row, last_column)
ElseIf MyTraName = "APAS" Then
    If Transaction = "APAS NEW" Then
        Call APAS_NEW(DataArray(), MyTraName, MyDelay, MyFORSApp, last_row, last_column)
    ElseIf Transaction = "APAS" Then
        Call APAS(DataArray(), MyTraName, MyDelay, MyFORSApp, last_row, last_column)
    End If
ElseIf MyTraName = "APZD" Then
    If Transaction = "APZD Bundle" Then
        Call APZD_Bundle(DataArray(), MyTraName, MyDelay, MyFORSApp, last_row, last_column)
    ElseIf Transaction = "APZD NullSerie" Then
        Call APZD_NullSerie(DataArray(), MyTraName, MyDelay, MyFORSApp, last_row, last_column)
    ElseIf Transaction = "APZD New" Then
        Call APZD_New(DataArray(), MyTraName, MyDelay, MyFORSApp, last_row, last_column)
    End If
ElseIf MyTraName = "APFW" Then
    If Transaction = "APFW CAO" Then
        Call APFW_CAO(IndexArray(), MyTraName, MyDelay, MyFORSApp, z)
    ElseIf Transaction = "APFW" Then
        Call APFW(IndexArray(), MyTraName, MyDelay, MyFORSApp, z)
    ElseIf Transaction = "APFW BLOCK" Then
        Call APFW_BLOCK(IndexArray(), MyTraName, MyDelay, MyFORSApp, z)
    End If
ElseIf MyTraName = "FFSC" Then
    If Transaction = "FFSCO004" Then
        Call FFSCO004(DataArray(), MyTraName, MyDelay, MyFORSApp, last_row)
    ElseIf Transaction = "FFSCO001" Then
        'Call FFSCO001(DataArray(), MyTraName, MyDelay, MyFORSApp, last_row)
    End If
End If

If MyTraName = "FFSC" Then GoTo Label100

If MyTraName = "APFW" Or MyTraName = "APPK" Or MyTraName = "MAKK" Then
    GoTo Label2
Else
    Response = MsgBox("'" & MyTraName & "' is done. When you ready click 'OK' to start SLUP & REL functions", vbInformation, "End '" & MyTraName & "' in FORS / Start SLUP & REL functions")
    AppActivate MyFORSApp ' Activate FORS window
End If

If MyTraName = "APAB" Or MyTraName = "APAG" Or MyTraName = "APAZ" Then
    Call APFW_CAO(IndexArray(), MyTraName, MyDelay, MyFORSApp, z)
ElseIf MyTraName = "APAS" Or MyTraName = "APZD" Then
    Call APFW(IndexArray(), MyTraName, MyDelay, MyFORSApp, z)
End If

Label2:
Response = MsgBox("Your new data is in FORS. Please check for errors", vbInformation, "FORS Insert Data")

Label100:

'Call [Form_frm_Select_FORS_Transaction].Form_Close

End Function

'------------------------------------------------------------
' FolderSelector
'
'------------------------------------------------------------
Function BrowseForFolder(Optional OpenAt As Variant) As Variant
     'Function purpose:  To Browser for a user selected folder.
     'If the "OpenAt" path is provided, open the browser at that directory
     'NOTE:  If invalid, it will open at the Desktop level
    Dim ShellApp As Object

    'If OpenAt = "" Then OpenAt = "\\SVBG1FILE01\Groups\AO\01-Change managment\01-C Class Project\03 Masterdata_LEPS\01-Database Master Data\06-Import\01 - Export_from_Drawings\"
    'OpenAt = "\\SVBG1FILE01\Groups\AO\01-Change managment\01-C Class Project\03 Masterdata_LEPS\02-FORS Data Loader\FORS Transactions File\"

     'Create a file browser window at the default folder
    Set ShellApp = CreateObject("Shell.Application"). _
    BrowseForFolder(0, "Please choose a folder", 0, OpenAt)

     'Set the folder to that selected.  (On error in case cancelled)
    On Error Resume Next
    BrowseForFolder = ShellApp.Self.Path
    On Error GoTo 0

     'Destroy the Shell Application
    Set ShellApp = Nothing

     'Check for invalid or non-entries and send to the Invalid error
     'handler if found
     'Valid selections can begin L: (where L is a letter) or
     '\\ (as in \\servername\sharename.  All others are invalid
    Select Case Mid(BrowseForFolder, 2, 1)
    Case Is = ":"
        If Left(BrowseForFolder, 1) = ":" Then GoTo Invalid
    Case Is = "\"
        If Not Left(BrowseForFolder, 1) = "\" Then GoTo Invalid
    Case Else
        GoTo Invalid
    End Select

    Exit Function

Invalid:
     'If it was determined that the selection was invalid, set to False
    BrowseForFolder = False
End Function

Function Run_MyFORSApp(MyFORSApp)

MyFORSApp = Shell("C:\Program Files (x86)\PuTTY\PuTTY.exe -load FORS fors_bg1", 1) ' Run FORS
'-load FORS_no_AD fors_bg1
End Function

Function Pause(NumberOfSeconds As Variant)
On Error GoTo Err_Pause
    
    Dim PauseTime As Variant, Start As Variant
    
    PauseTime = NumberOfSeconds
    Start = Timer
    Do While Timer < Start + PauseTime
    DoEvents
    Loop
    
Exit_Pause:
    Exit Function
    
Err_Pause:
    MsgBox Err.Number & " - " & Err.Description
    Resume Exit_Pause
    
End Function

Function APAB(DataArray() As Variant, MyTraName As String, MyDelay As Variant, MyFORSApp, last_row As Long, last_column As Integer)

Dim i As Long, j As Integer, k As Integer

For i = 2 To last_row
    If Len(DataArray(i, last_column)) >= 0 And Len(DataArray(i, last_column)) <= 60 Then
        j = 1
    ElseIf Len(DataArray(i, last_column)) > 60 And Len(DataArray(i, last_column)) <= 120 Then
        j = 2
    ElseIf Len(DataArray(i, last_column)) > 120 And Len(DataArray(i, last_column)) <= 180 Then
        j = 3
    ElseIf Len(DataArray(i, last_column)) > 180 And Len(DataArray(i, last_column)) <= 240 Then
        j = 4
    ElseIf Len(DataArray(i, last_column)) > 240 And Len(DataArray(i, last_column)) <= 300 Then
        j = 5
    End If
    SendKeys "{HOME}"
    SendKeys "{TAB}"
    SendKeys "{DEL 2}"
    SendKeys "61"
    SendKeys "{DEL 22}"
    SendKeys DataArray(i, last_column - 2) ' Part No.
    SendKeys "{ENTER 2}" 'SendKeys "{TAB 2}"
    SendKeys "{DEL 4}"
    SendKeys DataArray(i, last_column - 1) ' Operation No.
    SendKeys "{ENTER}" 'SendKeys "{TAB}"
    SendKeys "{F5}"
    SendKeys "{HOME}"
    SendKeys "{ENTER 8}"
    For k = 1 To 10
        SendKeys "{DEL 60}"
        SendKeys "{ENTER 2}"
    Call Pause(MyDelay)
    Next
    SendKeys "{HOME}"
    SendKeys "{ENTER 8}"
    If j = 1 Then
        SendKeys DataArray(i, last_column) ' Designation
    ElseIf j = 2 Then
        SendKeys Left(DataArray(i, last_column), 60) ' Designation
        SendKeys "{HOME}"
        SendKeys "{ENTER 10}"
        SendKeys "{DEL 60}"
        SendKeys Mid(DataArray(i, last_column), 61, Len(DataArray(i, last_column)) - 60) ' Designation
    ElseIf j = 3 Then
        SendKeys Left(DataArray(i, last_column), 60) ' Designation
        SendKeys "{HOME}"
        SendKeys "{ENTER 10}"
        SendKeys "{DEL 60}"
        SendKeys Mid(DataArray(i, last_column), 61, Len(DataArray(i, last_column)) - 60) ' Designation
        SendKeys "{HOME}"
        SendKeys "{ENTER 12}"
        SendKeys "{DEL 60}"
        SendKeys Mid(DataArray(i, last_column), 121, Len(DataArray(i, last_column)) - 120) ' Designation
    ElseIf j = 4 Then
        SendKeys Left(DataArray(i, last_column), 60) ' Designation
        SendKeys "{HOME}"
        SendKeys "{ENTER 10}"
        SendKeys "{DEL 60}"
        SendKeys Mid(DataArray(i, last_column), 61, Len(DataArray(i, last_column)) - 60) ' Designation
        SendKeys "{HOME}"
        SendKeys "{ENTER 12}"
        SendKeys "{DEL 60}"
        SendKeys Mid(DataArray(i, last_column), 121, Len(DataArray(i, last_column)) - 120) ' Designation
        SendKeys "{HOME}"
        SendKeys "{ENTER 14}"
        SendKeys "{DEL 60}"
        SendKeys Mid(DataArray(i, last_column), 181, Len(DataArray(i, last_column)) - 180) ' Designation
    ElseIf j = 5 Then
        SendKeys Left(DataArray(i, last_column), 60) ' Designation
        SendKeys "{HOME}"
        SendKeys "{ENTER 10}"
        SendKeys "{DEL 60}"
        SendKeys Mid(DataArray(i, last_column), 61, Len(DataArray(i, last_column)) - 60) ' Designation
        SendKeys "{HOME}"
        SendKeys "{ENTER 12}"
        SendKeys "{DEL 60}"
        SendKeys Mid(DataArray(i, last_column), 121, Len(DataArray(i, last_column)) - 120) ' Designation
        SendKeys "{HOME}"
        SendKeys "{ENTER 14}"
        SendKeys "{DEL 60}"
        SendKeys Mid(DataArray(i, last_column), 181, Len(DataArray(i, last_column)) - 180) ' Designation
        SendKeys "{HOME}"
        SendKeys "{ENTER 16}"
        SendKeys "{DEL 60}"
        SendKeys Mid(DataArray(i, last_column), 241, Len(DataArray(i, last_column)) - 240) ' Designation
    End If
    SendKeys "{HOME}"
    Call Pause(MyDelay)
    SendKeys "{F9}"
    SendKeys "{F10}"
    SendKeys "{F5}"

Next
End Function

Function APAG(DataArray() As Variant, MyTraName As String, MyDelay As Variant, MyFORSApp, last_row As Long, last_column As Integer)

Dim i As Long

For i = 2 To last_row

    SendKeys "{HOME}"
    SendKeys "{TAB}"
    SendKeys "{DEL 2}"
    SendKeys "61"
    SendKeys "{DEL 22}"
    SendKeys DataArray(i, last_column - 2) ' Part No.
    SendKeys "{ENTER 2}"
    Call Pause(MyDelay)
    SendKeys "{DEL 4}"
    SendKeys DataArray(i, last_column - 1) ' Operation No.
    SendKeys "{ENTER}"
    SendKeys "{F5}"
    SendKeys "{ENTER 9}"
    Call Pause(MyDelay)
    SendKeys "{DEL 10}"
    SendKeys DataArray(i, last_column) ' Min Qty
    SendKeys "{ENTER}"
    Call Pause(MyDelay)
    SendKeys "{F10}"

Next

End Function

Function APAZ(DataArray() As Variant, MyTraName As String, MyDelay As Variant, MyFORSApp, last_row As Long, last_column As Integer)

Dim i As Long

For i = 2 To last_row

    SendKeys "{HOME}"
    SendKeys "{TAB}"
    SendKeys "{DEL 2}"
    SendKeys "61"
    SendKeys "{DEL 22}"
    SendKeys DataArray(i, last_column - 3) ' Part No.
    SendKeys "{F5}"
    Call Pause(MyDelay)
    SendKeys "{DEL 4}"
    SendKeys DataArray(i, last_column - 2) ' Operation No.
    SendKeys "{F5}"
    SendKeys "{ENTER 6}"
    Call Pause(MyDelay)
    SendKeys "{DEL 9}"
    SendKeys DataArray(i, last_column - 1) ' Std. time
    SendKeys "{ENTER 3}"
    Call Pause(MyDelay)
    SendKeys "{DEL 10}"
    SendKeys DataArray(i, last_column) ' Qty
    SendKeys "{ENTER}"
    SendKeys "{HOME}"
    Call Pause(MyDelay)
    SendKeys "{F9}"
    SendKeys "{F10}"
    SendKeys "{F5}"

Next

End Function

Function MAKK(DataArray() As Variant, MyTraName As String, MyDelay As Variant, MyFORSApp, last_row As Long, last_column As Integer)

Dim i As Long

For i = 2 To last_row

    SendKeys "{HOME}"
    SendKeys "{TAB}"
    SendKeys "{DEL 21}"
    SendKeys DataArray(i, last_column - 1) ' Copy from
    SendKeys "{ENTER}"
    SendKeys "{DEL 21}"
    SendKeys DataArray(i, last_column) ' To
    SendKeys "{ENTER}"
    SendKeys "1"
    SendKeys "1"
    SendKeys "1"
    SendKeys "{F9 2}"
    Call Pause(MyDelay)

Next

End Function
Function APAR(DataArray() As Variant, MyTraName As String, MyDelay As Variant, MyFORSApp, last_row As Long, last_column As Integer)

End Function

Function APAS_NEW(DataArray() As Variant, MyTraName As String, MyDelay As Variant, MyFORSApp, last_row As Long, last_column As Integer)

Dim i As Long

For i = 2 To last_row

    SendKeys "{HOME}"
    SendKeys "{TAB}"
    SendKeys "{DEL 2}"
    SendKeys "61"
    SendKeys "{DEL 22}"
    SendKeys DataArray(i, last_column - 3) ' Part No.
    SendKeys "{ENTER 2}"
    Call Pause(MyDelay)
    SendKeys "{DEL 4}"
    SendKeys DataArray(i, last_column - 2) ' Operation No.
    SendKeys "{HOME}"
    SendKeys "{ENTER 5}"
    SendKeys "{DEL 4}"
    SendKeys "{ENTER 2}"
    Call Pause(MyDelay)
    SendKeys "{DEL 22}"
    SendKeys DataArray(i, last_column - 1) ' Component
    SendKeys "{HOME}"
    SendKeys "{ENTER 8}"
    Call Pause(MyDelay)
    SendKeys "{DEL 13}"
    SendKeys DataArray(i, last_column) ' Qty
    SendKeys "{ENTER}"
    SendKeys "{F9}"
    SendKeys "{F5}"
    Call Pause(MyDelay)

Next

End Function

Function APAS(DataArray() As Variant, MyTraName As String, MyDelay As Variant, MyFORSApp, last_row As Long, last_column As Integer)

Dim i As Long

For i = 2 To last_row

    SendKeys "{HOME}"
    SendKeys "{TAB}"
    SendKeys "{DEL 2}"
    SendKeys "61"
    SendKeys "{DEL 22}"
    SendKeys DataArray(i, last_column - 3) ' Part No.
    SendKeys "{ENTER 2}"
    Call Pause(MyDelay)
    SendKeys "{DEL 4}"
    SendKeys DataArray(i, last_column - 2) ' Operation No.
    SendKeys "{HOME}"
    SendKeys "{ENTER 5}"
    SendKeys "{DEL 4}"
    SendKeys "{ENTER 2}"
    Call Pause(MyDelay)
    SendKeys "{DEL 22}"
    SendKeys DataArray(i, last_column - 1) ' Component
    SendKeys "{F9}"
    SendKeys "{F5}"
    SendKeys "{HOME}"
    SendKeys "{ENTER 8}"
    Call Pause(MyDelay)
    SendKeys "{DEL 13}"
    SendKeys DataArray(i, last_column) ' Qty
    SendKeys "{ENTER}"
    SendKeys "{F10}"
    SendKeys "{F5}"
    Call Pause(MyDelay)

Next

End Function

Function APZD_Bundle(DataArray() As Variant, MyTraName As String, MyDelay As Variant, MyFORSApp, last_row As Long, last_column As Integer)

Dim i As Long

For i = 2 To last_row

    SendKeys "{HOME}"
    SendKeys "{TAB}"
    SendKeys "{DEL 2}"
    SendKeys "61"
    SendKeys "{DEL 22}"
    SendKeys DataArray(i, last_column - 2) ' Part No.
    SendKeys "{ENTER 2}"
    Call Pause(MyDelay)
    SendKeys "{DEL 12}"
    SendKeys "{ENTER 4}"
    SendKeys DataArray(i, last_column - 1) ' Wire No.
    SendKeys "{ENTER}"
    SendKeys "{F5}"
    SendKeys "{ENTER 7}"
    Call Pause(MyDelay)
    SendKeys "{DEL 8}"
    SendKeys DataArray(i, last_column) ' Qty
    SendKeys "{ENTER}"
    SendKeys "{F10}"
    SendKeys "{F5}"
    Call Pause(MyDelay)

Next

End Function

Function APZD_NullSerie(DataArray() As Variant, MyTraName As String, MyDelay As Variant, MyFORSApp, last_row As Long, last_column As Integer)

Dim i As Long

For i = 2 To last_row

    SendKeys "{HOME}"
    SendKeys "{TAB}"
    SendKeys "{DEL 2}"
    SendKeys "61"
    SendKeys "{DEL 22}"
    SendKeys DataArray(i, last_column - 2) ' Part No.
    SendKeys "{ENTER 2}"
    Call Pause(MyDelay)
    SendKeys "{DEL 12}"
    SendKeys "{ENTER 4}"
    SendKeys DataArray(i, last_column - 1) ' Wire No.
    SendKeys "{ENTER}"
    SendKeys "{F5}"
    SendKeys "{ENTER 6}"
    Call Pause(MyDelay)
    SendKeys "{DEL 14}"
    SendKeys DataArray(i, last_column) ' Length
    SendKeys "{ENTER}"
    SendKeys "{F10}"
    SendKeys "{F5}"
    Call Pause(MyDelay)

Next

End Function

Function APZD_New(DataArray() As Variant, MyTraName As String, MyDelay As Variant, MyFORSApp, last_row As Long, last_column As Integer)

Dim i As Long

For i = 2 To last_row

    SendKeys "{HOME}"
    SendKeys "{TAB}"
    SendKeys "{DEL 2}"
    SendKeys "61"
    SendKeys "{DEL 22}"
    SendKeys DataArray(i, last_column - 6) ' Part No.
    SendKeys "{ENTER 2}"
    Call Pause(MyDelay)
    SendKeys "{DEL 4}"
    SendKeys DataArray(i, last_column - 5) ' Operation No.
    SendKeys "{ENTER 1}"
    SendKeys "{HOME}"
    SendKeys "{ENTER 5}"
    SendKeys "{DEL 8}"
    SendKeys DataArray(i, last_column - 4) ' Time
    SendKeys "{ENTER 1}"
    SendKeys "{DEL 4}"
    SendKeys "{ENTER 2}"
    SendKeys "{DEL 12}"
    SendKeys DataArray(i, last_column - 3) ' Wire No.
    SendKeys "{ENTER}"
    SendKeys "{HOME}"
    SendKeys "{ENTER 9}"
    Call Pause(MyDelay)
    SendKeys "{DEL 21}"
    SendKeys DataArray(i, last_column - 2) ' Component
    SendKeys "{ENTER}"
    SendKeys "{DEL 8}"
    SendKeys DataArray(i, last_column - 1) ' Qty
    SendKeys "{ENTER}"
    SendKeys "{DEL 14}"
    SendKeys DataArray(i, last_column) ' Length
    SendKeys "{ENTER}"
    SendKeys "{F9}"
    Call Pause(MyDelay)
    SendKeys "{F2}"
    SendKeys "{DEL 4}"
    SendKeys "APFW"
    SendKeys "{ENTER}"
    SendKeys "{HOME}"
    SendKeys "{DEL 4}"
    SendKeys "SLUP"
    SendKeys "{F12}"
    Call Pause(MyDelay)
    SendKeys "{F5}"
    SendKeys "{HOME}"
    SendKeys "{DEL 4}"
    SendKeys "REL"
    SendKeys "{F12}"
    Call Pause(MyDelay)
    SendKeys "{F5}"
    If i <> last_row Then
        SendKeys "{F2}"
        SendKeys "{DEL 4}"
        SendKeys "APZD"
        SendKeys "{F12}"
        Call Pause(MyDelay)
    Else
        MyTraName = "APFW"
    End If

Next

End Function
Function FFSCO004(DataArray() As Variant, MyTraName As String, MyDelay As Variant, MyFORSApp, last_row As Long)

Dim i As Long, j As Integer, k As Integer

j = 0
k = (last_row - 1) \ 15
SendKeys "03"
SendKeys "{F12}"
SendKeys "{DEL 10}"
SendKeys Left(DataArray(1, 4), 8) ' Export Name
SendKeys "{ENTER 9}"
SendKeys "J"
SendKeys "{ENTER 9}"
SendKeys "J"
SendKeys "05"
SendKeys "{F9}"
SendKeys "{F4}"
SendKeys "{ENTER 2}"
For i = 2 To last_row
    SendKeys DataArray(i, 1) ' Part No.
    SendKeys "{ENTER}"
    SendKeys "1"    'Qty
    Call Pause(MyDelay)
    If k = 0 Then
        If i = last_row Then
            SendKeys "{F9}"
        Else
            SendKeys "{ENTER 2}"
        End If
        GoTo Label1
    Else
        j = j + 1
        If j = 15 Then
            j = 0
            k = k - 1
            SendKeys "{F9}"
            SendKeys "{F2}"
            SendKeys "{ENTER 2}"
        Else
            SendKeys "{ENTER 2}"
        End If
    End If
Label1:
Next
SendKeys "{F3}"
SendKeys "{F6}"
SendKeys "{F3}"
SendKeys "07"
SendKeys "{F12}"
SendKeys "{DEL 10}"
SendKeys Left(DataArray(1, 4), 8) ' Export Name
SendKeys "{ENTER 2}"
SendKeys Left(DataArray(1, 4), 8) ' Export Name
SendKeys "{F8}"
Call Pause(MyDelay)

End Function

Function APFW_CAO(IndexArray() As Variant, MyTraName As String, MyDelay As Variant, MyFORSApp, z As Long)

Dim i As Integer

For i = 0 To z - 1
    If MyTraName <> "APFW" Then
        SendKeys "{F2}"
        SendKeys "{DEL 4}"
        SendKeys "APFW"
        SendKeys "{ENTER}"
    End If
    SendKeys "{HOME}"
    SendKeys "{TAB}"
    SendKeys "{DEL 2}"
    SendKeys "61"
    SendKeys "{DEL 22}"
    SendKeys IndexArray(i) ' Part No.
    SendKeys "{ENTER}"
    SendKeys "{F5}"
    SendKeys "{HOME}"
    SendKeys "{DEL 4}"
    SendKeys "SLUP"
    SendKeys "{F12}"
    Call Pause(MyDelay)
    SendKeys "{F5}"
    SendKeys "{HOME}"
    SendKeys "{DEL 4}"
    SendKeys "REL"
    SendKeys "{F10}"
    Call Pause(MyDelay)
    SendKeys "{F5}"
    SendKeys "{F2}"
    SendKeys "APPK"
    SendKeys "{F5}"
    SendKeys "{HOME}"
    SendKeys "{DEL 4}"
    SendKeys "REL"
    SendKeys "{F12}"
    Call Pause(MyDelay)
    MyTraName = "APPK"

Next

End Function

Function APFW(IndexArray() As Variant, MyTraName As String, MyDelay As Variant, MyFORSApp, z As Long)

Dim i As Integer

For i = 0 To z - 1
    If MyTraName <> "APFW" Then
        SendKeys "{F2}"
        SendKeys "{DEL 4}"
        SendKeys "APFW"
        SendKeys "{ENTER}"
    End If
    SendKeys "{HOME}"
    SendKeys "{TAB}"
    SendKeys "{DEL 2}"
    SendKeys "61"
    SendKeys "{DEL 22}"
    SendKeys IndexArray(i) ' Part No.
    SendKeys "{ENTER}"
    SendKeys "{F5}"
    SendKeys "{HOME}"
    SendKeys "{DEL 4}"
    SendKeys "SLRE"
    SendKeys "{F12}"
    Call Pause(MyDelay)
    MyTraName = "APFW"

Next

End Function

Function APFW_BLOCK(IndexArray() As Variant, MyTraName As String, MyDelay As Variant, MyFORSApp, z As Long)

Dim i As Integer

For i = 0 To z - 1
    If MyTraName <> "APFW" Then
        SendKeys "{F2}"
        SendKeys "{DEL 4}"
        SendKeys "APFW"
        SendKeys "{ENTER}"
    End If
    SendKeys "{HOME}"
    SendKeys "{TAB}"
    SendKeys "{DEL 2}"
    SendKeys "61"
    SendKeys "{DEL 22}"
    SendKeys IndexArray(i) ' Part No.
    SendKeys "{F5}"
    SendKeys "{ENTER 8}"
    SendKeys "{DEL 1}"
    SendKeys "1"
    SendKeys "{DEL 24}"
    SendKeys "block index"
    SendKeys "{ENTER}"
    SendKeys "{F2}"
    SendKeys "{ENTER}"
    SendKeys "{F10}"
    Call Pause(MyDelay)
    MyTraName = "APFW"

Next

End Function

