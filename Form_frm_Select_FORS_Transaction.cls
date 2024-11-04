VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_Select_FORS_Transaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Dim Transaction As String, Combined As Boolean

Public Sub Combo_Select_Transaction_AfterUpdate()

Me.Requery

    If Me.Combo_Select_Transaction = "Select Transaction!" Then

        Answer = MsgBox("Please Select FORS Transaction", vbInformation, "FORS Transaction")
   
    End If
    
Me.Requery

End Sub

Public Sub Form_Current()
Me.Requery
'If Me.Combo_Select_Transaction = "Select Transaction!" Then

'    Answer = MsgBox("Please open EXCEL file at first, after that click 'OK'", vbInformation, "FORS Transaction")
    
'End If
  
End Sub

Public Sub SelectTransactionOK_Click()

If Me.Combo_Select_Transaction = "Select Transaction!" Then

    Answer = MsgBox("Please Select FORS Transaction", vbInformation, "FORS Transaction")
   
ElseIf Me.Combo_Select_Transaction = "APZD" Then

    Answer = MsgBox("Do you want to change Null Serie?", vbQuestion + vbYesNo, "FORS Transaction")
  
    If Answer = vbYes Then
    
        Transaction = Me.Combo_Select_Transaction & " NullSerie"
        
    Else
    
        Transaction = Me.Combo_Select_Transaction & " Bundle"
        Answer = MsgBox("Then you will change Bundle Size!", vbInformation, "FORS Transaction")
        
    End If

ElseIf Me.Combo_Select_Transaction = "APAS" Then

    Answer = MsgBox("Do you want to import NEW Part Number?", vbQuestion + vbYesNo, "FORS Transaction")
  
    If Answer = vbYes Then
    
        Transaction = Me.Combo_Select_Transaction & " NEW"
        
    Else
    
        Transaction = Me.Combo_Select_Transaction
        Answer = MsgBox("Then you will change OLD Part Number!", vbInformation, "FORS Transaction")
        
    End If

ElseIf Me.Combo_Select_Transaction = "APFW" Then

    Answer = MsgBox("Do you want to use SLUP/REL/APPK?", vbQuestion + vbYesNo, "FORS Transaction")
  
    If Answer = vbYes Then
    
        Transaction = Me.Combo_Select_Transaction & " CAO"
        
    Else
    
        Transaction = Me.Combo_Select_Transaction
        Answer = MsgBox("Then you will use SLUP/REL!", vbInformation, "FORS Transaction")
        
    End If


Else

    Transaction = Me.Combo_Select_Transaction
    
End If

    Delay = Me.Text7
    Call FORS_Data_Loader(Transaction, Delay)

End Sub
