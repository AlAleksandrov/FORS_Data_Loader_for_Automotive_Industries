VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_Main Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command388_Click()

Me.Label378.Visible = True
Me.Label369.Visible = True
Me.Label391.Visible = True
Me.Label392.Visible = True
Me.Label183.Visible = False
Me.Label186.Visible = False
Me.Label387.Visible = False
Me.Label390.Visible = False

End Sub

Private Sub Command389_Click()

Me.Label183.Visible = True
Me.Label186.Visible = True
Me.Label387.Visible = True
Me.Label390.Visible = True
Me.Label378.Visible = False
Me.Label369.Visible = False
Me.Label391.Visible = False
Me.Label392.Visible = False

End Sub

Private Sub Form_Current()
Call Command388_Click
End Sub

Private Sub Command121_Click()

Application.FollowHyperlink "\\SVBG1FILE01\Groups\AO\01-Change managment\01-C Class Project\03 Masterdata_LEPS\02-FORS Data Loader\Example\FORS_Transactions.xlsx"

End Sub

Public Sub Command122_Click()

DoCmd.OpenForm FormName:="frm_Select_FORS_Transaction"

End Sub
