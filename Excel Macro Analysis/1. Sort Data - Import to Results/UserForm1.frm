VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6720
   ClientLeft      =   240
   ClientTop       =   885
   ClientWidth     =   4185
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbRun_Click()
UserForm1.Hide
End Sub

Private Sub Row_1_Click()
    txtTitle1.Enabled = (Row_1.Value = True)
    txtPrefix1.Enabled = (Row_1.Value = True)
    txtSufix1.Enabled = (Row_1.Value = True)
    Row_2.Locked = False
End Sub
Private Sub Row_2_Click()
    txtTitle2.Enabled = (Row_2.Value = True)
    txtPrefix2.Enabled = (Row_2.Value = True)
    txtSufix2.Enabled = (Row_2.Value = True)
    Row_3.Locked = False
End Sub
Private Sub Row_3_Click()
    txtTitle3.Enabled = (Row_3.Value = True)
    txtPrefix3.Enabled = (Row_3.Value = True)
    txtSufix3.Enabled = (Row_3.Value = True)
End Sub

Private Sub UserForm_Activate()
    Row_1.Value = False
    Row_1.Locked = False
    txtTitle1.Enabled = False
    txtPrefix1.Enabled = False
    txtSufix1.Enabled = False
    Row_2.Value = False
    Row_2.Locked = True
    txtTitle2.Enabled = False
    txtPrefix2.Enabled = False
    txtSufix2.Enabled = False
    Row_3.Value = False
    Row_3.Locked = True
    txtTitle3.Enabled = False
    txtPrefix3.Enabled = False
    txtSufix3.Enabled = False
    Filled.Value = True
    Ring.Value = False
    Controls("Label6").Caption = "Please Select the Appropriate File Seporator - '\' or '/'. File Path:" & Application.ThisWorkbook.Path
End Sub
