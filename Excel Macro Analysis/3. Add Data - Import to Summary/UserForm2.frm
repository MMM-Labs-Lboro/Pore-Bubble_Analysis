VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Data Set-up Information"
   ClientHeight    =   3630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3030
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox2_Click()
If CheckBox2.Value = True Then
    CheckBox3.Enabled = True
    txtSeries2.Enabled = True
    Label4.Caption = 2
End If
If CheckBox2.Value = False Then
    CheckBox3.Value = False
    CheckBox3.Enabled = False
    txtSeries2.Enabled = False
    CheckBox4.Value = False
    CheckBox4.Enabled = False
    txtSeries3.Enabled = False
    CheckBox5.Value = False
    CheckBox5.Enabled = False
    txtSeries4.Enabled = False
    txtSeries5.Enabled = False
    Label4.Caption = 1
End If
End Sub
Private Sub CheckBox3_Click()
If CheckBox3.Value = True Then
    CheckBox4.Enabled = True
    txtSeries3.Enabled = True
    Label4.Caption = 3
End If
If CheckBox3.Value = False Then
    CheckBox4.Value = False
    CheckBox4.Enabled = False
    txtSeries3.Enabled = False
    CheckBox5.Value = False
    CheckBox5.Enabled = False
    txtSeries4.Enabled = False
    txtSeries5.Enabled = False
    Label4.Caption = 2
End If
End Sub
Private Sub CheckBox4_Click()
If CheckBox4.Value = True Then
    CheckBox5.Enabled = True
    txtSeries4.Enabled = True
    Label4.Caption = 4
End If
If CheckBox4.Value = False Then
    CheckBox5.Value = False
    CheckBox5.Enabled = False
    txtSeries4.Enabled = False
    txtSeries5.Enabled = False
    Label4.Caption = 3
End If
End Sub
Private Sub CheckBox5_Click()
If CheckBox5.Value = True Then
    txtSeries5.Enabled = True
    Label4.Caption = 5
End If
If CheckBox5.Value = False Then
    txtSeries5.Enabled = False
    Label4.Caption = 4
End If
End Sub
Private Sub cmdCopy_Click()
Label5.Caption = 0
UserForm2.Hide
End Sub
Private Sub ToggleButton1_Click()
Label5.Caption = 1
UserForm2.Hide
End Sub
