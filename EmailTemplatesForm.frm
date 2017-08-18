VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EmailTemplatesForm
   Caption         =   "Userform1"
   ClientHeight    =   4620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6075
   OleObjectBlob   =   "EmailTemplatesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EmailTemplatesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub ListBox1_Click()
    'Enables 'Ok' button when an item is selected
    CommandButton1.Enabled = True

End Sub

Public Sub UserForm_Initialize()

End Sub

Private Sub CommandButton1_Click()
    lstNum = EmailTemplatesForm.ListBox1.ListIndex

    'Disables Ok button for next occurrence
    CommandButton1.Enabled = False

    Me.Hide
End Sub
