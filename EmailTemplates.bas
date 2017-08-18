Attribute VB_Name = "EmailTemplates"
'=========================================================================
' Public variables cause lazy
'=========================================================================
Public CurrentEmail As Object
Public NewEmail As Outlook.MailItem
Public Template As Outlook.MailItem
'=========================================================================
' First folder
'=========================================================================
Sub ProjectClosures()
    CreateReplyFromTemplate "Email_Templates"
End Sub
'=========================================================================
' STEP BY STEP PROCESS
'=========================================================================
Sub CreateReplyFromTemplate(templateFolderPath As String)
    CheckIfValidEmail
    GrabEmailToReplyTo
    PopulateForm (templateFolderPath)
    DisplayForm (templateFolderPath)
    PrepareEmail
End Sub
'=========================================================================
' Check if it is an email
'=========================================================================
Function CheckIfValidEmail()
    'if email is open in its own window
    If TypeName(ActiveWindow) = "Explorer" Then
        'select it
        'if statement checks that an email item is selected.
        If ActiveExplorer.Selection.Count > 0 Then
            If TypeOf ActiveExplorer.Selection.Item(1) Is Outlook.MailItem Then
                On Error GoTo ReplyError:
            Else
                MsgBox ("Only works on emails. Does not work on Calendar, Contact or Other items.")
                End
            End If
        Else
            MsgBox ("Please select an email to use this function")
            End
        End If
    'if the email is open in a panel
    ElseIf TypeName(ActiveWindow) = "Inspector" Then
        'otherwise select the email being inspected from the menu
        'if statement checks that an email item is selected.
        If TypeOf ActiveInspector.CurrentItem Is Outlook.MailItem Then
            If ActiveInspector.CurrentItem.Sent = False Then
                MsgBox ("You are already replying")
                End
            End If
        Else
            MsgBox ("Only works on emails. Does not work on Calendar, Contact or Other items.")
            End
        End If
    Else
        End
    End If

Exit Function

ReplyError:
    MsgBox ("You are already replying")
    End
End Function
'=========================================================================
' GRAB EMAIL TO REPLY TO
'=========================================================================
Function GrabEmailToReplyTo()
    Set CurrentEmail = ActiveExplorer.Selection(1).Reply
    Set NewEmail = CurrentEmail.Forward
End Function
'=========================================================================
' CREATE THE USER FORM
'=========================================================================
Function PopulateForm(templateFolderPath As String)

    EmailTemplatesForm.UserForm_Initialize
    Dim MyFile As String
    Dim Counter As Long

    'Create a dynamic array variable, and then declare its initial size
    Dim DirectoryListArray() As String
    ReDim DirectoryListArray(50)

    'To change the caption of the user form
    EmailTemplatesForm.Caption = templateFolderPath + " - Templates"

    'Loop through all the files in the directory by using Dir$ function
    'If statement checks for files that are not .oft and excludes them.
    MyFile = Dir$(TemplatesFolder() & templateFolderPath & "\*.*")
    Do While MyFile <> ""
        If Right(MyFile, Len(MyFile) - InStrRev(MyFile, ".")) = "oft" Then
            DirectoryListArray(Counter) = MyFile
            Counter = Counter + 1
        End If
        MyFile = Dir$
    Loop

    'Reset the size of the array without losing its values by using Redim Preserve
    ReDim Preserve DirectoryListArray(Counter - 1)

    'Populate the combo box
    EmailTemplatesForm.ListBox1.List = DirectoryListArray
    'clearing
    Erase DirectoryListArray
End Function
'=========================================================================
' DISPLAY THE USERFORM AND SET TEMPLATES
'=========================================================================
Function DisplayForm(templateFolderPath)
    'Open the popup
    EmailTemplatesForm.Show

    'MsgBox (EmailTemplatesForm.ComboBox1.Value)

    ' If nothing selected or canceled
    If EmailTemplatesForm.ListBox1.Value <> "" Then
        ' Grab the template
        Set Template = Application.CreateItemFromTemplate(TemplatesFolder() & templateFolderPath & "\" & EmailTemplatesForm.ListBox1.Value)
    Else
        End
    End If
End Function
'=========================================================================
' PREPARE AND DISPLAY EMAIL
'=========================================================================
Function PrepareEmail()
    With NewEmail
        .SentOnBehalfOfName = "youremailhere@email.com" 'Change to the email you are sending as if required.
        .To = "" 'empty the to field'
        .CC = "whoyouwanttoCC@email.com"
        .Subject="Your Subject Line"
        .HTMLBody = Template.HTMLBody & CurrentEmail.HTMLBody 'Combine the template with the email we are replying to
        .Recipients.ResolveAll ' Resolve names
        .Display 'display our new email
    End With

    'Clear variables
    Set Template = Nothing
    Set CurrentEmail = Nothing
    Set NewEmail = Nothing
End Function
'=========================================================================
' This returns the folder path that templates should be in
'=========================================================================
Function TemplatesFolder() As String
    ' Fill in the absolute URL of the Templates Folder
    TemplatesFolder = "C:\.........\......\.......\Templates\"
    If Dir(TemplatesFolder, vbDirectory) = "" Then
        MsgBox "Could not find the folder at """ & TemplatesFolder & """. Did you follow the setup steps and change TemplatesFolder path in the VBA code?"
        End
    ElseIf Right(TemplatesFolder, 1) <> "\" Then
        TemplatesFolder = TemplatesFolder & "\"
    End If
End Function
