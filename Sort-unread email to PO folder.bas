Attribute VB_Name = "Module1"
Sub SortUnreadEmailsToPOFolders()
    Dim ns As Outlook.NameSpace
    Dim inbox As Outlook.MAPIFolder
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim poNumber As String
    Dim targetFolder As Outlook.MAPIFolder
    Dim regEx As Object
    Dim movedCount As Long
    Dim skippedCount As Long
    Dim missingFolders As String

    MsgBox "Sorting unread emails by PO number has started.", vbInformation

    Set ns = Application.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(olFolderInbox)

    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "PO\d+"
    regEx.IgnoreCase = True
    regEx.Global = False

    movedCount = 0
    skippedCount = 0
    missingFolders = ""

    For i = inbox.Items.Count To 1 Step -1
        If TypeOf inbox.Items(i) Is Outlook.MailItem Then
            Set mail = inbox.Items(i)
            If mail.UnRead Then
                If regEx.Test(mail.Subject) Then
                    poNumber = regEx.Execute(mail.Subject)(0)

                    On Error Resume Next
                    Set targetFolder = inbox.Folders(poNumber)
                    On Error GoTo 0

                    If Not targetFolder Is Nothing Then
                        mail.Move targetFolder
                        movedCount = movedCount + 1
                    Else
                        skippedCount = skippedCount + 1
                        missingFolders = missingFolders & vbCrLf & "- " & poNumber
                        Debug.Print "? Folder not found for: " & poNumber
                    End If
                End If
            End If
        End If
    Next i

    Dim resultMsg As String
    resultMsg = "? Sorting complete!" & vbCrLf & _
                "Emails moved: " & movedCount & vbCrLf & _
                "Skipped (folder not found): " & skippedCount

    If skippedCount > 0 Then
        resultMsg = resultMsg & vbCrLf & vbCrLf & "Missing folders:" & missingFolders
    End If

    MsgBox resultMsg, vbInformation
End Sub


