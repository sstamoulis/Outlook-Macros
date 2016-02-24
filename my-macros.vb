Public Sub GetItemsFolderPath()
  Dim obj As Object
  Dim F As Outlook.MAPIFolder
  Dim Msg$
  Set obj = Application.ActiveWindow
  If TypeOf obj Is Outlook.Inspector Then
    Set obj = obj.CurrentItem
  Else
    Set obj = obj.Selection(1)
  End If
  Set F = obj.Parent
  Msg = "The path is: " & F.FolderPath & vbCrLf
  Msg = Msg & "Switch to the folder?"
  If MsgBox(Msg, vbYesNo) = vbYes Then
    Set Application.ActiveExplorer.CurrentFolder = F
  End If
End Sub

Public Sub ChangeSender()
    Dim oExplorer As Outlook.Explorer
    Dim oSelection As Outlook.Selection
    Dim oDialog As SelectNamesDialog
    Dim oItem As Object
    Dim oMailItem As MailItem
    
    Set oExplorer = Application.ActiveExplorer
    Set oSelection = oExplorer.Selection
    Set oDialog = Application.Session.GetSelectNamesDialog
    
    With oDialog
        .AllowMultipleSelection = False
        .NumberOfRecipientSelectors = olShowNone
        .Caption = "Please select the sender:"
    End With
    
    If oDialog.Display And oDialog.Recipients.Count > 0 Then
        For i = 1 To oSelection.Count
            Set oItem = oSelection.Item(i)
            If oItem.MessageClass = "IPM.Note" Then
                Set oMailItem = oItem
                oMailItem.Sender = oDialog.Recipients.Item(1).AddressEntry
                oMailItem.Save
            End If
        Next
    End If
End Sub

Public Sub SendAll()
    Dim oExplorer As Outlook.Explorer
    Dim oSelection As Outlook.Selection
    Dim oItem As Object
    Dim oMailItem As MailItem
    
    Set oExplorer = Application.ActiveExplorer
    Set oSelection = oExplorer.Selection
    For i = 1 To oSelection.Count
        Set oItem = oSelection.Item(i)
        If oItem.MessageClass = "IPM.Note" Then
            Set oMailItem = oItem
            oMailItem.Send
        End If
    Next
End Sub
