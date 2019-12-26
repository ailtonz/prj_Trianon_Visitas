Attribute VB_Name = "modTitulo"
Option Compare Database

Function SetApplicationTitle(ByVal MyTitle As String)
   If SetStartupProperty("AppTitle", dbText, MyTitle) Then
      Application.RefreshTitleBar
   Else
      MsgBox "ERROR: Could not set Application Title"
   End If
End Function

Function SetStartupProperty(prpName As String, _
      prpType As Variant, prpValue As Variant) As Integer
   Dim DB As DAO.Database, PRP As DAO.Property, WS As Workspace
   Const ERROR_PROPNOTFOUND = 3270

   Set DB = CurrentDb()

   ' Set the startup property value.
   On Error GoTo Err_SetStartupProperty
   DB.Properties(prpName) = prpValue
   SetStartupProperty = True
   

Bye_SetStartupProperty:
   Exit Function

Err_SetStartupProperty:
   Select Case Err
   ' If the property does not exist, create it and try again.
   Case ERROR_PROPNOTFOUND
      Set PRP = DB.CreateProperty(prpName, prpType, prpValue)
      DB.Properties.Append PRP
      Resume
   Case Else
      SetStartupProperty = False
      Resume Bye_SetStartupProperty
   End Select
End Function

Public Function CurrentMDB() As String
   Dim i As Integer, FullPath As String
   FullPath = CurrentDb.Name
   ' Search backward in string for back slash character.
   For i = Len(FullPath) To 1 Step -1
      ' Return all characters to the right of the back slash.
      If Mid(FullPath, i, 1) = "\" Then
         CurrentMDB = Mid(FullPath, i + 1)
         Exit Function
      End If
   Next i
End Function


