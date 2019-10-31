Option Compare Database


Private Sub Form_Load()

    Dim NameImport As String
    Dim TextXML As String
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE FROM _tblIMPORT_XML"
    For i = 0 To CurrentProject.ImportExportSpecifications.Count - 1
        NameImport = CurrentProject.ImportExportSpecifications.Item(i).Name
        TextXML = CurrentProject.ImportExportSpecifications.Item(i).XML
        DoCmd.RunSQL "INSERT INTO _tblIMPORT_XML ([NameImport], [TextXML]) VALUES ('" & NameImport & "', '" & TextXML & "')"
    Next i
    DoCmd.SetWarnings True
    Me.Requery

End Sub


Private Sub txtTextXML_Change()

    Dim ss
    ss = txtTextXML.SelStart
    If Me.Dirty Then
        Me.Dirty = False
        txtTextXML.SelStart = ss
    End If

End Sub


Private Sub btnSaveXML_Click()

    Dim NameImport As String
    Dim TextXML As String
    NameImport = Me.Form![NameImport]
    TextXML = Me.Form![TextXML]
    CurrentProject.ImportExportSpecifications.Item(NameImport).XML = Me.Form![TextXML]

End Sub


Private Sub btnDeleteAllXML_Click()

    If MsgBox("Вы уверены что хотите удалить все XML импорта из БД?", vbYesNo + vbExclamation, "Внимание!!!") = vbYes Then
        RunAllDeleteExportSpecificationsXML
        MsgBox "Все XML импорта удалены из БД!", vbExclamation, "Внимание!!!"
        Form_Load
    End If

End Sub


Private Sub btnSaveAllXML_Click()
    
    RunAllSaveExportSpecificationsXML

End Sub


Private Sub btnLoadAllXML_Click()

    RunAllLoadExportSpecificationsXML
    Form_Load

End Sub


Sub RunAllDeleteExportSpecificationsXML()

    For i = CurrentProject.ImportExportSpecifications.Count - 1 To 0 Step -1
        CurrentProject.ImportExportSpecifications.Item(i).Delete
    Next i

End Sub


Sub RunAllSaveExportSpecificationsXML()

    Dim objShellApp As Object
    Dim objFolder As Object
    Dim Path As String
    Set objShellApp = CreateObject("Shell.Application")
    Set objFolder = objShellApp.BrowseForFolder(0, "Выбрать папку", 0, CurrentProject.Path)
    If (Not objFolder Is Nothing) Then
        Path = objFolder.Self.Path & "\"
        For i = 0 To CurrentProject.ImportExportSpecifications.Count - 1
            SaveImportSpecificationsXMLAsText Path, CurrentProject.ImportExportSpecifications.Item(i).Name
        Next i
        MsgBox "Все XML импорта сохранены в директории" & vbCrLf & Path, vbInformation, "Внимание"
    End If
    Set objFolder = Nothing
    Set objShellApp = Nothing

End Sub


Sub RunAllLoadExportSpecificationsXML()
    
    Dim fd As Object
    Set fd = Application.FileDialog(3) 'msoFileDialogOpen
    With fd
        .Title = "Выбрать XML файлы"
        .InitialFileName = ""
        '.ButtonName = ""
        .InitialView = 1
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "XML Files", "*.xml"
        If .Show = True Then
            Dim FileName As Variant
            For Each FileName In .SelectedItems
                LoadImportSpecificationsXMLFromText FileName
            Next
        End If
    End With
    Set fd = Nothing

End Sub


Sub SaveImportSpecificationsXMLAsText(ByVal Path As String, ByVal NameSpec As String)

    Dim text As String
    text = CurrentProject.ImportExportSpecifications.Item(NameSpec).XML
    Dim ff
    ff = FreeFile
    Open Path & NameSpec & ".xml" For Output As #ff
    Print #ff, NameSpec
    Print #ff, text
    Close #ff

End Sub


Sub LoadImportSpecificationsXMLFromText(ByVal FileName As String)
        
        On Error GoTo LoadImportSpecificationsXMLFromText_Err
        
        Dim text As String
100     text = ""
        Dim ff
102     ff = FreeFile
104     Open FileName For Input As ff
106     Line Input #ff, fText
108     NameSpec = fText
110     Do While Not EOF(ff)
112         Line Input #ff, fText
114         text = text & fText & vbCrLf
        Loop
116     Close ff
    
118     CurrentProject.ImportExportSpecifications.Add NameSpec, text

        Exit Sub
LoadImportSpecificationsXMLFromText_Err:
        If Err.Number = 31603 Then
            CurrentProject.ImportExportSpecifications.Item(NameSpec).XML = text
        Else
            MsgBox Err.Description & vbCrLf & _
                   "in Form_frmIMPORT_XML.LoadImportSpecificationsXMLFromText " & _
                   "at line " & Erl, _
                   vbExclamation + vbOKOnly, "Application Error"
        End If

End Sub

