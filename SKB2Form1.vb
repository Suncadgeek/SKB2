Imports System.Windows.Forms
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Text
Imports System.IO
Imports NXOpen
Imports NXOpen.UF
Public Class Form1

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        Dim fd As FolderBrowserDialog = New FolderBrowserDialog()
        Dim strFolder As String

        If fd.ShowDialog() = DialogResult.OK Then
            strFolder = fd.SelectedPath
            'TextBox4.Text = strFolder
        End If

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim strFileName As String

        fd.Title = "Charger tableur squelette machine"
        fd.InitialDirectory = "C:\"
        fd.Filter = "All Excel files|*.xls;*.xlsx;*.xlsxm"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.FileName
            TextBox1.Text = strFileName
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim path_to_excel_file As String = TextBox1.Text
        Label4.Visible = True
        Try
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Open(path_to_excel_file)
            xlWorkSheet = xlWorkBook.Worksheets(1)
        Catch
            MsgBox("Chemin vers tableur non valide, abandon")
            Exit Sub
        End Try

        Dim fileNbr As Integer = Split(TextBox5.Text, ";").Length - 1
        Dim fileIndex, SD, CL As Integer
        Dim StrtRng, EndRng As Integer
        Dim boundArray(1000) As String
        Dim fileName As String = "model"
        SD = 0 : CL = 0

        Dim CellsDictionnary As New Dictionary(Of String, Integer)
        Dim cellName As String
        If CheckBox3.Checked And CheckBox1.Checked Then
            Dim i As Integer = 2

            Do
                cellName = xlWorkSheet.Cells(i, 5).value & xlWorkSheet.Cells(i, 6).value
                If CellsDictionnary.ContainsKey(cellName) Then
                    CellsDictionnary.Item(cellName) += 1
                Else
                    CellsDictionnary.Add(cellName, 1)
                End If
                i += 1
            Loop While xlWorkSheet.Cells(i, 5).value <> ""
            'définir boundarray a partir du dictionnaire de cellule
            Dim j As Integer = 0
            Dim myoffset As Integer = 2
            For Each item In CellsDictionnary
                boundArray(j) = myoffset
                boundArray(j + 1) = boundArray(j) + item.Value - 1
                myoffset = boundArray(j + 1) + 1
                j += 2
            Next item
            fileNbr = j


        Else


            If CheckBox1.Checked = False Then

                boundArray(0) = CInt(TextBox2.Text)
                boundArray(1) = CInt(TextBox3.Text)
                fileNbr = 1

            Else
                If Strings.Right(TextBox5.Text, 1) = ";" Then
                    MsgBox("Erreur d'entrée, le code ne peut pas terminer par un ;")
                    Exit Sub
                End If
                boundArray = Split(TextBox5.Text, ";")
                If (fileNbr \ 2) * 2 = fileNbr Then
                    MsgBox("Erreur d'entrée, le code doit contenir un nombre paire de sections")
                    Exit Sub
                End If
            End If
        End If
        Dim mystep As Integer = 2
        'If CheckBox3.Checked And CheckBox1.Checked Then mystep = 2 Else mystep = 1


        Dim tableSectioName() As String = Split(TextBox6.Text, ";")


        For fileIndex = 1 To fileNbr Step mystep

            StrtRng = boundArray(fileIndex - 1)
            EndRng = boundArray(fileIndex)


            'Dim outputdir As String = TextBox4.Text


            Dim success As Integer
            'Dim filename As String





            If CheckBox3.Checked = True Then
                'fileName = xlWorkSheet.Cells(StrtRng + 2, 5).value.ToString
                fileName = tableSectioName((fileIndex - 1) / 2) & "_SQL"
                success = NewPart(fileName, "Modèle")
            Else

                If StrtRng = EndRng Then
                    fileName = "ANS-SD-"
                    SD += 1
                    success = NewPart(fileName & SD.ToString, "Modèle")

                Else
                    fileName = "ANS-SC-"
                    CL += 1
                    success = NewPart(fileName & CL.ToString & ".prt", "Modèle")
                End If

            End If

            'If System.IO.File.Exists(outputdir & "\" & fileName) Then
            'MsgBox("Un fichier existe déjà. Abandon...")
            'End
            'End If




            SKBuilder(fileName, xlWorkSheet, StrtRng, EndRng)

        Next



        xlWorkBook.Close()
        xlApp.Quit()
        'If CheckBox2.Checked = True And CheckBox2.Enabled = True And CheckBox3.Checked = False Then Call Create_Assembly(TextBox4.Text)
        If CheckBox2.Checked = True And CheckBox2.Enabled = True And CheckBox3.Checked = True Then
            Dim ListOfParts() As NXOpen.BasePart = theSession.Parts.ToArray()
            Array.Reverse(ListOfParts)
            ReDim Preserve ListOfParts(ListOfParts.Length - 1)
            Dim nbparts As Integer = ListOfParts.Length - 1
            'Dim isFirst = True
            For Each part In ListOfParts
                Create_parent(Create_parent(part, True), False)

            Next
            ListOfParts = theSession.Parts.ToArray()

            Dim it As Integer

            For it = 0 To nbparts
                ListOfParts(it) = ListOfParts(it * 2)
            Next

            ReDim Preserve ListOfParts(nbparts)
            Array.Reverse(ListOfParts)

            For i = 0 To ListOfParts.Length - 1 Step 2
                Dim comp1 As BasePart = ListOfParts(i)
                Dim comp2 As BasePart = ListOfParts(i + 1)
                Create_Assembly(comp1, comp2)

            Next


            ListOfParts = theSession.Parts.ToArray()
            ReDim Preserve ListOfParts((nbparts + 1) / 2 - 1)
            Array.Reverse(ListOfParts)



            Create_Master(ListOfParts)



        End If


        Label4.Visible = False
        MsgBox("Opération terminée")
        Close()


        'Else
        'MsgBox("Vérifier les donnée de sélection de zone ou le découpage des zones")

        'End If


    End Sub


    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub
    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            GroupBox1.Enabled = False
            GroupBox2.Enabled = True
            TextBox5.Enabled = True
            CheckBox2.Enabled = True
            CheckBox3.Enabled = True
        Else
            GroupBox1.Enabled = True
            GroupBox2.Enabled = False
            TextBox5.Enabled = False
            CheckBox2.Enabled = False
            CheckBox2.Checked = False
            CheckBox3.Enabled = False
        End If
    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call Read_config()




    End Sub
    Sub Read_config()
        Dim configfilepath As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\Documents\SKBuilder_config.ini"

        If System.IO.File.Exists(configfilepath) Then 'le fichier existe on remplit le form
            Dim sr As New System.IO.StreamReader(configfilepath)
            Dim configr(7) As String


            TextBox1.Text = sr.ReadLine()
            TextBox2.Text = sr.ReadLine()
            TextBox3.Text = sr.ReadLine()
            'TextBox4.Text = sr.ReadLine()
            TextBox5.Text = sr.ReadLine()
            TextBox6.Text = sr.ReadLine()
            CheckBox1.Checked = sr.ReadLine()
            CheckBox2.Checked = sr.ReadLine()
            CheckBox3.Checked = sr.ReadLine()

            sr.Close()


        End If
    End Sub
    Sub Write_config()
        Dim configfilepath As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\Documents\SKBuilder_config.ini"
        Dim sw As New System.IO.StreamWriter(configfilepath) 'le fichier est ensuite sauvegardé
        Dim i As Integer
        sw.Flush()
        Dim configw(7) As String
        configw(0) = TextBox1.Text
        configw(1) = TextBox2.Text
        configw(2) = TextBox3.Text
        'configw(3) = TextBox4.Text
        configw(3) = TextBox5.Text
        configw(4) = TextBox6.Text
        configw(5) = CheckBox1.Checked.ToString
        configw(6) = CheckBox2.Checked.ToString
        configw(7) = CheckBox3.Checked.ToString


        For i = 0 To 7
            sw.WriteLine(configw(i))
        Next
        sw.Close()
    End Sub
    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged, TextBox6.TextChanged

        TextBox5.Text = System.Text.RegularExpressions.Regex.Replace(TextBox5.Text, "[^\d;]", "")    'Removes all character except numbers
        TextBox5.Select(TextBox5.Text.Length + 1, 1)    'To bring the textbox focus to the right

    End Sub

    Private Sub Form1_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Call Write_config()
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then TextBox5.Enabled = False Else TextBox5.Enabled = True
        If CheckBox3.Checked = False Then TextBox6.Enabled = False Else TextBox6.Enabled = True
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        'If CheckBox2.Checked = False Then TextBox5.Enabled = True Else TextBox5.Enabled = False

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs)

    End Sub
    Function NewPart(partname As String, parttype As String)
        'parttype = "Modèle"
        'dim test As String = "changed"
        Dim markId1 As NXOpen.Session.UndoMarkId = Nothing
        markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Départ")

        Dim fileNew1 As NXOpen.FileNew = Nothing
        fileNew1 = theSession.Parts.FileNew()

        fileNew1.TemplateFileName = "@DB/000120/A"

        fileNew1.UseBlankTemplate = False

        If parttype = "Modèle" Then fileNew1.ApplicationName = "ModelTemplate" Else fileNew1.ApplicationName = "AssemblyTemplate"

        fileNew1.Units = NXOpen.Part.Units.Millimeters

        fileNew1.RelationType = "master"

        fileNew1.UsesMasterModel = "No"

        fileNew1.TemplateType = NXOpen.FileNewTemplateType.Item

        fileNew1.TemplatePresentationName = parttype

        fileNew1.ItemType = "SO8_CAD"

        fileNew1.Specialization = ""

        fileNew1.SetCanCreateAltrep(False)

        Dim partOperationCreateBuilder1 As NXOpen.PDM.PartOperationCreateBuilder = Nothing
        partOperationCreateBuilder1 = theSession.PdmSession.CreateCreateOperationBuilder(NXOpen.PDM.PartOperationBuilder.OperationType.Create)

        fileNew1.SetPartOperationCreateBuilder(partOperationCreateBuilder1)

        partOperationCreateBuilder1.SetOperationSubType(NXOpen.PDM.PartOperationCreateBuilder.OperationSubType.FromTemplate)

        partOperationCreateBuilder1.SetModelType("master")

        partOperationCreateBuilder1.SetItemType("SO8_CAD")

        Dim logicalobjects1() As NXOpen.PDM.LogicalObject
        partOperationCreateBuilder1.CreateLogicalObjects(logicalobjects1)

        Dim sourceobjects1() As NXOpen.NXObject
        sourceobjects1 = logicalobjects1(0).GetUserAttributeSourceObjects()

        partOperationCreateBuilder1.DefaultDestinationFolder = ":Newstuff"

        Dim sourceobjects2() As NXOpen.NXObject
        sourceobjects2 = logicalobjects1(0).GetUserAttributeSourceObjects()

        partOperationCreateBuilder1.SetOperationSubType(NXOpen.PDM.PartOperationCreateBuilder.OperationSubType.FromTemplate)

        Dim sourceobjects3() As NXOpen.NXObject
        sourceobjects3 = logicalobjects1(0).GetUserAttributeSourceObjects()

        theSession.SetUndoMarkName(markId1, "Boîte de dialogue Nouvel élément")

        Dim attributetitles1(0) As String
        attributetitles1(0) = "DB_PART_NO"
        Dim titlepatterns1(0) As String
        titlepatterns1(0) = """CAO""nnnnnnnnn"
        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = partOperationCreateBuilder1.CreateAttributeTitleToNamingPatternMap(attributetitles1, titlepatterns1)

        Dim objects1(0) As NXOpen.NXObject
        objects1(0) = logicalobjects1(0)
        Dim properties1(0) As NXOpen.NXObject
        properties1(0) = nXObject1
        Dim errorList1 As NXOpen.ErrorList = Nothing
        errorList1 = partOperationCreateBuilder1.AutoAssignAttributesWithNamingPattern(objects1, properties1)

        errorList1.Dispose()
        Dim errorMessageHandler1 As NXOpen.PDM.ErrorMessageHandler = Nothing
        errorMessageHandler1 = partOperationCreateBuilder1.GetErrorMessageHandler(True)

        Dim nullNXOpen_BasePart As NXOpen.BasePart = Nothing

        Dim objects2(-1) As NXOpen.NXObject
        Dim attributePropertiesBuilder1 As NXOpen.AttributePropertiesBuilder = Nothing
        attributePropertiesBuilder1 = theSession.AttributeManager.CreateAttributePropertiesBuilder(nullNXOpen_BasePart, objects2, NXOpen.AttributePropertiesBuilder.OperationType.None)

        Dim objects3(-1) As NXOpen.NXObject
        attributePropertiesBuilder1.SetAttributeObjects(objects3)

        Dim objects4(0) As NXOpen.NXObject
        objects4(0) = sourceobjects1(0)
        attributePropertiesBuilder1.SetAttributeObjects(objects4)

        attributePropertiesBuilder1.Title = "DB_PART_NAME"

        attributePropertiesBuilder1.Category = "SO8_CAD"

        attributePropertiesBuilder1.StringValue = partname

        attributePropertiesBuilder1.Category = "SO8_CAD"

        Dim changed1 As Boolean = Nothing
        changed1 = attributePropertiesBuilder1.CreateAttribute()

        Dim markId2 As NXOpen.Session.UndoMarkId = Nothing
        markId2 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Nouvel élément")

        theSession.DeleteUndoMark(markId2, Nothing)

        Dim markId3 As NXOpen.Session.UndoMarkId = Nothing
        markId3 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Nouvel élément")

        fileNew1.MasterFileName = ""

        fileNew1.MakeDisplayedPart = True

        fileNew1.DisplayPartOption = NXOpen.DisplayPartOption.AllowAdditional

        partOperationCreateBuilder1.ValidateLogicalObjectsToCommit()

        Dim logicalobjects2(0) As NXOpen.PDM.LogicalObject
        logicalobjects2(0) = logicalobjects1(0)
        partOperationCreateBuilder1.CreateSpecificationsForLogicalObjects(logicalobjects2)

        Dim errorMessageHandler2 As NXOpen.PDM.ErrorMessageHandler = Nothing
        errorMessageHandler2 = partOperationCreateBuilder1.GetErrorMessageHandler(True)

        Dim errorMessageHandler3 As NXOpen.PDM.ErrorMessageHandler = Nothing
        errorMessageHandler3 = partOperationCreateBuilder1.GetErrorMessageHandler(True)

        Dim nXObject2 As NXOpen.NXObject = Nothing
        nXObject2 = fileNew1.Commit()


        workPart = theSession.Parts.Work ' CAO000083658/AA-TEST
        displayPart = theSession.Parts.Display ' CAO000083658/AA-TEST
        Dim errorMessageHandler4 As NXOpen.PDM.ErrorMessageHandler = Nothing
        errorMessageHandler4 = partOperationCreateBuilder1.GetErrorMessageHandler(True)

        theSession.DeleteUndoMark(markId3, Nothing)

        fileNew1.Destroy()

        attributePropertiesBuilder1.Destroy()
        Dim markId4 As NXOpen.Session.UndoMarkId = Nothing
        markId4 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Delete")

        theSession.UpdateManager.ClearErrorList()

        Dim markId5 As NXOpen.Session.UndoMarkId = Nothing
        markId5 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Delete")

        Dim objects20(0) As NXOpen.TaggedObject
        Dim datumCsys1 As NXOpen.Features.DatumCsys = CType(workPart.Features.FindObject("DATUM_CSYS(0)"), NXOpen.Features.DatumCsys)

        objects20(0) = datumCsys1
        Dim nErrs1 As Integer = Nothing
        nErrs1 = theSession.UpdateManager.AddObjectsToDeleteList(objects20)

        Dim id1 As NXOpen.Session.UndoMarkId = Nothing
        id1 = theSession.NewestVisibleUndoMark

        Dim nErrs2 As Integer = Nothing
        nErrs2 = theSession.UpdateManager.DoUpdate(id1)

        theSession.DeleteUndoMark(markId4, Nothing)

        Return 0
    End Function
End Class

