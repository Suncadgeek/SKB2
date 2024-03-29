﻿Imports System.Windows.Forms
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Text
Imports System.IO
Imports NXOpen
Imports NXOpen.UF
Public Class Form1

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim fd As FolderBrowserDialog = New FolderBrowserDialog()
        Dim strFolder As String

        If fd.ShowDialog() = DialogResult.OK Then
            strFolder = fd.SelectedPath
            TextBox4.Text = strFolder
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


            Dim outputdir As String = TextBox4.Text


            Dim fileNew1 As NXOpen.FileNew = Nothing
            fileNew1 = theSession.Parts.FileNew()
            fileNew1.TemplateFileName = "model-plain-1-mm-template.prt"
            fileNew1.UseBlankTemplate = False
            fileNew1.ApplicationName = "ModelTemplate"
            fileNew1.Units = NXOpen.Part.Units.Millimeters
            fileNew1.RelationType = ""
            fileNew1.UsesMasterModel = "No"
            fileNew1.TemplateType = NXOpen.FileNewTemplateType.Item
            fileNew1.TemplatePresentationName = "Modèle"
            fileNew1.ItemType = ""
            fileNew1.Specialization = ""
            fileNew1.SetCanCreateAltrep(False)

            If CheckBox3.Checked = True Then
                'fileName = xlWorkSheet.Cells(StrtRng + 2, 5).value.ToString
                fileName = tableSectioName((fileIndex - 1) / 2)
                fileNew1.NewFileName = outputdir & "\" & fileName & ".prt"

            Else

                If StrtRng = EndRng Then
                    fileName = "ANS-SD-"
                    SD += 1
                    fileNew1.NewFileName = outputdir & "\" & fileName & SD.ToString & ".prt"

                Else
                    fileName = "ANS-SC-"
                    CL += 1
                    fileNew1.NewFileName = outputdir & "\" & fileName & CL.ToString & ".prt"
                End If

            End If

            If System.IO.File.Exists(outputdir & "\" & fileName) Then
                MsgBox("Un fichier existe déjà. Abandon...")
                End
            End If

            fileNew1.MasterFileName = ""
            fileNew1.MakeDisplayedPart = True
            fileNew1.DisplayPartOption = NXOpen.DisplayPartOption.AllowAdditional
            Dim nXObject1 As NXOpen.NXObject = Nothing
            nXObject1 = fileNew1.Commit()
            theSession.ApplicationSwitchImmediate("UG_APP_MODELING")


            Dim workPart As NXOpen.Part = theSession.Parts.Work
            Dim displayPart As NXOpen.Part = theSession.Parts.Display


            Dim markId1 As NXOpen.Session.UndoMarkId = Nothing
            markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Delete")
            Dim notifyOnDelete1 As Boolean = Nothing
            notifyOnDelete1 = theSession.Preferences.Modeling.NotifyOnDelete
            theSession.UpdateManager.ClearErrorList()
            Dim markId2 As NXOpen.Session.UndoMarkId = Nothing
            markId2 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Delete")
            Dim objects1(0) As NXOpen.TaggedObject
            Dim datumCsys1 As NXOpen.Features.DatumCsys = CType(workPart.Features.FindObject("DATUM_CSYS(0)"), NXOpen.Features.DatumCsys)
            objects1(0) = datumCsys1
            Dim nErrs1 As Integer = Nothing
            nErrs1 = theSession.UpdateManager.AddObjectsToDeleteList(objects1)
            Dim notifyOnDelete2 As Boolean = Nothing
            notifyOnDelete2 = theSession.Preferences.Modeling.NotifyOnDelete
            Dim nErrs2 As Integer = Nothing
            nErrs2 = theSession.UpdateManager.DoUpdate(markId2)
            theSession.DeleteUndoMark(markId1, Nothing)







            SKBuilder(xlWorkSheet, StrtRng, EndRng)

            fileNew1.Destroy()
        Next



        xlWorkBook.Close()
        xlApp.Quit()
        If CheckBox2.Checked = True And CheckBox2.Enabled = True Then Call Create_Assembly(TextBox4.Text)


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
            Dim configr(8) As String


            TextBox1.Text = sr.ReadLine()
            TextBox2.Text = sr.ReadLine()
            TextBox3.Text = sr.ReadLine()
            TextBox4.Text = sr.ReadLine()
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
        Dim configw(8) As String
        configw(0) = TextBox1.Text
        configw(1) = TextBox2.Text
        configw(2) = TextBox3.Text
        configw(3) = TextBox4.Text
        configw(4) = TextBox5.Text
        configw(5) = TextBox6.Text
        configw(6) = CheckBox1.Checked.ToString
        configw(7) = CheckBox2.Checked.ToString
        configw(8) = CheckBox3.Checked.ToString


        For i = 0 To 8
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
End Class

