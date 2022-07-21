Option Infer On
Option Strict Off
Imports System.Threading
Imports System
Imports NXOpen
Imports NXOpen.UF
Imports Excel = Microsoft.Office.Interop.Excel

Module _3DBuilderFork




    Public the3DSession As Session = Session.GetSession()
    Public ufs As UFSession = UFSession.GetUFSession()
    Public the3DworkPart As Part = the3DSession.Parts.Work
    Public lw As ListingWindow = the3DSession.ListingWindow
    Public the3DdisplayPart As NXOpen.Part = the3DSession.Parts.Display
    Public referenceSet1 As NXOpen.ReferenceSet

    Sub Godmode()
        Dim ListOfParts() As NXOpen.BasePart = the3DSession.Parts.ToArray()
        For i = 0 To ListOfParts.Length - 1
            Outline(ListOfParts(i))
        Next
    End Sub


    Public Sub Outline(skelbasePart1)

        ' opening all necessary parts
        Dim CAOcomp(200, 2)
        CAOcomp = Open_Parts(skelbasePart1)

        ' creating master assembly
        Dim Skelcomp As NXObject = Create_3DAssembly()

        ' importing magnets as component in master assembly
        Import_Magnets(Skelcomp, CAOcomp)



    End Sub

    Function Open_Parts(basePart1 As NXOpen.BasePart)



        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim path_to_excel_file As String = Form1.TextBox1.Text

        Try
            xlApp = New Excel.Application

            xlWorkBook = xlApp.Workbooks.Open(path_to_excel_file)

            xlWorkSheet = xlWorkBook.Worksheets(1)

        Catch
            MsgBox("Chemin vers tableur non valide, abandon")




            End
        End Try
        Dim CAOitem(200, 2) As String
        Dim CAOcomp(200, 2) As String
        Dim CAOdic As New Dictionary(Of String, Integer)

        Dim i As Integer = 0
        While xlWorkSheet.Cells(i + 1, 1).value <> ""
            CAOitem(i, 0) = xlWorkSheet.Cells(i + 1, 1).value.ToString
            CAOitem(i, 1) = xlWorkSheet.Cells(i + 1, 2).value.ToString
            CAOitem(i, 2) = xlWorkSheet.Cells(i + 1, 4).value.ToString

            i += 1
        End While
        Dim nbMagnets As Integer = i




        Dim magnetname As String
        'Dim basePart1 As NXOpen.BasePart = Nothing
        Dim partLoadStatus1 As NXOpen.PartLoadStatus = Nothing
        'loading skeleton
        'skeletonpath = "@DB/" & Form1.TextBox2.Text
        'basePart1 = theSession.Parts.OpenActiveDisplay(skeletonpath, NXOpen.DisplayPartOption.AllowAdditional, partLoadStatus1)
        the3DworkPart = the3DSession.Parts.Work
        the3DdisplayPart = the3DSession.Parts.Display
        partLoadStatus1.Dispose()

        'loading magnets

        ' Dim files() As String = IO.Directory.GetFiles(magnetspath)

        Dim SkelPart As Part = basePart1 'defining te skeleton part
        For i = 0 To SkelPart.Features.GetFeatures.Length - 1 'par
            Dim CurrentFeat As Features.Feature = SkelPart.Features.GetFeatures(i)
            Dim CurrentFeatName As String = Split(CurrentFeat.Name, ".")(0)
            If InStr(CurrentFeatName, "_") <> 0 Then CurrentFeatName = Split(CurrentFeatName, "_")(0)
            If CurrentFeat.FeatureType.ToString = "DATUM_CSYS" And InStr(1, CurrentFeat.Name, "Entrée", 1) = 0 And InStr(1, CurrentFeat.Name, "Sortie", 1) = 0 And InStr(1, CurrentFeat.Name, "1/3", 1) = 0 And InStr(1, CurrentFeat.Name, "3/3", 1) = 0 Then
                If CAOdic.ContainsKey(CurrentFeatName) Then
                    CAOdic.Item(CurrentFeatName) += 1
                Else
                    CAOdic.Add(CurrentFeatName, 1)
                End If
            End If
        Next
        i = 0
        For Each item In CAOdic



            For j = 0 To nbMagnets - 1
                If CAOitem(j, 2) = item.Key Then
                    CAOcomp(i, 0) = CAOitem(j, 0)
                    CAOcomp(i, 1) = CAOitem(j, 1)
                    CAOcomp(i, 2) = item.Key
                    i += 1
                End If
            Next

        Next
        nbMagnets = i

        For i = 0 To nbMagnets - 1
            magnetname = "@DB/" & CAOcomp(i, 0) & "/" & CAOcomp(i, 1)
            'MsgBox(magnetname)
            basePart1 = the3DSession.Parts.OpenActiveDisplay(magnetname, NXOpen.DisplayPartOption.AllowAdditional, partLoadStatus1)
            the3DworkPart = the3DSession.Parts.Work
            the3DdisplayPart = the3DSession.Parts.Display
            partLoadStatus1.Dispose()


        Next
        'MsgBox("all magnets loaded")

        xlWorkBook.Close()
        xlApp.Quit()

        Return CAOcomp
    End Function
    Function Create_3DAssembly()

        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        ' ----------------------------------------------
        '    Menu: Fichier->Nouveau->Elément...
        ' ----------------------------------------------
        Dim markId1 As NXOpen.Session.UndoMarkId
        markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Départ")
        Dim fileNew1 As NXOpen.FileNew
        fileNew1 = theSession.Parts.FileNew()
        fileNew1.TemplateFileName = "@DB/000120/A"
        fileNew1.UseBlankTemplate = False
        fileNew1.ApplicationName = "AssemblyTemplate"
        fileNew1.Units = NXOpen.Part.Units.Millimeters
        fileNew1.RelationType = "master"
        fileNew1.UsesMasterModel = "No"
        fileNew1.TemplateType = NXOpen.FileNewTemplateType.Item
        fileNew1.TemplatePresentationName = "Assemblage"
        fileNew1.ItemType = "SO8_CAD"
        fileNew1.Specialization = ""
        fileNew1.SetCanCreateAltrep(False)
        Dim partOperationCreateBuilder1 As NXOpen.PDM.PartOperationCreateBuilder
        partOperationCreateBuilder1 = theSession.PdmSession.CreateCreateOperationBuilder(NXOpen.PDM.PartOperationBuilder.OperationType.Create)
        fileNew1.SetPartOperationCreateBuilder(partOperationCreateBuilder1)
        partOperationCreateBuilder1.SetOperationSubType(NXOpen.PDM.PartOperationCreateBuilder.OperationSubType.FromTemplate)
        partOperationCreateBuilder1.SetModelType("master")
        partOperationCreateBuilder1.SetItemType("SO8_CAD")
        Dim logicalobjects1() As NXOpen.PDM.LogicalObject = Nothing
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
        Dim nXObject1 As NXOpen.NXObject
        nXObject1 = partOperationCreateBuilder1.CreateAttributeTitleToNamingPatternMap(attributetitles1, titlepatterns1)
        Dim objects1(0) As NXOpen.NXObject
        objects1(0) = logicalobjects1(0)
        Dim properties1(0) As NXOpen.NXObject
        properties1(0) = nXObject1
        Dim errorList1 As NXOpen.ErrorList
        errorList1 = partOperationCreateBuilder1.AutoAssignAttributesWithNamingPattern(objects1, properties1)
        errorList1.Dispose()
        Dim errorMessageHandler1 As NXOpen.PDM.ErrorMessageHandler
        errorMessageHandler1 = partOperationCreateBuilder1.GetErrorMessageHandler(True)
        Dim markId2 As NXOpen.Session.UndoMarkId
        markId2 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Nouvel élément")
        theSession.DeleteUndoMark(markId2, Nothing)
        Dim markId3 As NXOpen.Session.UndoMarkId
        markId3 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Nouvel élément")
        fileNew1.MasterFileName = "3DBuild"
        fileNew1.MakeDisplayedPart = True
        fileNew1.DisplayPartOption = NXOpen.DisplayPartOption.AllowAdditional
        partOperationCreateBuilder1.ValidateLogicalObjectsToCommit()
        Dim logicalobjects2(0) As NXOpen.PDM.LogicalObject
        logicalobjects2(0) = logicalobjects1(0)
        partOperationCreateBuilder1.CreateSpecificationsForLogicalObjects(logicalobjects2)
        Dim errorMessageHandler2 As NXOpen.PDM.ErrorMessageHandler
        errorMessageHandler2 = partOperationCreateBuilder1.GetErrorMessageHandler(True)
        Dim errorMessageHandler3 As NXOpen.PDM.ErrorMessageHandler
        errorMessageHandler3 = partOperationCreateBuilder1.GetErrorMessageHandler(True)




        Dim nXObject2 As NXOpen.NXObject
        nXObject2 = fileNew1.Commit()
        Dim workPart As NXOpen.Part = theSession.Parts.Work
        Dim displayPart As NXOpen.Part = theSession.Parts.Display
        Dim errorMessageHandler4 As NXOpen.PDM.ErrorMessageHandler
        errorMessageHandler4 = partOperationCreateBuilder1.GetErrorMessageHandler(True)
        theSession.DeleteUndoMark(markId3, Nothing)
        fileNew1.Destroy()
        theSession.ApplicationSwitchImmediate("UG_APP_MODELING")



        Dim ListOfParts() As NXOpen.BasePart = theSession.Parts.ToArray()
        Dim AssemblyPart As Part = ListOfParts(0)
        Array.Reverse(ListOfParts)
        Dim SkelPart As Part = ListOfParts(0)
        Dim addComponentBuilder1 As NXOpen.Assemblies.AddComponentBuilder
        addComponentBuilder1 = AssemblyPart.AssemblyManager.CreateAddComponentBuilder()
        addComponentBuilder1.ReferenceSet = "MODEL"
        Dim TempPartList(0) As BasePart
        TempPartList(0) = SkelPart
        addComponentBuilder1.SetPartsToAdd(TempPartList)
        Dim SkelComp As NXObject = addComponentBuilder1.Commit()
        addComponentBuilder1.Destroy()


        ' set fix anchor for skeleton

        Dim componentPositioner1 As NXOpen.Positioning.ComponentPositioner
        componentPositioner1 = AssemblyPart.ComponentAssembly.Positioner
        componentPositioner1.ClearNetwork()
        componentPositioner1.BeginAssemblyConstraints()
        Dim network1 As NXOpen.Positioning.Network
        network1 = componentPositioner1.EstablishNetwork()
        Dim componentNetwork1 As NXOpen.Positioning.ComponentNetwork = CType(network1, NXOpen.Positioning.ComponentNetwork)
        componentNetwork1.MoveObjectsState = True
        Dim nullNXOpen_Assemblies_Component As NXOpen.Assemblies.Component = Nothing
        componentNetwork1.DisplayComponent = nullNXOpen_Assemblies_Component
        componentNetwork1.NetworkArrangementsMode = NXOpen.Positioning.ComponentNetwork.ArrangementsMode.Existing
        componentNetwork1.MoveObjectsState = True
        'Dim constraint1 As NXOpen.Positioning.Constraint
        'constraint1 = componentPositioner1.CreateConstraint(True)
        'Dim componentConstraint1 As NXOpen.Positioning.ComponentConstraint = CType(constraint1, NXOpen.Positioning.ComponentConstraint)
        'componentConstraint1.ConstraintAlignment = NXOpen.Positioning.Constraint.Alignment.InferAlign
        'componentConstraint1.ConstraintType = NXOpen.Positioning.Constraint.Type.Touch
        Dim constraint2 As NXOpen.Positioning.Constraint
        constraint2 = componentPositioner1.CreateConstraint(True)
        Dim componentConstraint2 As NXOpen.Positioning.ComponentConstraint = CType(constraint2, NXOpen.Positioning.ComponentConstraint)
        componentConstraint2.ConstraintType = NXOpen.Positioning.Constraint.Type.Fix
        Dim constraintReference3 As NXOpen.Positioning.ConstraintReference = componentConstraint2.CreateConstraintReference(SkelComp, SkelComp, False, False, False)
        constraintReference3.SetFixHint(True)
        componentNetwork1.Solve()
        componentNetwork1.Solve()
        componentPositioner1.ClearNetwork()
        componentPositioner1.DeleteNonPersistentConstraints()
        componentPositioner1.EndAssemblyConstraints()
        Dim displayedConstraint2 As NXOpen.Positioning.DisplayedConstraint = constraint2.GetDisplayedConstraint
        displayedConstraint2.Blank()



        Return SkelComp

    End Function
    Sub Import_Magnets(SkelComp As NXObject, CAOcomp As Array)

        Dim i As Integer
        Dim partIndex As Integer

        Dim ListOfParts() As NXOpen.BasePart = the3DSession.Parts.ToArray()
        Dim AssemblyPart As Part
        Dim TempPartList(0) As BasePart

        AssemblyPart = ListOfParts(0) 'Defining the Main Assembly part
        'MsgBox("main assembly part is " & AssemblyPart.Name)
        Dim SkelPart As Part = ListOfParts(ListOfParts.Length - 1) 'defining te skeleton part
        'MsgBox("skel  part is " & SkelPart.Name)
        Dim nbmagnets As Integer = CInt((SkelPart.Features.GetFeatures.Length - 1) / 10)
        Dim pgBarInc As Integer = CInt(700 / nbmagnets)


        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim path_to_excel_file As String = Form1.TextBox1.Text
        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open(path_to_excel_file)
        xlWorkSheet = xlWorkBook.Worksheets(1)
        'Dim CAOitem(200) As String
        'Dim CAOrev(200) As String
        'Dim CAOcode(200) As String
        'i = 0
        'While xlWorkSheet.Cells(i + 2, 1).value <> ""
        '    CAOitem(i) = xlWorkSheet.Cells(i + 1, 1).value.ToString
        '    CAOcode(i) = xlWorkSheet.Cells(i + 1, 4).value.ToString
        '    CAOrev(i) = xlWorkSheet.Cells(i + 1, 2).value.ToString

        '    i += 1
        'End While
        'ReDim Preserve CAOcode(i)
        'ReDim Preserve CAOrev(i)
        'ReDim Preserve CAOitem(i)



        'MsgBox(ListOfParts.Length)

        Dim j, k As Integer
        For i = 0 To SkelPart.Features.GetFeatures.Length - 1 'parsing skeleton feature to import the right magnets
            Dim CurrentFeat As Features.Feature = SkelPart.Features.GetFeatures(i)
            ' when a magnet CSYS feature is parsed :
            'MsgBox(CurrentFeat.Name)
            If InStr(CurrentFeat.Name, "SHF") <> 0 Then
                Dim stopstr As String = "Stop"

            End If
            If CurrentFeat.FeatureType.ToString = "DATUM_CSYS" And InStr(1, CurrentFeat.Name, "Entrée", 1) = 0 And InStr(1, CurrentFeat.Name, "Sortie", 1) = 0 And InStr(1, CurrentFeat.Name, "Drift", 1) = 0 And InStr(1, CurrentFeat.Name, "1/3", 1) = 0 And InStr(1, CurrentFeat.Name, "3/3", 1) = 0 Then


                Dim FeatName As String = Split(CurrentFeat.Name, ".")(0)
                Dim MatchedFound As Boolean = False
                Dim test As String = "change"
                j = 0
                For Each item In CAOcomp

                    k = 0
                    If CAOcomp(j, 2) = Split(FeatName, "_")(0) Then

                        For Each partItem In ListOfParts
                            If CAOcomp(j, 0) = Split(partItem.Name, "/")(0) Then
                                partIndex = k
                                MatchedFound = True
                                Exit For
                            End If
                            k += 1
                        Next
                        If MatchedFound Then Exit For
                    End If
                    j += 1
                    If j > 200 Then Exit For
                Next

                'adding the right magnet as component to the main assembly
                If MatchedFound = True Then

                    Dim addComponentBuilder1 As NXOpen.Assemblies.AddComponentBuilder
                    addComponentBuilder1 = AssemblyPart.AssemblyManager.CreateAddComponentBuilder()

                    addComponentBuilder1.ReferenceSet = "MODEL"
                    TempPartList(0) = ListOfParts(partIndex)
                    ' MsgBox("featname Is " & FeatName) 'sextsextupole 1
                    'MsgBox("tempartlist(0).name Is " & TempPartList(0).Name) '73750/AA
                    'MsgBox("part index Is " & partIndex) '0
                    'MsgBox(AssemblyPart.Name) ''73750/AA
                    addComponentBuilder1.SetPartsToAdd(TempPartList)
                    Dim CurrentMagComp As NXObject = addComponentBuilder1.Commit()

                    addComponentBuilder1.Destroy()

                    'set constraint between skeleton comp and current magnet comp within Assembly at CurrentFeat (CSYS)")
                    'MsgBox(SkelComp.Name) '73714
                    'MsgBox(CurrentMagComp.Name)
                    'MsgBox(CurrentFeat.Name)
                    'MsgBox(AssemblyPart.Name)

                    SetConstraints(SkelComp, CurrentMagComp, AssemblyPart, CurrentFeat.Name)

                End If

            End If




        Next

    End Sub



    Sub SetConstraints(SkelComp As NXObject, CurrentMagComp As NXObject, AssemblyPart As Part, SkelCSYSName As String)
        Dim SkelCSYS As NXOpen.CartesianCoordinateSystem
        Dim SkelCSYSji As String = Nothing
        Dim MagnetCSYSji As String = Nothing
        'MsgBox(SkelComp.Name & CurrentMagComp.Name & AssemblyPart.Name & SkelCSYSName)
        Dim componentPositioner1 As NXOpen.Positioning.ComponentPositioner
        componentPositioner1 = AssemblyPart.ComponentAssembly.Positioner
        componentPositioner1.ClearNetwork()
        componentPositioner1.BeginAssemblyConstraints()
        Dim network1 As NXOpen.Positioning.Network
        network1 = componentPositioner1.EstablishNetwork()
        Dim componentNetwork1 As NXOpen.Positioning.ComponentNetwork = CType(network1, NXOpen.Positioning.ComponentNetwork)
        componentNetwork1.MoveObjectsState = True
        Dim nullNXOpen_Assemblies_Component As NXOpen.Assemblies.Component = Nothing
        componentNetwork1.DisplayComponent = nullNXOpen_Assemblies_Component
        componentNetwork1.NetworkArrangementsMode = NXOpen.Positioning.ComponentNetwork.ArrangementsMode.Existing
        componentNetwork1.MoveObjectsState = True
        Dim constraint1 As NXOpen.Positioning.Constraint
        constraint1 = componentPositioner1.CreateConstraint(True)
        Dim componentConstraint1 As NXOpen.Positioning.ComponentConstraint = CType(constraint1, NXOpen.Positioning.ComponentConstraint)
        componentConstraint1.ConstraintAlignment = NXOpen.Positioning.Constraint.Alignment.InferAlign
        componentConstraint1.ConstraintType = NXOpen.Positioning.Constraint.Type.Touch


        'parsing journal identifier of CSYS feature
        For Each tempfeature As Features.Feature In SkelComp.Prototype.OwningPart.Features
            If tempfeature.Name = SkelCSYSName And tempfeature.FeatureType.ToString = "DATUM_CSYS" And InStr(1, tempfeature.Name, "Entrée", 1) = 0 And InStr(1, tempfeature.Name, "Sortie", 1) = 0 And InStr(1, tempfeature.Name, "Drift", 1) = 0 And InStr(1, tempfeature.Name, "1/3", 1) = 0 And InStr(1, tempfeature.Name, "3/3", 1) = 0 Then
                SkelCSYSji = tempfeature.JournalIdentifier
            End If

        Next
        SkelCSYS = CType(SkelComp.FindObject("PROTO#.Features|" & SkelCSYSji & "|CSYSTEM 1"), NXOpen.CartesianCoordinateSystem)

        Dim constraintReference1 As NXOpen.Positioning.ConstraintReference = componentConstraint1.CreateConstraintReference(SkelComp, SkelCSYS, False, False)


        'parsing journal identifier of CSYS feature
        Dim MagnetCSYS As NXOpen.CartesianCoordinateSystem
        For Each tempfeature As Features.Feature In CurrentMagComp.Prototype.OwningPart.Features
            'MsgBox("currentfeat is " & tempfeature.Name & " and instrbool is " & InStr(tempfeature.Name, "RPM"))
            If InStr(tempfeature.Name, "RPM") <> 0 And tempfeature.FeatureType.ToString = "DATUM_CSYS" Then

                MagnetCSYSji = tempfeature.JournalIdentifier
            End If

        Next
        MagnetCSYS = CType(CurrentMagComp.FindObject("PROTO#.Features|" & MagnetCSYSji & "|CSYSTEM 1"), NXOpen.CartesianCoordinateSystem)
        'MsgBox(MagnetCSYS.Name)
        Dim constraintReference2 As NXOpen.Positioning.ConstraintReference = componentConstraint1.CreateConstraintReference(CurrentMagComp, MagnetCSYS, False, False)
        constraintReference2.SetFixHint(True)
        componentNetwork1.Solve()


        componentPositioner1.ClearNetwork()
        componentPositioner1.DeleteNonPersistentConstraints()
        componentPositioner1.EndAssemblyConstraints()
        Dim displayedConstraint1 As NXOpen.Positioning.DisplayedConstraint = constraint1.GetDisplayedConstraint
        displayedConstraint1.Blank()


    End Sub
    Public Function GetUnloadOption() As Integer

        'Unloads the image when the NX session terminates
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately

    End Function


End Module
