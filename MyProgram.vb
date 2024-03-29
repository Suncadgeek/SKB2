'NXJournaling.com
'create associative line example journal
'creates line from (1,2,0) to (4,5,0)
'tested on NX 7.5 and 8
'test commit 01

Option Infer On
Option Strict Off
Imports System.Threading
Imports System
Imports NXOpen
Imports NXOpen.UF
Imports Excel = Microsoft.Office.Interop.Excel
'Imports MSForms


Module Module1


    Public theSession As Session = Session.GetSession()
    Public ufs As UFSession = UFSession.GetUFSession()
    Public workPart As Part = theSession.Parts.Work
    Public lw As ListingWindow = theSession.ListingWindow
    Public displayPart As NXOpen.Part = theSession.Parts.Display
    Public referenceSet1, referenceSet2 As NXOpen.ReferenceSet
    Public multibend As Boolean = False
    Public Sub Main()

        Form1.Show()

    End Sub
    Public Sub SKBuilder(xlWorkSheet, startRange, endRange)

        ufs.Disp.SetDisplay(UFConstants.UF_DISP_SUPPRESS_DISPLAY)

        Dim nbdrift, nbquad, nbsext, nboctu, nbbend, nbrevbend As Integer
        nbdrift = 1 : nbquad = 1 : nbsext = 1 : nboctu = 1 : nbbend = 1 : nbrevbend = 1
        workPart = theSession.Parts.Work
        displayPart = theSession.Parts.Display
        referenceSet1 = workPart.CreateReferenceSet()
        referenceSet1.SetName("MODEL")
        referenceSet2 = workPart.CreateReferenceSet()
        referenceSet2.SetName("MAIN CSYS")

        'Dim x1, x2, z1, z2, xmid, zmid As Decimal
        Dim x2, z2, xmid, zmid As Decimal
        Dim x1exp, z1exp As NXOpen.Expression
        Dim unit1 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
        x1exp = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)
        z1exp = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)
        Dim teta1 As Decimal
        Dim LineName As String
        'x1 = 0
        'z1 = 0
        Dim CsysTriplet() As NXOpen.Features.Feature
        Dim featuresToGroup As New List(Of Tag)
        Dim theFeatureGroupTag As Tag
        Dim MagnetsDictionnary As New Dictionary(Of String, Integer)
        For rowIndex = startRange To endRange
            Dim long1 As Decimal = xlWorkSheet.Cells(rowIndex, 2).value
            teta1 = xlWorkSheet.Cells(rowIndex, 3).value / 1000

            LineName = xlWorkSheet.Cells(rowIndex, 7).value
            If LineName.Replace(" ", "") = "" Then LineName = "Drift"


            If MagnetsDictionnary.ContainsKey(LineName) Then
                If xlWorkSheet.Cells(rowIndex, 1).value = "Bend" Then
                    If xlWorkSheet.Cells(rowIndex, 1).value = "Bend" And xlWorkSheet.Cells(rowIndex + 1, 1).value = "Bend" And xlWorkSheet.Cells(rowIndex + 2, 1).value = "Bend" Then
                        'bas champ n°1, incrément
                        MagnetsDictionnary.Item(LineName) += 1
                    ElseIf rowIndex > 2 And xlWorkSheet.Cells(rowIndex, 1).value = "Bend" And xlWorkSheet.Cells(rowIndex - 1, 1).value = "Bend" And xlWorkSheet.Cells(rowIndex - 2, 1).value = "Bend" Then
                        'bas champ n°2, pas d'incrément
                    Else
                        'fort champ, pas d'incrément

                    End If
                Else
                    MagnetsDictionnary.Item(LineName) += 1
                End If

            Else
                    MagnetsDictionnary.Add(LineName, 1)
            End If

            LineName = LineName & "." & MagnetsDictionnary.Item(LineName).ToString("D2")


            'Select Case LineName
            '    Case "Drift"
            '        LineName = xlWorkSheet.Cells(rowIndex, 6).value & " " & nbdrift.ToString("D2")
            '        nbdrift += 1
            '    Case "Quadrupole"
            '        LineName = xlWorkSheet.Cells(rowIndex, 6).value & " " & nbquad.ToString("D2")
            '        nbquad += 1
            '    Case "Sextupole"
            '        LineName = xlWorkSheet.Cells(rowIndex, 6).value & " " & nbsext.ToString("D2")
            '        nbsext += 1
            '    Case "Bend"
            '        If teta1 > 0 Then
            '            LineName = xlWorkSheet.Cells(rowIndex, 6).value & " " & nbbend.ToString("D2")
            '            If xlWorkSheet.Cells(rowIndex, 1).value = "Bend" And xlWorkSheet.Cells(rowIndex + 1, 1).value = "Bend" And xlWorkSheet.Cells(rowIndex + 2, 1).value = "Bend" Then

            '            ElseIf rowIndex > 2 Then
            '                If xlWorkSheet.Cells(rowIndex, 1).value = "Bend" And xlWorkSheet.Cells(rowIndex - 1, 1).value = "Bend" And xlWorkSheet.Cells(rowIndex - 2, 1).value = "Bend" Then

            '                End If
            '            Else
            '                nbbend += 1
            '            End If



            '        Else
            '            LineName = "Reverse Bend" & " " & nbrevbend.ToString("D2")
            '            nbrevbend += 1
            '        End If

            '    Case "Octupole"
            '        LineName = xlWorkSheet.Cells(rowIndex, 1).value & " " & nboctu.ToString("D2")
            '        nboctu += 1
            'End Select

            'Select Case InStr(LineName, "Bend")
            'Case 0 'pas un arc


            If teta1 = 0 Then
                x2 = 0
                z2 = long1
                Dim LineTuplet = MakeLine(x1exp, z1exp, x2, z2, LineName.ToUpper)
                Dim Line1Feature As Features.Feature = LineTuplet.Item1
                x1exp = LineTuplet.Item2
                z1exp = LineTuplet.Item3
                Dim Line1 As Line = Line1Feature.GetEntities(0)

                Dim components1(0) As NXOpen.NXObject
                components1(0) = Line1
                referenceSet1.AddObjectsToReferenceSet(components1)
                referenceSet2.AddObjectsToReferenceSet(components1)
                featuresToGroup.Add(Line1Feature.Tag)
                CsysTriplet = MakeCSYS(Line1, LineName, teta1)
            Else 'est un dipole ou dipole inversé

                'Case Else 'bend ou reverse bend
                Dim r As Decimal = long1 / teta1
                'Dim r As Decimal = xlWorkSheet.Cells(rowIndex, 7).value

                x2 = r * Math.Cos(teta1) - r
                z2 = r * Math.Sin(teta1)
                xmid = r * Math.Cos(teta1 / 2) - r
                zmid = r * Math.Sin(teta1 / 2)
                Dim ArcTuplet = MakeArc(x1exp, z1exp, x2, z2, xmid, zmid, LineName.ToUpper)
                Dim Arc1Feature = ArcTuplet.Item1
                x1exp = ArcTuplet.Item2
                z1exp = ArcTuplet.Item3
                Dim Arc1 As Arc = Arc1Feature.getentities(0)
                Dim components1(0) As NXOpen.NXObject
                components1(0) = Arc1
                referenceSet1.AddObjectsToReferenceSet(components1)
                referenceSet2.AddObjectsToReferenceSet(components1)
                featuresToGroup.Add(Arc1Feature.Tag)
                If xlWorkSheet.Cells(rowIndex, 1).value = "Bend" And xlWorkSheet.Cells(rowIndex + 1, 1).value = "Bend" And xlWorkSheet.Cells(rowIndex + 2, 1).value = "Bend" Then
                    ' LineName = LineName & " Bas Champ"

                    LineName = Left(LineName, LineName.Length - 3) & "_1/3" & Right(LineName, 3)
                    multibend = True
                ElseIf rowIndex > 2 And xlWorkSheet.Cells(rowIndex, 1).value = "Bend" And xlWorkSheet.Cells(rowIndex - 1, 1).value = "Bend" And xlWorkSheet.Cells(rowIndex - 2, 1).value = "Bend" Then
                    LineName = Left(LineName, LineName.Length - 3) & "_3/3" & Right(LineName, 3)
                    multibend = False
                Else
                    If teta1 > 0 And multibend Then LineName = Left(LineName, LineName.Length - 3) & "_2/3" & Right(LineName, 3)
                End If





                CsysTriplet = MakeCSYS(Arc1, LineName, teta1)
            End If
            ' End Select


            For i As Integer = 0 To 2
                featuresToGroup.Add(CsysTriplet(i).Tag)
            Next



            ufs.Modl.CreateSetOfFeature(LineName.ToUpper, featuresToGroup.ToArray, featuresToGroup.Count, 1, theFeatureGroupTag)
            featuresToGroup.Clear()
        Next



        ufs.Disp.SetDisplay(UFConstants.UF_DISP_UNSUPPRESS_DISPLAY)
        ufs.Disp.RegenerateDisplay()


    End Sub
    Function MakeArc(x1exp As Expression, z1exp As Expression, coordx2 As Decimal, coordz2 As Decimal, coordxmid As Decimal, coordzmid As Decimal, LineName As String)
        'Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        'Dim workPart As NXOpen.Part = theSession.Parts.Work
        'Dim displayPart As NXOpen.Part = theSession.Parts.Display
        ' ----------------------------------------------
        '    Menu: Insertion->Courbe->Arc/Cercle...
        ' ----------------------------------------------
        'Dim startpoint3D As Point3d
        'startpoint3D.X = coordx1
        'startpoint3D.Y = 0
        'startpoint3D.Z = coordz1

        Dim midpoint3D As Point3d
        midpoint3D.X = coordxmid
        midpoint3D.Y = 0
        midpoint3D.Z = coordzmid

        Dim endpoint3D As Point3d
        endpoint3D.X = coordx2
        endpoint3D.Y = 0
        endpoint3D.Z = coordz2

        'coordx1 = WCS2Abs(startpoint3D).X
        'coordz1 = WCS2Abs(startpoint3D).Z
        coordx2 = WCS2Abs(endpoint3D).X
        coordz2 = WCS2Abs(endpoint3D).Z
        coordxmid = WCS2Abs(midpoint3D).X
        coordzmid = WCS2Abs(midpoint3D).Z

        Dim longX2 As Decimal = coordx2 - x1exp.Value
        Dim longZ2 As Decimal = coordz2 - z1exp.Value
        Dim longXmid As Decimal = coordxmid - x1exp.Value
        Dim longZmid As Decimal = coordzmid - z1exp.Value

        Dim nullNXOpen_Features_AssociativeArc As NXOpen.Features.AssociativeArc = Nothing
        Dim associativeArcBuilder1 As NXOpen.Features.AssociativeArcBuilder
        associativeArcBuilder1 = workPart.BaseFeatures.CreateAssociativeArcBuilder(nullNXOpen_Features_AssociativeArc)
        associativeArcBuilder1.Limits.StartLimit.LimitOption = NXOpen.GeometricUtilities.CurveExtendData.LimitOptions.AtPoint
        associativeArcBuilder1.Limits.EndLimit.LimitOption = NXOpen.GeometricUtilities.CurveExtendData.LimitOptions.AtPoint

        'Dim point1 As Point = MakePoint(coordx1, coordz1)
        Dim unit1 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
        'start point X coordinate
        Dim expressionX1 As Expression
        expressionX1 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)
        workPart.Expressions.EditWithUnits(expressionX1, unit1, x1exp.Name)
        Dim scalar1 As Scalar = workPart.Scalars.CreateScalarExpression(expressionX1, Scalar.DimensionalityType.Length, SmartObject.UpdateOption.WithinModeling)
        'start point Y coordinate
        Dim expressionY1 As Expression
        expressionY1 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)
        Dim scalar2 As Scalar
        scalar2 = workPart.Scalars.CreateScalarExpression(expressionY1, Scalar.DimensionalityType.Length, SmartObject.UpdateOption.WithinModeling)
        'start point Z coordinate
        Dim expressionZ1 As Expression
        expressionZ1 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)
        workPart.Expressions.EditWithUnits(expressionZ1, unit1, z1exp.Name)
        Dim scalar3 As Scalar = workPart.Scalars.CreateScalarExpression(expressionZ1, Scalar.DimensionalityType.Length, SmartObject.UpdateOption.WithinModeling)
        Dim point1 As Point
        point1 = workPart.Points.CreatePoint(scalar1, scalar2, scalar3, SmartObject.UpdateOption.WithinModeling)

        associativeArcBuilder1.StartPoint.Value = point1
        associativeArcBuilder1.StartPointOptions = NXOpen.Features.AssociativeArcBuilder.StartOption.Point


        'mid point X coordinate
        Dim expressionXmid As Expression
        expressionXmid = workPart.Expressions.CreateSystemExpressionWithUnits(x1exp.Name & " + " & Strings.Replace(longXmid.ToString, ",", "."), unit1)
        Dim scalar7 As Scalar
        scalar7 = workPart.Scalars.CreateScalarExpression(expressionXmid, Scalar.DimensionalityType.Length, SmartObject.UpdateOption.WithinModeling)
        'mid point Y coordinate
        Dim expressionYmid As Expression
        expressionYmid = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)
        Dim scalar8 As Scalar
        scalar8 = workPart.Scalars.CreateScalarExpression(expressionYmid, Scalar.DimensionalityType.Length, SmartObject.UpdateOption.WithinModeling)
        'mid point Z coordinate
        Dim expressionZmid As Expression
        expressionZmid = workPart.Expressions.CreateSystemExpressionWithUnits(z1exp.Name & " + " & Strings.Replace(longZmid.ToString, ",", "."), unit1)
        Dim scalar9 As Scalar
        scalar9 = workPart.Scalars.CreateScalarExpression(expressionZmid, Scalar.DimensionalityType.Length, SmartObject.UpdateOption.WithinModeling)
        Dim pointmid As Point
        pointmid = workPart.Points.CreatePoint(scalar7, scalar8, scalar9, SmartObject.UpdateOption.WithinModeling)
        associativeArcBuilder1.MidPoint.Value = pointmid
        associativeArcBuilder1.MidPointOptions = NXOpen.Features.AssociativeArcBuilder.MidOption.Point

        'end point X coordinate
        Dim expressionX2 As Expression
        expressionX2 = workPart.Expressions.CreateSystemExpressionWithUnits(x1exp.Name & " + " & Strings.Replace(longX2.ToString, ",", "."), unit1)
        Dim scalar4 As Scalar
        scalar4 = workPart.Scalars.CreateScalarExpression(expressionX2, Scalar.DimensionalityType.Length, SmartObject.UpdateOption.WithinModeling)
        'end point Y coordinate
        Dim expressionY2 As Expression
        expressionY2 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)
        Dim scalar5 As Scalar
        scalar5 = workPart.Scalars.CreateScalarExpression(expressionY2, Scalar.DimensionalityType.Length, SmartObject.UpdateOption.WithinModeling)
        'end point Z coordinate
        Dim expressionZ2 As Expression
        expressionZ2 = workPart.Expressions.CreateSystemExpressionWithUnits(z1exp.Name & " + " & Strings.Replace(longZ2.ToString, ",", "."), unit1)
        Dim scalar6 As Scalar
        scalar6 = workPart.Scalars.CreateScalarExpression(expressionZ2, Scalar.DimensionalityType.Length, SmartObject.UpdateOption.WithinModeling)
        Dim point2 As Point
        point2 = workPart.Points.CreatePoint(scalar4, scalar5, scalar6, SmartObject.UpdateOption.WithinModeling)

        associativeArcBuilder1.EndPoint.Value = point2
        associativeArcBuilder1.EndPointOptions = NXOpen.Features.AssociativeArcBuilder.EndOption.Point

        Dim myArcFeature As Features.AssociativeArc
        myArcFeature = associativeArcBuilder1.Commit
        'ufs.Tag.AskHandleFromTag(myArcFeature.Tag)

        associativeArcBuilder1.Destroy()
        Dim myEntities() As NXObject
        myEntities = myArcFeature.GetEntities()
        myArcFeature.SetName(LineName)

        Dim myArcTuple = (myArcFeature, expressionX2, expressionZ2)
        Return myArcTuple



    End Function

    Function MakeCSYS(Line1, CSYSname, teta1)


        Dim nullNXOpen_Features_Feature As NXOpen.Features.Feature = Nothing
        Dim nXObject1 As NXOpen.NXObject = Line1
        Dim datumCsysBuilderStartCSYS As NXOpen.Features.DatumCsysBuilder = workPart.Features.CreateDatumCsysBuilder(nullNXOpen_Features_Feature)
        Dim datumCsysBuilderEndCSYS As NXOpen.Features.DatumCsysBuilder = workPart.Features.CreateDatumCsysBuilder(nullNXOpen_Features_Feature)
        Dim datumCsysBuilderMainCSYS As NXOpen.Features.DatumCsysBuilder = workPart.Features.CreateDatumCsysBuilder(nullNXOpen_Features_Feature)
        Dim unit1 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
        Dim scalar1 As NXOpen.Scalar = workPart.Scalars.CreateScalar(0.0, NXOpen.Scalar.DimensionalityType.None, NXOpen.SmartObject.UpdateOption.WithinModeling)
        Dim scalar2 As NXOpen.Scalar = workPart.Scalars.CreateScalar(1.0, NXOpen.Scalar.DimensionalityType.None, NXOpen.SmartObject.UpdateOption.WithinModeling)
        Dim scalar3 As NXOpen.Scalar = workPart.Scalars.CreateScalar(0.5, NXOpen.Scalar.DimensionalityType.None, NXOpen.SmartObject.UpdateOption.WithinModeling)
        Dim point1 As NXOpen.Point = workPart.Points.CreatePoint(Line1, scalar1, NXOpen.SmartObject.UpdateOption.WithinModeling)
        Dim point2 As NXOpen.Point = workPart.Points.CreatePoint(Line1, scalar2, NXOpen.SmartObject.UpdateOption.WithinModeling)
        Dim point3 As NXOpen.Point = workPart.Points.CreatePoint(Line1, scalar3, NXOpen.SmartObject.UpdateOption.WithinModeling)

        Select Case Line1.GetType.Name
            Case "Line"
                Line1 = CType(nXObject1, NXOpen.Line)
            Case "Arc"
                Line1 = CType(nXObject1, NXOpen.Arc)
        End Select

        Dim ypoint As Point3d
        ypoint.X = 0
        ypoint.Y = 0
        ypoint.Z = 0

        Dim yvector As New NXOpen.Vector3d(0, 1, 0)
        Dim DirY As NXOpen.Direction = workPart.Directions.CreateDirection(ypoint, yvector, NXOpen.SmartObject.UpdateOption.WithinModeling)
        'Dim YAxis As NXOpen.DatumAxis

        'Dim YAxis As NXOpen.DatumAxis = CType(workPart.Datums.FindObject("DATUM_CSYS(0) Y axis"), NXOpen.DatumAxis)
        'Dim DirY As NXOpen.Direction = workPart.Directions.CreateDirection(YAxis, NXOpen.Sense.Forward, NXOpen.SmartObject.UpdateOption.WithinModeling)
        Dim dirZStart As NXOpen.Direction = Nothing
        Dim dirZEnd As NXOpen.Direction = Nothing
        Dim dirZMain As NXOpen.Direction = Nothing

        ' Create Direction Z of StartCSYS
        Select Case Line1.GetType.Name
            Case "Line"
                dirZStart = workPart.Directions.CreateDirectionOnPointParentCurve(point1, Line1, Direction.OnCurveOption.Tangent, NXOpen.Sense.Forward, NXOpen.SmartObject.UpdateOption.WithinModeling)
            Case "Arc"
                If teta1 > 0 Then
                    dirZStart = workPart.Directions.CreateDirectionOnPointParentCurve(point1, Line1, Direction.OnCurveOption.Tangent, NXOpen.Sense.Forward, NXOpen.SmartObject.UpdateOption.WithinModeling)
                Else
                    dirZStart = workPart.Directions.CreateDirectionOnPointParentCurve(point1, Line1, Direction.OnCurveOption.Tangent, NXOpen.Sense.Forward, NXOpen.SmartObject.UpdateOption.WithinModeling)
                End If

        End Select


        ' Create Direction Z of EndCSYS
        Select Case Line1.GetType.Name
            Case "Line"
                dirZEnd = workPart.Directions.CreateDirectionOnPointParentCurve(point2, Line1, Direction.OnCurveOption.Tangent, NXOpen.Sense.Forward, NXOpen.SmartObject.UpdateOption.WithinModeling)
            Case "Arc"
                If teta1 > 0 Then
                    dirZEnd = workPart.Directions.CreateDirectionOnPointParentCurve(point2, Line1, Direction.OnCurveOption.Tangent, NXOpen.Sense.Forward, NXOpen.SmartObject.UpdateOption.WithinModeling)
                Else
                    dirZEnd = workPart.Directions.CreateDirectionOnPointParentCurve(point2, Line1, Direction.OnCurveOption.Tangent, NXOpen.Sense.Forward, NXOpen.SmartObject.UpdateOption.WithinModeling)
                End If

        End Select

        ' Create Direction Z of MainCSYS
        Select Case Line1.GetType.Name
            Case "Line"
                dirZMain = workPart.Directions.CreateDirectionOnPointParentCurve(point2, Line1, Direction.OnCurveOption.Tangent, NXOpen.Sense.Forward, NXOpen.SmartObject.UpdateOption.WithinModeling)
            Case "Arc"
                If teta1 > 0 Then
                    dirZMain = workPart.Directions.CreateDirectionOnPointParentCurve(point3, Line1, Direction.OnCurveOption.Tangent, NXOpen.Sense.Forward, NXOpen.SmartObject.UpdateOption.WithinModeling)
                Else
                    dirZMain = workPart.Directions.CreateDirectionOnPointParentCurve(point3, Line1, Direction.OnCurveOption.Tangent, NXOpen.Sense.Forward, NXOpen.SmartObject.UpdateOption.WithinModeling)
                End If

        End Select

        'STARTCSYS definition
        Dim xformStart As NXOpen.Xform = Nothing
        Select Case Line1.GetType.Name
            Case "Line"
                xformStart = workPart.Xforms.CreateXformByPointYDirZDir(point1, DirY, dirZStart, NXOpen.SmartObject.UpdateOption.WithinModeling, 1.0)
            Case "Arc"
                If teta1 > 0 Then
                    xformStart = workPart.Xforms.CreateXformByPointYDirZDir(point1, DirY, dirZStart, NXOpen.SmartObject.UpdateOption.WithinModeling, 1.0)
                Else
                    xformStart = workPart.Xforms.CreateXformByPointYDirZDir(point1, DirY, dirZStart, NXOpen.SmartObject.UpdateOption.WithinModeling, 1.0)
                End If

        End Select

        'ENDCSYS definition
        Dim xformEnd As NXOpen.Xform = Nothing
        Select Case Line1.GetType.Name
            Case "Line"
                xformEnd = workPart.Xforms.CreateXformByPointYDirZDir(point2, DirY, dirZEnd, NXOpen.SmartObject.UpdateOption.WithinModeling, 1.0)
            Case "Arc"
                If teta1 > 0 Then
                    xformEnd = workPart.Xforms.CreateXformByPointYDirZDir(point2, DirY, dirZEnd, NXOpen.SmartObject.UpdateOption.WithinModeling, 1.0)
                Else
                    xformEnd = workPart.Xforms.CreateXformByPointYDirZDir(point2, DirY, dirZEnd, NXOpen.SmartObject.UpdateOption.WithinModeling, 1.0)
                End If

        End Select



        'MAINCSYS definition
        Dim xformMain As NXOpen.Xform = workPart.Xforms.CreateXformByPointYDirZDir(point3, DirY, dirZMain, NXOpen.SmartObject.UpdateOption.WithinModeling, 1.0)



        'CSYSs CREATION
        Dim StartCSYS As NXOpen.CartesianCoordinateSystem = workPart.CoordinateSystems.CreateCoordinateSystem(xformStart, NXOpen.SmartObject.UpdateOption.WithinModeling)
        Dim EndCSYS As NXOpen.CartesianCoordinateSystem = workPart.CoordinateSystems.CreateCoordinateSystem(xformEnd, NXOpen.SmartObject.UpdateOption.WithinModeling)
        Dim MainCSYS As NXOpen.CartesianCoordinateSystem = workPart.CoordinateSystems.CreateCoordinateSystem(xformMain, NXOpen.SmartObject.UpdateOption.WithinModeling)
        datumCsysBuilderStartCSYS.Csys = StartCSYS
        datumCsysBuilderMainCSYS.Csys = MainCSYS
        datumCsysBuilderEndCSYS.Csys = EndCSYS
        datumCsysBuilderStartCSYS.DisplayScaleFactor = 0.5
        datumCsysBuilderMainCSYS.DisplayScaleFactor = 1
        datumCsysBuilderEndCSYS.DisplayScaleFactor = 0.5
        datumCsysBuilderStartCSYS.Csys.Layer = 20
        datumCsysBuilderEndCSYS.Csys.Layer = 20
        Dim CSYSFeatureStart As Features.Feature = datumCsysBuilderStartCSYS.Commit()
        Dim CSYSFeatureMain As Features.Feature = datumCsysBuilderMainCSYS.Commit()
        Dim CSYSFeatureEnd As Features.Feature = datumCsysBuilderEndCSYS.Commit()
        CSYSFeatureStart.SetName(CSYSname & " Entrée")
        CSYSFeatureMain.SetName(CSYSname.ToString.ToUpper)
        CSYSFeatureEnd.SetName(CSYSname & " Sortie")
        FormatMainCSYS(datumCsysBuilderMainCSYS.GetFeature, CSYSname, Line1)
        FormatStartEndCSYS(datumCsysBuilderEndCSYS.GetFeature)
        FormatStartEndCSYS(datumCsysBuilderStartCSYS.GetFeature)
        displayPart.WCS.SetCoordinateSystemCartesianAtCsys(EndCSYS)
        Dim datumCsysStart As NXObject = datumCsysBuilderStartCSYS.GetObject
        Dim datumCsysMain As NXObject = datumCsysBuilderMainCSYS.GetObject
        Dim datumCsysEnd As NXObject = datumCsysBuilderEndCSYS.GetObject
        Call RenameDatumCsys(datumCsysStart, CSYSname & " Entrée")
        Call RenameDatumCsys(datumCsysMain, CSYSname)
        Call RenameDatumCsys(datumCsysEnd, CSYSname & " Sortie")

        datumCsysBuilderStartCSYS.Destroy()
        datumCsysBuilderMainCSYS.Destroy()
        datumCsysBuilderEndCSYS.Destroy()
        Dim CsysTriplet(2) As Features.Feature


        Dim components1(0) As NXOpen.NXObject
        components1(0) = CSYSFeatureMain
        referenceSet2.AddObjectsToReferenceSet(components1)

        CsysTriplet(0) = CSYSFeatureStart
        CsysTriplet(1) = CSYSFeatureMain
        CsysTriplet(2) = CSYSFeatureEnd

        Return CsysTriplet

    End Function
    Sub RenameDatumCsys(datumobj As NXOpen.NXObject, CSYSName As String)

        Dim objects1(0) As NXOpen.NXObject
        Dim datumCsys1 As NXOpen.Features.DatumCsys = datumobj 'CType(workPart.Features.FindObject("DATUM_CSYS(4)"), NXOpen.Features.DatumCsys)
        objects1(0) = datumCsys1
        Dim featureGeneralPropertiesBuilder1 As NXOpen.FeatureGeneralPropertiesBuilder = workPart.PropertiesManager.CreateFeatureGeneralPropertiesBuilder(objects1)
        featureGeneralPropertiesBuilder1.GeneralName = CSYSName
        Dim nXObject1 As NXOpen.NXObject
        nXObject1 = featureGeneralPropertiesBuilder1.Commit()
        featureGeneralPropertiesBuilder1.Destroy()

    End Sub

    Sub FormatStartEndCSYS(myCsys)


        Dim displayModification1 As NXOpen.DisplayModification
        displayModification1 = theSession.DisplayManager.NewDisplayModification()
        displayModification1.ApplyToAllFaces = True
        displayModification1.ApplyToOwningParts = False

        displayModification1.NewWidth = NXOpen.DisplayableObject.ObjectWidth.Nine
        displayModification1.NewLayer = 60
        displayModification1.NewColor = 130 'grey


        Dim dispObjcts As New List(Of DisplayableObject)

        For Each temp As DisplayableObject In myCsys.GetEntities
            dispObjcts.Add(temp)
        Next

        'objects1(7) = cartesianCoord
        displayModification1.Apply(dispObjcts.ToArray)

        displayModification1.Dispose()





    End Sub
    Sub FormatMainCSYS(myCsys, myCsysName, myLine)


        Dim displayModification1 As NXOpen.DisplayModification
        displayModification1 = theSession.DisplayManager.NewDisplayModification()
        displayModification1.ApplyToAllFaces = True
        displayModification1.ApplyToOwningParts = False

        displayModification1.NewWidth = NXOpen.DisplayableObject.ObjectWidth.Nine

        Select Case Strings.Left(myCsysName, 2)
            Case "Dr"
                displayModification1.NewColor = 130 'grey
            Case "QF"
                displayModification1.NewColor = 211 'blue
            Case "QD"
                displayModification1.NewColor = 211 'blue
            Case "SH"
                displayModification1.NewColor = 6 'yellow
            Case "SC"
                displayModification1.NewColor = 6 'yellow
            Case "OH"
                displayModification1.NewColor = 36 'green
            Case "OC"
                displayModification1.NewColor = 36 'green
            Case "DN"
                displayModification1.NewColor = 112  'light faded red
            Case "DI"
                displayModification1.NewColor = 31 'Cyan
            Case Else
                displayModification1.NewColor = 130 'grey

        End Select


        Dim dispObjcts As New List(Of DisplayableObject)

        For Each temp As DisplayableObject In myCsys.GetEntities
            dispObjcts.Add(temp)
        Next
        dispObjcts.Add(myLine)
        'objects1(7) = cartesianCoord
        displayModification1.Apply(dispObjcts.ToArray)

        displayModification1.Dispose()





    End Sub

    Public Function MakeLine(x1exp As Expression, z1exp As Expression, coordx2 As Decimal, coordz2 As Decimal, LineName As String)

        Dim startpoint3D As Point3d
        'startpoint3D.X = coordx1
        startpoint3D.Y = 0
        'startpoint3D.Z = coordz1

        Dim endpoint3D As Point3d
        endpoint3D.X = coordx2
        endpoint3D.Y = 0
        endpoint3D.Z = coordz2

        'coordx1 = WCS2Abs(startpoint3D).X
        'coordz1 = WCS2Abs(startpoint3D).Z
        coordx2 = WCS2Abs(endpoint3D).X
        coordz2 = WCS2Abs(endpoint3D).Z
        Dim longX As Decimal = coordx2 - x1exp.Value
        Dim longZ As Decimal = coordz2 - z1exp.Value

        Dim nullFeatures_AssociativeLine As Features.AssociativeLine = Nothing

        Dim associativeLineBuilder1 As Features.AssociativeLineBuilder
        associativeLineBuilder1 = workPart.BaseFeatures.CreateAssociativeLineBuilder(nullFeatures_AssociativeLine)

        Dim unit1 As Unit
        unit1 = associativeLineBuilder1.Limits.StartLimit.Distance.Units

        associativeLineBuilder1.StartPointOptions = Features.AssociativeLineBuilder.StartOption.Point
        'start point X coordinate
        Dim expressionX1 As Expression
        expressionX1 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)
        workPart.Expressions.EditWithUnits(expressionX1, unit1, x1exp.Name)
        Dim scalar1 As Scalar = workPart.Scalars.CreateScalarExpression(expressionX1, Scalar.DimensionalityType.Length, SmartObject.UpdateOption.WithinModeling)

        'start point Y coordinate
        Dim expressionY1 As Expression
        expressionY1 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)

        Dim scalar2 As Scalar
        scalar2 = workPart.Scalars.CreateScalarExpression(expressionY1, Scalar.DimensionalityType.Length, SmartObject.UpdateOption.WithinModeling)

        'start point Z coordinate
        Dim expressionZ1 As Expression
        expressionZ1 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)
        workPart.Expressions.EditWithUnits(expressionZ1, unit1, z1exp.Name)
        Dim scalar3 As Scalar = workPart.Scalars.CreateScalarExpression(expressionZ1, Scalar.DimensionalityType.Length, SmartObject.UpdateOption.WithinModeling)

        Dim point1 As Point
        point1 = workPart.Points.CreatePoint(scalar1, scalar2, scalar3, SmartObject.UpdateOption.WithinModeling)

        associativeLineBuilder1.StartPoint.Value = point1

        associativeLineBuilder1.EndPointOptions = Features.AssociativeLineBuilder.EndOption.Point

        'end point X coordinate
        Dim expressionX2 As Expression

        expressionX2 = workPart.Expressions.CreateSystemExpressionWithUnits(x1exp.Name & " + " & Strings.Replace(longX.ToString, ",", "."), unit1)

        Dim scalar4 As Scalar
        scalar4 = workPart.Scalars.CreateScalarExpression(expressionX2, Scalar.DimensionalityType.Length, SmartObject.UpdateOption.WithinModeling)

        'end point Y coordinate
        Dim expressionY2 As Expression
        expressionY2 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1)

        Dim scalar5 As Scalar
        scalar5 = workPart.Scalars.CreateScalarExpression(expressionY2, Scalar.DimensionalityType.Length, SmartObject.UpdateOption.WithinModeling)

        'end point Z coordinate
        Dim expressionZ2 As Expression
        expressionZ2 = workPart.Expressions.CreateSystemExpressionWithUnits(z1exp.Name & " + " & Strings.Replace(longZ.ToString, ",", "."), unit1)

        Dim scalar6 As Scalar
        scalar6 = workPart.Scalars.CreateScalarExpression(expressionZ2, Scalar.DimensionalityType.Length, SmartObject.UpdateOption.WithinModeling)

        Dim point2 As Point
        point2 = workPart.Points.CreatePoint(scalar4, scalar5, scalar6, SmartObject.UpdateOption.WithinModeling)

        associativeLineBuilder1.EndPoint.Value = point2

        'associativeLineBuilder1.Limits.EndLimit.Distance.RightHandSide =
        'Dim expLim As NXOpen.Expression = associativeLineBuilder1.Limits.EndLimit.Distance

        Dim myLineFeature As Features.AssociativeLine
        myLineFeature = associativeLineBuilder1.Commit

        associativeLineBuilder1.Destroy()
        myLineFeature.SetName(LineName)
        'myLineFeature.
        'ufs.tag(myLineFeature.Tag)


        Dim myLineTuple = (myLineFeature, expressionX2, expressionZ2)
        Return myLineTuple


    End Function


    Public Function GetUnloadOption() As Integer

        'Unloads the image when the NX session terminates
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately

    End Function

    Function WCS2Abs(ByVal inPt As Point3d) As Point3d
        Dim pt1(2), pt2(2) As Double

        pt1(0) = inPt.X
        pt1(1) = inPt.Y
        pt1(2) = inPt.Z
        ' UFConstants.UF_CSYS_ROOT_WCS_COORDS(0)

        ufs.Csys.MapPoint(UFConstants.UF_CSYS_ROOT_WCS_COORDS, pt1, UFConstants.UF_CSYS_ROOT_COORDS, pt2)

        WCS2Abs.X = pt2(0)
        WCS2Abs.Y = pt2(1)
        WCS2Abs.Z = pt2(2)

    End Function

    Sub Create_Assembly(outputdir)
        Dim fileNew1 As NXOpen.FileNew = Nothing

        Dim partCount As Integer = theSession.Parts.ToArray.Length
        fileNew1 = theSession.Parts.FileNew()
        With fileNew1
            .TemplateFileName = "assembly-mm-template.prt"
            .UseBlankTemplate = False
            .ApplicationName = "AssemblyTemplate"
            .Units = NXOpen.Part.Units.Millimeters
            .RelationType = ""
            .UsesMasterModel = "No"
            .TemplateType = NXOpen.FileNewTemplateType.Item
            .TemplatePresentationName = "Assemblage"
            .ItemType = ""
            .Specialization = ""
            .SetCanCreateAltrep(False)
            .NewFileName = outputdir & "\" & "SQUELETTE MACHINE.prt"
            .MasterFileName = ""
            .MakeDisplayedPart = True
            .DisplayPartOption = NXOpen.DisplayPartOption.AllowAdditional
            .Commit()
            .Destroy()
        End With

        Dim ListOfParts() As NXOpen.BasePart = theSession.Parts.ToArray()
        Array.Reverse(ListOfParts)

        Dim part1 As Part = theSession.Parts.Work
        theSession.Parts.SetWork(part1)

        Dim addComponentBuilder1 As NXOpen.Assemblies.AddComponentBuilder = Nothing
        addComponentBuilder1 = part1.AssemblyManager.CreateAddComponentBuilder()
        addComponentBuilder1.ReferenceSet = "MODEL"
        ReDim Preserve ListOfParts(ListOfParts.Length - 1)
        addComponentBuilder1.SetPartsToAdd(ListOfParts)
        addComponentBuilder1.Commit()
        addComponentBuilder1.Destroy()

        Dim isFirst = True
        For i = 0 To ListOfParts.Length - 3
            Dim comp1 As BasePart = ListOfParts(i)
            Dim comp2 As BasePart = ListOfParts(i + 1)


            SetConstraints(comp1, comp2, part1, isFirst)
            isFirst = False
        Next
    End Sub
    Sub SetConstraints(comp1, comp2, part1, isFirst)

        ' ----------------------------------------------
        '    Menu: Assemblages->Position des composants->Contraintes d'assemblage...


        Dim componentPositioner1 As NXOpen.Positioning.ComponentPositioner
        componentPositioner1 = part1.ComponentAssembly.Positioner
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
        Dim component1 As NXOpen.Assemblies.Component = CType(part1.ComponentAssembly.RootComponent.FindObject("COMPONENT " & comp1.Name & " 1"), NXOpen.Assemblies.Component)

        If isFirst = True Then
            Dim constraint2 As NXOpen.Positioning.Constraint
            constraint2 = componentPositioner1.CreateConstraint(True)
            Dim componentConstraint2 As NXOpen.Positioning.ComponentConstraint = CType(constraint2, NXOpen.Positioning.ComponentConstraint)
            componentConstraint2.ConstraintType = NXOpen.Positioning.Constraint.Type.Fix
            Dim constraintReference3 As NXOpen.Positioning.ConstraintReference = componentConstraint2.CreateConstraintReference(component1, component1, False, False, False)
            constraintReference3.SetFixHint(True)
            componentNetwork1.Solve()
            componentNetwork1.Solve()
            componentPositioner1.ClearNetwork()
            componentPositioner1.DeleteNonPersistentConstraints()
            componentPositioner1.EndAssemblyConstraints()
            Dim displayedConstraint2 As NXOpen.Positioning.DisplayedConstraint = constraint2.GetDisplayedConstraint
            displayedConstraint2.Blank()

        End If

        Dim compFeatArray() As NXOpen.Features.Feature
        compFeatArray = component1.Prototype.OwningPart.BaseFeatures.ToArray
        Dim component2 As NXOpen.Assemblies.Component = CType(part1.ComponentAssembly.RootComponent.FindObject("COMPONENT " & comp2.Name & " 1"), NXOpen.Assemblies.Component)
        Dim cartesianCoordinateSystem1 As NXOpen.CartesianCoordinateSystem = CType(component1.FindObject("PROTO#.Features|DATUM_CSYS(" & (compFeatArray.Length - 1).ToString & ")|CSYSTEM 1"), NXOpen.CartesianCoordinateSystem)
        Dim constraintReference1 As NXOpen.Positioning.ConstraintReference = componentConstraint1.CreateConstraintReference(component1, cartesianCoordinateSystem1, False, False)
        Dim cartesianCoordinateSystem2 As NXOpen.CartesianCoordinateSystem = CType(component2.FindObject("PROTO#.Features|DATUM_CSYS(2)|CSYSTEM 1"), NXOpen.CartesianCoordinateSystem)
        Dim constraintReference2 As NXOpen.Positioning.ConstraintReference = componentConstraint1.CreateConstraintReference(component2, cartesianCoordinateSystem2, False, False)
        constraintReference2.SetFixHint(True)
        componentNetwork1.Solve()
        componentNetwork1.Solve()
        componentPositioner1.ClearNetwork()
        componentPositioner1.DeleteNonPersistentConstraints()
        componentPositioner1.EndAssemblyConstraints()
        Dim displayedConstraint1 As NXOpen.Positioning.DisplayedConstraint = constraint1.GetDisplayedConstraint
        displayedConstraint1.Blank()


    End Sub
End Module

