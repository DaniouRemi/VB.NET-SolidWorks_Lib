Imports System.Math
Imports SwConst

Public Class Functions
    Inherits Variables

    Enum Box_e
        x = 0
        y = 1
        z = 2
    End Enum
    Protected vPos As New List(Of Object)

    'Protected Friend _ScaleUsed As ScaleMode
    'Protected Friend Function TemplateAuto(ByVal _RotateViewActivate As Boolean, Optional ByVal ScaleValue As Double = 0.5) As Templates_p
    '    Dim Template_p As New Templates_p
    '    Dim _Scale As New List(Of Double)
    '    Dim _TemplateUsed As Templates_p

    '    'calcul de la boite englobante et définition de la valeur la plus grande
    '    DefineBox = DefineBox_fc()

    '    TemplateOrientation = Orientation_e.Landscape

    '    'calcul de la rotation de la première view et redéfinie la boite englobante et définition de la valeur la plus grande
    '    If _RotateViewActivate Then
    '        View_Rotation = SetRotationView()
    '    End If

    '    For Each Template As Templates_p In Templates_p
    '        'calcul de toutes les échelles suivant les formats de plan (A4, A3, A2)
    '        If Template.Name <> "DXF" Then
    '            If DefineBox > Template.ScaleMode.Item(DrawingCreationType).ViewZone Then
    '                _Scale.Add(DefineBox / Template.ScaleMode.Item(DrawingCreationType).ViewZone)
    '            Else
    '                _Scale.Add(Template.ScaleMode.Item(DrawingCreationType).ViewZone / DefineBox)
    '            End If
    '        End If

    '    Next

    '    'décision du format utilisé (au plus près de 0 car -1 à l'étape d'avant)
    '    DecScale = _Scale.Item(0)
    '    _TemplateUsed = Templates_p.Item(1)

    '    For i As Integer = 1 To _Scale.Count - 1
    '        Debug.Print(Sqrt(Pow(_Scale.Item(i) - ScaleValue, 2)))
    '        Debug.Print(Sqrt(Pow(DecScale - ScaleValue, 2)))
    '        If Sqrt(Pow(_Scale.Item(i) - ScaleValue, 2)) < Sqrt(Pow(DecScale - ScaleValue, 2)) Then
    '            DecScale = _Scale.Item(i)
    '            _TemplateUsed = Templates_p.Item(i)
    '            ScaleUsed = _TemplateUsed.ScaleMode(DrawingCreationType)
    '        End If
    '    Next

    '    Return _TemplateUsed
    'End Function

    Protected Friend Function TemplateAuto() As Templates_p
        Dim Template_p As New Templates_p
        Dim _Scale As New List(Of Double)
        Dim _TemplateUsed As Templates_p
        Dim ComponentHeight As Double
        Dim ComponentWidth As Double

        Dim A4Width As Double = 0.21
        Dim A4Height As Double = 0.297
        Dim A3Width As Double = 0.297
        Dim A3Height As Double = 0.42
        Dim A2Width As Double = 0.42
        Dim A2Height As Double = 0.594

        DefineBox = DefineBox_fc()
        ComponentHeight = (VarScale2 + VarScale3) * 1.5
        ComponentWidth = (VarScale1 + VarScale3) * 1.5

        Dim A4WidthScale As Double = A4Width / ComponentWidth
        Dim A4HeightScale As Double = A4Height / ComponentHeight
        Dim A3WidthScale As Double = A3Width / ComponentWidth
        Dim A3HeightScale As Double = A3Height / ComponentHeight
        Dim A2WidthScale As Double = A2Width / ComponentWidth
        Dim A2HeightScale As Double = A2Height / ComponentHeight

        If A4WidthScale < 0.75 And A4HeightScale < 0.75 Then
            If A3WidthScale < 1.5 And A3HeightScale < 1.5 Then
                _TemplateUsed = Templates_p.Item(1)
            Else
                _TemplateUsed = Templates_p.Item(2)
            End If
        Else
            _TemplateUsed = Templates_p.Item(3)
        End If

        Return _TemplateUsed
    End Function

    Protected Friend Function ScaleAuto(ByVal _RotateViewActivate As Boolean, ByVal ScaleValue As String) As Double

        Return 0

    End Function

    Public Function Dec2Frac_fc(ByVal f As Double) As String

        Dim df As Double
        Dim lUpperPart As Long
        Dim lLowerPart As Long

        lUpperPart = 1
        lLowerPart = 1

        df = lUpperPart / lLowerPart
        While (df <> f)
            If (df < f) Then
                lUpperPart = lUpperPart + 1
            Else
                lLowerPart = lLowerPart + 1
                lUpperPart = f * lLowerPart
            End If
            df = lUpperPart / lLowerPart
        End While
        Dec2Frac_fc = CStr(lUpperPart) & "/" & CStr(lLowerPart)
    End Function

    ''' <summary>
    ''' Définie la boîte englobante et retourne la longueur la plus grande
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DefineBox_fc() As Double

        Dim swAssy As SldWorks.AssemblyDoc
        Dim swPart As SldWorks.PartDoc
        Dim vBox As Object

        If LCase(Right(swModel.GetPathName, 6)) = "sldasm" Then
            swAssy = swModel
            vBox = swAssy.GetBox(True)
        ElseIf LCase(Right(swModel.GetPathName, 6)) = "sldprt" Then
            swPart = swModel
            vBox = swPart.GetPartBox(True)
        Else
            vBox = Nothing
            Return 0
        End If

        Box(Box_e.x) = vBox(3) - vBox(0)
        Box(Box_e.y) = vBox(4) - vBox(1)
        Box(Box_e.z) = vBox(5) - vBox(2)
        Debug.Print("Dimensions de la boîte englobante : ")
        Debug.Print("x = " & Box(Box_e.x))
        Debug.Print("y = " & Box(Box_e.y))
        Debug.Print("z = " & Box(Box_e.z))


        DefineBox = Box(Box_e.x)
        If Box(Box_e.y) > DefineBox Then
            DefineBox = Box(Box_e.y)
        End If

        If Box(Box_e.z) > DefineBox Then
            DefineBox = Box(Box_e.z)
        End If

        If Box(Box_e.x) >= Box(Box_e.y) And Box(Box_e.y) >= Box(Box_e.z) Then
            orderBox = "xyz"
            VarScale1 = Box(Box_e.x)
            VarScale2 = Box(Box_e.y)
            VarScale3 = Box(Box_e.z)
        ElseIf Box(Box_e.y) >= Box(Box_e.x) And Box(Box_e.x) >= Box(Box_e.z) Then
            orderBox = "yxz"
            VarScale1 = Box(Box_e.y)
            VarScale2 = Box(Box_e.x)
            VarScale3 = Box(Box_e.z)
        ElseIf Box(Box_e.z) >= Box(Box_e.x) And Box(Box_e.x) >= Box(Box_e.y) Then
            orderBox = "zxy"
            VarScale1 = Box(Box_e.z)
            VarScale2 = Box(Box_e.x)
            VarScale3 = Box(Box_e.y)
        ElseIf Box(Box_e.x) >= Box(Box_e.z) And Box(Box_e.z) >= Box(Box_e.y) Then
            orderBox = "xzy"
            VarScale1 = Box(Box_e.x)
            VarScale2 = Box(Box_e.z)
            VarScale3 = Box(Box_e.y)
        ElseIf Box(Box_e.y) >= Box(Box_e.z) And Box(Box_e.z) >= Box(Box_e.x) Then
            orderBox = "yzx"
            VarScale1 = Box(Box_e.y)
            VarScale2 = Box(Box_e.z)
            VarScale3 = Box(Box_e.x)
        ElseIf Box(Box_e.z) >= Box(Box_e.y) And Box(Box_e.y) >= Box(Box_e.x) Then
            orderBox = "zyx"
            VarScale1 = Box(Box_e.z)
            VarScale2 = Box(Box_e.y)
            VarScale3 = Box(Box_e.x)
        End If

        Debug.Print("DefineBox = " & DefineBox)
        Return DefineBox

    End Function

    Public Function SetRotationView() As Double
        Dim _View_Rotation As Double

        If Left(orderBox, 2) = "zy" Or Left(orderBox, 1) = "x" Then
            _View_Rotation = 0
        Else
            _View_Rotation = 90
        End If

        Return _View_Rotation
    End Function

    Public Function SetOrientation() As String

        Select Case Left(orderBox, 2)
            Case "xy", "yx"
                Return "*Face"
            Case "xz", "zx"
                Return "*Dessus"
            Case Else
                Return "*Gauche"
        End Select

    End Function

    Public Sub SetScale()
        Dim Height As Double
        Dim Width As Double
        Dim SizeSheetWidth As Object
        Dim SizeSheetHeight As Object
        Dim ScaleWidth As Double
        Dim ScaleHeight As Double

        Height = (VarScale2 + VarScale3) * 1.6
        Width = (VarScale1 + VarScale3) * 1.6

        swSheet = swDraw.GetCurrentSheet
        swSheet.GetSize(SizeSheetWidth, SizeSheetHeight)
        Debug.Print(SizeSheetWidth)
        Debug.Print(SizeSheetHeight)

        'SizeSheetWidth = Sqrt(Pow(SizeSheetWidth, 2))
        'SizeSheetHeight = Sqrt(Pow(SizeSheetHeight, 2))

        ScaleWidth = Sqrt(Pow(SizeSheetWidth / Width, 2))
        ScaleHeight = Sqrt(Pow(SizeSheetHeight / Height, 2))


        Dim vSheetProps As Object
        vSheetProps = swSheet.GetProperties

        If ScaleWidth < ScaleHeight Then
            ScaleUpper = Split(Dec2Frac_fc(Round(ScaleWidth, 2)), "/")(0)
            ScaleLower = Split(Dec2Frac_fc(Round(ScaleWidth, 2)), "/")(1)
        Else
            ScaleUpper = Split(Dec2Frac_fc(Round(ScaleHeight, 2)), "/")(0)
            ScaleLower = Split(Dec2Frac_fc(Round(ScaleHeight, 2)), "/")(1)
        End If

        swSheet.SetProperties( _
                                    vSheetProps(0), _
                                    vSheetProps(1), _
                                    ScaleUpper, _
                                    ScaleLower, _
                                    vSheetProps(4), _
                                    vSheetProps(5), _
                                    vSheetProps(6))

    End Sub

    Public Function CreateViews(ByVal DrawingCreationType As Integer) As Boolean

        Select Case DrawingCreationType
            Case DrawingCreationType_e.NoView

            Case DrawingCreationType_e.ThreeViews
                CreateThreeViews()

            Case DrawingCreationType_e.ThreeViewsAndIso
                CreateThreeViews()
                CreateView(IsometricView, True, SetPositionThreeViews(0), SetPositionThreeViews(1))

            Case DrawingCreationType_e.ThreeViewsAndIsoAndFlatPattern
                'swDraw.Create1stAngleViews2(swModel.GetPathName)
                CreateThreeViews()
                CreateView(IsometricView, True, SetPositionThreeViews(0), SetPositionThreeViews(1))
                CreateFlatPatternView(0, 0)

            Case DrawingCreationType_e.FlatPatternView
                CreateFlatPatternView(0, 0)
                'Case DrawingCreationType_e.CustomView
                '    CreateView(CustomViewName, False, SetPositionThreeViews(0), SetPositionThreeViews(1))

            Case Else
                Return False
        End Select

        Return True
    End Function

    Public Sub CreateView(ByVal _ViewName As String, ByVal Colored As Boolean, Optional ByVal _IsopositionX As Double = 0, Optional ByVal _IsopositionY As Double = 0)
        Dim swView As SldWorks.View
        Try
            swView = swDraw.CreateDrawViewFromModelView3(swModel.GetPathName, _ViewName, _IsopositionX, _IsopositionY, 0)

            If Colored Then
                swView.SetDisplayMode3(False, 3, True, True)
            End If
        Catch ex As Exception
            'swDraw.InsertNewNote2("COOL", "cool", False, False, swArrowStyle_e.swNO_ARROWHEAD, swLeaderStyle_e.swNO_LEADER, 0, swBalloonStyle_e.swBS_None, _
            '                      swBalloonFit_e.swBF_1Char, 1, 1)
        End Try

    End Sub 'création de vue

    Public Sub CreateThreeViews()
        Dim swView As SldWorks.View
        Dim ViewName As String = ""
        Dim t_swView As SldWorks.View
        Dim swViewSupress As SldWorks.View

        Dim swModelDrawing As SldWorks.ModelDoc2 = swApp.ActiveDoc
        Dim Orientation As String = SetOrientation()

        swDraw.Create1stAngleViews2(swModel.GetPathName)

        If Not Orientation = "*Face" Then

            'récupération des positions et suppression des vues
            swView = swDraw.GetFirstView
            swView = swView.GetNextView
            For i As Integer = 0 To 2
                vPos.Add(swView.Position)
                Debug.Print(swView.Name & " => x:" & swView.Position(0) & " y:" & swView.Position(1))
                swViewSupress = swView
                swView = swView.GetNextView
                swModelDrawing.Extension.SelectByID2(swViewSupress.Name, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
                swModelDrawing.EditDelete()
            Next

            'Créer la première vue
            swDraw.CreateDrawViewFromModelView3(swModel.GetPathName, Orientation, vPos.Item(0)(0), vPos.Item(0)(1), 0)

            'récupère le nom de la dernière vue insérée
            t_swView = swDraw.GetFirstView
            t_swView = t_swView.GetNextView
            While Not IsNothing(t_swView)
                ViewName = t_swView.Name
                swView = t_swView
                t_swView = t_swView.GetNextView
            End While

            'créer les 3 vues projetées
            swDraw.CreateUnfoldedViewAt3(vPos.Item(1)(0), vPos.Item(1)(1), 0, False)
            swDraw.ActivateView(ViewName)
            swModelDrawing.Extension.SelectByID2(ViewName, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
            swDraw.CreateUnfoldedViewAt3(vPos.Item(2)(0), vPos.Item(2)(1), 0, False)

            'lie les vues aux parents
            swView = swDraw.GetFirstView
            While Not IsNothing(swView)
                swView.LinkParentConfiguration = True
                swView = swView.GetNextView
            End While

        End If

        'rotation
        Try
            swView = swDraw.GetFirstView
            swView = swView.GetNextView
            swView.Angle = PI * SetRotationView() / 180 'en rad
        Catch ex As Exception

        End Try

        SetScale()

    End Sub

    Public Function CreateFlatPatternView(Optional ByVal _IsopositionX As Double = 0, Optional ByVal _IsopositionY As Double = 0, Optional ByVal _ConfigurationName As String = Nothing)
        Dim Width, Height As Double

        If _ConfigurationName = Nothing Then
            Dim swConfMgr As SldWorks.ConfigurationManager
            Dim swConf As SldWorks.Configuration
            swConfMgr = swModel.ConfigurationManager
            swConf = swConfMgr.ActiveConfiguration
            _ConfigurationName = swConf.Name
        End If

        swSheet = swDraw.GetCurrentSheet
        swSheet.GetSize(Width, Height)
        Return swDraw.CreateFlatPatternViewFromModelView3(swModel.GetPathName, _ConfigurationName, _IsopositionX, _IsopositionY, 0, False, False)

    End Function

    Private Function SetPositionThreeViews() As Double()
        Dim _Value(2) As Double
        Dim swView As SldWorks.View
        Dim nPtData(2, 1) As Double
        Dim IsoPositionX As Double
        Dim IsoPositionY As Double

        swView = swDraw.GetFirstView
        swView = swView.GetNextView

        For nPtDataPos = 0 To 2
            nPtData(nPtDataPos, 0) = swView.Position(0)
            nPtData(nPtDataPos, 1) = swView.Position(1)

            swView = swView.GetNextView
        Next

        If nPtData(0, 0) = nPtData(1, 0) Then
            IsoPositionX = nPtData(2, 0)
        ElseIf nPtData(0, 0) = nPtData(2, 0) Then
            IsoPositionX = nPtData(1, 0)
        Else
            IsoPositionX = nPtData(0, 0)
        End If

        If nPtData(0, 1) = nPtData(1, 1) Then
            IsoPositionY = nPtData(2, 1)
        ElseIf nPtData(0, 1) = nPtData(2, 1) Then
            IsoPositionY = nPtData(1, 1)
        Else
            IsoPositionY = nPtData(0, 1)
        End If

        _Value(0) = IsoPositionX
        _Value(1) = IsoPositionY

        Return _Value
    End Function

    Public Function CreateFlatPatternForDXF(Optional ByVal _ConfigurationName As String = Nothing) As Boolean
        Dim TemplateAdress As String = ""
        Dim swSheet As SldWorks.Sheet
        Dim Width, Height As Double

        For Each Template As Templates_p In Templates_p
            If Template.Name = "DXF" Then
                TemplateAdress = Template.Adress
                Exit For
            End If
        Next

        If _ConfigurationName = Nothing Then
            Dim swConfMgr As SldWorks.ConfigurationManager
            Dim swConf As SldWorks.Configuration
            swConfMgr = swModel.ConfigurationManager
            swConf = swConfMgr.ActiveConfiguration
            _ConfigurationName = swConf.Name
        End If

        swDraw = swApp.NewDocument(TemplateAdress, swDwgPaperSizes_e.swDwgPaperAsize, 0, 0)
        swSheet = swDraw.GetCurrentSheet

        swSheet.GetSize(Width, Height)
        swDraw.CreateFlatPatternViewFromModelView3(swModel.GetPathName, _ConfigurationName, Width / 2, Height / 2, 0, False, False)

        Return True

    End Function

    'Public Sub SetScaleInSW(ByVal _ScaleUpper As Double, ByVal _ScaleLower As Double, ByVal _ViewZone As Double)
    '    Dim swSheet As SldWorks.Sheet
    '    Dim vSheetProps As Object

    '    swSheet = swDraw.GetCurrentSheet
    '    vSheetProps = swSheet.GetProperties

    '    If DefineBox < _ViewZone Then
    '        swSheet.SetProperties( _
    '                                vSheetProps(0), _
    '                                vSheetProps(1), _
    '                                _ScaleUpper, _
    '                                _ScaleLower, _
    '                                vSheetProps(4), _
    '                                vSheetProps(5), _
    '                                vSheetProps(6))
    '    Else
    '        swSheet.SetProperties( _
    '                                vSheetProps(0), _
    '                                vSheetProps(1), _
    '                                _ScaleLower, _
    '                                _ScaleUpper, _
    '                                vSheetProps(4), _
    '                                vSheetProps(5), _
    '                                vSheetProps(6))
    '    End If

    'End Sub

End Class