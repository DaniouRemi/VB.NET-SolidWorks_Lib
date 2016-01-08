Imports System.Math
Imports SwConst

Public Class Drawing
    Inherits Variables

    Private Templates As Templates
    Private Functions As Functions

    Sub New()

        DrawingGeneral = New General
        Numerical = New FE_Library.Numerical
        Templates = New Templates
        Functions = New Functions

    End Sub

#Region "NewDrawing"

    Public Sub NewDrawing(_PatchFileName As String, DrawingCreationType As Short, _TemplateAdress As String, _Scale As Double, _RotateView As Boolean)

        ViewRotationActivate = _RotateView
        AutoTemplateActivate = False
        AutoScaleActivate = False
        CustomTemplatePath = _TemplateAdress
        CustomScaleValue = _Scale
        swModel = SetswModel(_PatchFileName)

        NewDrawing()

    End Sub 'création du plan avec le template et l'échelle définis

    Public Sub NewDrawing(_swModel As SldWorks.ModelDoc2, _DrawingCreationType As Short, _TemplateAdress As String, _Scale As Double, _RotateView As Boolean)

        ViewRotationActivate = _RotateView
        AutoTemplateActivate = False
        AutoScaleActivate = False
        CustomTemplatePath = _TemplateAdress
        CustomScaleValue = _Scale
        DrawingCreationType = _DrawingCreationType
        swModel = _swModel

        NewDrawing()

    End Sub 'création du plan avec le template et l'échelle définis

    Public Sub NewDrawing(_swModel As SldWorks.ModelDoc2, _DrawingCreationType As Integer)

        ViewRotationActivate = True
        AutoTemplateActivate = True
        AutoScaleActivate = True
        DrawingCreationType = _DrawingCreationType
        swModel = _swModel

        NewDrawing()

    End Sub 'création du plan en automatique

    Public Sub NewDrawing(_PatchFileName As String, _DrawingCreationType As Short)

        ViewRotationActivate = True
        AutoTemplateActivate = True
        AutoScaleActivate = True
        swModel = SetswModel(_PatchFileName)
        DrawingCreationType = _DrawingCreationType

        NewDrawing()

    End Sub 'création du plan en automatique

    Public Sub NewDrawing(_swModel As SldWorks.ModelDoc2, _DrawingCreationType As Short, ByVal _TemplateAdress As String, ByVal _RotateView As Boolean)

        ViewRotationActivate = _RotateView
        AutoTemplateActivate = False
        AutoScaleActivate = True
        CustomTemplatePath = _TemplateAdress
        DrawingCreationType = _DrawingCreationType
        swModel = _swModel

        NewDrawing()

    End Sub 'création du plan avec le template et la rotation des vues

    Private TemplateUsed As Templates_p
    Private Sub Newdrawing()

        If IsNothing(swApp) Then
            swApp = CreateObject("SldWorks.Application")
        End If

        If DrawingCreationType = DrawingCreationType_e.FlatPatternViewForDXF Then
            Functions.CreateFlatPatternForDXF()
        Else
            'set template and scale
            If AutoTemplateActivate Then

                If AutoScaleActivate Then
                    'TemplateUsed = Functions.TemplateAuto(ViewRotationActivate)
                    TemplateUsed = Functions.TemplateAuto()
                    ScaleUsed = ScaleUsed
                    DecScale = DecScale

                Else
                    'TemplateUsed = Functions.TemplateAuto(ViewRotationActivate, CustomScaleValue)
                    TemplateUsed = Functions.TemplateAuto()
                    DecScale = CustomScaleValue
                End If

                swDraw = swApp.NewDocument(TemplateUsed.Adress, swDwgPaperSizes_e.swDwgPaperAsize, 0, 0)

                ScaleUpper = Split(Functions.Dec2Frac_fc(Round(DecScale, 1)), "/")(0)
                ScaleLower = Split(Functions.Dec2Frac_fc(Round(DecScale, 1)), "/")(1)
                'Functions.SetScaleInSW(ScaleUpper, ScaleLower, TemplateUsed.ScaleMode.Item(DrawingCreationType).ViewZone)
                Functions.CreateViews(DrawingCreationType)

            Else

                If AutoScaleActivate Then
                    Functions.ScaleAuto(ViewRotationActivate, CustomTemplatePath)
                Else

                End If
            End If
        End If
    End Sub

#End Region

    Private Function SetswModel(_PatchFileName As String)
        Dim ExtentionFile As String
        Dim TypeFile As Integer
        Dim nErrors As Long

        If IsNothing(swApp) Then
            swApp = CreateObject("SldWorks.Application")
        End If

        ExtentionFile = LCase(Right(_PatchFileName, 7))

        Select Case ExtentionFile
            Case ".sldasm"
                TypeFile = 2
            Case ".sldprt"
                TypeFile = 1
        End Select

        Return swApp.OpenDoc2(_PatchFileName, TypeFile, True, True, True, nErrors)

    End Function

End Class
