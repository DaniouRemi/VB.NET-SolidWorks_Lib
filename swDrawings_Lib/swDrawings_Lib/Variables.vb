Imports SldWorks
Imports FE_Library

Public Class Variables

    Public Enum Scale_e
        RealScale = 0   'échelle 1:1
        SheetScale = 1  'échelle de la feuille
        ForceScale = 2  'échelle forcé
        AutoScale = 3   'échelle auto
    End Enum

    Public Enum Orientation_e
        Portrait = 0
        Landscape = 1
    End Enum

    Public Enum DrawingCreationType_e
        ThreeViews = 0                     '3 vues standards
        ThreeViewsAndIso = 1               '3 vues standards + vue iso
        ThreeViewsAndIsoAndFlatPattern = 2 '3 vues standards + vue iso + déplié
        FrontView = 3                      'vue de fase
        BackView = 4                       'vue de derrière
        TopView = 5                        'vue de dessus
        BottomView = 6                     'vue de dessous
        LeftView = 7                       'vue de gauche
        RightView = 8                      'vue de droite
        FlatPatternView = 9                'vue déplié
        FlatPatternViewForDXF = 10         'vue déplié pour DXF
        CustomView = 11                    'vue personalisée
        NoView = 12                        'sans vue
    End Enum

    Public Shared bRet As Boolean

    Public Shared swApp As SldWorks.SldWorks
    Public Shared swModel As SldWorks.ModelDoc2
    Public Shared swDraw As SldWorks.DrawingDoc
    Public Shared swSheet As SldWorks.Sheet

    Public Shared Numerical As FE_Library.Numerical

    Public Shared DrawingGeneral As General
    Protected Friend IsometricView As String = "*Isométrique"

    Public Shared Templates_p As List(Of Templates_p)

    Public Shared vBox As Object
    Public Shared Box(3) As Double
    Public Shared orderBox As String
    Public Shared DefineBox As Double

    Public VarScale1 As Double
    Public VarScale2 As Double
    Public VarScale3 As Double

    Public Shared TemplateOrientation As Integer
    Public Shared TemplateSize As Integer

    Public Shared View_Rotation As Double

    Public Shared ScaleUsed As ScaleMode
    Public Shared DecScale As Double
    Public Shared ScaleUpper As Double
    Public Shared ScaleLower As Double

    'Paramètres utilisateur
    Public Shared DrawingCreationType As Integer
    Public Shared ViewRotationActivate As Boolean
    Public Shared AutoTemplateActivate As Boolean
    Public Shared AutoScaleActivate As Boolean
    Public Shared CustomScaleValue As Double
    Public Shared CustomTemplatePath As String
End Class
