Module Test

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

    Sub main()

        Dim swApp As SldWorks.SldWorks
        Dim swModel As SldWorks.ModelDoc2
        Dim Drawing As New Drawing()

        swApp = CreateObject("SldWorks.Application")
        swModel = swApp.ActiveDoc

        'Drawing.NewDrawing(swModel, DrawingCreationType_e.FlatPatternViewForDXF) 'ajouter une échelles forcée
        Drawing.NewDrawing(swModel, DrawingCreationType_e.ThreeViewsAndIso)

    End Sub

End Module
