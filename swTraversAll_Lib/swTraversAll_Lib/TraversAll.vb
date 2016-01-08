Imports System.IO
Imports SwConst
Imports System.Text.RegularExpressions

Public Class TraverseAll
    Inherits TraversAllVariables

    Private Checks As New Checks
    Private PropertiesRecovery As New PropertiesRecovery

    Protected swcomponent As SldWorks.Component2
    Protected ChemComp As String
    Protected NomConf As String
    Protected NomComp As String
    Protected TypeFile As Long

    Protected PropertiesOptions As PropertiesOptions
    Protected TreatThis As Boolean

#Region "Énumérations SolidWorks"

    Protected Enum swDocumentTypes_e
        swDocASSEMBLY = 2
        swDocDRAWING = 3
        swDocNONE = 0
        swDocPART = 1
        swDocSDM = 4
    End Enum

#End Region

    Protected tParent As New List(Of tParents)
    Protected childID As Integer = ID
    Protected TreatID As Integer

    'Friend Sub Traverse(Optional _PropertiesOptions As PropertiesOptions = Nothing)
    '    PropertiesOptions = _PropertiesOptions
    '    Debug.Print(Now.ToLongTimeString & " - " & "DEBUT définition du type de composant d'entrée")

    '    Select Case LCase(Right(swModel.GetPathName, 7))
    '        Case ".sldasm"
    '            Debug.Print(Now.ToLongTimeString & " - " & "    - Composant type 'assemblage'")
    '            TraverseASM()

    '        Case ".sldprt"

    '            If PropertiesOptions.FEMode Then

    '            End If
    '            Debug.Print(Now.ToLongTimeString & " - " & "    - Composant type 'pièce'")

    '            With PRT_BOM
    '                .Patch = swModel.GetPathName
    '                .Type = swDocumentTypes_e.swDocPART
    '                .swConfMgr = swModel.ConfigurationManager
    '                .swConf = PRT_BOM.swConfMgr.ActiveConfiguration
    '                .swRootComp = PRT_BOM.swConf.GetRootComponent3(True)
    '                .vComponent = PRT_BOM.swRootComp.GetChildren
    '                .LastConfiguration = PRT_BOM.swConf.Name
    '            End With

    '            TraverseFunctions()
    '            TraverseBodies()

    '        Case ".slddrw"
    '            Debug.Print(Now.ToLongTimeString & " - " & "    - Composant type 'mise en plan'")

    '    End Select

    'End Sub

    ''' <summary>
    ''' Début du TraverseAll
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub TraverseASM(Optional _PropertiesOptions As PropertiesOptions = Nothing)
        Debug.Print(Now.ToLongTimeString & " - " & "Début du TraverseAll de l'assemblage")

        PropertiesOptions = _PropertiesOptions
        Dim vParent As New Parents
        Dim vID As Integer = Nothing
        Dim swConfMgr As SldWorks.ConfigurationManager
        Dim swConf As SldWorks.Configuration
        Dim swRootComp As SldWorks.Component2
        Dim vComponent As Object
        Dim ShortName As String

        ASM_BOM.Add(New Components)
        'PRT_BOM = New Components
        ID += 1
        swConfMgr = swModel.ConfigurationManager
        swConf = swConfMgr.ActiveConfiguration
        swRootComp = swConf.GetRootComponent3(True)
        vComponent = swRootComp.GetChildren
        TreatThis = True

        ASM_BOM.Item(ID).swConfMgr = swConfMgr
        ASM_BOM.Item(ID).swConf = swConf
        ASM_BOM.Item(ID).swRootComp = swRootComp
        ASM_BOM.Item(ID).vComponent = vComponent

        'With PRT_BOM
        '    .swConfMgr = swConfMgr
        '    .swConf = swConf
        '    .swRootComp = swRootComp
        '    .vComponent = vComponent
        'End With

        'ASM_BOM.Add(PRT_BOM)

        'définition du composant initial
        If ASM_BOM.Count <= 1 Then
            Debug.Print(Now.ToLongTimeString & " - " & "Définition du composant initial (type ASM)")
            ASM_BOM.Item(ID).Patch = swModel.GetPathName
            ASM_BOM.Item(ID).LastConfiguration = swConf.Name
            ASM_BOM.Add(New Components)
            ID += 1
        End If

        'If ASM_BOM.Count <= 0 Then
        '    Debug.Print(Now.ToLongTimeString & " - " & "Définition du composant initial (type ASM)")
        '    PRT_BOM.Patch = swModel.GetPathName
        '    PRT_BOM.LastConfiguration = swConf.Name
        '    ASM_BOM.Add(PRT_BOM)
        '    PRT_BOM = New Components
        '    ID += 1
        'End If

        'ASM_BOM.Add(PRT_BOM)

        ' Ajout du nouveau parent dans la liste temporaire
        tParent.Add(New tParents With {.Adress = swModel.GetPathName, _
                                           .ConfigurationName = swConf.Name, _
                                           .ID = ID - 1})

        Debug.Print(Now.ToLongTimeString & " - " & "Initialisation du parent temporaire :")
        Debug.Print(Now.ToLongTimeString & " - " & "    - Chemin : '" & swModel.GetPathName & "'")
        Debug.Print(Now.ToLongTimeString & " - " & "    - Configuration : '" & swConf.Name & "'")
        Debug.Print(Now.ToLongTimeString & " - " & "    - Nombre de composants : '" & UBound(vComponent) & "'")

        ShortName = Split(swModel.GetPathName, "\")(Split(swModel.GetPathName, "\").Count - 1) 'Pour le Debug

        For i As Integer = 0 To UBound(vComponent)

            Debug.Print(Now.ToLongTimeString & " - " & "composant ID " & i & " / " & UBound(vComponent) & " (dans '" & ShortName & "') :")

            ASM_BOM.Item(ID).swComponent = vComponent(i)
            ASM_BOM.Item(ID).Patch = ASM_BOM.Item(ID).swComponent.GetPathName
            ASM_BOM.Item(ID).LastConfiguration = ASM_BOM.Item(ID).swComponent.ReferencedConfiguration

            'PRT_BOM.swComponent = vComponent(i)
            'PRT_BOM.Patch = PRT_BOM.swComponent.GetPathName
            'PRT_BOM.LastConfiguration = PRT_BOM.swComponent.ReferencedConfiguration

            'ASM_BOM.Add(PRT_BOM)

            Debug.Print(Now.ToLongTimeString & " - " & "    - Adresse : '" & ASM_BOM.Item(ID).Patch & "'")
            Debug.Print(Now.ToLongTimeString & " - " & "    - Configuration : '" & ASM_BOM.Item(ID).LastConfiguration & "'")

            childID = ID 'récupérer ID de l'enfant

            '================== VERIFICATION DOUBLON ==================
            Checks.DoubloonCheck(ASM_BOM, tParent.Last, ID)
            ASM_BOM = Checks.BOM_Checked
            Dim NewConfiguration As Boolean = Checks.NewConfiguration
            Dim NewComponent As Boolean = Checks.NewComponent

            If Not NewConfiguration Then
                ASM_BOM = Checks.BOM_Checked
            End If

            '================== Création composant ==================
            If NewConfiguration Then
                Debug.Print(Now.ToLongTimeString & " - " & "    DEBUT de l'insertion de la configuration dans la liste")
                If Not NewComponent Then
                    Debug.Print(Now.ToLongTimeString & " - " & "        - Composant déjà existant (Index = " & Checks.ComponentID & ")")
                    ASM_BOM.Remove(ASM_BOM.Last)
                    ID -= 1
                    TreatID = Checks.ComponentID
                Else
                    Debug.Print(Now.ToLongTimeString & " - " & "        - Nouveau composant (Index = " & ID & ")")

                    TreatID = ID
                End If

                With ASM_BOM.Item(TreatID)
                    .swComponent = vComponent(i)
                    .LastConfiguration = ASM_BOM.Item(TreatID).swComponent.ReferencedConfiguration
                    .Configuration.Add(New Configurations() With {.Name = ASM_BOM.Item(TreatID).LastConfiguration, _
                                                                              .Quantity = 1})
                End With

                Debug.Print(Now.ToLongTimeString & " - " & "        - Nom de la configuration : '" & ASM_BOM.Item(TreatID).LastConfiguration & "' (Index = " & ASM_BOM.Item(TreatID).Configuration.Count - 1 & ")")
                Debug.Print(Now.ToLongTimeString & " - " & "    FIN de l'insertion de la configuration dans la liste")

                'Insertion du parent dans le composant
                Debug.Print(Now.ToLongTimeString & " - " & "    DEBUT de l'insertion du parent dans la configuration")
                ASM_BOM.Item(TreatID).Configuration.Last.Parent.Add(New Parents With {.Adress = tParent.Last.Adress, _
                                                                            .ConfigurationName = tParent.Last.ConfigurationName, _
                                                                             .Quantity = 1})
                Debug.Print(Now.ToLongTimeString & " - " & "        - Adresse : '" & tParent.Last.Adress & "'")
                Debug.Print(Now.ToLongTimeString & " - " & "        - Configuration : '" & tParent.Last.ConfigurationName & "'")
                Debug.Print(Now.ToLongTimeString & " - " & "    FIN de l'insertion du parent dans la configuration")

                '================== Définition du type de composant ==================
                Debug.Print(Now.ToLongTimeString & " - " & "    DEBUT de la définition du type du composant")
                Dim ExtentionFile As String
                ExtentionFile = LCase(Right(ASM_BOM.Item(TreatID).Patch, 7))
                Select Case ExtentionFile
                    Case ".sldasm"
                        Debug.Print(Now.ToLongTimeString & " - " & "        - Type du composant : Assemblage (*.SLDASM)")
                        ASM_BOM.Item(TreatID).Type = swDocumentTypes_e.swDocASSEMBLY

                    Case ".sldprt"
                        Debug.Print(Now.ToLongTimeString & " - " & "        - Type du composant : Pièce (*.SLDPRT)")
                        ASM_BOM.Item(TreatID).Type = swDocumentTypes_e.swDocPART
                End Select
                Debug.Print(Now.ToLongTimeString & " - " & "    FIN de la définition du type du composant")

                If _PropertiesOptions.FEMode Then
                    TreatFranceEtuves() 'traitement France-Etuves
                End If

                Treatsuppressed() 'état de suppression

                '================== Traitement du composant ==================
                Debug.Print(Now.ToLongTimeString & " - " & "    DEBUT du traitemnt du composant")

                Debug.Print(Now.ToLongTimeString & " - " & "        DEBUT ouverture du fichier")

                If File.Exists(ASM_BOM.Item(TreatID).Patch) Then
                    swModel = swApp.OpenDoc6(ASM_BOM.Item(TreatID).Patch, ASM_BOM.Item(TreatID).Type, 0, ASM_BOM.Item(TreatID).LastConfiguration, nErrors, nWarnings)
                    ASM_BOM.Item(TreatID).swModel = swModel
                    Debug.Print(Now.ToLongTimeString & " - " & "            - Fichier ouvert")

                    Debug.Print(Now.ToLongTimeString & " - " & "        FIN ouverture du fichier")

                    PropertiesRecovery.PropertiesComponent(ASM_BOM, TreatID)
                    ASM_BOM = PropertiesRecovery._BOM

                    If ASM_BOM.Item(TreatID).Type = swDocumentTypes_e.swDocASSEMBLY Then
                        Debug.Print(Now.ToLongTimeString & " - " & "        DEBUT ouverture du composant (type ASM) '(" & Split(ASM_BOM.Item(TreatID).Patch, "\")(Split(ASM_BOM.Item(TreatID).Patch, "\").Count - 1) & "')")
                        vID = TreatID
                        Debug.Print(Now.ToLongTimeString & " - " & "            - Lancement de la procédure 'TraverseASM()'")
                        TraverseASM(PropertiesOptions)
                        Debug.Print(Now.ToLongTimeString & " - " & "            - Suppression du parent temporaire ('" & Split(tParent.Last.Adress, "\")(Split(tParent.Last.Adress, "\").Count - 1))
                        tParent.Remove(tParent.Last)
                        Debug.Print(Now.ToLongTimeString & " - " & "        FIN ouverture du composant (type ASM) '(" & Split(ASM_BOM.Item(TreatID).Patch, "\")(Split(ASM_BOM.Item(TreatID).Patch, "\").Count - 1) & "')")

                    Else
                        PRT_BOM = ASM_BOM.Item(TreatID)
                        TraverseFunctions()
                        TraverseBodies()
                        ASM_BOM.Item(TreatID) = PRT_BOM
                        ASM_BOM.Add(New Components)

                        ID += 1
                        swApp.QuitDoc(ASM_BOM.Item(TreatID).Patch)
                    End If

                Else
                    Debug.Print(Now.ToLongTimeString & " - " & "Fichier non existant")
                End If

            Else
                'Debug.Print(Now.ToLongTimeString & " - " & "Composant non renseigné dans la nomenclature")
            End If

        Next

    End Sub

    Public BobiesFeature As SldWorks.Feature
    ''' <summary>
    ''' Traversée des fonctions
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub TraverseFunctions()
        Debug.Print(Now.ToLongTimeString & " - " & "        DEBUT Traversée des fonctions")
        'TODO créer une boucle pour récupérer toutes les fonctions et sous-fonctions

        'Feat_Obj = ASM_BOM.Item(TreatID).swModel.FirstFeature
        Feat_Obj = PRT_BOM.swModel.FirstFeature

        Do While Not Feat_Obj Is Nothing


            'ASM_BOM.Item(TreatID).Configuration.Last.Functions.Add(New Functions With {.Name = Feat_Obj.GetTypeName2, _
            '                                                                                   .Feature = Feat_Obj})
            PRT_BOM.Configuration.Last.Functions.Add(New Functions With {.Name = Feat_Obj.GetTypeName2, _
                                                                                               .Feature = Feat_Obj})
            Debug.Print(Now.ToLongTimeString & " - " & "            " & Feat_Obj.Name & " ('" & Feat_Obj.GetTypeName2 & "')")

            If Feat_Obj.GetTypeName2 = "SolidBodyFolder" Then
                Debug.Print(Now.ToLongTimeString & " -                   " & "- Mise à jour de '" & Feat_Obj.Name & "'")
                UpdateCutList(Feat_Obj)
                BobiesFeature = Feat_Obj

                If PropertiesOptions.FEMode Then
                    Exit Do
                ElseIf Not PropertiesOptions.comp.fct Then
                    Exit Do
                End If

            End If

            TraverseSubFunctions(Feat_Obj, PRT_BOM.Configuration.Last.Functions)

            Feat_Obj = Feat_Obj.GetNextFeature
        Loop

        Debug.Print(Now.ToLongTimeString & " - " & "        FIN Traversée des fonctions")
    End Sub

    Private Sub TraverseSubFunctions(Feat_Obj As SldWorks.Feature, Features As List(Of Functions), Optional SpaceForDebug As String = "   ")

        Dim SubFeat As Object
        SubFeat = Feat_Obj.GetFirstSubFeature

        Do While Not SubFeat Is Nothing
            Features.Last.SubFunction.Add(New Functions With {.Name = SubFeat.GetTypeName2, _
                                                              .Feature = SubFeat})

            Debug.Print(Now.ToLongTimeString & " - " & "            " & SpaceForDebug & SubFeat.name & " ('" & SubFeat.GetTypeName2 & "')")

            SpaceForDebug = SpaceForDebug & "   "
            TraverseSubFunctions(SubFeat, Features.Last.SubFunction, SpaceForDebug)


            SubFeat = SubFeat.GetNextSubFeature
            SpaceForDebug = Left(SpaceForDebug, SpaceForDebug.Length - 3)
        Loop
    End Sub

    ''' <summary>
    ''' Traversée des corps
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub TraverseBodies()
        Debug.Print(Now.ToLongTimeString & " - " & "        DEBUT Traversée des corps")

        Debug.Print(Now.ToLongTimeString & " - " & "            - Recherche de 'SolidBodyFolder'")
        'For Each Ftc As Functions In ASM_BOM.Item(TreatID).Configuration.Last.Functions
        '    If Ftc.Feature.GetTypeName2 = "SolidBodyFolder" Then
        '        Debug.Print(Now.ToLongTimeString & " - " & "            - 'SolidBodyFolder' trouvé")

        '        Debug.Print(Now.ToLongTimeString & " - " & "            - 'Recherche des corps")

        '        For Each SubFtc As Functions In Ftc.SubFunction
        '            If SubFtc.Feature.GetTypeName2 = "CutListFolder" Then
        '                Debug.Print(Now.ToLongTimeString & " - " & "            - 'corps trouvé ('" & SubFtc.Feature.Name & "')")

        '                'FlatPattern(SubFtc.Feature)

        '                Debug.Print(Now.ToLongTimeString & " - " & "                - Initialisation du corps")
        '                ASM_BOM.Item(TreatID).swBodyFolder = SubFtc.Feature.GetSpecificFeature2
        '                ASM_BOM.Item(TreatID).swCustPropMgr = SubFtc.Feature.CustomPropertyManager

        '                If Not ASM_BOM.Item(TreatID).swCustPropMgr Is Nothing Then
        '                    FlatPattern(SubFtc.Feature)
        '                    Debug.Print(Now.ToLongTimeString & " - " & "            - Mise à jour de la liste de pièces soudées")
        '                    UpdateCutList(Ftc.Feature)
        '                    PropertiesRecovery.PropertiesBodies(ASM_BOM, TreatID, SubFtc.Feature)
        '                    ASM_BOM = PropertiesRecovery._BOM
        '                End If

        '            End If
        '        Next
        '    End If
        'Next

        For Each Ftc As Functions In PRT_BOM.Configuration.Last.Functions
            If Ftc.Feature.GetTypeName2 = "SolidBodyFolder" Then
                Debug.Print(Now.ToLongTimeString & " - " & "            - 'SolidBodyFolder' trouvé")

                Debug.Print(Now.ToLongTimeString & " - " & "            - 'Recherche des corps")

                For Each SubFtc As Functions In Ftc.SubFunction
                    If SubFtc.Feature.GetTypeName2 = "CutListFolder" Then
                        Debug.Print(Now.ToLongTimeString & " - " & "            - 'corps trouvé ('" & SubFtc.Feature.Name & "')")

                        'FlatPattern(SubFtc.Feature)

                        Debug.Print(Now.ToLongTimeString & " - " & "                - Initialisation du corps")
                        PRT_BOM.swBodyFolder = SubFtc.Feature.GetSpecificFeature2
                        PRT_BOM.swCustPropMgr = SubFtc.Feature.CustomPropertyManager

                        If Not PRT_BOM.swCustPropMgr Is Nothing Then
                            FlatPattern(SubFtc.Feature)
                            Debug.Print(Now.ToLongTimeString & " - " & "            - Mise à jour de la liste de pièces soudées")
                            UpdateCutList(Ftc.Feature)
                            'PropertiesRecovery.PropertiesBodies(ASM_BOM, TreatID, SubFtc.Feature)
                            'ASM_BOM = PropertiesRecovery._BOM

                            PropertiesRecovery.PropertiesBodies(PRT_BOM, SubFtc.Feature)
                            'ASM_BOM(TreatID) = PRT_BOM

                        End If

                    End If
                Next
            End If
        Next

        Debug.Print(Now.ToLongTimeString & " - " & "        FIN Traversée des corps")
    End Sub

#Region "Bodies"

    ''' <summary>
    ''' dépliage et repliage d'une pièce de tôlerie
    ''' </summary>
    ''' <param name="Feature">fonction (Feature.GetTypeName2 = "CutListFolder")</param>
    ''' <remarks></remarks>
    Private Sub FlatPattern(Feature As SldWorks.Feature)
        Debug.Print(Now.ToLongTimeString & " - " & "                DEBUT dépliage/pliage de du corps")
        Dim SubFeat As SldWorks.Feature = Feature.GetNextFeature
        Debug.Print(Now.ToLongTimeString & " - " & "                    - Recherche de 'FlatPattern'")

        'Do While Not SubFeat Is Nothing
        '    If SubFeat.GetTypeName2 = "FlatPattern" Then
        '        Debug.Print(Now.ToLongTimeString & " - " & "                    - 'FlatPattern' trouvé")
        '        'Feature.SetSuppression2(swFeatureSuppressionAction_e.swUnSuppressFeature, swInConfigurationOpts_e.swThisConfiguration, NomConf)
        '        'ASM_BOM.Item(TreatID).swModel.ForceRebuild3(True)
        '        'Feature.SetSuppression2(swFeatureSuppressionAction_e.swSuppressFeature, swInConfigurationOpts_e.swThisConfiguration, NomConf)
        '        'ASM_BOM.Item(TreatID).swModel.ForceRebuild3(True)
        '        Feature.SetSuppression2(swFeatureSuppressionAction_e.swUnSuppressFeature, swInConfigurationOpts_e.swThisConfiguration, NomConf)
        '        PRT_BOM.swModel.ForceRebuild3(True)
        '        Feature.SetSuppression2(swFeatureSuppressionAction_e.swSuppressFeature, swInConfigurationOpts_e.swThisConfiguration, NomConf)
        '        PRT_BOM.swModel.ForceRebuild3(True)
        '        Debug.Print(Now.ToLongTimeString & " - " & "                    - corps déplié et replié")
        '        Exit Do
        '    End If

        '    SubFeat = SubFeat.GetNextFeature
        'Loop

        Debug.Print(Now.ToLongTimeString & " - " & "                FIN dépliage/pliage de du corps")
    End Sub

    ''' <summary>
    ''' Mise à jour de la liste de pièces soudées
    ''' </summary>
    ''' <param name="_swFeature">Fonction</param>
    ''' <remarks></remarks>
    Public Sub UpdateCutList(ByVal _swFeature As SldWorks.Feature)
        Debug.Print(Now.ToLongTimeString & " -                   " & "DEBUT MAJ de la liste de pièces soudées")
        Dim _swBodyFolder As SldWorks.BodyFolder = Nothing

        Try
            _swBodyFolder = _swFeature.GetSpecificFeature2
            _swBodyFolder.SetAutomaticCutList(True)
            _swBodyFolder.UpdateCutList()
            Debug.Print(Now.ToLongTimeString & " -                   " & "   - MAJ OK")
        Catch ex As Exception
            Debug.Print(Now.ToLongTimeString & " -                   " & "ERROR MAJ - ex : " & ex.ToString)
        End Try

        Debug.Print(Now.ToLongTimeString & " -                   " & "FIN MAJ de la liste de pièces soudées")
    End Sub

#End Region

#Region "Traitements divers"

    Private Sub Treatsuppressed()
        If ASM_BOM.Item(TreatID).swComponent.GetSuppression = 0 Then
            ASM_BOM.Item(TreatID).Suppressed = True
        Else
            ASM_BOM.Item(TreatID).Suppressed = False
        End If
    End Sub

    Private Sub TreatFranceEtuves()
        Dim PartName As String
        TreatThis = False

        If ASM_BOM.Item(TreatID).Type = swDocumentTypes_e.swDocPART Then
            PartName = ASM_BOM.Item(TreatID).Patch
            If Regex.IsMatch(PartName, "^(\d{1,3}\-\d{1,3})|\d{6}") Then
                TreatThis = True
            End If
        End If
    End Sub

#End Region

#Region "Calcul des quantités totales"

    Private QuantityByComponent As List(Of Double)
    Private QuantityTotal As List(Of Double)


    Protected Friend Sub Quantity()
        'TODO Quantités pas bonnes
        Debug.Print(Now.ToLongTimeString & " - " & "    DEBUT Calcul des quantités totales")
        QuantityByComponent = New List(Of Double)

        For Each Component As Components In ASM_BOM

            QuantityTotal = New List(Of Double)
            QuantityByComponent = New List(Of Double)

            For Each Configuration As Configurations In Component.Configuration
                Debug.Print(Now.ToLongTimeString & " - " & "        - Ajout de la quantité totale de " & Split(Component.Patch, "\") _
                            (Split(Component.Patch, "\").Count - 1) & " [" & Configuration.Name & "] (" & Configuration.Parent.Count & " parents)")
                GetParent(Configuration, "---")
            Next

            For Each Qty As Double In QuantityTotal
                Component.Quantity = Component.Quantity + Qty
            Next
            Debug.Print(Now.ToLongTimeString & " - " & "            - Quantité = " & Component.Quantity)

        Next

        Debug.Print(Now.ToLongTimeString & " - " & "    FIN Calcul des quantités totales")
    End Sub 'pour parcourir les composants

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Configuration"></param>
    ''' <param name="Space"></param>
    ''' <remarks></remarks>
    Private Sub GetParent(Configuration As Configurations, Space As String)

        If Not IsNothing(Configuration) Then

            Space = Space + "---"
            For Each Par As Parents In Configuration.Parent
                Debug.Print(Par.Adress)
                QuantityByComponent.Add(Par.Quantity)

                If ParentsToComponents(Par).Quantity <> 0 Then

                    QuantityByComponent.Add(ParentsToComponents(Par).Quantity)

                    Dim Total As Double = 1
                    For Each Qty As Double In QuantityByComponent
                        Total = Total * Qty
                    Next

                    QuantityTotal.Add(Total)

                    QuantityByComponent = New List(Of Double)

                Else

                    GetParent(ParentsToConfigurations(Par), Space)
                    'Private QuantityTotal As List(Of List(Of Double)) ??!!
                End If

            Next

        Else

            Dim Total As Double = 1
            For Each Qty As Double In QuantityByComponent
                Total = Total * Qty
            Next

            QuantityTotal.Add(Total)
            'Debug.Print(Now.ToLongTimeString & " - " & "            - Total = " & Total)

        End If

        If QuantityByComponent.Count >= 1 Then
            QuantityByComponent.RemoveAt(QuantityByComponent.Count - 1)
        End If

    End Sub 'pour monter dans les parents

    ''' <summary>
    ''' Conversion du type 'Parents' en tant que type 'Configurations'
    ''' </summary>
    ''' <param name="_Parent">'Parents' à récupérer en tant que 'Configurations'</param>
    ''' <returns>Le 'parent' en tant que 'Configurations'</returns>
    ''' <remarks></remarks>
    Private Function ParentsToConfigurations(_Parent As Parents) As Configurations

        For Each Component As Components In ASM_BOM

            If Component.Patch = _Parent.Adress Then

                For Each Configuration As Configurations In Component.Configuration
                    If Configuration.Name = _Parent.ConfigurationName Then
                        Return Configuration
                    End If
                Next

            End If

        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' Conversion du type 'Parents' en tant que type 'Components'
    ''' </summary>
    ''' <param name="_Parent">'Parents' à récupérer en tant que 'components'</param>
    ''' <returns>Le 'parent' en tant que 'Components'</returns>
    ''' <remarks></remarks>
    Private Function ParentsToComponents(_Parent As Parents) As Components

        For Each Component As Components In ASM_BOM

            If Component.Patch = _Parent.Adress Then

                Return Component

            End If

        Next
        Return Nothing
    End Function

#End Region



End Class