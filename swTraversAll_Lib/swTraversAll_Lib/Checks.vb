Public Class Checks
    Inherits TraversAllVariables

    Friend BOM_Checked As New List(Of Components)

    Public Enum TypeReport
        Information = 0
        Alert = 1
        Critical = 2
    End Enum
#Region "Report"
    Friend Sub AddReport(Type As Integer, Text As String)

        Report.Add(New Report With {.Time = Date.Now.ToLongTimeString, _
                                    .type = Type, _
                                   .Name = "", _
                                   .Description = Text})
        'Console.WriteLine(Report.Last.Time & " : " & Report.Last.Name & ", " & Report.Last.Description)
    End Sub

    Friend Sub endtest()

    End Sub

#End Region

#Region "Checks"

    Friend Sub SW_Open()
        Dim proc() As System.Diagnostics.Process

        proc = System.Diagnostics.Process.GetProcessesByName("SLDWORKS")
        If proc.Count < 1 Then
            Debug.Print(Now.ToLongTimeString & " - " & "SLDWORKS no lunched")
            endtest()
        End If

    End Sub

    Friend Sub Check(swModel As SldWorks.ModelDoc2)

        If IsNothing(swModel) Then
            Debug.Print(Now.ToLongTimeString & " - " & "swModel is Nothing")
            endtest()
        Else
            Debug.Print(Now.ToLongTimeString & " - " & "swModel is " & swModel.GetPathName)
        End If

    End Sub
#End Region

#Region "Parents/enfants"

    Friend NewComponent As Boolean
    Friend NewConfiguration As Boolean
    Friend ComponentID As Integer
    Friend ConfigurationID As Integer

    ''' <summary>
    ''' Vérification d'un doublon dans la liste
    ''' </summary>
    ''' <param name="BOM">Liste de composants</param>
    ''' <param name="tParent">le parent du composant à vérifier</param>
    ''' <param name="ID">Index du composant à vérifier</param>
    ''' <remarks></remarks>
    Friend Sub DoubloonCheck(BOM As List(Of Components), tParent As tParents, ID As Integer)
        Debug.Print(Now.ToLongTimeString & " -     " & "DEBUT vérification du doublon")

        BOM_Checked = BOM
        NewComponent = True
        NewConfiguration = True

        If BOM.Count <= 2 Then
            Debug.Print(Now.ToLongTimeString & " -         " & "- Aucune vérification à effectuer")
        End If
        Debug.Print(Now.ToLongTimeString & " -         " & "- Recherche d'un composant identique")
        For i As Integer = 1 To BOM.Count - 2

            If BOM_Checked.Item(i).Patch = BOM_Checked.Item(ID).Patch Then 'si les chemins sont identiques
                Debug.Print(Now.ToLongTimeString & " -         " & "- Composant similaire trouvé en index " & i)
                NewComponent = False
                'BOM_Checked.Item(i).Quantity += 1
                Debug.Print(Now.ToLongTimeString & " -         " & "- Incrémentation de la quantité :" & BOM_Checked.Item(i).Quantity)
                ComponentID = i

                For j As Integer = 0 To BOM_Checked.Item(i).Configuration.Count - 1
                    Debug.Print(Now.ToLongTimeString & " -         " & "- Recherche d'une configuration identique")
                    If BOM_Checked.Item(ComponentID).Configuration.Item(j).Name = BOM_Checked.Item(ID).LastConfiguration Then
                        Debug.Print(Now.ToLongTimeString & " -         " & "- configuration similaire trouvée en Index " & j)
                        NewConfiguration = False
                        BOM_Checked.Item(ComponentID).Configuration.Item(j).Quantity += 1
                        Debug.Print(Now.ToLongTimeString & " -         " & "- Incrémentation de la quantité :" & BOM_Checked.Item(ComponentID).Configuration.Item(j).Quantity)
                        ConfigurationID = j
                        'Vérification du parent ==================
                        ParentCheck(BOM, tParent, ID, ComponentID, ConfigurationID)
                        BOM = BOM_Checked
                        Exit For
                    End If

                Next
                Exit For

            End If

        Next
        Debug.Print(Now.ToLongTimeString & " -     " & "FIN vérification du doublon")
    End Sub

    Friend ChildID As Integer
    Friend Sub ParentCheck(BOM As List(Of Components), tParent As tParents, ID As Integer, PartID As Integer, ConfigID As Integer)
        Debug.Print(Now.ToLongTimeString & " -     " & "DEBUT vérification du parent renseigné")
        BOM_Checked = BOM

        Dim tVerif As String = tParent.Adress & tParent.ConfigurationName
        Dim ttVerif As String
        Dim NewParent As Boolean = True

        Debug.Print(Now.ToLongTimeString & " -         " & "- Recherche d'un composant parent identique")
        For Each Parent As Parents In BOM_Checked.Item(PartID).Configuration.Item(ConfigID).Parent
            ttVerif = Parent.Adress & Parent.ConfigurationName

            If tParent.Adress = BOM_Checked.Item(PartID).Patch Then
                ChildID = PartID
                Debug.Print(Now.ToLongTimeString & " -         " & "- Composant parent similaire trouvé")
            End If

            Debug.Print(Now.ToLongTimeString & " -         " & "- Recherche d'une configuration du composant parent identique")
            If tVerif = ttVerif Then
                NewParent = False
                Debug.Print(Now.ToLongTimeString & " -         " & "- Configuration du composant parent similaire trouvé")
                Parent.Quantity += 1
                Debug.Print(Now.ToLongTimeString & " -         " & "    - Incrémentation de la quantité dans le parent")
                Exit For
            End If

        Next

        If NewParent Then
            Debug.Print(Now.ToLongTimeString & " -         " & "    - Composant parent similaire trouvé non trouvé")
            Debug.Print(Now.ToLongTimeString & " -         " & "    - Création du parent")
            BOM_Checked.Item(PartID).Configuration.Item(ConfigID).Parent.Add(New Parents With {.Adress = tParent.Adress, _
                                                                              .ConfigurationName = tParent.ConfigurationName, _
                                                                              .Quantity = 1})

            Debug.Print(Now.ToLongTimeString & " -         " & "        - Adresse : '" & BOM_Checked.Item(PartID).Configuration.Item(ConfigID).Parent.Last.Adress & "'")
            Debug.Print(Now.ToLongTimeString & " -         " & "        - Configuration : '" & BOM_Checked.Item(PartID).Configuration.Item(ConfigID).Parent.Last.ConfigurationName & "'")
        Else

        End If

        Debug.Print(Now.ToLongTimeString & " -     " & "FIN vérification du parent renseigné")
    End Sub

#End Region

End Class
