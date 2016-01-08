Public Class PropertiesRecovery
    Inherits TraversAllVariables

    Private Checks As New Checks
    Friend _BOM As New List(Of Components)

    ''' <summary>
    ''' récupération des propriétés du composant
    ''' </summary>
    ''' <param name="BOM">la liste des composants</param>
    ''' <param name="ID">Index du composant dans BOM</param>
    ''' <remarks></remarks>
    Friend Sub PropertiesComponent(ByVal BOM As List(Of Components), ByVal ID As Integer)
        Debug.Print(Now.ToLongTimeString & " - " & "        DEBUT récupération des propriétés du composant")
        _BOM = BOM

        _BOM.Item(ID).swModelDocExt = _BOM.Item(ID).swModel.Extension
        _BOM.Item(ID).swCustProp = _BOM.Item(ID).swModelDocExt.CustomPropertyManager("")

        Dim Bool As Boolean
        Dim val As String = ""
        Dim Tab_Temp As String = ""

        Dim CustNames As Object
        CustNames = _BOM.Item(ID).swCustProp.GetNames

        Dim z As Integer = 0
        If Not IsNothing(CustNames) Then
            Debug.Print(Now.ToLongTimeString & " - " & "            - " & CustNames.length & " propriété(s) trouvée(s)")
            For Each vname As Object In CustNames

                Bool = _BOM.Item(ID).swCustProp.Get4(vname, False, val, Tab_Temp)
                _BOM.Item(ID).Properties.Add(New Properties With { _
                                                     .Name = vname, _
                                                     .textexp = val, _
                                                     .Value = Tab_Temp})

                Debug.Print(Now.ToLongTimeString & " - " & "            - Propriété index " & _BOM.Item(ID).Properties.Count - 1 & " :")
                Debug.Print(Now.ToLongTimeString & " - " & "                - Name : '" & _BOM.Item(ID).Properties.Last.Name & "'")
                Debug.Print(Now.ToLongTimeString & " - " & "                - textexp : '" & _BOM.Item(ID).Properties.Last.textexp & "'")
                Debug.Print(Now.ToLongTimeString & " - " & "                - Value : '" & _BOM.Item(ID).Properties.Last.Value & "'")

                z += 1
            Next
        Else
            Debug.Print(Now.ToLongTimeString & " - " & "            - Aucune propriété trouvée")
        End If

        Debug.Print(Now.ToLongTimeString & " - " & "        FIN récupération des propriétés du composant")
    End Sub

    ''' <summary>
    ''' Récupération des propriétés du coprs
    ''' </summary>
    ''' <param name="BOM">la liste des composants</param>
    ''' <param name="ID">Index du composant dans BOM</param>
    ''' <param name="SubFeat_Obj">Fonction contenant les corps</param>
    ''' <remarks></remarks>
    Friend Sub PropertiesBodies(ByVal BOM As List(Of Components), ByVal ID As Integer, ByVal SubFeat_Obj As SldWorks.Feature)
        Debug.Print(Now.ToLongTimeString & " - " & "                    DEBUT récupération des propriétés du corps")

        Dim name As Object
        Dim textexp As String
        Dim evalval As String

        _BOM = BOM

        If Not IsNothing(_BOM.Item(ID).swCustPropMgr.GetNames) Then

            'Try
            '    _BOM.Item(ID).Configuration.Last.Bodies.Names = _BOM.Item(ID).swCustPropMgr.GetNames
            'Catch ex As Exception
            _BOM.Item(ID).Configuration.Last.Bodies = New Bodies
            _BOM.Item(ID).Configuration.Last.Bodies.Names = _BOM.Item(ID).swCustPropMgr.GetNames
            'End Try

            _BOM.Item(ID).Configuration.Last.Bodies.Body.Add(New Body() With { _
                                        .Name = SubFeat_Obj.Name, _
                                        .Count = _BOM.Item(ID).swBodyFolder.GetBodyCount})
            Debug.Print(Now.ToLongTimeString & " - " & "                        - Corps ajouté à la liste")
            Debug.Print(Now.ToLongTimeString & " - " & "                            - Nom : '" & _BOM.Item(ID).Configuration.Last.Bodies.Body.Last.Name & "'")
            Debug.Print(Now.ToLongTimeString & " - " & "                            - Quantité : " & _BOM.Item(ID).Configuration.Last.Bodies.Body.Last.Count)

            textexp = ""
            evalval = ""
            Debug.Print(Now.ToLongTimeString & " - " & "                            - Propriétés : (" & _BOM.Item(ID).Configuration.Last.Bodies.Names.Length & " propriété(s))")

            For Each name In _BOM.Item(ID).Configuration.Last.Bodies.Names
                _BOM.Item(ID).swCustPropMgr.Get2(name, textexp, evalval)
                _BOM.Item(ID).Configuration.Last.Bodies.Body.Last.Properties.Add(New Properties With {.Name = name, _
                                                                                       .textexp = textexp, _
                                                                                       .Value = evalval})

                Debug.Print(Now.ToLongTimeString & " - " & "                                - propriété ID" & _BOM.Item(ID).Configuration.Last.Bodies.Body.Last.Properties.Count - 1 & " : ")
                Debug.Print(Now.ToLongTimeString & " - " & "                                    - Nom : '" & _BOM.Item(ID).Configuration.Last.Bodies.Body.Last.Properties.Last.Name & "'")
                Debug.Print(Now.ToLongTimeString & " - " & "                                    - Expression : '" & _BOM.Item(ID).Configuration.Last.Bodies.Body.Last.Properties.Last.textexp & "'")
                Debug.Print(Now.ToLongTimeString & " - " & "                                    - Valeur : '" & _BOM.Item(ID).Configuration.Last.Bodies.Body.Last.Properties.Last.Value & "'")

            Next
        End If

        Debug.Print(Now.ToLongTimeString & " - " & "                    DEBUT récupération des propriétés du corps")
    End Sub

    Friend Sub PropertiesBodies(PRT_BOM As Components, ByVal SubFeat_Obj As SldWorks.Feature)
        Debug.Print(Now.ToLongTimeString & " - " & "                    DEBUT récupération des propriétés du corps")

        Dim name As Object
        Dim textexp As String
        Dim evalval As String

        If Not IsNothing(PRT_BOM.swCustPropMgr.GetNames) Then

            'Try
            '    _BOM.Item(ID).Configuration.Last.Bodies.Names = _BOM.Item(ID).swCustPropMgr.GetNames
            'Catch ex As Exception
            PRT_BOM.Configuration.Last.Bodies = New Bodies
            PRT_BOM.Configuration.Last.Bodies.Names = PRT_BOM.swCustPropMgr.GetNames
            'End Try

            PRT_BOM.Configuration.Last.Bodies.Body.Add(New Body() With { _
                                        .Name = SubFeat_Obj.Name, _
                                        .Count = PRT_BOM.swBodyFolder.GetBodyCount})
            Debug.Print(Now.ToLongTimeString & " - " & "                        - Corps ajouté à la liste")
            Debug.Print(Now.ToLongTimeString & " - " & "                            - Nom : '" & PRT_BOM.Configuration.Last.Bodies.Body.Last.Name & "'")
            Debug.Print(Now.ToLongTimeString & " - " & "                            - Quantité : " & PRT_BOM.Configuration.Last.Bodies.Body.Last.Count)

            textexp = ""
            evalval = ""
            Debug.Print(Now.ToLongTimeString & " - " & "                            - Propriétés : (" & PRT_BOM.Configuration.Last.Bodies.Names.Length & " propriété(s))")

            For Each name In PRT_BOM.Configuration.Last.Bodies.Names
                PRT_BOM.swCustPropMgr.Get2(name, textexp, evalval)
                PRT_BOM.Configuration.Last.Bodies.Body.Last.Properties.Add(New Properties With {.Name = name, _
                                                                                       .textexp = textexp, _
                                                                                       .Value = evalval})

                Debug.Print(Now.ToLongTimeString & " - " & "                                - propriété ID" & PRT_BOM.Configuration.Last.Bodies.Body.Last.Properties.Count - 1 & " : ")
                Debug.Print(Now.ToLongTimeString & " - " & "                                    - Nom : '" & PRT_BOM.Configuration.Last.Bodies.Body.Last.Properties.Last.Name & "'")
                Debug.Print(Now.ToLongTimeString & " - " & "                                    - Expression : '" & PRT_BOM.Configuration.Last.Bodies.Body.Last.Properties.Last.textexp & "'")
                Debug.Print(Now.ToLongTimeString & " - " & "                                    - Valeur : '" & PRT_BOM.Configuration.Last.Bodies.Body.Last.Properties.Last.Value & "'")

            Next
        End If

        Debug.Print(Now.ToLongTimeString & " - " & "                    DEBUT récupération des propriétés du corps")
    End Sub

End Class
