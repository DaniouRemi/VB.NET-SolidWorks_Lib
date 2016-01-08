Imports System.IO
Imports System.Diagnostics

Module TraversAllTest
    Private test As New SolidWorksComponent
    Sub main()

        test.Traverse("FE")
        'test.Traverse()
        writeLines()
        'TextFile()
        Console.Read()
    End Sub

    Const NeutralLine As String = "   │ "
    Const DerivLine As String = "   ├ "
    Const LastSub As String = "   └ "
    Const null As String = "     "
    Sub writeLines()

        Dim Y As Integer = 0
        Dim X As Integer = 0

        'test.BOM.RemoveAt(test.BOM.Count - 1)
        Debug.Print("Nombres composants : " & test.ASM_BOM.Count)

        '===== énoncer les composants
        For Each Component As Components In test.ASM_BOM

            Console.WriteLine(DerivLine & Split(Component.Patch, "\")(Split(Component.Patch, "\").Count - 1))
            Console.WriteLine(NeutralLine & DerivLine & "Adresse : " & Component.Patch)
            Console.WriteLine(NeutralLine & DerivLine & "Type : " & Component.Type)
            'TODO la quantité n'est pas bonne si le parent est multiple
            Console.WriteLine(NeutralLine & DerivLine & "Quantité totale: " & Component.Quantity)

            'Debug.Print(Split(Component.Patch, "\")(Split(Component.Patch, "\").Count - 1))
            '===== énoncer des enfants
            Dim U As Integer = 0
            Console.WriteLine(NeutralLine & DerivLine & "Children")
            For Each TChildren As Children In Component.Children
                Console.WriteLine(NeutralLine & NeutralLine & DerivLine & "Children " & U)
                Console.WriteLine(NeutralLine & NeutralLine & NeutralLine & DerivLine & " Adress = " & TChildren.Adress)

                Dim V As Integer = 0
                For Each TConfig As ChildrenConfigurations In TChildren.Configurations
                    Console.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & DerivLine & "Configuration " & V & " : " & TConfig.Name & " Qty : " & TConfig.Quantity)
                    V += 1
                Next
                U += 1

            Next
            'Console.WriteLine("")

            U = 0
            Console.WriteLine(NeutralLine & LastSub & "Properties Part: " & Component.Properties.Count)
            For Each TProperty As Properties In Component.Properties
                Console.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & DerivLine & "Property ID " & U)
                Console.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & DerivLine & " Name = " & TProperty.Name)
                Console.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & DerivLine & " Expression = " & TProperty.textexp)
                Console.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & DerivLine & " Value = " & TProperty.Value)
                U += 1
            Next
            'Console.WriteLine("")

            '===== énoncer les configurations
            For Each Configuration As Configurations In Component.Configuration

                Console.WriteLine(NeutralLine & DerivLine & "Configuration : " & Configuration.Name)
                Console.WriteLine(NeutralLine & DerivLine & "Quantité : " & Configuration.Quantity)

                If Not IsNothing(Configuration.Properties) Then
                    Console.WriteLine(NeutralLine & LastSub & "Properties configuration : " & Configuration.Properties.Count)

                    '===== énoncer les propriétés de la configuration
                    Dim V As Integer = 0
                    For Each Properties As Properties In Configuration.Properties
                        Console.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & DerivLine & "Property ID " & V)
                        Console.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & DerivLine & " Name = " & Properties.Name)
                        Console.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & DerivLine & " Expression = " & Properties.textexp)
                        Console.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & DerivLine & " Value = " & Properties.Value)
                        V += 1
                    Next
                End If

                Console.WriteLine(NeutralLine & DerivLine & "Parents : " & Configuration.Parent.Count)
                Dim W As Integer = 0
                For Each Parent As Parents In Configuration.Parent
                    Console.WriteLine(NeutralLine & NeutralLine & DerivLine & "Parents n°" & W & " :")
                    Console.WriteLine(NeutralLine & NeutralLine & NeutralLine & DerivLine & "Adress : " & Configuration.Parent.Item(W).Adress)
                    Console.WriteLine(NeutralLine & NeutralLine & NeutralLine & DerivLine & "Configuration : " & Configuration.Parent.Item(W).ConfigurationName)
                    Console.WriteLine(NeutralLine & NeutralLine & NeutralLine & DerivLine & "Quantité : " & Configuration.Parent.Item(W).Quantity)
                    W += 1
                Next


                If Not IsNothing(Configuration.Bodies) Then
                    'Console.WriteLine("")
                    Console.WriteLine(NeutralLine & LastSub & "Corps : ")

                    '===== énoncer les corps dans configuration
                    W = 0
                    For Each Body As Body In Configuration.Bodies.Body
                        'Console.WriteLine("")
                        Console.WriteLine(NeutralLine & null & DerivLine & "Corps ID " & W & " :")
                        Console.WriteLine(NeutralLine & null & NeutralLine & DerivLine & "Nom : " & Body.Name)
                        Console.WriteLine(NeutralLine & null & NeutralLine & DerivLine & "Quantité : " & Body.Count)
                        Console.WriteLine(NeutralLine & null & NeutralLine & DerivLine & "Properties : " & Body.Properties.Count)
                        '===== énoncer les propriértés du corps
                        Dim Z As Integer = 0
                        For Each Properties As Properties In Body.Properties
                            Console.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & LastSub & _
                                              "Propriété ID " & Z & " :")
                            Console.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & LastSub & _
                                             "Nom : " & Properties.Name)
                            Console.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & LastSub & _
                                             "Expression : " & Properties.textexp)
                            Console.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & LastSub & _
                                             "Valeur : " & Properties.Value)
                            Z += 1
                        Next
                        W += 1
                    Next
                End If
                'Console.WriteLine("")
            Next
            'Console.WriteLine(NeutralLine)
            'Console.WriteLine("=============================================================================================")
            Y += 1
        Next

        Console.WriteLine("Nombre de lignes : " & Y)

    End Sub

    Sub TextFile()

        Dim Text As New StreamWriter("F:\Programmation\Bibliothèques de classes\SolidWorks_Lib\SolidWorks_Lib\test.txt")
        Dim Y As Integer = 0
        Dim X As Integer = 0

        test.ASM_BOM.RemoveAt(test.ASM_BOM.Count - 1)

        '===== énoncer les composants
        For Each Component As Components In test.ASM_BOM

            Text.WriteLine(DerivLine & Split(Component.Patch, "\")(Split(Component.Patch, "\").Count - 1))
            Text.WriteLine(NeutralLine & DerivLine & "Adresse : " & Component.Patch)
            Text.WriteLine(NeutralLine & DerivLine & "Type : " & Component.Type)
            'TODO la quantité n'est pas bonne si le parent est multiple
            Text.WriteLine(NeutralLine & DerivLine & "Quantité totale: " & Component.Quantity)

            '===== énoncer des enfants
            Dim U As Integer = 0
            Text.WriteLine(NeutralLine & DerivLine & "Children")
            For Each TChildren As Children In Component.Children
                Text.WriteLine(NeutralLine & NeutralLine & DerivLine & "Children " & U)
                Text.WriteLine(NeutralLine & NeutralLine & NeutralLine & DerivLine & " Adress = " & TChildren.Adress)

                Dim V As Integer = 0
                For Each TConfig As ChildrenConfigurations In TChildren.Configurations
                    Text.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & DerivLine & "Configuration " & V & " : " & TConfig.Name & " Qty : " & TConfig.Quantity)
                    V += 1
                Next
                U += 1

            Next
            'Text.WriteLine("")

            U = 0
            Text.WriteLine(NeutralLine & LastSub & "Properties Part: " & Component.Properties.Count)
            For Each TProperty As Properties In Component.Properties
                Text.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & DerivLine & "Property ID " & U)
                Text.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & DerivLine & " Name = " & TProperty.Name)
                Text.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & DerivLine & " Expression = " & TProperty.textexp)
                Text.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & DerivLine & " Value = " & TProperty.Value)
                U += 1
            Next
            'Text.WriteLine("")

            '===== énoncer les configurations
            For Each Configuration As Configurations In Component.Configuration

                Text.WriteLine(NeutralLine & DerivLine & "Configuration : " & Configuration.Name)
                Text.WriteLine(NeutralLine & DerivLine & "Quantité : " & Configuration.Quantity)

                If Not IsNothing(Configuration.Properties) Then
                    Text.WriteLine(NeutralLine & LastSub & "Properties configuration : " & Configuration.Properties.Count)

                    '===== énoncer les propriétés de la configuration
                    Dim V As Integer = 0
                    For Each Properties As Properties In Configuration.Properties
                        Text.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & DerivLine & "Property ID " & V)
                        Text.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & DerivLine & " Name = " & Properties.Name)
                        Text.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & DerivLine & " Expression = " & Properties.textexp)
                        Text.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & DerivLine & " Value = " & Properties.Value)
                        V += 1
                    Next
                End If

                Text.WriteLine(NeutralLine & DerivLine & "Parents : " & Configuration.Parent.Count)
                Dim W As Integer = 0
                For Each Parent As Parents In Configuration.Parent
                    Text.WriteLine(NeutralLine & NeutralLine & DerivLine & "Parents n°" & W & " :")
                    Text.WriteLine(NeutralLine & NeutralLine & NeutralLine & DerivLine & "Adress : " & Configuration.Parent.Item(W).Adress)
                    Text.WriteLine(NeutralLine & NeutralLine & NeutralLine & DerivLine & "Configuration : " & Configuration.Parent.Item(W).ConfigurationName)
                    Text.WriteLine(NeutralLine & NeutralLine & NeutralLine & DerivLine & "Quantité : " & Configuration.Parent.Item(W).Quantity)
                    W += 1
                Next


                If Not IsNothing(Configuration.Bodies) Then
                    'Text.WriteLine("")
                    Text.WriteLine(NeutralLine & LastSub & "Corps : ")

                    '===== énoncer les corps dans configuration
                    W = 0
                    For Each Body As Body In Configuration.Bodies.Body
                        'Text.WriteLine("")
                        Text.WriteLine(NeutralLine & null & DerivLine & "Corps ID " & W & " :")
                        Text.WriteLine(NeutralLine & null & NeutralLine & DerivLine & "Nom : " & Body.Name)
                        Text.WriteLine(NeutralLine & null & NeutralLine & DerivLine & "Quantité : " & Body.Count)
                        Text.WriteLine(NeutralLine & null & NeutralLine & DerivLine & "Properties : " & Body.Properties.Count)
                        '===== énoncer les propriértés du corps
                        Dim Z As Integer = 0
                        For Each Properties As Properties In Body.Properties
                            Text.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & LastSub & _
                                              "Propriété ID " & Z & " :")
                            Text.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & LastSub & _
                                             "Nom : " & Properties.Name)
                            Text.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & LastSub & _
                                             "Expression : " & Properties.textexp)
                            Text.WriteLine(NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & NeutralLine & LastSub & _
                                             "Valeur : " & Properties.Value)
                            Z += 1
                        Next
                        W += 1
                    Next
                End If
                'Text.WriteLine("")
            Next
            'Text.WriteLine(NeutralLine)
            'Text.WriteLine("=============================================================================================")
            Y += 1
        Next

        Text.WriteLine("Nombre de lignes : " & Y)

    End Sub
End Module
