Imports System.Text.RegularExpressions

Module Module1
    Dim AllOptions As New List(Of String)
    Dim ArrangeRegex As New List(Of String)

    Sub Main()

        AllOptions.Add("comp")
        AllOptions.Add("bod")
        AllOptions.Add("ppt")
        AllOptions.Add("conf")

        Dim Options As String = "comp(bod,^ppt,^conf);bod(ppt);conf(ppt)"
        Dim m As Match
        Dim OPTline As String
        Dim OPTDomain As String
        Dim RegexOPT, RegexOPT1 As String

        For I As Integer = 0 To Split(Options, ";").Count - 1
            OPTline = Split(Options, ";")(I)
            RegexOPT = "\w*"

            If Regex.IsMatch(OPTline, RegexOPT) Then

                m = Regex.Match(OPTline, RegexOPT)
                Dim m_Matches As MatchCollection = Regex.Matches(OPTline, RegexOPT)

                If m.Success Then

                    For Each _M As Match In m_Matches
                        If _M.Value <> "" Then
                            ArrangeRegex.Add(_M.Value)
                        End If

                    Next
                    OPTDomain = m_Matches.Item(0).Value 'élément concerné

                    Debug.Print("Élément concerné = " & OPTDomain)
                    RegexOPT1 = "\((.*)\)"

                    If Regex.IsMatch(OPTline, RegexOPT1) Then

                        m = Regex.Match(OPTline, RegexOPT1)
                        Dim m1_Matches As MatchCollection = Regex.Matches(OPTline, RegexOPT1)

                        If m.Success Then

                            SubSearch(m1_Matches.Item(0).Value)

                        End If
                    End If

                End If
                Debug.Print("")
            End If

        Next

        Debug.Print("***********************************************************************")
    End Sub


    Sub SubSearch(ByVal Parameters As String)
        Dim Parameter As String
        Dim Result As String = ""

        Parameters = Left(Parameters, Parameters.Length - 1)
        Parameters = Right(Parameters, Parameters.Length - 1)

        For I As Integer = 0 To Split(Parameters, ",").Count - 1
            Parameter = Split(Parameters, ",")(I).ToLower

            If Left(Parameter, 1) = "^" Then
                Result = "   Ne prend pas en compte "
            Else
                Result = "   Prend en compte "
            End If

            For Each J As String In AllOptions
                If Parameter.Contains(J.ToLower) Then
                    Result = Result & J
                    Debug.Print(Result)
                End If
            Next

        Next

    End Sub

End Module
