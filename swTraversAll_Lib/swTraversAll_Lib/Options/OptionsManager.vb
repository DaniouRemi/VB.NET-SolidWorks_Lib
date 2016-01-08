Imports System.Text.RegularExpressions

Public Class OptionsManager

    Private bod As Boolean
    Private ppt As Boolean
    Private conf As Boolean
    Private fct As Boolean
    Private DefaultValue As Boolean

    Friend Function Initialize(Optional OptionsAsked As String = Nothing) As PropertiesOptions

        Dim PropertiesOptions As New PropertiesOptions

        If Not IsNothing(OptionsAsked) Then


            Dim m As Match
            Dim OPTline As String
            Dim RegexOPT, RegexOPT1 As String
            Dim OPTDomain As String = ""
            Dim Options As String() = Split(OptionsAsked, ";")

            RegexOPT = "\w*"
            RegexOPT1 = "\((.*)\)"

            For I As Integer = 0 To Options.Count - 1

                OPTline = Options(I)

                If Regex.IsMatch(OPTline, RegexOPT) Then

                    m = Regex.Match(OPTline, RegexOPT)
                    Dim m_Matches As MatchCollection = Regex.Matches(OPTline, RegexOPT)

                    If m.Success Then

                        For Each _M As Match In m_Matches
                            OPTDomain = _M.Value
                            Debug.Print(OPTDomain)

                            If OPTDomain.Contains("/") Then
                                bod = ppt = conf = fct = DefaultValue = True
                            Else
                                bod = ppt = conf = fct = DefaultValue = False
                            End If

                            If Regex.IsMatch(OPTline, RegexOPT1) Then
                                m = Regex.Match(OPTline, RegexOPT1)
                                Dim m1_Matches As MatchCollection = Regex.Matches(OPTline, RegexOPT1)

                                If m.Success Then
                                    SubParameters(m1_Matches.Item(0).Value)
                                End If
                            End If

                            Select Case OPTDomain.ToLower
                                Case "comp"
                                    With PropertiesOptions.comp
                                        .bod = bod
                                        .conf = conf
                                        .fct = fct
                                        .ppt = ppt
                                    End With

                                Case "bod"
                                    With PropertiesOptions.bod
                                        .ppt = ppt
                                    End With

                                Case "Fct"
                                    PropertiesOptions.fct = fct

                                Case Else
                            End Select

                        Next

                    End If
                    '    Debug.Print("")
                End If

            Next

            If OptionsAsked = "FE" Then 'si récupération FE
                PropertiesOptions.FEMode = True
            Else
                PropertiesOptions.FEMode = False
            End If

            Debug.Print("OPTIONS :")
            Debug.Print("   comp :")
            Debug.Print("       bod = " & PropertiesOptions.comp.bod)
            Debug.Print("       conf = " & PropertiesOptions.comp.conf)
            Debug.Print("       fct = " & PropertiesOptions.comp.fct)
            Debug.Print("       ppt = " & PropertiesOptions.comp.ppt)
            Debug.Print("   bod :")
            Debug.Print("       ppt = " & PropertiesOptions.bod.ppt)
            Debug.Print("   fct :")
            Debug.Print("       fct = " & PropertiesOptions.fct)

        Else

            With PropertiesOptions.comp
                .bod = True
                .conf = True
                .fct = True
                .ppt = True
            End With

            With PropertiesOptions.bod
                .ppt = True
            End With

            PropertiesOptions.fct = True

        End If

        Return PropertiesOptions
    End Function

    Private Sub SubParameters(Parameters As String)

        Parameters = Left(Parameters, Parameters.Length - 1)
        Parameters = Right(Parameters, Parameters.Length - 1)

        For J As Integer = 0 To Split(Parameters, ",").Count - 1

            Select Case Split(Parameters, ",")(J).ToLower

                Case "bod"
                    bod = Not DefaultValue
                Case "ppt"
                    ppt = Not DefaultValue
                Case "conf"
                    conf = Not DefaultValue
                Case "fct"
                    fct = Not DefaultValue
            End Select

        Next

    End Sub

End Class
