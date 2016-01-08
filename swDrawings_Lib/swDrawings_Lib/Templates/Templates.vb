Imports System.IO
Imports FE_Library

Public Class Templates
    Inherits Variables

    Private TemplateFile As String = My.Application.Info.DirectoryPath & "\SolidWorks_Lib\Templates.ini"
    Private ScalesProperties As String = My.Application.Info.DirectoryPath & "\SolidWorks_Lib\ScalesProperties.ini"

    Sub New()

        Templates_p = New List(Of Templates_p)

        Dim _Name As String = Nothing
        Dim _Adress As String = Nothing
        Dim _Orientation As String = Nothing
        Dim _Size As String = Nothing

        For Each lign As String In File.ReadAllLines(TemplateFile)

            Select Case Split(lign, "=")(0)
                Case "Name"
                    _Name = Split(lign, "=")(1)
                Case "Adress"
                    _Adress = Split(lign, "=")(1)
                Case "Orientation"
                    _Orientation = Split(lign, "=")(1)
                Case "Size"
                    _Size = Split(lign, "=")(1)
            End Select

            If Not IsNothing(_Name) And Not IsNothing(_Adress) And Not IsNothing(_Orientation) And Not IsNothing(_Size) Then
                Templates_p.Add(New Templates_p With {.Name = _Name, _
                                                    .Adress = _Adress, _
                                                    .Orientation = _Orientation, _
                                                    .Size = _Size, _
                                                      .ScaleMode = InitializeScaleModes(_Name, _Orientation)})

                _Name = Nothing
                _Adress = Nothing
                _Orientation = Nothing
                _Size = Nothing

            End If

        Next

    End Sub

    Private Function InitializeScaleModes(ByVal _Name As String, ByVal _Orientation As String) As List(Of ScaleMode)
        Dim ScaleModeList As New List(Of ScaleMode)
        Dim PartSizeColumn As Integer
        Dim ViewZoneColumn As Integer
        Dim BeginRead As Boolean = False
        Dim Insert As Boolean

        For Each Lign As String In File.ReadAllLines(ScalesProperties)
            Insert = False
            If BeginRead Then

                Select Case _Name
                    Case "A4H"
                        PartSizeColumn = 1
                        ViewZoneColumn = 7
                        Insert = True
                    Case "A4V"
                        PartSizeColumn = 2
                        ViewZoneColumn = 8
                        Insert = True
                    Case "A3H"
                        PartSizeColumn = 3
                        ViewZoneColumn = 9
                        Insert = True
                    Case "A3V"
                        PartSizeColumn = 4
                        ViewZoneColumn = 10
                        Insert = True
                    Case "A2H"
                        PartSizeColumn = 5
                        ViewZoneColumn = 11
                        Insert = True
                    Case "A2V"
                        PartSizeColumn = 6
                        ViewZoneColumn = 12
                        Insert = True
                    Case Else

                        Insert = False
                End Select

                If Insert Then

                    Debug.Print(Split(Lign, "|")(PartSizeColumn))
                    If IsNumeric(Numerical.ConvertToNumerical(Replace(Trim(Split(Lign, "|")(PartSizeColumn)), Chr(9), ""))) Then
                        Debug.Print(Numerical.ConvertToNumerical(Replace(Trim(Split(Lign, "|")(PartSizeColumn)), Chr(9), "")))
                    End If

                    Debug.Print(Numerical.ConvertToNumerical(Replace(Trim(Split(Lign, "|")(PartSizeColumn)), Chr(9), "")))
                    Debug.Print(Numerical.ConvertToNumerical(Replace(Trim(Split(Lign, "|")(ViewZoneColumn)), Chr(9), "")))

                    ScaleModeList.Add(New ScaleMode With {.PartSize = Numerical.ConvertToNumerical(Replace(Trim(Split(Lign, "|")(PartSizeColumn)), Chr(9), "")), _
                                                          .ViewZone = Numerical.ConvertToNumerical(Replace(Trim(Split(Lign, "|")(ViewZoneColumn)), Chr(9), ""))})

                End If

            End If

            If Lign.Contains("******") Then
                BeginRead = True
            End If

        Next

        Return ScaleModeList

    End Function

End Class
