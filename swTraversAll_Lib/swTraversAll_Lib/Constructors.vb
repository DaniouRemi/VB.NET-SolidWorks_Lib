Option Explicit On

Imports SldWorks
Imports SwConst
Imports System.IO
Imports System.Threading

Public Class SolidWorksComponent
    Inherits TraversAllVariables

    Private Checks As New Checks
    Private TraverseAll As New TraverseAll
    Private PropertiesOptions As New PropertiesOptions
    Private OptionsManager As New OptionsManager

    ''' <summary>
    ''' Initialise le Document ouvert
    ''' </summary>
    ''' <remarks></remarks>
    Sub New()
        Debug.Print("*********************************************************************************************")
        Debug.Print(Now.ToLongTimeString & " - " & "new()")

        Checks.SW_Open()

        Try
            swApp = New SldWorks.SldWorks
            swModel = swApp.ActiveDoc
            If IsNothing(swModel) Then
                Debug.Print(Now.ToLongTimeString & " - " & "no file open")
                Checks.endtest()

            ElseIf swModel.GetType = swDocumentTypes_e.swDocASSEMBLY Or swModel.GetType = swDocumentTypes_e.swDocPART Then
            Else
                Debug.Print(Now.ToLongTimeString & " - " & "no file compatible")
            End If
        Catch ex As Exception
            Debug.Print(Now.ToLongTimeString & " - " & " ERROR : " & ex.ToString)
            Checks.endtest()
        End Try

        Checks.Check(swModel)



    End Sub

    ''' <summary>
    ''' Initialise le document ciblé et l'ouvre
    ''' </summary>
    ''' <param name="Doc"></param>
    ''' <remarks></remarks>
    Sub New(Doc As String)
        Debug.Print(Now.ToLongTimeString & " - " & "New(ByVal Doc As String) with 'Doc' = " & Doc)

        Try
            If System.IO.File.Exists(Doc) Then
                If LCase(Right(Doc, 7)) = ".sldasm" Then
                    swApp = New SldWorks.SldWorks
                    swModel = swApp.OpenDoc(Doc, swDocumentTypes_e.swDocASSEMBLY)
                ElseIf LCase(Right(Doc, 7)) = ".sldprt" Then
                    swApp = New SldWorks.SldWorks
                    swModel = swApp.OpenDoc(Doc, swDocumentTypes_e.swDocPART)
                Else
                    Debug.Print(Now.ToLongTimeString & " - " & "Incompatible file")
                    Checks.endtest()
                End If
                Checks.Check(swModel)
            Else
                Debug.Print(Now.ToLongTimeString & " - " & "no file exist")
                Checks.endtest()
            End If

        Catch ex As Exception
            Debug.Print(Now.ToLongTimeString & " - " & "ERROR : " & ex.ToString)
            Checks.endtest()
        End Try


    End Sub

    'Sub New(ByVal swModel As SldWorks.ModelDoc2)
    '    Report(TypeReport.Information, "New(ByVal swModel As SldWorks.ModelDoc2) with 'swModel' = " & swModel.GetPathName)
    '    Check(swModel)
    'End Sub

    'Sub New(ByVal swPart As SldWorks.PartDoc)
    '    Report(TypeReport.Information, "New(ByVal swModel As SldWorks.ModelDoc2) with 'swModel' = " & swModel.GetPathName)
    '    Check(swModel)
    'End Sub

    ''' <summary>
    ''' Traverse le document
    ''' </summary>
    ''' <param name="Options">défini la profondeur de traversée :
    ''' components (comp) :
    '''     - corps (bod)
    '''     - propriétés (ppt)
    '''     - configurations (conf)
    '''     - Fonctions (fct)
    ''' 
    ''' Corps(bod) :
    ''' propriétés(ppt)
    ''' 
    ''' fonctions (fct)
    ''' 
    ''' ^	-> Tout
    ''' ^Name() -> récupère que ce qu'il y a entre paranthèses
    ''' Name() -> récupère tout sauf ce qu'il y a entre paranthèses
    ''' ^comp(ppt)		->	dans les composants, récupère les propriétés (que les propriétés)
    ''' comp(ppt)		->	dans les composants, récupère tout sauf les propriétés
    ''' comp(ppt,bod)	->	dans les composants, récupère tout sauf les corps et les propriétés
    ''' comp(ppt,^bod)	->	dans les composants, récupère tout sauf les corps et les propriétés
    ''' ^comp(bod);bod(ppt)	->	dans les composants, récupère que les corps; récupère tout dans les corps sauf les propriétés</param>
    ''' <param name="MultiThread">le TraversAll se fait dans un thread différent de celui avec lequel il a été appelé</param>
    ''' <remarks></remarks>
    Sub Traverse(Optional Options As String = Nothing, Optional MultiThread As Boolean = False)

        PropertiesOptions = OptionsManager.Initialize(Options)

        TraverseAll = New TraverseAll With {.swApp = swApp, _
                                         .swModel = swModel, _
                                         .ID = -1}
        If MultiThread Then
            Dim tThread As New Thread(AddressOf TraverseAll.TraverseASM)
            tThread.Start()

        Else
            TraverseAll.TraverseASM(PropertiesOptions)
            TraverseAll.ASM_BOM.RemoveAt(TraverseAll.ASM_BOM.Count - 1)
            TraverseAll.Quantity()
        End If
        ASM_BOM = TraverseAll.ASM_BOM

    End Sub

End Class



