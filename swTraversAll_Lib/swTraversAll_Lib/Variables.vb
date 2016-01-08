'Imports Microsoft.VisualStudio.TestTools.UnitTesting

Public Class TraversAllVariables
    Inherits SolidWorksVar


    Public ASM_BOM As New List(Of Components)
    Public PRT_BOM As New Components
    Friend Report As New List(Of Report)
    Public ID As Integer

End Class

Public Class SolidWorksVar

    Protected Friend swApp As SldWorks.SldWorks
    Protected Friend swModel As SldWorks.ModelDoc2
    Protected SubFeat_Obj As SldWorks.Feature
    Protected Feat_Obj As SldWorks.Feature
    Protected SubSubFeat_Obj As SldWorks.Feature
    Protected nErrors As Long
    Protected nWarnings As Long

End Class
