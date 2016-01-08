Public Class Components

    Public swModel As SldWorks.ModelDoc2   'swModel du composant

    Public swConfMgr As SldWorks.ConfigurationManager

    Public swConf As SldWorks.Configuration  'swConf du composant

    Public swRootComp As SldWorks.Component2  'swRootComp du composant

    Public vComponent As Object  'vComponent du composant

    Public swModelDocExt As SldWorks.ModelDocExtension  'swModelDocExt du composant

    Public swCustProp As SldWorks.CustomPropertyManager  'swCustProp du composant

    Public swselmgr As SldWorks.SelectionMgr  'swselmgr du composant

    Public swBodyFolder As SldWorks.BodyFolder  'swBodyFolder du composant

    Public swCustPropMgr As SldWorks.CustomPropertyManager  'CustomPropertyManager du composant

    Public I As Integer 'Index du vcomponent

    Public swComponent As SldWorks.Component2  'swComponent du composant

    Public LastConfiguration As String   'nom de la configuration active du composant

    Public Configuration As New List(Of Configurations)   'Configuration du composant

    Public Patch As String   'répertoire du composant

    Public Type As Integer  'type du composant (PRT ou ASM)

    Public Quantity As Integer  'Quantité du composant (toutes config)

    Public Properties As New List(Of Properties)   'Propriétés du composant

    Public Children As New List(Of Children)   'Enfants dans le composant

    Public Suppressed As Boolean

End Class
