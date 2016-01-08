Public Class Configurations

    Public Name As String   'Nom de la configuration

    Public Parent As New List(Of Parents)   'Parent de la configuration

    Public Children As New List(Of Children)   'Enfants dans la configuration

    Public Bodies As Bodies   'Corps du composant

    Public Properties As New List(Of Properties)   'Propriétés du composant

    Public Quantity As String   'Quantité (combien de fois on a rencontré cette configuration)

    Public Functions As New List(Of Functions)

End Class
