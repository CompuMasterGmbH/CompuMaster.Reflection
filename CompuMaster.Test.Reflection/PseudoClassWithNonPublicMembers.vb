Public Class PseudoClassWithNonPublicMembers

#Region "Private constructor and instance members"
    Private Sub New()
        Me.PrivateReadOnlyField = "This is private (instance)"
    End Sub

    Private ReadOnly PrivateReadOnlyField As String

    Private PrivateField As String = "This is private (instance)"

    Private Property PrivateProperty As String = "This is private (instance)"

    Private Function PrivateFunction() As String
        Return "This is private (instance)"
    End Function

    Private _ValueToBeChangedByPrivateMethodOnly As String
    Private ReadOnly Property ValueToBeChangedByPrivateMethodOnly As String
        Get
            Return _ValueToBeChangedByPrivateMethodOnly
        End Get
    End Property

    Private Sub PrivateMethod()
        _ValueToBeChangedByPrivateMethodOnly = "This is private (instance)"
    End Sub

    Private Event PrivateEvent()
#End Region

#Region "Private static members"
    Private Shared PrivateStaticField As String = "This is private (static)"

    Private Shared Property PrivateStaticProperty As String = "This is private (static)"

    Private Shared Function PrivateStaticFunction() As String
        Return "This is private (static)"
    End Function

    Private Shared _StaticValueToBeChangedByPrivateMethodOnly As String
    Private Shared ReadOnly Property StaticValueToBeChangedByPrivateMethodOnly As String
        Get
            Return _StaticValueToBeChangedByPrivateMethodOnly
        End Get
    End Property

    Private Shared Sub PrivateStaticMethod()
        _StaticValueToBeChangedByPrivateMethodOnly = "This is private (static)"
    End Sub

    Private Shared Event PrivateStaticEvent()
#End Region

End Class
