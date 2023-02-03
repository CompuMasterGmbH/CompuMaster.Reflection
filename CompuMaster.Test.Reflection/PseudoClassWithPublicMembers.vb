Public Class PseudoClassWithPublicMembers

#Region "Public constructor and instance members"
    Public Sub New()
        Me.PublicReadOnlyField = "This is public (instance)"
    End Sub

    Public ReadOnly PublicReadOnlyField As String

    Public PublicField As String = "This is public (instance)"

    Public Property PublicProperty As String = "This is public (instance)"

    Public Function PublicFunction() As String
        Return "This is public (instance)"
    End Function

    Private _ValueToBeChangedByPublicMethodOnly As String
    Public ReadOnly Property ValueToBeChangedByPublicMethodOnly As String
        Get
            Return _ValueToBeChangedByPublicMethodOnly
        End Get
    End Property

    Public Sub PublicMethod()
        _ValueToBeChangedByPublicMethodOnly = "This is public (instance)"
    End Sub

    Public Event PublicEvent()
#End Region

#Region "Public static members"
    Public Shared PublicStaticField As String = "This is public (static)"

    Public Shared Property PublicStaticProperty As String = "This is public (static)"

    Public Shared Function PublicStaticFunction() As String
        Return "This is public (static)"
    End Function

    Private Shared _StaticValueToBeChangedByPublicMethodOnly As String
    Public Shared ReadOnly Property StaticValueToBeChangedByPublicMethodOnly As String
        Get
            Return _StaticValueToBeChangedByPublicMethodOnly
        End Get
    End Property

    Public Shared Sub PublicStaticMethod()
        _StaticValueToBeChangedByPublicMethodOnly = "This is public (static)"
    End Sub

    Public Shared Event PublicStaticEvent()
#End Region

End Class
