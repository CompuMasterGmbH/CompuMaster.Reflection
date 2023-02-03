Imports NUnit.Framework
Imports CompuMaster.Reflection

<NonParallelizable>
Public Class ReflectionTest

    <Test>
    Public Sub TestGetMembers()
        Assert.Throws(Of NotSupportedException)(Sub() PublicInstanceMembers.GetMembers(Of System.Reflection.MemberInfo)(GetType(Object), GetType(String)))

        Dim mp As List(Of System.Reflection.PropertyInfo)
        mp = PublicInstanceMembers.GetMembers(Of System.Reflection.PropertyInfo)(GetType(Object), GetType(String))
        Assert.AreEqual(0, mp.Count)

        Dim mf As List(Of System.Reflection.MethodInfo)
        mf = PublicInstanceMembers.GetMembers(Of System.Reflection.MethodInfo)(GetType(Object), GetType(String))
        Assert.AreEqual(1, mf.Count)
        Assert.AreEqual("ToString", mf(0).Name)
    End Sub

    <Test>
    Public Sub TestPublicInstanceMembers()
        Const ExpectedString As String = "This is public (instance)"

        Dim Instance As Object = PublicInstanceMembers.InvokeConstructor(Of PseudoClassWithPublicMembers)

        Assert.AreEqual(ExpectedString, PublicInstanceMembers.InvokeFieldGet(Of String)(Instance, Instance.GetType, "PublicReadOnlyField"))
        Assert.AreEqual(ExpectedString, PublicInstanceMembers.InvokeFieldGet(Of String)(Instance, Instance.GetType, "PublicField"))

        Assert.AreEqual(ExpectedString, PublicInstanceMembers.InvokePropertyGet(Of String)(Instance, Instance.GetType, "PublicProperty"))

        Assert.AreEqual(ExpectedString, PublicInstanceMembers.InvokeFunction(Of String)(Instance, Instance.GetType, "PublicFunction"))

        Assert.AreEqual(Nothing, PublicInstanceMembers.InvokePropertyGet(Of String)(Instance, Instance.GetType, "ValueToBeChangedByPublicMethodOnly"))
        PublicInstanceMembers.InvokeMethod(Instance, Instance.GetType, "PublicMethod", Nothing)
        Assert.AreEqual(ExpectedString, PublicInstanceMembers.InvokePropertyGet(Of String)(Instance, Instance.GetType, "ValueToBeChangedByPublicMethodOnly"))
    End Sub

    <Test>
    Public Sub TestPrivateInstanceMembers()
        Const ExpectedString As String = "This is private (instance)"
        'Dim Instance As Object = 

        Dim Instance As Object = NonPublicInstanceMembers.InvokeConstructor(Of PseudoClassWithNonPublicMembers)

        Assert.AreEqual(ExpectedString, NonPublicInstanceMembers.InvokeFieldGet(Of String)(Instance, Instance.GetType, "PrivateReadOnlyField"))
        Assert.AreEqual(ExpectedString, NonPublicInstanceMembers.InvokeFieldGet(Of String)(Instance, Instance.GetType, "PrivateField"))

        Assert.AreEqual(ExpectedString, NonPublicInstanceMembers.InvokePropertyGet(Of String)(Instance, Instance.GetType, "PrivateProperty"))

        Assert.AreEqual(ExpectedString, NonPublicInstanceMembers.InvokeFunction(Of String)(Instance, Instance.GetType, "PrivateFunction"))

        Assert.AreEqual(Nothing, NonPublicInstanceMembers.InvokeFieldGet(Of String)(Instance, Instance.GetType, "_ValueToBeChangedByPrivateMethodOnly"))
        Assert.AreEqual(Nothing, NonPublicInstanceMembers.InvokePropertyGet(Of String)(Instance, Instance.GetType, "ValueToBeChangedByPrivateMethodOnly"))
        NonPublicInstanceMembers.InvokeMethod(Instance, Instance.GetType, "PrivateMethod", Nothing)
        Assert.AreEqual(ExpectedString, NonPublicInstanceMembers.InvokeFieldGet(Of String)(Instance, Instance.GetType, "_ValueToBeChangedByPrivateMethodOnly"))
        Assert.AreEqual(ExpectedString, NonPublicInstanceMembers.InvokePropertyGet(Of String)(Instance, Instance.GetType, "ValueToBeChangedByPrivateMethodOnly"))
    End Sub

    <Test>
    Public Sub TestPublicStaticMembers()
        Const ExpectedString As String = "This is public (static)"
        Dim StaticType As Type = GetType(PseudoClassWithPublicMembers)

        Assert.AreEqual(ExpectedString, PublicStaticMembers.InvokeFieldGet(Of String)(StaticType, "PublicStaticField"))

        Assert.AreEqual(ExpectedString, PublicStaticMembers.InvokePropertyGet(Of String)(StaticType, "PublicStaticProperty"))

        Assert.AreEqual(ExpectedString, PublicStaticMembers.InvokeFunction(Of String)(StaticType, "PublicStaticFunction"))

        Assert.AreEqual(Nothing, PublicStaticMembers.InvokePropertyGet(Of String)(StaticType, "StaticValueToBeChangedByPublicMethodOnly"))
        PublicStaticMembers.InvokeMethod(StaticType, "PublicStaticMethod", Nothing)
        Assert.AreEqual(ExpectedString, PublicStaticMembers.InvokePropertyGet(Of String)(StaticType, "StaticValueToBeChangedByPublicMethodOnly"))
    End Sub

    <Test>
    Public Sub TestPrivateStaticMembers()
        Const ExpectedString As String = "This is private (static)"
        Dim StaticType As Type = GetType(PseudoClassWithNonPublicMembers)

        Assert.AreEqual(ExpectedString, NonPublicStaticMembers.InvokeFieldGet(Of String)(StaticType, "PrivateStaticField"))

        Assert.AreEqual(ExpectedString, NonPublicStaticMembers.InvokePropertyGet(Of String)(StaticType, "PrivateStaticProperty"))

        Assert.AreEqual(ExpectedString, NonPublicStaticMembers.InvokeFunction(Of String)(StaticType, "PrivateStaticFunction"))

        Assert.AreEqual(Nothing, NonPublicStaticMembers.InvokeFieldGet(Of String)(StaticType, "_StaticValueToBeChangedByPrivateMethodOnly"))
        Assert.AreEqual(Nothing, NonPublicStaticMembers.InvokePropertyGet(Of String)(StaticType, "StaticValueToBeChangedByPrivateMethodOnly"))
        NonPublicStaticMembers.InvokeMethod(StaticType, "PrivateStaticMethod", Nothing)
        Assert.AreEqual(ExpectedString, NonPublicStaticMembers.InvokeFieldGet(Of String)(StaticType, "_StaticValueToBeChangedByPrivateMethodOnly"))
        Assert.AreEqual(ExpectedString, NonPublicStaticMembers.InvokePropertyGet(Of String)(StaticType, "StaticValueToBeChangedByPrivateMethodOnly"))
    End Sub

End Class
