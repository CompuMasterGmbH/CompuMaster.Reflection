Option Explicit On
Option Strict On

Namespace CompuMaster.Reflection

    ''' <summary>
    ''' Provides simplified InvokeMember for public members of an instance object
    ''' </summary>
    Public NotInheritable Class NonPublicInstanceMembers

        Private Const MEMBERFILTER As System.Reflection.BindingFlags = System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.NonPublic
        Private Const EXCEPTION_FILTERINFO As String = "non-public instance"

        Public Shared Function InvokeGetMembers(Of TSearchedMembers As System.Reflection.MemberInfo)(objtype As Type, expectedType As Type) As List(Of TSearchedMembers)
            Return ReflectionWorker.InvokeGetMembers(Of TSearchedMembers)(MEMBERFILTER, EXCEPTION_FILTERINFO, objtype, expectedType)
        End Function

        Public Shared Function InvokeGetMembers(objtype As Type) As System.Reflection.MemberInfo()
            Return ReflectionWorker.InvokeGetMembers(MEMBERFILTER, EXCEPTION_FILTERINFO, objtype)
        End Function

        Public Shared Function InvokeFunction(Of T)(obj As Object, objType As Type, name As String, ParamArray values As Object()) As T
            Return ReflectionWorker.InvokeFunction(Of T)(MEMBERFILTER, EXCEPTION_FILTERINFO, obj, objType, name, values)
        End Function

        Public Shared Sub InvokeMethod(obj As Object, objType As Type, name As String, ParamArray values As Object())
            ReflectionWorker.InvokeMethod(MEMBERFILTER, EXCEPTION_FILTERINFO, obj, objType, name, values)
        End Sub

        Public Shared Function InvokePropertyGet(Of T)(obj As Object, objtype As Type, name As String) As T
            Return ReflectionWorker.InvokePropertyGet(Of T)(MEMBERFILTER, EXCEPTION_FILTERINFO, obj, objtype, name)
        End Function

        Public Shared Function InvokePropertyGet(Of T)(obj As Object, objtype As Type, name As String, propertyArrayItem As Object) As T
            Return ReflectionWorker.InvokePropertyGet(Of T)(MEMBERFILTER, EXCEPTION_FILTERINFO, obj, objtype, name, propertyArrayItem)
        End Function

        Public Shared Sub InvokePropertySet(Of T)(obj As Object, objType As Type, name As String, value As T)
            ReflectionWorker.InvokePropertySet(Of T)(MEMBERFILTER, EXCEPTION_FILTERINFO, obj, objType, name, value)
        End Sub

        Public Shared Sub InvokePropertySet(Of T)(obj As Object, objType As Type, name As String, values As T())
            ReflectionWorker.InvokePropertySet(Of T)(MEMBERFILTER, EXCEPTION_FILTERINFO, obj, objType, name, values)
        End Sub

        Public Shared Function InvokeFieldGet(Of T)(obj As Object, objtype As Type, name As String) As T
            Return ReflectionWorker.InvokeFieldGet(Of T)(MEMBERFILTER, EXCEPTION_FILTERINFO, obj, objtype, name)
        End Function

        Public Shared Sub InvokeFieldSet(Of T)(obj As Object, objType As Type, name As String, value As T)
            ReflectionWorker.InvokeFieldSet(Of T)(MEMBERFILTER, EXCEPTION_FILTERINFO, obj, objType, name, value)
        End Sub

        Public Shared Sub InvokeFieldSet(Of T)(obj As Object, objType As Type, name As String, values As T())
            ReflectionWorker.InvokeFieldSet(Of T)(MEMBERFILTER, EXCEPTION_FILTERINFO, obj, objType, name, values)
        End Sub

        Public Shared Function InvokeConstructor(Of T)(ParamArray args As Object()) As T
            Dim M As System.Reflection.BindingFlags = MEMBERFILTER And Not System.Reflection.BindingFlags.Instance
            Return ReflectionWorker.InvokeConstructor(Of T)(MEMBERFILTER, EXCEPTION_FILTERINFO, args)
        End Function

    End Class

End Namespace