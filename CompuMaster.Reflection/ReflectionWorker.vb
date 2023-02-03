Option Explicit On
Option Strict On

Namespace CompuMaster.Reflection

    ''' <summary>
    ''' Provides simplified InvokeMember for public members of an instance object
    ''' </summary>
    Friend NotInheritable Class ReflectionWorker

        Public Shared Function GetMembers(Of TSearchedMembers As System.Reflection.MemberInfo)(memberFilter As System.Reflection.BindingFlags, exceptionFilterInfo As String, objtype As Type, expectedType As Type) As List(Of TSearchedMembers)
            Dim FoundMemberInfos As System.Reflection.MemberInfo() = GetMembers(memberFilter, exceptionFilterInfo, objtype)
            Dim Result As New List(Of TSearchedMembers)
            For Each Member In FoundMemberInfos
                Select Case GetType(TSearchedMembers)
                    Case GetType(System.Reflection.MemberInfo)
                        Throw New NotSupportedException("Filtering not supported for type MemberInfo, please use types derived from MemberInfo like FieldInfo, PropertyInfo, MethodInfo, etc. ")
                    Case GetType(System.Reflection.FieldInfo)
                        If Member.MemberType = System.Reflection.MemberTypes.Field AndAlso CType(Member, System.Reflection.FieldInfo).FieldType Is expectedType Then
                            Result.Add(CType(Member, TSearchedMembers))
                        End If
                    Case GetType(System.Reflection.PropertyInfo)
                        If Member.MemberType = System.Reflection.MemberTypes.Property AndAlso CType(Member, System.Reflection.PropertyInfo).PropertyType Is expectedType Then
                            Result.Add(CType(Member, TSearchedMembers))
                        End If
                    Case GetType(System.Reflection.MethodInfo)
                        If Member.MemberType = System.Reflection.MemberTypes.Method AndAlso CType(Member, System.Reflection.MethodInfo).ReturnType Is expectedType Then
                            Result.Add(CType(Member, TSearchedMembers))
                        End If
                    Case GetType(System.Reflection.EventInfo)
                        If Member.MemberType = System.Reflection.MemberTypes.Event AndAlso CType(Member, System.Reflection.EventInfo).EventHandlerType Is expectedType Then
                            Result.Add(CType(Member, TSearchedMembers))
                        End If
                    Case GetType(System.Reflection.ConstructorInfo)
                        If expectedType IsNot Nothing Then Throw New ArgumentException("Constructor search doesn't support expectedType argument", NameOf(expectedType))
                        If Member.MemberType = System.Reflection.MemberTypes.Constructor Then
                            Result.Add(CType(Member, TSearchedMembers))
                        End If
                    Case Else
                        Throw New NotImplementedException("MemberType not yet supported, please file a feature request ticket at https://github.com/CompuMasterGmbH/CompuMaster.Reflection")
                End Select
            Next
            Return Result
        End Function

        Public Shared Function GetMembers(memberFilter As System.Reflection.BindingFlags, exceptionFilterInfo As String, objtype As Type) As System.Reflection.MemberInfo()
            Try
                Return objtype.GetMembers(memberFilter)
            Catch ex As Exception
                Throw New InvokeException("GetMembers failed for " & exceptionFilterInfo & " members", ex)
            End Try
        End Function

        Public Shared Function InvokeFunction(Of T)(memberFilter As System.Reflection.BindingFlags, exceptionFilterInfo As String, obj As Object, objType As Type, name As String, ParamArray values As Object()) As T
            Try
                Return CType(objType.InvokeMember(name, memberFilter Or System.Reflection.BindingFlags.InvokeMethod, Nothing, obj, values), T)
            Catch ex As Exception
                Throw New InvokeException("InvokeMember failed for " & exceptionFilterInfo & " function for """ & name & """", ex)
            End Try
        End Function

        Public Shared Sub InvokeMethod(memberFilter As System.Reflection.BindingFlags, exceptionFilterInfo As String, obj As Object, objType As Type, name As String, ParamArray values As Object())
            Try
                objType.InvokeMember(name, memberFilter Or System.Reflection.BindingFlags.InvokeMethod, Nothing, obj, values)
            Catch ex As Exception
                Throw New InvokeException("InvokeMember failed for " & exceptionFilterInfo & " method """ & name & """", ex)
            End Try
        End Sub

        Public Shared Function InvokePropertyGet(Of T)(memberFilter As System.Reflection.BindingFlags, exceptionFilterInfo As String, obj As Object, objtype As Type, name As String) As T
            Try
                Return CType(objtype.InvokeMember(name, memberFilter Or System.Reflection.BindingFlags.GetProperty, Nothing, obj, Nothing), T)
            Catch ex As Exception
                Throw New InvokeException("InvokeMember failed for " & exceptionFilterInfo & " property get """ & name & """", ex)
            End Try
        End Function

        Public Shared Function InvokePropertyGet(Of T)(memberFilter As System.Reflection.BindingFlags, exceptionFilterInfo As String, obj As Object, objtype As Type, name As String, propertyArrayItem As Object) As T
            Try
                Return CType(objtype.InvokeMember(name, memberFilter Or System.Reflection.BindingFlags.GetProperty, Nothing, obj, New Object() {propertyArrayItem}), T)
            Catch ex As Exception
                Throw New InvokeException("InvokeMember failed for " & exceptionFilterInfo & " property get """ & name & """", ex)
            End Try
        End Function

        Public Shared Sub InvokePropertySet(Of T)(memberFilter As System.Reflection.BindingFlags, exceptionFilterInfo As String, obj As Object, objType As Type, name As String, value As T)
            Try
                objType.InvokeMember(name, memberFilter Or System.Reflection.BindingFlags.SetProperty, Nothing, obj, ReflectionTools.ConvertArgumentsToObjectArray(value))
            Catch ex As Exception
                Throw New InvokeException("InvokeMember failed for " & exceptionFilterInfo & " property set """ & name & """", ex)
            End Try
        End Sub

        Public Shared Sub InvokePropertySet(Of T)(memberFilter As System.Reflection.BindingFlags, exceptionFilterInfo As String, obj As Object, objType As Type, name As String, values As T())
            Try
                objType.InvokeMember(name, memberFilter Or System.Reflection.BindingFlags.SetProperty, Nothing, obj, ReflectionTools.ConvertArgumentsToObjectArray(values))
            Catch ex As Exception
                Throw New InvokeException("InvokeMember failed for " & exceptionFilterInfo & " property set """ & name & """", ex)
            End Try
        End Sub

        Public Shared Function InvokeFieldGet(Of T)(memberFilter As System.Reflection.BindingFlags, exceptionFilterInfo As String, obj As Object, objtype As Type, name As String) As T
            Try
                Return CType(objtype.InvokeMember(name, memberFilter Or System.Reflection.BindingFlags.GetField, Nothing, obj, Nothing), T)
            Catch ex As Exception
                Throw New InvokeException("InvokeMember failed for " & exceptionFilterInfo & " field get """ & name & """", ex)
            End Try
        End Function

        Public Shared Sub InvokeFieldSet(Of T)(memberFilter As System.Reflection.BindingFlags, exceptionFilterInfo As String, obj As Object, objType As Type, name As String, value As T)
            Try
                objType.InvokeMember(name, memberFilter Or System.Reflection.BindingFlags.SetField, Nothing, obj, ReflectionTools.ConvertArgumentsToObjectArray(value))
            Catch ex As Exception
                Throw New InvokeException("InvokeMember failed for " & exceptionFilterInfo & " field set """ & name & """", ex)
            End Try
        End Sub

        Public Shared Sub InvokeFieldSet(Of T)(memberFilter As System.Reflection.BindingFlags, exceptionFilterInfo As String, obj As Object, objType As Type, name As String, values As T())
            Try
                objType.InvokeMember(name, memberFilter Or System.Reflection.BindingFlags.SetField, Nothing, obj, ReflectionTools.ConvertArgumentsToObjectArray(values))
            Catch ex As Exception
                Throw New InvokeException("InvokeMember failed for " & exceptionFilterInfo & " field set """ & name & """", ex)
            End Try
        End Sub

        Public Shared Function InvokeConstructor(Of T)(memberFilter As System.Reflection.BindingFlags, exceptionFilterInfo As String, ParamArray args As Object()) As T
            Try
                Return CType(GetType(T).InvokeMember(Nothing, memberFilter Or System.Reflection.BindingFlags.CreateInstance, Nothing, Nothing, args), T)
            Catch ex As Exception
                Throw New InvokeException("InvokeMember failed for " & exceptionFilterInfo & " constructor with " & args.Length & " argument(s)", ex)
            End Try
        End Function

    End Class

End Namespace