Module mdGlobals

	Public Function IsAnyOf(str As String, ParamArray values As String()) As Boolean

		For Each strValue In values
			If str.Equals(strValue, StringComparison.CurrentCultureIgnoreCase) Then Return True
		Next

		Return False
	End Function

	Public Function IsEmpty(str As String)
		Return str Is Nothing OrElse str.Trim().Length = 0
	End Function

	Public Function IsNotEmpty(str As String)
		Return str IsNot Nothing AndAlso str.Trim().Length > 0
	End Function

	Public Function MakeInt(ByVal o As Object) As Integer
		Try
			If o IsNot Nothing And IsNotEmpty(o.ToString()) Then
				If IsNumeric(o) Then Return CInt(o)

				Dim nResult As Integer = 0
				If Integer.TryParse(o.ToString(), nResult) Then Return nResult
			End If
		Catch ex As Exception
		End Try

		Return 0
	End Function

End Module
