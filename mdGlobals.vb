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

End Module
