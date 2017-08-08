Public Class clContactList
	Inherits Generic.LinkedList(Of clContact)

	'Private moLogger As New System.Text.StringBuilder()

	Public Sub Add(oContact As clContact)
		AddLast(oContact)
	End Sub

	'Public Sub Log(strMsg As String)
	'	moLogger.AppendLine(strMsg)
	'End Sub

	'Public Function GetLog() As String
	'	Return moLogger.ToString()
	'End Function

	Public ReadOnly Property IsEmpty() As Boolean
		Get
			Return Count = 0
		End Get
	End Property

End Class
