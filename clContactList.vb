Public Class clContactList
	Inherits Generic.LinkedList(Of clContact)


	Public Sub Add(oContact As clContact)
		AddLast(oContact)
	End Sub

	Public ReadOnly Property IsEmpty() As Boolean
		Get
			Return Count = 0
		End Get
	End Property

End Class
