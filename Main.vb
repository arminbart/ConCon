Imports ConCon

Public Class Main
	Implements Logger

	Private Sub Log(strMsg As String) Implements Logger.Log
		txtLog.AppendText(If(IsEmpty(txtLog.Text), strMsg, vbCrLf & strMsg))
		txtLog.Refresh()
	End Sub

	Private Sub cmdNewList_Click(sender As System.Object, e As System.EventArgs) Handles cmdNewList.Click
		Dim frmContacts As New frmContactList(Nothing, Me)

		frmContacts.Show(Me)
	End Sub

	Private Sub cmdNewListFromFile_Click(sender As System.Object, e As System.EventArgs) Handles cmdNewListFromFile.Click
		Dim frm As New System.Windows.Forms.OpenFileDialog()

		frm.Filter = "vCard files (*.vcf)|*.vcf|HTC HSM-Backup files (*.xml)|*.xml|All Files (*.*)|*.*"

		If frm.ShowDialog(Me) = DialogResult.OK AndAlso IsNotEmpty(frm.FileName) Then
			Dim oContacts As New clContactList()

			For Each oContact As clContact In clContactFile.ReadFile(frm.FileName, Me)
				oContacts.AddLast(oContact)
			Next

			Dim frmContacts As New frmContactList(oContacts, Me)

			frmContacts.Show(Me)
		End If
	End Sub

	Private Sub cmdNewListFromFolder_Click(sender As Object, e As EventArgs) Handles cmdNewListFromFolder.Click
		Dim frm As New System.Windows.Forms.FolderBrowserDialog()

		If frm.ShowDialog(Me) = DialogResult.OK AndAlso IsNotEmpty(frm.SelectedPath) Then
			Dim oContacts As clContactList = clContactFile.ReadFilesFromFolder(frm.SelectedPath, Me)
			Dim frmContacts As New frmContactList(oContacts, Me)

			frmContacts.Show(Me)
		End If
	End Sub

End Class
