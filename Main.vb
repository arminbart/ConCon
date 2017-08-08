﻿Imports ConCon

Public Class Main
	Implements Logger

	Private Sub cmdNewListFromFolder_Click(sender As Object, e As EventArgs) Handles cmdNewListFromFolder.Click
		Dim frm As New System.Windows.Forms.FolderBrowserDialog()

		If frm.ShowDialog(Me) = DialogResult.OK Then
			'System.Diagnostics.Debugger.Break()
			Dim strPath As String = frm.SelectedPath
			'Dim strPath As String = "C:\Users\Armin\Downloads\Contact"
			If Not String.IsNullOrEmpty(strPath) Then
				Dim oContacts As clContactList = ReadFilesFromFolder(strPath)
				Dim frmContacts As New frmContactList(oContacts)

				frmContacts.Show(Me)
			End If
		End If
	End Sub

	Private Function ReadFilesFromFolder(strPath As String) As clContactList
		Dim oContacts As New clContactList()
		Dim arrFiles As String() = System.IO.Directory.GetFiles(strPath)

		For Each strFile As String In arrFiles
			For Each oContact As clContact In clContactFile.ReadFile(strFile, Me)
				oContacts.AddLast(oContact)
			Next
		Next

		Return oContacts
	End Function

	Private Sub Log(strMsg As String) Implements Logger.Log
		txtLog.AppendText(If(String.IsNullOrEmpty(txtLog.Text), strMsg, vbCrLf & strMsg))
		txtLog.Refresh()
	End Sub
End Class
