﻿Imports ConCon

Public Class Main
	Implements Logger

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

	Private Sub cmdNewList_Click(sender As System.Object, e As System.EventArgs) Handles cmdNewList.Click
		Dim frmContacts As New frmContactList(Nothing)

		frmContacts.Show(Me)
	End Sub

	Private Sub cmdNewListFromFile_Click(sender As System.Object, e As System.EventArgs) Handles cmdNewListFromFile.Click
		Dim frm As New System.Windows.Forms.OpenFileDialog()

		frm.Filter = "HTC HSM-Backup files (*.xml)|*.xml|vCard files (*.vcf)|*.vcf|All Files (*.*)|*.*"

		If frm.ShowDialog(Me) = DialogResult.OK Then
			Dim oContacts As New clContactList()
			Dim strFile As String = frm.FileName

			For Each oContact As clContact In clContactFile.ReadFile(strFile, Me)
				oContacts.AddLast(oContact)
			Next

			Dim frmContacts As New frmContactList(oContacts)

			frmContacts.Show(Me)
		End If
	End Sub

	Private Sub cmdNewListFromFolder_Click(sender As Object, e As EventArgs) Handles cmdNewListFromFolder.Click
		Dim frm As New System.Windows.Forms.FolderBrowserDialog()

		If True Then 'frm.ShowDialog(Me) = DialogResult.OK Then
			'System.Diagnostics.Debugger.Break()
			'Dim strPath As String = frm.SelectedPath
			'Dim strPath As String = "C:\Users\Armin\Downloads\Contact"
			'Dim strPath As String = "C:\Users\Armin.XPECTO\Downloads\Contacts"
			Dim strPath As String = "G:\Backup\HTC\HSM\BR_485433355657393039323139\B37456C4530BE810DC040F50DA72EDA09ADDFB0A"

			If Not String.IsNullOrEmpty(strPath) Then
				Dim oContacts As clContactList = ReadFilesFromFolder(strPath)
				Dim frmContacts As New frmContactList(oContacts)

				frmContacts.Show(Me)
			End If
		End If
	End Sub

End Class
