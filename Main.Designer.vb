<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Main
	Inherits System.Windows.Forms.Form

	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> _
	Protected Overrides Sub Dispose(ByVal disposing As Boolean)
		Try
			If disposing AndAlso components IsNot Nothing Then
				components.Dispose()
			End If
		Finally
			MyBase.Dispose(disposing)
		End Try
	End Sub

	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer

	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.  
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> _
	Private Sub InitializeComponent()
		Me.cmdNewList = New System.Windows.Forms.Button()
		Me.cmdNewListFromFile = New System.Windows.Forms.Button()
		Me.cmdNewListFromFolder = New System.Windows.Forms.Button()
		Me.txtLog = New System.Windows.Forms.TextBox()
		Me.SuspendLayout()
		'
		'cmdNewList
		'
		Me.cmdNewList.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
				  Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.cmdNewList.Location = New System.Drawing.Point(12, 12)
		Me.cmdNewList.Name = "cmdNewList"
		Me.cmdNewList.Size = New System.Drawing.Size(710, 70)
		Me.cmdNewList.TabIndex = 0
		Me.cmdNewList.Text = "Open &Empty Contact List"
		Me.cmdNewList.UseVisualStyleBackColor = True
		'
		'cmdNewListFromFile
		'
		Me.cmdNewListFromFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
				  Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.cmdNewListFromFile.Location = New System.Drawing.Point(12, 88)
		Me.cmdNewListFromFile.Name = "cmdNewListFromFile"
		Me.cmdNewListFromFile.Size = New System.Drawing.Size(710, 70)
		Me.cmdNewListFromFile.TabIndex = 1
		Me.cmdNewListFromFile.Text = "Open Contact List from &File"
		Me.cmdNewListFromFile.UseVisualStyleBackColor = True
		'
		'cmdNewListFromFolder
		'
		Me.cmdNewListFromFolder.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
				  Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.cmdNewListFromFolder.Location = New System.Drawing.Point(12, 164)
		Me.cmdNewListFromFolder.Name = "cmdNewListFromFolder"
		Me.cmdNewListFromFolder.Size = New System.Drawing.Size(710, 70)
		Me.cmdNewListFromFolder.TabIndex = 2
		Me.cmdNewListFromFolder.Text = "Open Contact List from Fol&der"
		Me.cmdNewListFromFolder.UseVisualStyleBackColor = True
		'
		'txtLog
		'
		Me.txtLog.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
				  Or System.Windows.Forms.AnchorStyles.Left) _
				  Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtLog.Location = New System.Drawing.Point(12, 240)
		Me.txtLog.Multiline = True
		Me.txtLog.Name = "txtLog"
		Me.txtLog.ReadOnly = True
		Me.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Both
		Me.txtLog.Size = New System.Drawing.Size(710, 210)
		Me.txtLog.TabIndex = 3
		'
		'Main
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(734, 462)
		Me.Controls.Add(Me.txtLog)
		Me.Controls.Add(Me.cmdNewListFromFolder)
		Me.Controls.Add(Me.cmdNewListFromFile)
		Me.Controls.Add(Me.cmdNewList)
		Me.Name = "Main"
		Me.Text = "ConCon - Contact Converter for Mobile Phone Contacts"
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents cmdNewList As Button
	Friend WithEvents cmdNewListFromFile As Button
	Friend WithEvents cmdNewListFromFolder As Button
	Friend WithEvents txtLog As TextBox
End Class
