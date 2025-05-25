<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class fMain
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
        Me.txtOpen = New System.Windows.Forms.Button()
        Me.txtConsole = New System.Windows.Forms.TextBox()
        Me.btnExportCpp = New System.Windows.Forms.Button()
        Me.btnExportJson = New System.Windows.Forms.Button()
        Me.btnExportTxt = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txtOpen
        '
        Me.txtOpen.Location = New System.Drawing.Point(12, 12)
        Me.txtOpen.Name = "txtOpen"
        Me.txtOpen.Size = New System.Drawing.Size(94, 23)
        Me.txtOpen.TabIndex = 0
        Me.txtOpen.Text = "Open OPM"
        Me.txtOpen.UseVisualStyleBackColor = True
        '
        'txtConsole
        '
        Me.txtConsole.Location = New System.Drawing.Point(12, 53)
        Me.txtConsole.Multiline = True
        Me.txtConsole.Name = "txtConsole"
        Me.txtConsole.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtConsole.Size = New System.Drawing.Size(776, 435)
        Me.txtConsole.TabIndex = 1
        '
        'btnExportCpp
        '
        Me.btnExportCpp.Location = New System.Drawing.Point(169, 12)
        Me.btnExportCpp.Name = "btnExportCpp"
        Me.btnExportCpp.Size = New System.Drawing.Size(94, 23)
        Me.btnExportCpp.TabIndex = 2
        Me.btnExportCpp.Text = "Export C++"
        Me.btnExportCpp.UseVisualStyleBackColor = True
        '
        'btnExportJson
        '
        Me.btnExportJson.Location = New System.Drawing.Point(269, 12)
        Me.btnExportJson.Name = "btnExportJson"
        Me.btnExportJson.Size = New System.Drawing.Size(94, 23)
        Me.btnExportJson.TabIndex = 3
        Me.btnExportJson.Text = "Export JSON"
        Me.btnExportJson.UseVisualStyleBackColor = True
        '
        'btnExportTxt
        '
        Me.btnExportTxt.Location = New System.Drawing.Point(369, 12)
        Me.btnExportTxt.Name = "btnExportTxt"
        Me.btnExportTxt.Size = New System.Drawing.Size(94, 23)
        Me.btnExportTxt.TabIndex = 4
        Me.btnExportTxt.Text = "Export Report"
        Me.btnExportTxt.UseVisualStyleBackColor = True
        '
        'fMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 500)
        Me.Controls.Add(Me.btnExportTxt)
        Me.Controls.Add(Me.btnExportJson)
        Me.Controls.Add(Me.btnExportCpp)
        Me.Controls.Add(Me.txtConsole)
        Me.Controls.Add(Me.txtOpen)
        Me.Name = "fMain"
        Me.Text = "OPM Importer"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtOpen As Button
    Friend WithEvents txtConsole As TextBox
    Friend WithEvents btnExportCpp As Button
    Friend WithEvents btnExportJson As Button
    Friend WithEvents btnExportTxt As Button
End Class
