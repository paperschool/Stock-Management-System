<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Label_Output
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
        Me.lblStockID = New System.Windows.Forms.Label()
        Me.LabelPrintForm = New System.Windows.Forms.PrintDialog()
        Me.LabelPrintDoc = New System.Drawing.Printing.PrintDocument()
        Me.SuspendLayout()
        '
        'lblStockID
        '
        Me.lblStockID.AutoSize = True
        Me.lblStockID.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStockID.Location = New System.Drawing.Point(10, 9)
        Me.lblStockID.Name = "lblStockID"
        Me.lblStockID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblStockID.Size = New System.Drawing.Size(148, 20)
        Me.lblStockID.TabIndex = 5
        Me.lblStockID.Text = "Sample Information"
        '
        'LabelPrintForm
        '
        Me.LabelPrintForm.UseEXDialog = True
        '
        'Label_Output
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ClientSize = New System.Drawing.Size(333, 171)
        Me.Controls.Add(Me.lblStockID)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "Label_Output"
        Me.Text = "Label_Output"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblStockID As System.Windows.Forms.Label
    Friend WithEvents LabelPrintForm As System.Windows.Forms.PrintDialog
    Friend WithEvents LabelPrintDoc As System.Drawing.Printing.PrintDocument
End Class
