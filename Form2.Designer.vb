<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
          Me.DataGridView1 = New System.Windows.Forms.DataGridView
          Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn
          Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn
          Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn
          CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
          Me.SuspendLayout()
          '
          'DataGridView1
          '
          Me.DataGridView1.AllowUserToOrderColumns = True
          Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
          Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column2, Me.Column3})
          Me.DataGridView1.Location = New System.Drawing.Point(12, 12)
          Me.DataGridView1.Name = "DataGridView1"
          Me.DataGridView1.Size = New System.Drawing.Size(497, 292)
          Me.DataGridView1.TabIndex = 17
          '
          'Column1
          '
          Me.Column1.HeaderText = "Size"
          Me.Column1.MaxInputLength = 32
          Me.Column1.Name = "Column1"
          Me.Column1.ReadOnly = True
          '
          'Column2
          '
          Me.Column2.HeaderText = ""
          Me.Column2.MaxInputLength = 32
          Me.Column2.Name = "Column2"
          Me.Column2.ReadOnly = True
          Me.Column2.Width = 24
          '
          'Column3
          '
          Me.Column3.HeaderText = "Path"
          Me.Column3.Name = "Column3"
          Me.Column3.ReadOnly = True
          Me.Column3.Width = 490
          '
          'Form2
          '
          Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
          Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
          Me.ClientSize = New System.Drawing.Size(552, 382)
          Me.Controls.Add(Me.DataGridView1)
          Me.Name = "Form2"
          Me.Text = "Form2"
          CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
          Me.ResumeLayout(False)

     End Sub
     Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
     Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
     Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
     Friend WithEvents Column3 As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
