<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class From_ShowJson
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.s_JsonText = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        's_JsonText
        '
        Me.s_JsonText.BackColor = System.Drawing.SystemColors.Window
        Me.s_JsonText.Dock = System.Windows.Forms.DockStyle.Fill
        Me.s_JsonText.Location = New System.Drawing.Point(0, 0)
        Me.s_JsonText.Multiline = True
        Me.s_JsonText.Name = "s_JsonText"
        Me.s_JsonText.ReadOnly = True
        Me.s_JsonText.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.s_JsonText.Size = New System.Drawing.Size(762, 459)
        Me.s_JsonText.TabIndex = 0
        '
        'From_ShowJson
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(762, 459)
        Me.Controls.Add(Me.s_JsonText)
        Me.Name = "From_ShowJson"
        Me.Text = "0"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents s_JsonText As TextBox
End Class
