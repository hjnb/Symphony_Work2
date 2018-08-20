<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TopForm
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
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

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnWork = New System.Windows.Forms.Button()
        Me.btnWeek = New System.Windows.Forms.Button()
        Me.btnCsv = New System.Windows.Forms.Button()
        Me.rbtnPreview = New System.Windows.Forms.RadioButton()
        Me.rbtnPrint = New System.Windows.Forms.RadioButton()
        Me.SuspendLayout()
        '
        'btnWork
        '
        Me.btnWork.Font = New System.Drawing.Font("MS UI Gothic", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnWork.ForeColor = System.Drawing.Color.Blue
        Me.btnWork.Location = New System.Drawing.Point(176, 58)
        Me.btnWork.Name = "btnWork"
        Me.btnWork.Size = New System.Drawing.Size(340, 161)
        Me.btnWork.TabIndex = 0
        Me.btnWork.Text = "勤務割表"
        Me.btnWork.UseVisualStyleBackColor = True
        '
        'btnWeek
        '
        Me.btnWeek.Font = New System.Drawing.Font("MS UI Gothic", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnWeek.ForeColor = System.Drawing.Color.Blue
        Me.btnWeek.Location = New System.Drawing.Point(176, 223)
        Me.btnWeek.Name = "btnWeek"
        Me.btnWeek.Size = New System.Drawing.Size(340, 161)
        Me.btnWeek.TabIndex = 1
        Me.btnWeek.Text = "週間表"
        Me.btnWeek.UseVisualStyleBackColor = True
        '
        'btnCsv
        '
        Me.btnCsv.Font = New System.Drawing.Font("MS UI Gothic", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnCsv.ForeColor = System.Drawing.Color.Blue
        Me.btnCsv.Location = New System.Drawing.Point(176, 389)
        Me.btnCsv.Name = "btnCsv"
        Me.btnCsv.Size = New System.Drawing.Size(340, 161)
        Me.btnCsv.TabIndex = 2
        Me.btnCsv.Text = "勤務割CSV作成"
        Me.btnCsv.UseVisualStyleBackColor = True
        '
        'rbtnPreview
        '
        Me.rbtnPreview.AutoSize = True
        Me.rbtnPreview.Checked = True
        Me.rbtnPreview.Location = New System.Drawing.Point(539, 176)
        Me.rbtnPreview.Name = "rbtnPreview"
        Me.rbtnPreview.Size = New System.Drawing.Size(63, 16)
        Me.rbtnPreview.TabIndex = 3
        Me.rbtnPreview.TabStop = True
        Me.rbtnPreview.Text = "ﾌﾟﾚﾋﾞｭｰ"
        Me.rbtnPreview.UseVisualStyleBackColor = True
        '
        'rbtnPrint
        '
        Me.rbtnPrint.AutoSize = True
        Me.rbtnPrint.Location = New System.Drawing.Point(619, 176)
        Me.rbtnPrint.Name = "rbtnPrint"
        Me.rbtnPrint.Size = New System.Drawing.Size(47, 16)
        Me.rbtnPrint.TabIndex = 4
        Me.rbtnPrint.Text = "印刷"
        Me.rbtnPrint.UseVisualStyleBackColor = True
        '
        'TopForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(963, 649)
        Me.Controls.Add(Me.rbtnPrint)
        Me.Controls.Add(Me.rbtnPreview)
        Me.Controls.Add(Me.btnCsv)
        Me.Controls.Add(Me.btnWeek)
        Me.Controls.Add(Me.btnWork)
        Me.Name = "TopForm"
        Me.Text = "特養勤務割"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnWork As System.Windows.Forms.Button
    Friend WithEvents btnWeek As System.Windows.Forms.Button
    Friend WithEvents btnCsv As System.Windows.Forms.Button
    Friend WithEvents rbtnPreview As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnPrint As System.Windows.Forms.RadioButton

End Class
