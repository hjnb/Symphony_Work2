<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 勤務割
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
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.btnRegist = New System.Windows.Forms.Button()
        Me.btnRowDelete = New System.Windows.Forms.Button()
        Me.btnRowAdd = New System.Windows.Forms.Button()
        Me.rbtn2F = New System.Windows.Forms.RadioButton()
        Me.rbtn3F = New System.Windows.Forms.RadioButton()
        Me.dgvWork = New System.Windows.Forms.DataGridView()
        CType(Me.dgvWork, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(591, 12)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 40)
        Me.btnDelete.TabIndex = 15
        Me.btnDelete.Text = "削除"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(677, 12)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(80, 40)
        Me.btnPrint.TabIndex = 14
        Me.btnPrint.Text = "印刷"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'btnRegist
        '
        Me.btnRegist.Location = New System.Drawing.Point(505, 12)
        Me.btnRegist.Name = "btnRegist"
        Me.btnRegist.Size = New System.Drawing.Size(80, 40)
        Me.btnRegist.TabIndex = 13
        Me.btnRegist.Text = "登録"
        Me.btnRegist.UseVisualStyleBackColor = True
        '
        'btnRowDelete
        '
        Me.btnRowDelete.Location = New System.Drawing.Point(425, 29)
        Me.btnRowDelete.Name = "btnRowDelete"
        Me.btnRowDelete.Size = New System.Drawing.Size(55, 23)
        Me.btnRowDelete.TabIndex = 12
        Me.btnRowDelete.Text = "行削除"
        Me.btnRowDelete.UseVisualStyleBackColor = True
        '
        'btnRowAdd
        '
        Me.btnRowAdd.Location = New System.Drawing.Point(364, 29)
        Me.btnRowAdd.Name = "btnRowAdd"
        Me.btnRowAdd.Size = New System.Drawing.Size(55, 23)
        Me.btnRowAdd.TabIndex = 11
        Me.btnRowAdd.Text = "行挿入"
        Me.btnRowAdd.UseVisualStyleBackColor = True
        '
        'rbtn2F
        '
        Me.rbtn2F.AutoSize = True
        Me.rbtn2F.Checked = True
        Me.rbtn2F.Location = New System.Drawing.Point(173, 39)
        Me.rbtn2F.Name = "rbtn2F"
        Me.rbtn2F.Size = New System.Drawing.Size(43, 16)
        Me.rbtn2F.TabIndex = 10
        Me.rbtn2F.TabStop = True
        Me.rbtn2F.Text = "２階"
        Me.rbtn2F.UseVisualStyleBackColor = True
        '
        'rbtn3F
        '
        Me.rbtn3F.AutoSize = True
        Me.rbtn3F.Location = New System.Drawing.Point(173, 12)
        Me.rbtn3F.Name = "rbtn3F"
        Me.rbtn3F.Size = New System.Drawing.Size(43, 16)
        Me.rbtn3F.TabIndex = 9
        Me.rbtn3F.Text = "３階"
        Me.rbtn3F.UseVisualStyleBackColor = True
        '
        'dgvWork
        '
        Me.dgvWork.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvWork.Location = New System.Drawing.Point(12, 80)
        Me.dgvWork.Name = "dgvWork"
        Me.dgvWork.RowTemplate.Height = 21
        Me.dgvWork.Size = New System.Drawing.Size(977, 630)
        Me.dgvWork.TabIndex = 8
        '
        '勤務割
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1000, 723)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnRegist)
        Me.Controls.Add(Me.btnRowDelete)
        Me.Controls.Add(Me.btnRowAdd)
        Me.Controls.Add(Me.rbtn2F)
        Me.Controls.Add(Me.rbtn3F)
        Me.Controls.Add(Me.dgvWork)
        Me.Name = "勤務割"
        Me.Text = "勤務割"
        CType(Me.dgvWork, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnRegist As System.Windows.Forms.Button
    Friend WithEvents btnRowDelete As System.Windows.Forms.Button
    Friend WithEvents btnRowAdd As System.Windows.Forms.Button
    Friend WithEvents rbtn2F As System.Windows.Forms.RadioButton
    Friend WithEvents rbtn3F As System.Windows.Forms.RadioButton
    Friend WithEvents dgvWork As System.Windows.Forms.DataGridView
End Class
