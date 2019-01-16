<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class passwordForm
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
        Me.errorLabel = New System.Windows.Forms.Label()
        Me.newPassBox = New System.Windows.Forms.TextBox()
        Me.confirmPassBox = New System.Windows.Forms.TextBox()
        Me.passBox = New System.Windows.Forms.TextBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'errorLabel
        '
        Me.errorLabel.AutoSize = True
        Me.errorLabel.ForeColor = System.Drawing.Color.Red
        Me.errorLabel.Location = New System.Drawing.Point(48, 140)
        Me.errorLabel.Name = "errorLabel"
        Me.errorLabel.Size = New System.Drawing.Size(38, 12)
        Me.errorLabel.TabIndex = 26
        Me.errorLabel.Text = "Label4"
        Me.errorLabel.Visible = False
        '
        'newPassBox
        '
        Me.newPassBox.Location = New System.Drawing.Point(128, 69)
        Me.newPassBox.Name = "newPassBox"
        Me.newPassBox.Size = New System.Drawing.Size(130, 19)
        Me.newPassBox.TabIndex = 19
        '
        'confirmPassBox
        '
        Me.confirmPassBox.Location = New System.Drawing.Point(128, 103)
        Me.confirmPassBox.Name = "confirmPassBox"
        Me.confirmPassBox.Size = New System.Drawing.Size(130, 19)
        Me.confirmPassBox.TabIndex = 20
        '
        'passBox
        '
        Me.passBox.Location = New System.Drawing.Point(128, 26)
        Me.passBox.Name = "passBox"
        Me.passBox.Size = New System.Drawing.Size(130, 19)
        Me.passBox.TabIndex = 18
        '
        'btnCancel
        '
        Me.btnCancel.ForeColor = System.Drawing.Color.Blue
        Me.btnCancel.Location = New System.Drawing.Point(184, 174)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 32)
        Me.btnCancel.TabIndex = 22
        Me.btnCancel.Text = "キャンセル"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOk
        '
        Me.btnOk.ForeColor = System.Drawing.Color.Blue
        Me.btnOk.Location = New System.Drawing.Point(100, 174)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(75, 32)
        Me.btnOk.TabIndex = 21
        Me.btnOk.Text = "OK"
        Me.btnOk.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(48, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 12)
        Me.Label3.TabIndex = 24
        Me.Label3.Text = "新パスワード"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(48, 108)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 12)
        Me.Label2.TabIndex = 25
        Me.Label2.Text = "変更確認"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(48, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(52, 12)
        Me.Label1.TabIndex = 23
        Me.Label1.Text = "パスワード"
        '
        'passwordForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(306, 232)
        Me.Controls.Add(Me.errorLabel)
        Me.Controls.Add(Me.newPassBox)
        Me.Controls.Add(Me.confirmPassBox)
        Me.Controls.Add(Me.passBox)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "passwordForm"
        Me.Text = "パスワード入力"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents errorLabel As System.Windows.Forms.Label
    Friend WithEvents newPassBox As System.Windows.Forms.TextBox
    Friend WithEvents confirmPassBox As System.Windows.Forms.TextBox
    Friend WithEvents passBox As System.Windows.Forms.TextBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
