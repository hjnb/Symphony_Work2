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
        Me.components = New System.ComponentModel.Container()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.btnRegist = New System.Windows.Forms.Button()
        Me.btnRowDelete = New System.Windows.Forms.Button()
        Me.btnRowAdd = New System.Windows.Forms.Button()
        Me.rbtn2F = New System.Windows.Forms.RadioButton()
        Me.rbtn3F = New System.Windows.Forms.RadioButton()
        Me.ymBox = New ymdBox.ymdBox()
        Me.wordPanel = New System.Windows.Forms.Panel()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.dgvWork = New Symphony_Work2.workDataGridView(Me.components)
        Me.wordPanel.SuspendLayout()
        CType(Me.dgvWork, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(545, 15)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 35)
        Me.btnDelete.TabIndex = 14
        Me.btnDelete.Text = "削除"
        Me.btnDelete.UseVisualStyleBackColor = True
        Me.btnDelete.Visible = False
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(631, 15)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(80, 35)
        Me.btnPrint.TabIndex = 13
        Me.btnPrint.Text = "印刷"
        Me.btnPrint.UseVisualStyleBackColor = True
        Me.btnPrint.Visible = False
        '
        'btnRegist
        '
        Me.btnRegist.Location = New System.Drawing.Point(459, 15)
        Me.btnRegist.Name = "btnRegist"
        Me.btnRegist.Size = New System.Drawing.Size(80, 35)
        Me.btnRegist.TabIndex = 12
        Me.btnRegist.Text = "登録"
        Me.btnRegist.UseVisualStyleBackColor = True
        Me.btnRegist.Visible = False
        '
        'btnRowDelete
        '
        Me.btnRowDelete.Location = New System.Drawing.Point(379, 26)
        Me.btnRowDelete.Name = "btnRowDelete"
        Me.btnRowDelete.Size = New System.Drawing.Size(55, 23)
        Me.btnRowDelete.TabIndex = 11
        Me.btnRowDelete.Text = "行削除"
        Me.btnRowDelete.UseVisualStyleBackColor = True
        Me.btnRowDelete.Visible = False
        '
        'btnRowAdd
        '
        Me.btnRowAdd.Location = New System.Drawing.Point(318, 26)
        Me.btnRowAdd.Name = "btnRowAdd"
        Me.btnRowAdd.Size = New System.Drawing.Size(55, 23)
        Me.btnRowAdd.TabIndex = 10
        Me.btnRowAdd.Text = "行挿入"
        Me.btnRowAdd.UseVisualStyleBackColor = True
        Me.btnRowAdd.Visible = False
        '
        'rbtn2F
        '
        Me.rbtn2F.AutoSize = True
        Me.rbtn2F.Location = New System.Drawing.Point(152, 36)
        Me.rbtn2F.Name = "rbtn2F"
        Me.rbtn2F.Size = New System.Drawing.Size(43, 16)
        Me.rbtn2F.TabIndex = 9
        Me.rbtn2F.Text = "２階"
        Me.rbtn2F.UseVisualStyleBackColor = True
        '
        'rbtn3F
        '
        Me.rbtn3F.AutoSize = True
        Me.rbtn3F.Location = New System.Drawing.Point(152, 9)
        Me.rbtn3F.Name = "rbtn3F"
        Me.rbtn3F.Size = New System.Drawing.Size(43, 16)
        Me.rbtn3F.TabIndex = 8
        Me.rbtn3F.Text = "３階"
        Me.rbtn3F.UseVisualStyleBackColor = True
        '
        'ymBox
        '
        Me.ymBox.boxType = 5
        Me.ymBox.DateText = ""
        Me.ymBox.EraLabelText = "H31"
        Me.ymBox.EraText = ""
        Me.ymBox.Location = New System.Drawing.Point(41, 9)
        Me.ymBox.MonthLabelText = "02"
        Me.ymBox.MonthText = ""
        Me.ymBox.Name = "ymBox"
        Me.ymBox.Size = New System.Drawing.Size(95, 40)
        Me.ymBox.TabIndex = 15
        '
        'wordPanel
        '
        Me.wordPanel.Controls.Add(Me.Label8)
        Me.wordPanel.Controls.Add(Me.Label1)
        Me.wordPanel.Controls.Add(Me.Label10)
        Me.wordPanel.Controls.Add(Me.Label2)
        Me.wordPanel.Controls.Add(Me.Label18)
        Me.wordPanel.Controls.Add(Me.Label4)
        Me.wordPanel.Controls.Add(Me.Label20)
        Me.wordPanel.Controls.Add(Me.Label3)
        Me.wordPanel.Controls.Add(Me.Label21)
        Me.wordPanel.Controls.Add(Me.Label7)
        Me.wordPanel.Controls.Add(Me.Label22)
        Me.wordPanel.Controls.Add(Me.Label6)
        Me.wordPanel.Controls.Add(Me.Label11)
        Me.wordPanel.Controls.Add(Me.Label5)
        Me.wordPanel.Controls.Add(Me.Label12)
        Me.wordPanel.Controls.Add(Me.Label9)
        Me.wordPanel.Controls.Add(Me.Label13)
        Me.wordPanel.Controls.Add(Me.Label17)
        Me.wordPanel.Controls.Add(Me.Label14)
        Me.wordPanel.Controls.Add(Me.Label16)
        Me.wordPanel.Controls.Add(Me.Label15)
        Me.wordPanel.Location = New System.Drawing.Point(62, 650)
        Me.wordPanel.Name = "wordPanel"
        Me.wordPanel.Size = New System.Drawing.Size(999, 57)
        Me.wordPanel.TabIndex = 33
        Me.wordPanel.Visible = False
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.ForeColor = System.Drawing.Color.Blue
        Me.Label8.Location = New System.Drawing.Point(465, 8)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(45, 12)
        Me.Label8.TabIndex = 17
        Me.Label8.Text = "8： 深夜"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.Color.Blue
        Me.Label1.Location = New System.Drawing.Point(5, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(44, 12)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "0 ： ｸﾘｱ"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.ForeColor = System.Drawing.Color.Blue
        Me.Label10.Location = New System.Drawing.Point(870, 31)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(55, 12)
        Me.Label10.TabIndex = 30
        Me.Label10.Text = "35 ： 産休"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.ForeColor = System.Drawing.Color.Blue
        Me.Label2.Location = New System.Drawing.Point(58, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(49, 12)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "1 ： 早出"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.ForeColor = System.Drawing.Color.Blue
        Me.Label18.Location = New System.Drawing.Point(802, 31)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(55, 12)
        Me.Label18.TabIndex = 29
        Me.Label18.Text = "34 ： 希休"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ForeColor = System.Drawing.Color.Blue
        Me.Label4.Location = New System.Drawing.Point(118, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(49, 12)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "2 ： 日早"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.ForeColor = System.Drawing.Color.Blue
        Me.Label20.Location = New System.Drawing.Point(670, 31)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(53, 12)
        Me.Label20.TabIndex = 27
        Me.Label20.Text = "33 ： 明け"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.ForeColor = System.Drawing.Color.Blue
        Me.Label3.Location = New System.Drawing.Point(176, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(49, 12)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "3 ： 日勤"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.ForeColor = System.Drawing.Color.Blue
        Me.Label21.Location = New System.Drawing.Point(607, 31)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(51, 12)
        Me.Label21.TabIndex = 26
        Me.Label21.Text = "32： 公休"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.ForeColor = System.Drawing.Color.Blue
        Me.Label7.Location = New System.Drawing.Point(234, 8)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(49, 12)
        Me.Label7.TabIndex = 13
        Me.Label7.Text = "4 ： 日遅"
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.ForeColor = System.Drawing.Color.Blue
        Me.Label22.Location = New System.Drawing.Point(538, 31)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(55, 12)
        Me.Label22.TabIndex = 25
        Me.Label22.Text = "31 ： 有休"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.ForeColor = System.Drawing.Color.Blue
        Me.Label6.Location = New System.Drawing.Point(294, 8)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 12)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "5 ： 遅出"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.ForeColor = System.Drawing.Color.Blue
        Me.Label11.Location = New System.Drawing.Point(938, 31)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(55, 12)
        Me.Label11.TabIndex = 24
        Me.Label11.Text = "36 ： 特休"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.ForeColor = System.Drawing.Color.Blue
        Me.Label5.Location = New System.Drawing.Point(352, 8)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(49, 12)
        Me.Label5.TabIndex = 15
        Me.Label5.Text = "6 ： 遅々"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.ForeColor = System.Drawing.Color.Blue
        Me.Label12.Location = New System.Drawing.Point(870, 8)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(55, 12)
        Me.Label12.TabIndex = 23
        Me.Label12.Text = "22 ： 研修"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.ForeColor = System.Drawing.Color.Blue
        Me.Label9.Location = New System.Drawing.Point(407, 8)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(49, 12)
        Me.Label9.TabIndex = 16
        Me.Label9.Text = "7 ： 夜勤"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.ForeColor = System.Drawing.Color.Blue
        Me.Label13.Location = New System.Drawing.Point(802, 8)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(55, 12)
        Me.Label13.TabIndex = 22
        Me.Label13.Text = "21 ： 半行"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.ForeColor = System.Drawing.Color.Blue
        Me.Label17.Location = New System.Drawing.Point(538, 8)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(43, 12)
        Me.Label17.TabIndex = 18
        Me.Label17.Text = "10 ： 半"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.ForeColor = System.Drawing.Color.Blue
        Me.Label14.Location = New System.Drawing.Point(735, 8)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(55, 12)
        Me.Label14.TabIndex = 21
        Me.Label14.Text = "13 ： 半夜"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.ForeColor = System.Drawing.Color.Blue
        Me.Label16.Location = New System.Drawing.Point(607, 8)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(47, 12)
        Me.Label16.TabIndex = 19
        Me.Label16.Text = "11： 半A"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.ForeColor = System.Drawing.Color.Blue
        Me.Label15.Location = New System.Drawing.Point(670, 8)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(51, 12)
        Me.Label15.TabIndex = 20
        Me.Label15.Text = "12 ： 半B"
        '
        'dgvWork
        '
        Me.dgvWork.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvWork.Location = New System.Drawing.Point(36, 56)
        Me.dgvWork.Name = "dgvWork"
        Me.dgvWork.RowTemplate.Height = 21
        Me.dgvWork.Size = New System.Drawing.Size(1044, 593)
        Me.dgvWork.TabIndex = 16
        '
        '勤務割
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1115, 754)
        Me.Controls.Add(Me.wordPanel)
        Me.Controls.Add(Me.dgvWork)
        Me.Controls.Add(Me.ymBox)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnRegist)
        Me.Controls.Add(Me.btnRowDelete)
        Me.Controls.Add(Me.btnRowAdd)
        Me.Controls.Add(Me.rbtn2F)
        Me.Controls.Add(Me.rbtn3F)
        Me.Name = "勤務割"
        Me.Text = "勤務割"
        Me.wordPanel.ResumeLayout(False)
        Me.wordPanel.PerformLayout()
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
    Friend WithEvents ymBox As ymdBox.ymdBox
    Friend WithEvents dgvWork As Symphony_Work2.workDataGridView
    Friend WithEvents wordPanel As System.Windows.Forms.Panel
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
End Class
