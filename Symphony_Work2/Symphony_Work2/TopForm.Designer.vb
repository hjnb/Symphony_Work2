﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
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
        Me.rbtnPrintout = New System.Windows.Forms.RadioButton()
        Me.lblday = New System.Windows.Forms.Label()
        Me.lblFloor = New System.Windows.Forms.Label()
        Me.topPicture = New System.Windows.Forms.PictureBox()
        CType(Me.topPicture, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.rbtnPreview.Location = New System.Drawing.Point(539, 176)
        Me.rbtnPreview.Name = "rbtnPreview"
        Me.rbtnPreview.Size = New System.Drawing.Size(63, 16)
        Me.rbtnPreview.TabIndex = 3
        Me.rbtnPreview.Text = "ﾌﾟﾚﾋﾞｭｰ"
        Me.rbtnPreview.UseVisualStyleBackColor = True
        '
        'rbtnPrintout
        '
        Me.rbtnPrintout.AutoSize = True
        Me.rbtnPrintout.Location = New System.Drawing.Point(619, 176)
        Me.rbtnPrintout.Name = "rbtnPrintout"
        Me.rbtnPrintout.Size = New System.Drawing.Size(47, 16)
        Me.rbtnPrintout.TabIndex = 4
        Me.rbtnPrintout.Text = "印刷"
        Me.rbtnPrintout.UseVisualStyleBackColor = True
        '
        'lblday
        '
        Me.lblday.AutoSize = True
        Me.lblday.Location = New System.Drawing.Point(822, 59)
        Me.lblday.Name = "lblday"
        Me.lblday.Size = New System.Drawing.Size(29, 12)
        Me.lblday.TabIndex = 5
        Me.lblday.Text = "日付"
        Me.lblday.Visible = False
        '
        'lblFloor
        '
        Me.lblFloor.AutoSize = True
        Me.lblFloor.Location = New System.Drawing.Point(820, 85)
        Me.lblFloor.Name = "lblFloor"
        Me.lblFloor.Size = New System.Drawing.Size(11, 12)
        Me.lblFloor.TabIndex = 6
        Me.lblFloor.Text = "2"
        Me.lblFloor.Visible = False
        '
        'topPicture
        '
        Me.topPicture.Location = New System.Drawing.Point(551, 58)
        Me.topPicture.Name = "topPicture"
        Me.topPicture.Size = New System.Drawing.Size(106, 103)
        Me.topPicture.TabIndex = 7
        Me.topPicture.TabStop = False
        '
        'TopForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(963, 649)
        Me.Controls.Add(Me.topPicture)
        Me.Controls.Add(Me.lblFloor)
        Me.Controls.Add(Me.lblday)
        Me.Controls.Add(Me.rbtnPrintout)
        Me.Controls.Add(Me.rbtnPreview)
        Me.Controls.Add(Me.btnCsv)
        Me.Controls.Add(Me.btnWeek)
        Me.Controls.Add(Me.btnWork)
        Me.Name = "TopForm"
        Me.Text = "特養勤務割"
        CType(Me.topPicture, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnWork As System.Windows.Forms.Button
    Friend WithEvents btnWeek As System.Windows.Forms.Button
    Friend WithEvents btnCsv As System.Windows.Forms.Button
    Friend WithEvents rbtnPreview As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnPrintout As System.Windows.Forms.RadioButton
    Friend WithEvents lblday As System.Windows.Forms.Label
    Friend WithEvents lblFloor As System.Windows.Forms.Label
    Friend WithEvents topPicture As System.Windows.Forms.PictureBox

End Class
