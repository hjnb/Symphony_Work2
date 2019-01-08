﻿Public Class TopForm
    'データベースのパス
    Public dbFilePath As String = My.Application.Info.DirectoryPath & "\Work2.mdb"
    Public DB_Work As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbFilePath

    'エクセルのパス
    Public excelFilePass As String = My.Application.Info.DirectoryPath & "\Work2.xls"

    '.iniファイルのパス
    Public iniFilePath As String = My.Application.Info.DirectoryPath & "\Work2.ini"

    '画像パス
    'Public imageFilePath As String = My.Application.Info.DirectoryPath & "\"


    Private workForm As 勤務割
    Private weekForm As 週間表
    Private csvForm As CSV作成

    Private Sub TopForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'データベース、エクセル、構成ファイルの存在チェック
        If Not System.IO.File.Exists(dbFilePath) Then
            MsgBox("データベースファイルが存在しません。ファイルを配置して下さい。")
            Me.Close()
            Exit Sub
        End If

        If Not System.IO.File.Exists(excelFilePass) Then
            MsgBox("エクセルファイルが存在しません。ファイルを配置して下さい。")
            Me.Close()
            Exit Sub
        End If

        If Not System.IO.File.Exists(iniFilePath) Then
            MsgBox("構成ファイルが存在しません。ファイルを配置して下さい。")
            Me.Close()
            Exit Sub
        End If

        'If Not System.IO.File.Exists(imageFilePath) Then
        '    MsgBox("画像ファイルが存在しません。ファイルを配置して下さい。")
        '    Me.Close()
        '    Exit Sub
        'End If

        Me.WindowState = FormWindowState.Maximized
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        '画像の配置処理
        '
        '

    End Sub

    Private Sub btnWork_Click(sender As Object, e As EventArgs) Handles btnWork.Click
        If IsNothing(workForm) OrElse workForm.IsDisposed Then
            workForm = New 勤務割()
            workForm.Owner = Me
            workForm.Show()
        End If
    End Sub

    Private Sub btnWeek_Click(sender As System.Object, e As System.EventArgs) Handles btnWeek.Click
        If IsNothing(weekForm) OrElse weekForm.IsDisposed Then
            weekForm = New 週間表(lblday.Text, lblFloor.Text)
            weekForm.Owner = Me
            weekForm.Show()
        End If
    End Sub

    Private Sub btnCSV_Click(sender As System.Object, e As System.EventArgs) Handles btnCsv.Click
        If IsNothing(csvForm) OrElse csvForm.IsDisposed Then
            csvForm = New CSV作成()
            csvForm.Owner = Me
            csvForm.Show()
        End If
    End Sub
End Class
