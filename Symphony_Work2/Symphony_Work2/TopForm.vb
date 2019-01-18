Public Class TopForm
    'データベースのパス
    Public dbFilePath As String = My.Application.Info.DirectoryPath & "\Work2.mdb"
    Public DB_Work As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbFilePath

    'エクセルのパス
    Public excelFilePass As String = My.Application.Info.DirectoryPath & "\Work2.xls"

    '.iniファイルのパス
    Public iniFilePath As String = My.Application.Info.DirectoryPath & "\Work2.ini"

    '画像パス
    Public imageFilePath As String = My.Application.Info.DirectoryPath & "\Work2.png"

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

        If Not System.IO.File.Exists(imageFilePath) Then
            MsgBox("画像ファイルが存在しません。ファイルを配置して下さい。")
            Me.Close()
            Exit Sub
        End If

        '画面サイズ等
        Me.WindowState = FormWindowState.Maximized
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        '画像の配置処理
        topPicture.ImageLocation = imageFilePath

        '印刷ラジオボタンの初期設定
        initPrintState()

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

    Private Sub initPrintState()
        Dim state As String = Util.getIniString("System", "Printer", iniFilePath)
        If state = "Y" Then
            rbtnPrintout.Checked = True
        Else
            rbtnPreview.Checked = True
        End If
    End Sub

    Private Sub rbtnPreview_CheckedChanged(sender As Object, e As System.EventArgs) Handles rbtnPreview.CheckedChanged
        If rbtnPreview.Checked = True Then
            Util.putIniString("System", "Printer", "N", iniFilePath)
        End If
    End Sub

    Private Sub rbtnPrint_CheckedChanged(sender As Object, e As System.EventArgs) Handles rbtnPrintout.CheckedChanged
        If rbtnPrintout.Checked = True Then
            Util.putIniString("System", "Printer", "Y", iniFilePath)
        End If
    End Sub

    Private Sub topPicture_Click(sender As System.Object, e As System.EventArgs) Handles topPicture.Click
        Me.Close()
    End Sub
End Class

