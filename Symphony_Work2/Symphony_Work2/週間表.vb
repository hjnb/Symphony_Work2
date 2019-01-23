Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Public Class 週間表
    Private DGV1Table As DataTable
    Private DGV2Table As DataTable
    Private cellStyle1 As DataGridViewCellStyle
    Private cellStyle2 As DataGridViewCellStyle
    Private clearCellStyle As DataGridViewCellStyle
    Private whiteCellStyle As DataGridViewCellStyle
    Private pinkCellStyle As DataGridViewCellStyle
    Private Const HEISEI_Str As String = "H"
    Private Const NEXT_WAREKI As String = "X"
    Public clr1, clr2, clr3, clr4, clr5, clr6, clr7, clr8, clr9, clr10, clr11, clr12, clr13, clr14, clr15, clr16, clr17, clr18, clr19, clr20, clr21, clr22, clr23, clr24, clr25, clr26, clr27, clr28, clr29, clr30, clr31, clr32, clr33, clr34, clr35, clr36, clr37, clr38, clr39, clr40, clr41, clr42, clr43, clr44, clr45, clr46, clr47, clr48, clr49, clr50, clr51, clr52, clr53, clr54, clr55, clr56, clr57, clr58, clr59, clr60, clr61, clr62, clr63, clr64, clr65, clr66, clr67, clr68, clr69, clr70, clr71, clr72, clr73, clr74, clr75, clr76, clr77, clr78 As Integer
    Private startday As String
    Private floor As Integer

    Public Sub New(startday As String, floor As Integer)
        InitializeComponent()

        Me.startday = startday
        Me.floor = floor

    End Sub

    Private Sub MadeStyle()
        '文字の大きさ指定
        cellStyle1 = New DataGridViewCellStyle()
        cellStyle1.Font = New Font("MS UI Gothic", 6)
        cellStyle1.ForeColor = Color.Blue
        cellStyle1.BackColor = Color.FromArgb(234, 234, 234)
        cellStyle1.Alignment = DataGridViewContentAlignment.MiddleCenter

        cellStyle2 = New DataGridViewCellStyle()
        cellStyle2.Font = New Font("MS UI Gothic", 8)
        cellStyle2.ForeColor = Color.Blue
        cellStyle2.Alignment = DataGridViewContentAlignment.MiddleCenter

        clearCellStyle = New DataGridViewCellStyle()
        clearCellStyle.BackColor = Color.FromArgb(234, 234, 234)
        clearCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        whiteCellStyle = New DataGridViewCellStyle()
        whiteCellStyle.Font = New Font("MS UI Gothic", 8)
        whiteCellStyle.BackColor = Color.FromArgb(255, 255, 255)
        whiteCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        pinkCellStyle = New DataGridViewCellStyle()
        pinkCellStyle.Font = New Font("MS UI Gothic", 8)
        pinkCellStyle.BackColor = Color.FromArgb(255, 192, 255)
        pinkCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

    End Sub

    Private Sub 週間表_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.Alt = True Then
            If e.KeyCode = Keys.F12 Then
                btnTouroku.Visible = True
                btnSakujo.Visible = True
                btnInnsatu.Visible = True
                btnTorikomi.Visible = True
                Dim Staff As New 職員リスト()
                Staff.Owner = Me
                Staff.Show()
            End If
        End If
    End Sub

    Private Sub 週間表_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        MadeStyle()
        Me.WindowState = FormWindowState.Maximized
        'DataGridView1の設定
        Dim Cn As New OleDbConnection(TopForm.DB_Work2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        DGV1Table = New DataTable()
        Util.EnableDoubleBuffering(DataGridView1)

        With DataGridView1
            .RowTemplate.Height = 20
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .ColumnHeadersVisible = False
            .RowHeadersVisible = False
            .DefaultCellStyle.SelectionForeColor = Color.Black
            .DefaultCellStyle.Font = New Font("MS UI Gothic", 7)
        End With

        'DataGridView1列作成
        For i As Integer = 0 To 28
            DGV1Table.Columns.Add("a" & i, Type.GetType("System.String"))
        Next

        'DataGridView1行作成
        For i As Integer = 0 To 39
            DGV1Table.Rows.Add(DGV1Table.NewRow())
        Next

        'DataGridView1空を表示
        DataGridView1.DataSource = DGV1Table

        'DataGridView1列の設定
        For c As Integer = 0 To 28
            If c = 0 Then
                DataGridView1.Columns(c).Width = 30
            ElseIf c Mod 2 = 0 Then '偶数列
                DataGridView1.Columns(c).Width = 55
                DataGridView1.Columns(c).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            ElseIf c = 1 OrElse c = 5 OrElse c = 9 OrElse c = 13 OrElse c = 17 OrElse c = 21 OrElse c = 25 Then     '各日の左列
                DataGridView1.Columns(c).Width = 30
            ElseIf c = 3 OrElse c = 7 OrElse c = 11 OrElse c = 15 OrElse c = 19 OrElse c = 23 OrElse c = 27 Then    '各日の右列
                DataGridView1.Columns(c).Width = 22
            End If
        Next

        'DataGridView1の行の設定
        For r As Integer = 0 To 39
            DataGridView1.Rows(r).Height = 15
        Next

        'DataGridView2の設定
        Dim SQLCm2 As OleDbCommand = Cn.CreateCommand
        Dim Adapter2 As New OleDbDataAdapter(SQLCm2)
        DGV2Table = New DataTable()
        Util.EnableDoubleBuffering(DataGridView2)

        With DataGridView2
            '.RowTemplate.Height = 20
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .ColumnHeadersVisible = False
            .RowHeadersVisible = False
            .DefaultCellStyle.SelectionForeColor = Color.Black
            .DefaultCellStyle.Font = New Font("MS UI Gothic", 8)
        End With

        'DataGridView2の列作成
        For i As Integer = 0 To 6
            DGV2Table.Columns.Add("a" & i, Type.GetType("System.String"))
        Next

        'DataGridView2の行作成
        For i As Integer = 1 To 5
            DGV2Table.Rows.Add(DGV2Table.NewRow())
        Next

        'DataGridView2の空を表示
        DataGridView2.DataSource = DGV2Table

        'DataGridView2の列の設定
        For c As Integer = 0 To 6
            DataGridView2.Columns(c).Width = 162
        Next

        'DataGridView2の行の設定
        For r As Integer = 0 To 4
            DataGridView2.Rows(r).Height = 15
        Next

        KeyPreview = True

        '日付の設定
        If startday = "日付" OrElse startday = "" Then
            Dim ymd As Date = Today
            Dim weekNumber As DayOfWeek = ymd.DayOfWeek
            For i As Integer = 0 To 6
                If weekNumber = i Then
                    ymd = ymd.AddDays(-i)
                    lblYmd.Text = ChangeWareki(ymd)
                End If
            Next
            lblYmd.Text = lblYmd.Text & "（日）"
        Else
            lblYmd.Text = startday
        End If


        '各セルのスタイルの設定
        With DataGridView1
            With .Rows(0)
                .ReadOnly = True
                .DefaultCellStyle = clearCellStyle
            End With
            With .Columns(0)
                .ReadOnly = True
                .DefaultCellStyle = clearCellStyle
            End With
            For i As Integer = 1 To 28 Step 2
                .Columns(i).ReadOnly = True
                .Columns(i).DefaultCellStyle = cellStyle1
            Next
        End With

        For row As Integer = 10 To 22
            If row = 10 OrElse row = 16 OrElse row = 22 Then
                For col As Integer = 1 To 28
                    DataGridView1(col, row).Style = whiteCellStyle
                    DataGridView1(col, row).ReadOnly = False
                Next
            End If
        Next

        '各セルの固定値部分の設定
        Dim Youbi() As String = {"日", "月", "火", "水", "木", "金", "土"}
        For i As Integer = 0 To 6
            DataGridView1(i * 4 + 3, 0).Value = Youbi(i)
            DataGridView1(i * 4 + 1, 1).Value = "風呂"
            DataGridView1(i * 4 + 3, 1).Value = "誘"
            DataGridView1(i * 4 + 1, 2).Value = "風呂"
            DataGridView1(i * 4 + 3, 2).Value = "誘"
            DataGridView1(i * 4 + 1, 3).Value = "ｱｸﾞﾚ"
            DataGridView1(i * 4 + 3, 3).Value = "１"
            DataGridView1(i * 4 + 1, 4).Value = "ﾎﾟｼﾃ"
            DataGridView1(i * 4 + 3, 4).Value = "２"
            DataGridView1(i * 4 + 1, 38).Value = "夜"
            DataGridView1(i * 4 + 1, 39).Value = "深"
            DataGridView1(i * 4 + 3, 0).Style = cellStyle2
            DataGridView1(i * 4 + 1, 1).Style = cellStyle2
            DataGridView1(i * 4 + 3, 1).Style = cellStyle2
            DataGridView1(i * 4 + 1, 2).Style = cellStyle2
            DataGridView1(i * 4 + 3, 2).Style = cellStyle2
            DataGridView1(i * 4 + 1, 3).Style = cellStyle2
            DataGridView1(i * 4 + 3, 3).Style = cellStyle2
            DataGridView1(i * 4 + 1, 4).Style = cellStyle2
            DataGridView1(i * 4 + 3, 4).Style = cellStyle2
            DataGridView1(i * 4 + 1, 38).Style = cellStyle2
            DataGridView1(i * 4 + 1, 39).Style = cellStyle2
            For r As Integer = 1 To 3
                DataGridView1(i * 4 + 1, r * 6 - 1).Value = "早"
                DataGridView1(i * 4 + 1, r * 6).Value = "日早"
                DataGridView1(i * 4 + 1, r * 6 + 1).Value = "日遅"
                DataGridView1(i * 4 + 1, r * 6 + 2).Value = "遅"
                DataGridView1(i * 4 + 1, r * 6 + 3).Value = "遅々"
                DataGridView1(i * 4 + 1, r * 5 + 18).Value = "朝"
                DataGridView1(i * 4 + 1, r * 5 + 20).Value = "昼"
                DataGridView1(i * 4 + 1, r * 5 + 22).Value = "夕"
                DataGridView1(i * 4 + 1, r * 6 - 1).Style = cellStyle2
                DataGridView1(i * 4 + 1, r * 6).Style = cellStyle2
                DataGridView1(i * 4 + 1, r * 6 + 1).Style = cellStyle2
                DataGridView1(i * 4 + 1, r * 6 + 2).Style = cellStyle2
                DataGridView1(i * 4 + 1, r * 6 + 3).Style = cellStyle2
                DataGridView1(i * 4 + 1, r * 5 + 18).Style = cellStyle2
                DataGridView1(i * 4 + 1, r * 5 + 20).Style = cellStyle2
                DataGridView1(i * 4 + 1, r * 5 + 22).Style = cellStyle2
            Next
        Next

        Dim Moji As String() = {"ＡＭ", "ＰＭ", "学", "習", "丘", "P", "虹", "P", "光", "P", "丘", "虹", "光", "夜勤", "深夜"}
        Dim Gyo As Integer() = {1, 2, 3, 4, 7, 10, 13, 16, 19, 22, 25, 30, 35, 38, 39}

        For n As Integer = 0 To 14
            DataGridView1(0, Gyo(n)).Style = cellStyle2
            DataGridView1(0, Gyo(n)).Value = Moji(n)
        Next

        If floor = 2 Then

        Else
            rbn3F.Checked = True
        End If

        'DataGridView2(0, 0).Selected = False
    End Sub

    Private Sub DataGridView_CellPainting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles DataGridView1.CellPainting, DataGridView2.CellPainting
        '選択したセルに枠を付ける
        If e.ColumnIndex >= 0 AndAlso e.RowIndex >= 0 AndAlso (e.PaintParts And DataGridViewPaintParts.Background) = DataGridViewPaintParts.Background Then
            e.Graphics.FillRectangle(New SolidBrush(e.CellStyle.BackColor), e.CellBounds)

            If (e.PaintParts And DataGridViewPaintParts.SelectionBackground) = DataGridViewPaintParts.SelectionBackground AndAlso (e.State And DataGridViewElementStates.Selected) = DataGridViewElementStates.Selected Then
                e.Graphics.DrawRectangle(New Pen(Color.Black, 2I), e.CellBounds.X + 1I, e.CellBounds.Y + 1I, e.CellBounds.Width - 3I, e.CellBounds.Height - 3I)
            End If

            Dim pParts As DataGridViewPaintParts = e.PaintParts And Not DataGridViewPaintParts.Background
            e.Paint(e.ClipBounds, pParts)
            e.Handled = True
        End If
    End Sub

    Public Function ChangeWareki(ymd As Date) As String
        Dim wareki As String = ""
        Dim Result As String = ""
        If ymd <= "2019/04/30" Then
            wareki = "H"
            Dim YY As String = (Val(Strings.Left(ymd, 4)) - 1988)
            If YY.Length = 1 Then
                YY = "0" & (Val(Strings.Left(ymd, 4)) - 1988)
            End If
            Result = wareki & YY & "/" & Strings.Right(ymd, 5)
        ElseIf ymd > "2019/04/30" Then
            wareki = NEXT_WAREKI
            Dim YY As String = (Val(Strings.Left(ymd, 4)) - 2018)
            If YY.Length = 1 Then
                YY = "0" & (Val(Strings.Left(ymd, 4)) - 2018)
            End If
            Result = wareki & YY & "/" & Strings.Right(ymd, 5)
        End If
        Return Result
    End Function

    Public Shared Function ChangeSeireki(ymd As String) As String
        Dim Seireki As Integer
        If Strings.Left(ymd, 1) = "H" Then
            Seireki = Val(Strings.Mid(ymd, 2, 2) + 1988)
        ElseIf Strings.Left(ymd, 1) = NEXT_WAREKI Then
            Seireki = Val(Strings.Mid(ymd, 2, 2) + 2018)
        End If

        Return Seireki
    End Function

    Private Sub rbn2F_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbn2F.CheckedChanged
        rbn2F.BackColor = Color.Yellow
        rbn3F.BackColor = SystemColors.Control

        TopForm.lblday.Text = lblYmd.Text
        TopForm.lblFloor.Text = "2"
    End Sub

    Private Sub rbn3F_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbn3F.CheckedChanged
        rbn2F.BackColor = SystemColors.Control
        rbn3F.BackColor = Color.Yellow

        TopForm.lblday.Text = lblYmd.Text
        TopForm.lblFloor.Text = "3"
        ChangeForm()
    End Sub

    Private Sub ChangeForm()
        '2階と3階で共通の部分
        DataGridView1.Columns.Clear()

        Dim Cn As New OleDbConnection(TopForm.DB_Work2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        DGV1Table = New DataTable()
        Util.EnableDoubleBuffering(DataGridView1)

        With DataGridView1
            .RowTemplate.Height = 20
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .ColumnHeadersVisible = False
            .RowHeadersVisible = False
            .DefaultCellStyle.SelectionForeColor = Color.Black
            .DefaultCellStyle.Font = New Font("MS UI Gothic", 7)
        End With

        'DataGridView1列作成
        For i As Integer = 0 To 28
            DGV1Table.Columns.Add("a" & i, Type.GetType("System.String"))
        Next

        'DataGridViewのサイズを決める
        If rbn2F.Checked = True Then    '2階の情報を表示
            Label8.Visible = True
            Label9.Visible = True
            Label2.Location = New Point(18, 95)
            Label3.Location = New Point(18, 125)
            Label4.Location = New Point(18, 214)
            Label5.Location = New Point(18, 304)
            Label6.Location = New Point(18, 394)
            Label7.Location = New Point(18, 469)
            Label8.Location = New Point(18, 544)
            Label9.Location = New Point(18, 620)

            For i As Integer = 10 To 16
                Controls("Label" & i).Size = New Size(2, 603)
            Next

            DataGridView1.Location = New Point(18, 49)
            DataGridView1.Size = New Size(1167, 603)
            DataGridView2.Location = New Point(48, 651)

            'DataGridView1行作成
            For i As Integer = 0 To 39
                DGV1Table.Rows.Add(DGV1Table.NewRow())
            Next

        ElseIf rbn3F.Checked = True Then    '3階の情報を表示
            Label2.Location = New Point(18, 95)
            Label3.Location = New Point(18, 125)
            Label4.Location = New Point(18, 214)
            Label5.Location = New Point(18, 304)
            Label6.Location = New Point(18, 379)
            Label7.Location = New Point(18, 454)
            Label8.Visible = False
            Label9.Visible = False
            For i As Integer = 10 To 16
                Controls("Label" & i).Size = New Size(2, 438)
            Next

            DataGridView1.Location = New Point(18, 49)
            DataGridView1.Size = New Size(1167, 438)
            DataGridView2.Location = New Point(48, 485)

            'DataGridView1行作成
            For i As Integer = 0 To 28
                DGV1Table.Rows.Add(DGV1Table.NewRow())
            Next
        End If

        '2階と3階で共通の部分

        'DataGridView1空を表示
        DataGridView1.DataSource = DGV1Table

        'DataGridView1列の設定
        For c As Integer = 0 To 28
            If c = 0 Then
                DataGridView1.Columns(c).Width = 30
            ElseIf c Mod 2 = 0 Then '偶数列
                DataGridView1.Columns(c).Width = 55
                DataGridView1.Columns(c).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            ElseIf c = 1 OrElse c = 5 OrElse c = 9 OrElse c = 13 OrElse c = 17 OrElse c = 21 OrElse c = 25 Then     '各日の左列
                DataGridView1.Columns(c).Width = 30
            ElseIf c = 3 OrElse c = 7 OrElse c = 11 OrElse c = 15 OrElse c = 19 OrElse c = 23 OrElse c = 27 Then    '各日の右列
                DataGridView1.Columns(c).Width = 22
            End If
        Next

        '2階と3階で共通の部分

        '各セルのスタイルの設定
        With DataGridView1
            With .Rows(0)
                .ReadOnly = True
                .DefaultCellStyle = clearCellStyle
            End With
            With .Columns(0)
                .ReadOnly = True
                .DefaultCellStyle = clearCellStyle
            End With
            For i As Integer = 1 To 28 Step 2
                .Columns(i).ReadOnly = True
                .Columns(i).DefaultCellStyle = cellStyle1
            Next
        End With

        If rbn2F.Checked = True Then    '2階の情報を表示
            'DataGridView1行の設定
            For r As Integer = 0 To 39
                DataGridView1.Rows(r).Height = 15
            Next

            For row As Integer = 10 To 22
                If row = 10 OrElse row = 16 OrElse row = 22 Then
                    For col As Integer = 1 To 28
                        DataGridView1(col, row).Style = whiteCellStyle
                        DataGridView1(col, row).ReadOnly = False
                    Next
                End If
            Next
        ElseIf rbn3F.Checked = True Then    '3階の情報を表示
            'DataGridView1行の設定
            For r As Integer = 0 To 28
                DataGridView1.Rows(r).Height = 15
            Next

            For row As Integer = 10 To 16
                If row = 10 OrElse row = 16 Then
                    For col As Integer = 1 To 28
                        DataGridView1(col, row).Style = whiteCellStyle
                        DataGridView1(col, row).ReadOnly = False
                    Next
                End If
            Next
        End If

        If rbn2F.Checked = True Then    '2階の情報を表示
            '各セルの固定値部分の設定
            Dim Youbi() As String = {"日", "月", "火", "水", "木", "金", "土"}
            For i As Integer = 0 To 6
                DataGridView1(i * 4 + 3, 0).Value = Youbi(i)
                DataGridView1(i * 4 + 1, 1).Value = "風呂"
                DataGridView1(i * 4 + 3, 1).Value = "誘"
                DataGridView1(i * 4 + 1, 2).Value = "風呂"
                DataGridView1(i * 4 + 3, 2).Value = "誘"
                DataGridView1(i * 4 + 1, 3).Value = "ｱｸﾞﾚ"
                DataGridView1(i * 4 + 3, 3).Value = "１"
                DataGridView1(i * 4 + 1, 4).Value = "ﾎﾟｼﾃ"
                DataGridView1(i * 4 + 3, 4).Value = "２"
                DataGridView1(i * 4 + 1, 38).Value = "夜"
                DataGridView1(i * 4 + 1, 39).Value = "深"
                DataGridView1(i * 4 + 3, 0).Style = cellStyle2
                DataGridView1(i * 4 + 1, 1).Style = cellStyle2
                DataGridView1(i * 4 + 3, 1).Style = cellStyle2
                DataGridView1(i * 4 + 1, 2).Style = cellStyle2
                DataGridView1(i * 4 + 3, 2).Style = cellStyle2
                DataGridView1(i * 4 + 1, 3).Style = cellStyle2
                DataGridView1(i * 4 + 3, 3).Style = cellStyle2
                DataGridView1(i * 4 + 1, 4).Style = cellStyle2
                DataGridView1(i * 4 + 3, 4).Style = cellStyle2
                DataGridView1(i * 4 + 1, 38).Style = cellStyle2
                DataGridView1(i * 4 + 1, 39).Style = cellStyle2
                For r As Integer = 1 To 3
                    DataGridView1(i * 4 + 1, r * 6 - 1).Value = "早"
                    DataGridView1(i * 4 + 1, r * 6).Value = "日早"
                    DataGridView1(i * 4 + 1, r * 6 + 1).Value = "日遅"
                    DataGridView1(i * 4 + 1, r * 6 + 2).Value = "遅"
                    DataGridView1(i * 4 + 1, r * 6 + 3).Value = "遅々"
                    DataGridView1(i * 4 + 1, r * 5 + 18).Value = "朝"
                    DataGridView1(i * 4 + 1, r * 5 + 20).Value = "昼"
                    DataGridView1(i * 4 + 1, r * 5 + 22).Value = "夕"
                    DataGridView1(i * 4 + 1, r * 6 - 1).Style = cellStyle2
                    DataGridView1(i * 4 + 1, r * 6).Style = cellStyle2
                    DataGridView1(i * 4 + 1, r * 6 + 1).Style = cellStyle2
                    DataGridView1(i * 4 + 1, r * 6 + 2).Style = cellStyle2
                    DataGridView1(i * 4 + 1, r * 6 + 3).Style = cellStyle2
                    DataGridView1(i * 4 + 1, r * 5 + 18).Style = cellStyle2
                    DataGridView1(i * 4 + 1, r * 5 + 20).Style = cellStyle2
                    DataGridView1(i * 4 + 1, r * 5 + 22).Style = cellStyle2
                Next
            Next

            Dim Moji As String() = {"ＡＭ", "ＰＭ", "学", "習", "丘", "P", "虹", "P", "光", "P", "丘", "虹", "光", "夜勤", "深夜"}
            Dim Gyo As Integer() = {1, 2, 3, 4, 7, 10, 13, 16, 19, 22, 25, 30, 35, 38, 39}

            For n As Integer = 0 To 14
                DataGridView1(0, Gyo(n)).Style = cellStyle2
                DataGridView1(0, Gyo(n)).Value = Moji(n)
            Next

        ElseIf rbn3F.Checked = True Then    '3階の情報を表示
            '各セルの固定値部分の設定
            Dim Youbi() As String = {"日", "月", "火", "水", "木", "金", "土"}
            For i As Integer = 0 To 6
                DataGridView1(i * 4 + 3, 0).Value = Youbi(i)
                DataGridView1(i * 4 + 1, 1).Value = "風呂"
                DataGridView1(i * 4 + 3, 1).Value = "誘"
                DataGridView1(i * 4 + 1, 2).Value = "風呂"
                DataGridView1(i * 4 + 3, 2).Value = "誘"
                DataGridView1(i * 4 + 1, 3).Value = "ｱｸﾞﾚ"
                DataGridView1(i * 4 + 3, 3).Value = "１"
                DataGridView1(i * 4 + 1, 4).Value = "ﾎﾟｼﾃ"
                DataGridView1(i * 4 + 3, 4).Value = "２"
                DataGridView1(i * 4 + 1, 27).Value = "夜"
                DataGridView1(i * 4 + 1, 28).Value = "深"
                DataGridView1(i * 4 + 3, 0).Style = cellStyle2
                DataGridView1(i * 4 + 1, 1).Style = cellStyle2
                DataGridView1(i * 4 + 3, 1).Style = cellStyle2
                DataGridView1(i * 4 + 1, 2).Style = cellStyle2
                DataGridView1(i * 4 + 3, 2).Style = cellStyle2
                DataGridView1(i * 4 + 1, 3).Style = cellStyle2
                DataGridView1(i * 4 + 3, 3).Style = cellStyle2
                DataGridView1(i * 4 + 1, 4).Style = cellStyle2
                DataGridView1(i * 4 + 3, 4).Style = cellStyle2
                DataGridView1(i * 4 + 1, 27).Style = cellStyle2
                DataGridView1(i * 4 + 1, 28).Style = cellStyle2
                For r As Integer = 1 To 2
                    DataGridView1(i * 4 + 1, r * 6 - 1).Value = "早"
                    DataGridView1(i * 4 + 1, r * 6).Value = "日早"
                    DataGridView1(i * 4 + 1, r * 6 + 1).Value = "日遅"
                    DataGridView1(i * 4 + 1, r * 6 + 2).Value = "遅"
                    DataGridView1(i * 4 + 1, r * 6 + 3).Value = "遅々"
                    DataGridView1(i * 4 + 1, r * 5 + 12).Value = "朝"
                    DataGridView1(i * 4 + 1, r * 5 + 14).Value = "昼"
                    DataGridView1(i * 4 + 1, r * 5 + 16).Value = "夕"
                    DataGridView1(i * 4 + 1, r * 6 - 1).Style = cellStyle2
                    DataGridView1(i * 4 + 1, r * 6).Style = cellStyle2
                    DataGridView1(i * 4 + 1, r * 6 + 1).Style = cellStyle2
                    DataGridView1(i * 4 + 1, r * 6 + 2).Style = cellStyle2
                    DataGridView1(i * 4 + 1, r * 6 + 3).Style = cellStyle2
                    DataGridView1(i * 4 + 1, r * 5 + 12).Style = cellStyle2
                    DataGridView1(i * 4 + 1, r * 5 + 14).Style = cellStyle2
                    DataGridView1(i * 4 + 1, r * 5 + 16).Style = cellStyle2
                Next
            Next

            Dim Moji As String() = {"ＡＭ", "ＰＭ", "学", "習", "雪", "P", "風", "P", "雪", "風", "夜勤", "深夜"}
            Dim Gyo As Integer() = {1, 2, 3, 4, 7, 10, 13, 16, 19, 24, 27, 28}

            For n As Integer = 0 To 11
                DataGridView1(0, Gyo(n)).Style = cellStyle2
                DataGridView1(0, Gyo(n)).Value = Moji(n)
            Next
        End If

        '2階と3階で共通の部分
        For i As Integer = 0 To 6
            DataGridView1(4 * i + 2, 0).Value = Val(Strings.Mid(lblYmd.Text, 8, 2)) + i
        Next

        Dim Getumatu As Integer = Date.DaysInMonth(ChangeSeireki(Strings.Left(lblYmd.Text, 9)), Val(Strings.Mid(lblYmd.Text, 5, 2)))

        For i As Integer = 0 To 6
            If Val(DataGridView1(4 * i + 2, 0).Value) > Getumatu Then
                DataGridView1(4 * i + 2, 0).Value = Val(DataGridView1(4 * i + 2, 0).Value) - Getumatu
            End If
        Next

        DataIndication()

    End Sub

    Private Sub btnUp_Click(sender As System.Object, e As System.EventArgs) Handles btnUp.Click
        Dim ymd As Date = ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 5)
        ymd = ymd.AddDays(7)
        lblYmd.Text = ChangeWareki(ymd) & "（日）"
    End Sub

    Private Sub btnDown_Click(sender As System.Object, e As System.EventArgs) Handles btnDown.Click
        Dim ymd As Date = ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 5)
        ymd = ymd.AddDays(-7)
        lblYmd.Text = ChangeWareki(ymd) & "（日）"
    End Sub

    Private Sub lblYmd_TextChanged(sender As Object, e As System.EventArgs) Handles lblYmd.TextChanged
        If Strings.Left(lblYmd.Text, 9) = "H30.12.31" Then
            Return
        End If
        For i As Integer = 0 To 6
            DataGridView1(4 * i + 2, 0).Value = Val(Strings.Mid(lblYmd.Text, 8, 2)) + i
        Next

        Dim Getumatu As Integer = Date.DaysInMonth(ChangeSeireki(Strings.Left(lblYmd.Text, 9)), Val(Strings.Mid(lblYmd.Text, 5, 2)))

        For i As Integer = 0 To 6
            If Val(DataGridView1(4 * i + 2, 0).Value) > Getumatu Then
                DataGridView1(4 * i + 2, 0).Value = Val(DataGridView1(4 * i + 2, 0).Value) - Getumatu
            End If
        Next

        DataIndication()
    End Sub

    Private Sub DataClear()
        If rbn2F.Checked = True Then
            For column As Integer = 1 To 28
                If column Mod 2 = 0 Then    '偶数列
                    For row As Integer = 1 To 39    '入力可能な行をチェック
                        DataGridView1(column, row).Value = ""
                        DataGridView1(column, row).Style = whiteCellStyle
                    Next
                Else
                    For row As Integer = 10 To 22
                        If row = 10 OrElse row = 16 OrElse row = 22 Then
                            DataGridView1(column, row).Value = ""
                        End If
                    Next
                End If
            Next
        ElseIf rbn3F.Checked = True Then
            For column As Integer = 1 To 28
                If column Mod 2 = 0 Then    '偶数列
                    For row As Integer = 1 To 28    '入力可能な行をチェック
                        DataGridView1(column, row).Value = ""
                        DataGridView1(column, row).Style = whiteCellStyle
                    Next
                Else
                    For row As Integer = 10 To 16
                        If row = 10 OrElse row = 16 Then
                            DataGridView1(column, row).Value = ""
                        End If
                    Next
                End If
            Next
        End If


        For c2 As Integer = 0 To 6
            For r2 As Integer = 0 To 4
                DataGridView2(c2, r2).Value = ""
            Next
        Next
    End Sub

    Private Sub DataIndication()
        DataClear()

        If rbn2F.Checked = True Then    '2階
            Dim Ymd As Date = ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 5)
            Dim YmdAdd7 As Date = Ymd.AddDays(6)

            Dim cnn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim sql As String = "select * from ASHyo WHERE #" & Ymd & "# <= Ymd and Ymd <= #" & YmdAdd7 & "# order by Ymd"
            cnn.Open(TopForm.DB_Work2)
            rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

            'Datagridview1への表示
            Dim ColumnsNo As Integer = 0
            While Not rs.EOF
                For RowNo As Integer = 1 To 39
                    If RowNo = 1 OrElse RowNo = 2 Then
                        'データ表示
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 0).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 2).Value
                        '色付け処理
                        If rs.Fields(RowNo + 83).Value = "1" Then
                            DataGridView1(ColumnsNo * 4 + 2, RowNo).Style = pinkCellStyle
                        ElseIf rs.Fields(RowNo + 83).Value = "0" Then
                            DataGridView1(ColumnsNo * 4 + 2, RowNo).Style = whiteCellStyle
                        End If
                        If rs.Fields(RowNo + 85).Value = "1" Then
                            DataGridView1(ColumnsNo * 4 + 4, RowNo).Style = pinkCellStyle
                        ElseIf rs.Fields(RowNo + 85).Value = "0" Then
                            DataGridView1(ColumnsNo * 4 + 4, RowNo).Style = whiteCellStyle
                        End If
                    ElseIf RowNo = 3 OrElse RowNo = 4 Then
                        'データ表示
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 2).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 4).Value
                        '色付け処理
                        If rs.Fields(RowNo + 85).Value = "1" Then
                            DataGridView1(ColumnsNo * 4 + 2, RowNo).Style = pinkCellStyle
                        ElseIf rs.Fields(RowNo + 85).Value = "0" Then
                            DataGridView1(ColumnsNo * 4 + 2, RowNo).Style = whiteCellStyle
                        End If
                        If rs.Fields(RowNo + 87).Value = "1" Then
                            DataGridView1(ColumnsNo * 4 + 4, RowNo).Style = pinkCellStyle
                        ElseIf rs.Fields(RowNo + 87).Value = "0" Then
                            DataGridView1(ColumnsNo * 4 + 4, RowNo).Style = whiteCellStyle
                        End If
                    ElseIf 5 <= RowNo And RowNo <= 39 Then
                        'データ表示
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo * 2 - 1).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo * 2 + 0).Value
                        If RowNo = 10 Then
                            DGV1Table.Rows(10).Item("a" & ColumnsNo * 4 + 1) = rs.Fields(162).Value
                            DGV1Table.Rows(10).Item("a" & ColumnsNo * 4 + 3) = rs.Fields(163).Value
                        ElseIf RowNo = 16 Then
                            DGV1Table.Rows(16).Item("a" & ColumnsNo * 4 + 1) = rs.Fields(164).Value
                            DGV1Table.Rows(16).Item("a" & ColumnsNo * 4 + 3) = rs.Fields(165).Value
                        ElseIf RowNo = 22 Then
                            DGV1Table.Rows(22).Item("a" & ColumnsNo * 4 + 1) = rs.Fields(166).Value
                            DGV1Table.Rows(22).Item("a" & ColumnsNo * 4 + 3) = rs.Fields(167).Value
                        End If
                        '色付け処理
                        If rs.Fields((RowNo + 41) * 2).Value = "1" Then
                            DataGridView1(ColumnsNo * 4 + 2, RowNo).Style = pinkCellStyle
                        ElseIf rs.Fields((RowNo + 41) * 2).Value = "0" Then
                            DataGridView1(ColumnsNo * 4 + 2, RowNo).Style = whiteCellStyle
                        End If
                        If rs.Fields((RowNo + 41) * 2 + 1).Value = "1" Then
                            DataGridView1(ColumnsNo * 4 + 4, RowNo).Style = pinkCellStyle
                        ElseIf rs.Fields((RowNo + 41) * 2 + 1).Value = "0" Then
                            DataGridView1(ColumnsNo * 4 + 4, RowNo).Style = whiteCellStyle
                        End If
                    End If
                Next

                'Datagridview2への表示
                For rowno2 As Integer = 1 To 5
                    DataGridView2(ColumnsNo, rowno2 - 1).Value = rs.Fields(rowno2 + 78).Value
                Next

                rs.MoveNext()

                ColumnsNo = ColumnsNo + 1
            End While
            cnn.Close()

        ElseIf rbn3F.Checked = True Then    '3階
            Dim Ymd As Date = ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 5)
            Dim YmdAdd7 As Date = Ymd.AddDays(6)

            Dim cnn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim sql As String = "select * from ASHyo3 WHERE #" & Ymd & "# <= Ymd and Ymd <= #" & YmdAdd7 & "# order by Ymd"
            cnn.Open(TopForm.DB_Work2)
            rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

            'Datagridview1への表示
            Dim ColumnsNo As Integer = 0
            While Not rs.EOF
                For RowNo As Integer = 1 To 28
                    If RowNo = 1 OrElse RowNo = 2 Then
                        'データ表示
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 0).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 2).Value
                        '色付け処理
                        If rs.Fields(RowNo + 61).Value = "1" Then
                            DataGridView1(ColumnsNo * 4 + 2, RowNo).Style = pinkCellStyle
                        ElseIf rs.Fields(RowNo + 61).Value = "0" Then
                            DataGridView1(ColumnsNo * 4 + 2, RowNo).Style = whiteCellStyle
                        End If
                        If rs.Fields(RowNo + 63).Value = "1" Then
                            DataGridView1(ColumnsNo * 4 + 4, RowNo).Style = pinkCellStyle
                        ElseIf rs.Fields(RowNo + 63).Value = "0" Then
                            DataGridView1(ColumnsNo * 4 + 4, RowNo).Style = whiteCellStyle
                        End If
                    ElseIf RowNo = 3 OrElse RowNo = 4 Then
                        'データ表示
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 2).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 4).Value
                        '色付け処理
                        If rs.Fields(RowNo + 63).Value = "1" Then
                            DataGridView1(ColumnsNo * 4 + 2, RowNo).Style = pinkCellStyle
                        ElseIf rs.Fields(RowNo + 63).Value = "0" Then
                            DataGridView1(ColumnsNo * 4 + 2, RowNo).Style = whiteCellStyle
                        End If
                        If rs.Fields(RowNo + 65).Value = "1" Then
                            DataGridView1(ColumnsNo * 4 + 4, RowNo).Style = pinkCellStyle
                        ElseIf rs.Fields(RowNo + 65).Value = "0" Then
                            DataGridView1(ColumnsNo * 4 + 4, RowNo).Style = whiteCellStyle
                        End If
                    ElseIf 5 <= RowNo And RowNo <= 28 Then
                        'データ表示
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo * 2 - 1).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo * 2 + 0).Value
                        If RowNo = 10 Then
                            DGV1Table.Rows(10).Item("a" & ColumnsNo * 4 + 1) = rs.Fields(118).Value
                            DGV1Table.Rows(10).Item("a" & ColumnsNo * 4 + 3) = rs.Fields(119).Value
                        ElseIf RowNo = 16 Then
                            DGV1Table.Rows(16).Item("a" & ColumnsNo * 4 + 1) = rs.Fields(120).Value
                            DGV1Table.Rows(16).Item("a" & ColumnsNo * 4 + 3) = rs.Fields(121).Value
                        End If
                        '色付け処理
                        If rs.Fields((RowNo + 30) * 2).Value = "1" Then
                            DataGridView1(ColumnsNo * 4 + 2, RowNo).Style = pinkCellStyle
                        ElseIf rs.Fields((RowNo + 30) * 2).Value = "0" Then
                            DataGridView1(ColumnsNo * 4 + 2, RowNo).Style = whiteCellStyle
                        End If
                        If rs.Fields((RowNo + 30) * 2 + 1).Value = "1" Then
                            DataGridView1(ColumnsNo * 4 + 4, RowNo).Style = pinkCellStyle
                        ElseIf rs.Fields((RowNo + 30) * 2 + 1).Value = "0" Then
                            DataGridView1(ColumnsNo * 4 + 4, RowNo).Style = whiteCellStyle
                        End If
                    End If
                Next

                'Datagridview2への表示
                For rowno2 As Integer = 1 To 5
                    DataGridView2(ColumnsNo, rowno2 - 1).Value = rs.Fields(rowno2 + 56).Value
                Next

                rs.MoveNext()

                ColumnsNo = ColumnsNo + 1
            End While
            cnn.Close()
        End If

    End Sub

    Private Sub btnTouroku_Click(sender As System.Object, e As System.EventArgs) Handles btnTouroku.Click
        If MsgBox("登録してよろしいですか？", MsgBoxStyle.YesNo + vbExclamation, "登録確認") = MsgBoxResult.No Then
            Return
        End If

        Dim cnn As New ADODB.Connection
        cnn.Open(TopForm.DB_Work2)
        Dim Honjitu As Date = ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 5)
        'データ削除
        Dim DelYmd As Date = ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 5)
        Dim DelYmdAdd7 As Date = DelYmd.AddDays(6)
        Dim SQL As String = ""
        If rbn2F.Checked = True Then
            SQL = "DELETE FROM ASHyo WHERE #" & DelYmd & "# <= Ymd and Ymd <= #" & DelYmdAdd7 & "#"
        ElseIf rbn3F.Checked = True Then
            SQL = "DELETE FROM ASHyo3 WHERE #" & DelYmd & "# <= Ymd and Ymd <= #" & DelYmdAdd7 & "#"
        End If
        cnn.Execute(SQL)
        'データ登録
        If rbn2F.Checked = True Then    '2階の登録
            Dim ymd, amhuro, pmhuro, amyu, pmyu, gaka, gakp, no1, no2, oka1, oka2, oka3, oka4, oka5, oka6, oka7, oka8, oka9, oka10, oka11, oka12, nij1, nij2, nij3, nij4, nij5, nij6, nij7, nij8, nij9, nij10, nij11, nij12, hik1, hik2, hik3, hik4, hik5, hik6, hik7, hik8, hik9, hik10, hik11, hik12, okaa1, okaa2, okaa3, okaa4, okah1, okah2, okah3, okah4, okay1, okay2, nija1, nija2, nija3, nija4, nijh1, nijh2, nijh3, nijh4, nijy1, nijy2, hika1, hika2, hika3, hika4, hikh1, hikh2, hikh3, hikh4, hiky1, hiky2, yak1, yak2, sin1, sin2, text1, text2, text3, text4, text5, okap1, okap2, nijp1, nijp2, hikp1, hikp2 As String

            For dd As Integer = 0 To 6
                If dd = 0 Then
                    ymd = Honjitu
                Else
                    ymd = Honjitu.AddDays(dd)
                End If
                amhuro = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 1).Value)
                pmhuro = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 2).Value)
                amyu = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 1).Value)
                pmyu = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 2).Value)
                gaka = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 3).Value)
                gakp = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 4).Value)
                no1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 3).Value)
                no2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 4).Value)
                oka1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 5).Value)
                oka2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 5).Value)
                oka3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 6).Value)
                oka4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 6).Value)
                oka5 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 7).Value)
                oka6 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 7).Value)
                oka7 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 8).Value)
                oka8 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 8).Value)
                oka9 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 9).Value)
                oka10 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 9).Value)
                oka11 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 10).Value)
                oka12 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 10).Value)
                nij1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 11).Value)
                nij2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 11).Value)
                nij3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 12).Value)
                nij4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 12).Value)
                nij5 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 13).Value)
                nij6 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 13).Value)
                nij7 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 14).Value)
                nij8 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 14).Value)
                nij9 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 15).Value)
                nij10 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 15).Value)
                nij11 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 16).Value)
                nij12 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 16).Value)
                hik1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 17).Value)
                hik2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 17).Value)
                hik3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 18).Value)
                hik4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 18).Value)
                hik5 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 19).Value)
                hik6 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 19).Value)
                hik7 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 20).Value)
                hik8 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 20).Value)
                hik9 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 21).Value)
                hik10 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 21).Value)
                hik11 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 22).Value)
                hik12 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 22).Value)
                okaa1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 23).Value)
                okaa2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 23).Value)
                okaa3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 24).Value)
                okaa4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 24).Value)
                okah1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 25).Value)
                okah2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 25).Value)
                okah3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 26).Value)
                okah4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 26).Value)
                okay1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 27).Value)
                okay2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 27).Value)
                nija1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 28).Value)
                nija2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 28).Value)
                nija3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 29).Value)
                nija4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 29).Value)
                nijh1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 30).Value)
                nijh2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 30).Value)
                nijh3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 31).Value)
                nijh4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 31).Value)
                nijy1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 32).Value)
                nijy2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 32).Value)
                hika1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 33).Value)
                hika2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 33).Value)
                hika3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 34).Value)
                hika4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 34).Value)
                hikh1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 35).Value)
                hikh2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 35).Value)
                hikh3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 36).Value)
                hikh4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 36).Value)
                hiky1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 37).Value)
                hiky2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 37).Value)
                yak1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 38).Value)
                yak2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 38).Value)
                sin1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 39).Value)
                sin2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 39).Value)
                text1 = Util.checkDBNullValue(DataGridView2(dd, 0).Value)
                text2 = Util.checkDBNullValue(DataGridView2(dd, 1).Value)
                text3 = Util.checkDBNullValue(DataGridView2(dd, 2).Value)
                text4 = Util.checkDBNullValue(DataGridView2(dd, 3).Value)
                text5 = Util.checkDBNullValue(DataGridView2(dd, 4).Value)
                okap1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 1, 10).Value)
                okap2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 3, 10).Value)
                nijp1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 1, 16).Value)
                nijp2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 3, 16).Value)
                hikp1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 1, 22).Value)
                hikp2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 3, 22).Value)

                For r As Integer = 1 To 39
                    Dim u As Type = GetType(週間表)
                    If r = 1 OrElse r = 2 Then
                        If DataGridView1(dd * 4 + 2, r).Style.BackColor = Color.FromArgb(255, 192, 255) Then
                            u.InvokeMember("clr" & r, Reflection.BindingFlags.SetField, Nothing, Me, {1})
                        Else
                            u.InvokeMember("clr" & r, Reflection.BindingFlags.SetField, Nothing, Me, {0})
                        End If
                        If DataGridView1(dd * 4 + 4, r).Style.BackColor = Color.FromArgb(255, 192, 255) Then
                            u.InvokeMember("clr" & r + 2, Reflection.BindingFlags.SetField, Nothing, Me, {1})
                        Else
                            u.InvokeMember("clr" & r + 2, Reflection.BindingFlags.SetField, Nothing, Me, {0})
                        End If
                    ElseIf r = 3 OrElse r = 4 Then
                        If DataGridView1(dd * 4 + 2, r).Style.BackColor = Color.FromArgb(255, 192, 255) Then
                            u.InvokeMember("clr" & r + 2, Reflection.BindingFlags.SetField, Nothing, Me, {1})
                        Else
                            u.InvokeMember("clr" & r + 2, Reflection.BindingFlags.SetField, Nothing, Me, {0})
                        End If
                        If DataGridView1(dd * 4 + 4, r).Style.BackColor = Color.FromArgb(255, 192, 255) Then
                            u.InvokeMember("clr" & r + 4, Reflection.BindingFlags.SetField, Nothing, Me, {1})
                        Else
                            u.InvokeMember("clr" & r + 4, Reflection.BindingFlags.SetField, Nothing, Me, {0})
                        End If
                    Else
                        If DataGridView1(dd * 4 + 2, r).Style.BackColor = Color.FromArgb(255, 192, 255) Then
                            u.InvokeMember("clr" & r * 2 - 1, Reflection.BindingFlags.SetField, Nothing, Me, {1})
                        Else
                            u.InvokeMember("clr" & r * 2 - 1, Reflection.BindingFlags.SetField, Nothing, Me, {0})
                        End If
                        If DataGridView1(dd * 4 + 4, r).Style.BackColor = Color.FromArgb(255, 192, 255) Then
                            u.InvokeMember("clr" & r * 2, Reflection.BindingFlags.SetField, Nothing, Me, {1})
                        Else
                            u.InvokeMember("clr" & r * 2, Reflection.BindingFlags.SetField, Nothing, Me, {0})
                        End If
                    End If
                Next

                SQL = "INSERT INTO ASHyo VALUES ('" & ymd & "', '" & amhuro & "', '" & pmhuro & "', '" & amyu & "', '" & pmyu & "', '" & gaka & "', '" & gakp & "', '" & no1 & "', '" & no2 & "', '" & oka1 & "', '" & oka2 & "', '" & oka3 & "', '" & oka4 & "', '" & oka5 & "', '" & oka6 & "', '" & oka7 & "', '" & oka8 & "', '" & oka9 & "', '" & oka10 & "', '" & oka11 & "', '" & oka12 & "', '" & nij1 & "', '" & nij2 & "', '" & nij3 & "', '" & nij4 & "', '" & nij5 & "', '" & nij6 & "', '" & nij7 & "', '" & nij8 & "', '" & nij9 & "', '" & nij10 & "', '" & nij11 & "', '" & nij12 & "', '" & hik1 & "', '" & hik2 & "', '" & hik3 & "', '" & hik4 & "', '" & hik5 & "', '" & hik6 & "', '" & hik7 & "', '" & hik8 & "', '" & hik9 & "', '" & hik10 & "', '" & hik11 & "', '" & hik12 & "', '" & okaa1 & "', '" & okaa2 & "', '" & okaa3 & "', '" & okaa4 & "', '" & okah1 & "', '" & okah2 & "', '" & okah3 & "', '" & okah4 & "', '" & okay1 & "', '" & okay2 & "', '" & nija1 & "', '" & nija2 & "', '" & nija3 & "', '" & nija4 & "', '" & nijh1 & "', '" & nijh2 & "', '" & nijh3 & "', '" & nijh4 & "', '" & nijy1 & "', '" & nijy2 & "', '" & hika1 & "', '" & hika2 & "', '" & hika3 & "', '" & hika4 & "', '" & hikh1 & "', '" & hikh2 & "', '" & hikh3 & "', '" & hikh4 & "', '" & hiky1 & "', '" & hiky2 & "', '" & yak1 & "', '" & yak2 & "', '" & sin1 & "', '" & sin2 & "', '" & text1 & "', '" & text2 & "', '" & text3 & "', '" & text4 & "', '" & text5 & "', '" & clr1 & "', '" & clr2 & "', '" & clr3 & "', '" & clr4 & "', '" & clr5 & "', '" & clr6 & "', '" & clr7 & "', '" & clr8 & "', '" & clr9 & "', '" & clr10 & "', '" & clr11 & "', '" & clr12 & "', '" & clr13 & "', '" & clr14 & "', '" & clr15 & "', '" & clr16 & "', '" & clr17 & "', '" & clr18 & "', '" & clr19 & "', '" & clr20 & "', '" & clr21 & "', '" & clr22 & "', '" & clr23 & "', '" & clr24 & "', '" & clr25 & "', '" & clr26 & "', '" & clr27 & "', '" & clr28 & "', '" & clr29 & "', '" & clr30 & "', '" & clr31 & "', '" & clr32 & "', '" & clr33 & "', '" & clr34 & "', '" & clr35 & "', '" & clr36 & "', '" & clr37 & "', '" & clr38 & "', '" & clr39 & "', '" & clr40 & "', '" & clr41 & "', '" & clr42 & "', '" & clr43 & "', '" & clr44 & "', '" & clr45 & "', '" & clr46 & "', '" & clr47 & "', '" & clr48 & "', '" & clr49 & "', '" & clr50 & "', '" & clr51 & "', '" & clr52 & "', '" & clr53 & "', '" & clr54 & "', '" & clr55 & "', '" & clr56 & "', '" & clr57 & "', '" & clr58 & "', '" & clr59 & "', '" & clr60 & "', '" & clr61 & "', '" & clr62 & "', '" & clr63 & "', '" & clr64 & "', '" & clr65 & "', '" & clr66 & "', '" & clr67 & "', '" & clr68 & "', '" & clr69 & "', '" & clr70 & "', '" & clr71 & "', '" & clr72 & "', '" & clr73 & "', '" & clr74 & "', '" & clr75 & "', '" & clr76 & "', '" & clr77 & "', '" & clr78 & "', '" & okap1 & "', '" & okap2 & "', '" & nijp1 & "', '" & nijp2 & "', '" & hikp1 & "', '" & hikp2 & "')"
                cnn.Execute(SQL)
            Next
        ElseIf rbn3F.Checked = True Then    '3階の登録
            Dim ymd, amhuro, pmhuro, amyu, pmyu, gaka, gakp, no1, no2, yuk1, yuk2, yuk3, yuk4, yuk5, yuk6, yuk7, yuk8, yuk9, yuk10, yuk11, yuk12, kaz1, kaz2, kaz3, kaz4, kaz5, kaz6, kaz7, kaz8, kaz9, kaz10, kaz11, kaz12, yuka1, yuka2, yuka3, yuka4, yukh1, yukh2, yukh3, yukh4, yuky1, yuky2, kaza1, kaza2, kaza3, kaza4, kazh1, kazh2, kazh3, kazh4, kazy1, kazy2, yak1, yak2, sin1, sin2, text1, text2, text3, text4, text5, yukp1, yukp2, kazp1, kazp2 As String

            For dd As Integer = 0 To 6
                If dd = 0 Then
                    ymd = Honjitu
                Else
                    ymd = Honjitu.AddDays(dd)
                End If
                amhuro = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 1).Value)
                pmhuro = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 2).Value)
                amyu = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 1).Value)
                pmyu = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 2).Value)
                gaka = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 3).Value)
                gakp = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 4).Value)
                no1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 3).Value)
                no2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 4).Value)
                yuk1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 5).Value)
                yuk2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 5).Value)
                yuk3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 6).Value)
                yuk4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 6).Value)
                yuk5 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 7).Value)
                yuk6 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 7).Value)
                yuk7 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 8).Value)
                yuk8 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 8).Value)
                yuk9 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 9).Value)
                yuk10 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 9).Value)
                yuk11 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 10).Value)
                yuk12 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 10).Value)
                kaz1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 11).Value)
                kaz2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 11).Value)
                kaz3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 12).Value)
                kaz4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 12).Value)
                kaz5 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 13).Value)
                kaz6 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 13).Value)
                kaz7 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 14).Value)
                kaz8 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 14).Value)
                kaz9 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 15).Value)
                kaz10 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 15).Value)
                kaz11 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 16).Value)
                kaz12 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 16).Value)
                yuka1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 17).Value)
                yuka2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 17).Value)
                yuka3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 18).Value)
                yuka4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 18).Value)
                yukh1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 19).Value)
                yukh2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 19).Value)
                yukh3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 20).Value)
                yukh4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 20).Value)
                yuky1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 21).Value)
                yuky2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 21).Value)
                kaza1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 22).Value)
                kaza2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 22).Value)
                kaza3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 23).Value)
                kaza4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 23).Value)
                kazh1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 24).Value)
                kazh2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 24).Value)
                kazh3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 25).Value)
                kazh4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 25).Value)
                kazy1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 26).Value)
                kazy2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 26).Value)
                yak1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 27).Value)
                yak2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 27).Value)
                sin1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 28).Value)
                sin2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 28).Value)
                text1 = Util.checkDBNullValue(DataGridView2(dd, 0).Value)
                text2 = Util.checkDBNullValue(DataGridView2(dd, 1).Value)
                text3 = Util.checkDBNullValue(DataGridView2(dd, 2).Value)
                text4 = Util.checkDBNullValue(DataGridView2(dd, 3).Value)
                text5 = Util.checkDBNullValue(DataGridView2(dd, 4).Value)
                yukp1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 1, 10).Value)
                yukp2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 3, 10).Value)
                kazp1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 1, 16).Value)
                kazp2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 3, 16).Value)

                For r As Integer = 1 To 28
                    Dim u As Type = GetType(週間表)
                    If r = 1 OrElse r = 2 Then
                        If DataGridView1(dd * 4 + 2, r).Style.BackColor = Color.FromArgb(255, 192, 255) Then
                            u.InvokeMember("clr" & r, Reflection.BindingFlags.SetField, Nothing, Me, {1})
                        Else
                            u.InvokeMember("clr" & r, Reflection.BindingFlags.SetField, Nothing, Me, {0})
                        End If
                        If DataGridView1(dd * 4 + 4, r).Style.BackColor = Color.FromArgb(255, 192, 255) Then
                            u.InvokeMember("clr" & r + 2, Reflection.BindingFlags.SetField, Nothing, Me, {1})
                        Else
                            u.InvokeMember("clr" & r + 2, Reflection.BindingFlags.SetField, Nothing, Me, {0})
                        End If
                    ElseIf r = 3 OrElse r = 4 Then
                        If DataGridView1(dd * 4 + 2, r).Style.BackColor = Color.FromArgb(255, 192, 255) Then
                            u.InvokeMember("clr" & r + 2, Reflection.BindingFlags.SetField, Nothing, Me, {1})
                        Else
                            u.InvokeMember("clr" & r + 2, Reflection.BindingFlags.SetField, Nothing, Me, {0})
                        End If
                        If DataGridView1(dd * 4 + 4, r).Style.BackColor = Color.FromArgb(255, 192, 255) Then
                            u.InvokeMember("clr" & r + 4, Reflection.BindingFlags.SetField, Nothing, Me, {1})
                        Else
                            u.InvokeMember("clr" & r + 4, Reflection.BindingFlags.SetField, Nothing, Me, {0})
                        End If
                    Else
                        If DataGridView1(dd * 4 + 2, r).Style.BackColor = Color.FromArgb(255, 192, 255) Then
                            u.InvokeMember("clr" & r * 2 - 1, Reflection.BindingFlags.SetField, Nothing, Me, {1})
                        Else
                            u.InvokeMember("clr" & r * 2 - 1, Reflection.BindingFlags.SetField, Nothing, Me, {0})
                        End If
                        If DataGridView1(dd * 4 + 4, r).Style.BackColor = Color.FromArgb(255, 192, 255) Then
                            u.InvokeMember("clr" & r * 2, Reflection.BindingFlags.SetField, Nothing, Me, {1})
                        Else
                            u.InvokeMember("clr" & r * 2, Reflection.BindingFlags.SetField, Nothing, Me, {0})
                        End If
                    End If
                Next

                SQL = "INSERT INTO ASHyo3 VALUES ('" & ymd & "', '" & amhuro & "', '" & pmhuro & "', '" & amyu & "', '" & pmyu & "', '" & gaka & "', '" & gakp & "', '" & no1 & "', '" & no2 & "', '" & yuk1 & "', '" & yuk2 & "', '" & yuk3 & "', '" & yuk4 & "', '" & yuk5 & "', '" & yuk6 & "', '" & yuk7 & "', '" & yuk8 & "', '" & yuk9 & "', '" & yuk10 & "', '" & yuk11 & "', '" & yuk12 & "', '" & kaz1 & "', '" & kaz2 & "', '" & kaz3 & "', '" & kaz4 & "', '" & kaz5 & "', '" & kaz6 & "', '" & kaz7 & "', '" & kaz8 & "', '" & kaz9 & "', '" & kaz10 & "', '" & kaz11 & "', '" & kaz12 & "', '" & yuka1 & "', '" & yuka2 & "', '" & yuka3 & "', '" & yuka4 & "', '" & yukh1 & "', '" & yukh2 & "', '" & yukh3 & "', '" & yukh4 & "', '" & yuky1 & "', '" & yuky2 & "', '" & kaza1 & "', '" & kaza2 & "', '" & kaza3 & "', '" & kaza4 & "', '" & kazh1 & "', '" & kazh2 & "', '" & kazh3 & "', '" & kazh4 & "', '" & kazy1 & "', '" & kazy2 & "','" & yak1 & "', '" & yak2 & "', '" & sin1 & "', '" & sin2 & "', '" & text1 & "', '" & text2 & "', '" & text3 & "', '" & text4 & "', '" & text5 & "', '" & clr1 & "', '" & clr2 & "', '" & clr3 & "', '" & clr4 & "', '" & clr5 & "', '" & clr6 & "', '" & clr7 & "', '" & clr8 & "', '" & clr9 & "', '" & clr10 & "', '" & clr11 & "', '" & clr12 & "', '" & clr13 & "', '" & clr14 & "', '" & clr15 & "', '" & clr16 & "', '" & clr17 & "', '" & clr18 & "', '" & clr19 & "', '" & clr20 & "', '" & clr21 & "', '" & clr22 & "', '" & clr23 & "', '" & clr24 & "', '" & clr25 & "', '" & clr26 & "', '" & clr27 & "', '" & clr28 & "', '" & clr29 & "', '" & clr30 & "', '" & clr31 & "', '" & clr32 & "', '" & clr33 & "', '" & clr34 & "', '" & clr35 & "', '" & clr36 & "', '" & clr37 & "', '" & clr38 & "', '" & clr39 & "', '" & clr40 & "', '" & clr41 & "', '" & clr42 & "', '" & clr43 & "', '" & clr44 & "', '" & clr45 & "', '" & clr46 & "', '" & clr47 & "', '" & clr48 & "', '" & clr49 & "', '" & clr50 & "', '" & clr51 & "', '" & clr52 & "', '" & clr53 & "', '" & clr54 & "', '" & clr55 & "', '" & clr56 & "', '" & yukp1 & "', '" & yukp2 & "', '" & kazp1 & "', '" & kazp2 & "')"
                cnn.Execute(SQL)
            Next
        End If
        cnn.Close()

        KinnmuwariTouroku()

    End Sub

    Private Sub KinnmuwariTouroku()
        If MsgBox("パートの勤務割に登録してよろしいですか？", MsgBoxStyle.YesNo + vbExclamation, "ﾊﾟｰﾄ勤務割登録確認") = MsgBoxResult.No Then
            Return
        End If

        Dim cnn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim rsnextmonth As New ADODB.Recordset
        Dim rs2 As New ADODB.Recordset
        cnn.Open(TopForm.DB_Work2)
        Dim SQL As String = ""
        Dim SQLnextmonth As String = ""
        Dim SQL2 As String = ""
        Dim updateSQL As String = ""

        Dim M As Date = ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 5)
        M = M.AddMonths(1)
        Dim a As String = M.ToString("yyyy/MM/dd")

        SQL2 = "SELECT * FROM SNam"
        rs2.Open(SQL2, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

        If rbn2F.Checked = True Then
            Dim floar As String = 2
            Dim okaPwork1, okaPname1, okaPwork2, okaPname2, nijPwork1, nijPname1, nijPwork2, nijPname2, hikPwork1, hikPname1, hikPwork2, hikPname2 As String

            SQL = "SELECT * FROM KinD WHERE Ym='" & ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 2) & "' AND (Seq2='00' OR ('" & floar & "0' <= Seq2 AND Seq2 <= '" & floar & "9')) and Rdr = '' order by Seq"
            rs.Open(SQL, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

            If DataGridView1(2, 0).Value > 22 Then
                SQLnextmonth = "SELECT * FROM KinD WHERE YM='" & ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(a, 6, 2) & "' AND (Seq2='00' OR ('" & floar & "0' <= Seq2 AND Seq2 <= '" & floar & "9')) and Rdr = '' order by Seq"
                rsnextmonth.Open(SQLnextmonth, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            End If

            If rs.RecordCount <= 1 Then
                MsgBox("勤務割に該当月の登録データがありません")
            Else
                Dim listSQL As List(Of String) = New List(Of String)
                For dd As Integer = 0 To 6
                    okaPwork1 = Util.checkDBNullValue(DataGridView1(4 * dd + 1, 10).Value)
                    okaPname1 = Util.checkDBNullValue(DataGridView1(4 * dd + 2, 10).Value)
                    okaPwork2 = Util.checkDBNullValue(DataGridView1(4 * dd + 3, 10).Value)
                    okaPname2 = Util.checkDBNullValue(DataGridView1(4 * dd + 4, 10).Value)
                    hikPwork1 = Util.checkDBNullValue(DataGridView1(4 * dd + 1, 16).Value)
                    hikPname1 = Util.checkDBNullValue(DataGridView1(4 * dd + 2, 16).Value)
                    hikPwork2 = Util.checkDBNullValue(DataGridView1(4 * dd + 3, 16).Value)
                    hikPname2 = Util.checkDBNullValue(DataGridView1(4 * dd + 4, 16).Value)
                    nijPwork1 = Util.checkDBNullValue(DataGridView1(4 * dd + 1, 22).Value)
                    nijPname1 = Util.checkDBNullValue(DataGridView1(4 * dd + 2, 22).Value)
                    nijPwork2 = Util.checkDBNullValue(DataGridView1(4 * dd + 3, 22).Value)
                    nijPname2 = Util.checkDBNullValue(DataGridView1(4 * dd + 4, 22).Value)

                    Dim partname() As String = {okaPname1, okaPname2, hikPname1, hikPname2, nijPname1, nijPname2}
                    Dim partwork() As String = {okaPwork1, okaPwork2, hikPwork1, hikPwork2, nijPwork1, nijPwork2}

                    For i As Integer = 0 To 5
                        If partname(i) = "" Then

                        Else
                            Dim Hi As Integer = DataGridView1(4 * dd + 2, 0).Value
                            If Val(DataGridView1(2, 0).Value) - Hi > 20 Then   '月が替わった日以降
                                If partnamecheck(rsnextmonth, rs2, partname(i)) = True Then
                                    updateSQL = "UPDATE KinD SET Y" & Hi & " = '" & partwork(i) & "', J" & Hi & " = '" & partwork(i) & "' WHERE (Nam LIKE '%" & partname(i) & "%') And (YM='" & ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(a, 6, 2) & "')"
                                    listSQL.Add(updateSQL)
                                Else
                                    If i = 0 OrElse i = 1 Then
                                        MsgBox("ﾊﾟｰﾄの勤務割が登録されていません。" & vbCrLf & DataGridView1(4 * dd + 2, 0).Value & "日：丘：" & partname(i))
                                    ElseIf i = 2 OrElse i = 3 Then
                                        MsgBox("ﾊﾟｰﾄの勤務割が登録されていません。" & vbCrLf & DataGridView1(4 * dd + 2, 0).Value & "日：虹" & partname(i))
                                    ElseIf i >= 4 Then
                                        MsgBox("ﾊﾟｰﾄの勤務割が登録されていません。" & vbCrLf & DataGridView1(4 * dd + 2, 0).Value & "日：光：" & partname(i))
                                    End If
                                    Exit Sub
                                End If
                            Else    '同じ月
                                If partnamecheck(rs, rs2, partname(i)) = True Then
                                    updateSQL = "UPDATE KinD SET Y" & Hi & " = '" & partwork(i) & "', J" & Hi & " = '" & partwork(i) & "' WHERE (Nam LIKE '%" & partname(i) & "%') And (YM='" & ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 2) & "')"
                                    listSQL.Add(updateSQL)
                                Else
                                    If i = 0 OrElse i = 1 Then
                                        MsgBox("ﾊﾟｰﾄの勤務割が登録されていません。" & vbCrLf & DataGridView1(4 * dd + 2, 0).Value & "日：丘：" & partname(i))
                                    ElseIf i = 2 OrElse i = 3 Then
                                        MsgBox("ﾊﾟｰﾄの勤務割が登録されていません。" & vbCrLf & DataGridView1(4 * dd + 2, 0).Value & "日：虹" & partname(i))
                                    ElseIf i >= 4 Then
                                        MsgBox("ﾊﾟｰﾄの勤務割が登録されていません。" & vbCrLf & DataGridView1(4 * dd + 2, 0).Value & "日：光：" & partname(i))
                                    End If
                                    Exit Sub
                                End If
                            End If
                        End If
                    Next

                Next

                For Each i As String In listSQL
                    cnn.Execute(i)
                Next

            End If
            cnn.Close()

        ElseIf rbn3F.Checked = True Then
            Dim floar As String = 3
            Dim yukPwork1, yukPname1, yukPwork2, yukPname2, kazPwork1, kazPname1, kazPwork2, kazPname2 As String

            SQL = "SELECT * FROM KinD WHERE Ym='" & ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 2) & "' AND (Seq2='00' OR ('" & floar & "0' <= Seq2 AND Seq2 <= '" & floar & "9')) and Rdr = '' order by Seq"
            rs.Open(SQL, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

            If DataGridView1(2, 0).Value > 22 Then
                SQLnextmonth = "SELECT * FROM KinD WHERE YM='" & ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(a, 6, 2) & "' AND (Seq2='00' OR ('" & floar & "0' <= Seq2 AND Seq2 <= '" & floar & "9')) and Rdr = '' order by Seq"
                rsnextmonth.Open(SQLnextmonth, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            End If

            If rs.RecordCount <= 1 Then
                MsgBox("勤務割に該当月の登録データがありません")
            Else
                Dim listSQL As List(Of String) = New List(Of String)
                For dd As Integer = 0 To 6
                    yukPwork1 = Util.checkDBNullValue(DataGridView1(4 * dd + 1, 10).Value)
                    yukPname1 = Util.checkDBNullValue(DataGridView1(4 * dd + 2, 10).Value)
                    yukPwork2 = Util.checkDBNullValue(DataGridView1(4 * dd + 3, 10).Value)
                    yukPname2 = Util.checkDBNullValue(DataGridView1(4 * dd + 4, 10).Value)
                    kazPwork1 = Util.checkDBNullValue(DataGridView1(4 * dd + 1, 16).Value)
                    kazPname1 = Util.checkDBNullValue(DataGridView1(4 * dd + 2, 16).Value)
                    kazPwork2 = Util.checkDBNullValue(DataGridView1(4 * dd + 3, 16).Value)
                    kazPname2 = Util.checkDBNullValue(DataGridView1(4 * dd + 4, 16).Value)

                    Dim partname() As String = {yukPname1, yukPname2, kazPname1, kazPname2}
                    Dim partwork() As String = {yukPwork1, yukPwork2, kazPwork1, kazPwork2}
                    For i As Integer = 0 To 3
                        If partname(i) = "" Then

                        Else
                            Dim Hi As Integer = DataGridView1(4 * dd + 2, 0).Value
                            If Val(DataGridView1(2, 0).Value) - Hi > 20 Then   '月が替わった日以降
                                If partnamecheck(rsnextmonth, rs2, partname(i)) = True Then
                                    updateSQL = "UPDATE KinD SET Y" & Hi & " = '" & partwork(i) & "', J" & Hi & " = '" & partwork(i) & "' WHERE (Nam LIKE '%" & partname(i) & "%') And (YM='" & ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(a, 6, 2) & "')"
                                    listSQL.Add(updateSQL)
                                Else
                                    If i = 0 OrElse i = 1 Then
                                        MsgBox("ﾊﾟｰﾄの勤務割が登録されていません。" & vbCrLf & DataGridView1(4 * dd + 2, 0).Value & "日：雪：" & partname(i))
                                    ElseIf i = 2 OrElse i = 3 Then
                                        MsgBox("ﾊﾟｰﾄの勤務割が登録されていません。" & vbCrLf & DataGridView1(4 * dd + 2, 0).Value & "日：風：" & partname(i))
                                    End If
                                    Exit Sub
                                End If
                            Else    '同じ月
                                If partnamecheck(rs, rs2, partname(i)) = True Then
                                    updateSQL = "UPDATE KinD SET Y" & Hi & " = '" & partwork(i) & "', J" & Hi & " = '" & partwork(i) & "' WHERE (Nam LIKE '%" & partname(i) & "%') And (YM='" & ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 2) & "')"
                                    listSQL.Add(updateSQL)
                                Else
                                    If i = 0 OrElse i = 1 Then
                                        MsgBox("ﾊﾟｰﾄの勤務割が登録されていません。" & vbCrLf & DataGridView1(4 * dd + 2, 0).Value & "日：雪：" & partname(i))
                                    ElseIf i = 2 OrElse i = 3 Then
                                        MsgBox("ﾊﾟｰﾄの勤務割が登録されていません。" & vbCrLf & DataGridView1(4 * dd + 2, 0).Value & "日：風：" & partname(i))
                                    End If
                                    Exit Sub
                                End If
                            End If
                        End If
                    Next
                Next
                For Each i As String In listSQL
                    cnn.Execute(i)
                Next
            End If
            cnn.Close()
        End If


    End Sub

    Private Function partnamecheck(rs As ADODB.Recordset, rs2 As ADODB.Recordset, Pname As String) As Boolean

        rs.MoveFirst()
        rs2.MoveFirst()

        If Pname <> "" Then
            '勤務割のほうに名前があるか
            While Not rs.EOF
                If System.Text.RegularExpressions.Regex.IsMatch(rs.Fields("Nam").Value, "^" & Pname) = True Then
                    Return True
                End If
                rs.MoveNext()
            End While

            '名前がない場合
            Return False

        End If

        Return False

    End Function

    Private Sub btnSakujo_Click(sender As System.Object, e As System.EventArgs) Handles btnSakujo.Click
        If MsgBox("削除してよろしいですか？", MsgBoxStyle.YesNo + vbExclamation, "削除確認") = MsgBoxResult.Yes Then
            Dim cnn As New ADODB.Connection
            cnn.Open(TopForm.DB_Work2)

            Dim Ymd As Date = ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 5)
            Dim YmdAdd7 As Date = Ymd.AddDays(6)

            Dim SQL As String = ""
            If rbn2F.Checked = True Then
                SQL = "DELETE FROM ASHyo WHERE #" & Ymd & "# <= Ymd and Ymd <= #" & YmdAdd7 & "#"
            ElseIf rbn3F.Checked = True Then
                SQL = "DELETE FROM ASHyo3 WHERE #" & Ymd & "# <= Ymd and Ymd <= #" & YmdAdd7 & "#"
            End If

            cnn.Execute(SQL)
            cnn.Close()

            DataIndication()

        End If
    End Sub

    Private Sub btnTorikomi_Click(sender As System.Object, e As System.EventArgs) Handles btnTorikomi.Click
        Dim Ymd As Date = ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 5)
        Dim YmdAdd7 As Date = Ymd.AddDays(6)

        Dim cnn As New ADODB.Connection
        Dim rs As New ADODB.Recordset

        DataGridView1.CurrentCell = Nothing
        DataGridView2.CurrentCell = Nothing

        If rbn2F.Checked = True Then
            If MsgBox("週間表3階の'学習、夜勤等、備考'を取り込みますか？", MsgBoxStyle.YesNo + vbExclamation, "確認") = MsgBoxResult.Yes Then
                Dim sql As String = "select * from ASHyo3 WHERE #" & Ymd & "# <= Ymd and Ymd <= #" & YmdAdd7 & "# order by Ymd"
                cnn.Open(TopForm.DB_Work2)
                rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

                'Datagridview1への表示
                Dim ColumnsNo As Integer = 0
                While Not rs.EOF
                    For RowNo As Integer = 3 To 39
                        If RowNo = 3 OrElse RowNo = 4 Then
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 2).Value
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 4).Value
                            If rs.Fields(RowNo + 63).Value = 1 Then
                                DataGridView1(ColumnsNo * 4 + 2, RowNo).Style = pinkCellStyle
                            End If
                            If rs.Fields(RowNo + 65).Value = 1 Then
                                DataGridView1(ColumnsNo * 4 + 4, RowNo).Style = pinkCellStyle
                            End If
                        ElseIf RowNo = 38 Then
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 15).Value
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 16).Value
                        ElseIf RowNo = 39 Then
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 16).Value
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 17).Value
                        End If
                    Next

                    'Datagridview2への表示
                    For rowno2 As Integer = 1 To 5
                        DGV2Table.Rows(rowno2 - 1).Item("a" & ColumnsNo) = rs.Fields(rowno2 + 56).Value
                    Next

                    rs.MoveNext()

                    ColumnsNo = ColumnsNo + 1
                End While
                cnn.Close()

            End If
        ElseIf rbn3F.Checked = True Then
            If MsgBox("週間表2階の'学習、夜勤等、備考'を取り込みますか？", MsgBoxStyle.YesNo + vbExclamation, "確認") = MsgBoxResult.Yes Then

                Dim sql As String = "select * from ASHyo WHERE #" & Ymd & "# <= Ymd and Ymd <= #" & YmdAdd7 & "# order by Ymd"
                cnn.Open(TopForm.DB_Work2)
                rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

                'Datagridview1への表示
                Dim ColumnsNo As Integer = 0
                While Not rs.EOF
                    For RowNo As Integer = 3 To 28
                        If RowNo = 3 OrElse RowNo = 4 Then
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 2).Value
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 4).Value
                            If rs.Fields(RowNo + 85).Value = 1 Then
                                DataGridView1(ColumnsNo * 4 + 2, RowNo).Style = pinkCellStyle
                            End If
                            If rs.Fields(RowNo + 87).Value = 1 Then
                                DataGridView1(ColumnsNo * 4 + 4, RowNo).Style = pinkCellStyle
                            End If
                        ElseIf RowNo = 27 Then
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 48).Value
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 49).Value
                        ElseIf RowNo = 28 Then
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 49).Value
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 50).Value
                        End If
                    Next

                    'Datagridview2への表示
                    For rowno2 As Integer = 1 To 5
                        DGV2Table.Rows(rowno2 - 1).Item("a" & ColumnsNo) = rs.Fields(rowno2 + 78).Value
                    Next

                    rs.MoveNext()

                    ColumnsNo = ColumnsNo + 1
                End While
                cnn.Close()

            End If
        End If
    End Sub

    Private Sub btnInnsatu_Click(sender As System.Object, e As System.EventArgs) Handles btnInnsatu.Click
        Dim Ymd As Date = ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 5)
        Dim YmdAdd7 As Date = Ymd.AddDays(6)

        If rbn2F.Checked = True Then        '2階の印刷
            Dim cnn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim sql As String = "select * from ASHyo WHERE #" & Ymd & "# <= Ymd and Ymd <= #" & YmdAdd7 & "# order by Ymd"
            cnn.Open(TopForm.DB_Work2)
            rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

            If rs.RecordCount > 0 Then
                Dim objExcel As Object
                Dim objWorkBooks As Object
                Dim objWorkBook As Object
                Dim oSheets As Object
                Dim oSheet As Object

                objExcel = CreateObject("Excel.Application")
                objWorkBooks = objExcel.Workbooks
                objWorkBook = objWorkBooks.Open(TopForm.excelFilePass)
                oSheets = objWorkBook.Worksheets
                oSheet = objWorkBook.Worksheets("週間表改")

                oSheet.Range("F1").Value = Strings.Left(lblYmd.Text, 6) & "月"

                Dim Cell() As String = {"C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD"}
                Dim ColumnsNo As Integer = 0
                While Not rs.EOF
                    oSheet.Range(Cell(ColumnsNo * 4 + 1) & "2").Value = DataGridView1((ColumnsNo * 4) + 2, 0).Value
                    For RowNo As Integer = 1 To 39
                        If RowNo = 1 OrElse RowNo = 2 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo + 2).Value = rs.Fields(RowNo + 0).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 3) & RowNo + 2).Value = rs.Fields(RowNo + 2).Value
                        ElseIf RowNo = 3 OrElse RowNo = 4 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo + 2).Value = rs.Fields(RowNo + 2).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 3) & RowNo + 2).Value = rs.Fields(RowNo + 4).Value
                        ElseIf 5 <= RowNo And RowNo <= 39 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo + 2).Value = rs.Fields(RowNo * 2 - 1).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo + 2).Value = rs.Fields(RowNo * 2).Value
                        End If
                    Next

                    'Datagridview2への表示
                    For rowno2 As Integer = 1 To 5
                        oSheet.Range(Cell(ColumnsNo * 4) & rowno2 + 41).Value = rs.Fields(rowno2 + 78).Value
                    Next

                    rs.MoveNext()

                    ColumnsNo = ColumnsNo + 1
                End While
                cnn.Close()

                '保存
                objExcel.DisplayAlerts = False

                ' エクセル表示
                objExcel.Visible = True

                '印刷
                If TopForm.rbtnPreview.Checked = True Then
                    oSheet.PrintPreview(1)
                ElseIf TopForm.rbtnPrintout.Checked = True Then
                    oSheet.Printout(1)
                End If

                ' EXCEL解放
                objExcel.Quit()
                Marshal.ReleaseComObject(oSheet)
                Marshal.ReleaseComObject(objWorkBook)
                Marshal.ReleaseComObject(objExcel)
                oSheet = Nothing
                objWorkBook = Nothing
                objExcel = Nothing
            Else
                MsgBox("出力するデータがありません")
            End If

        ElseIf rbn3F.Checked = True Then        '3階の印刷
            Dim cnn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim sql As String = "select * from ASHyo3 WHERE #" & Ymd & "# <= Ymd and Ymd <= #" & YmdAdd7 & "# order by Ymd"
            cnn.Open(TopForm.DB_Work2)
            rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

            If rs.RecordCount > 0 Then
                Dim objExcel As Object
                Dim objWorkBooks As Object
                Dim objWorkBook As Object
                Dim oSheets As Object
                Dim oSheet As Object

                objExcel = CreateObject("Excel.Application")
                objWorkBooks = objExcel.Workbooks
                objWorkBook = objWorkBooks.Open(TopForm.excelFilePass)
                oSheets = objWorkBook.Worksheets
                oSheet = objWorkBook.Worksheets("週間表３改")

                oSheet.Range("F1").Value = Strings.Left(lblYmd.Text, 6) & "月"

                Dim Cell() As String = {"C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD"}
                Dim ColumnsNo As Integer = 0
                While Not rs.EOF
                    oSheet.Range(Cell(ColumnsNo * 4 + 1) & "2").Value = DataGridView1((ColumnsNo * 4) + 2, 0).Value
                    For RowNo As Integer = 1 To 28
                        If RowNo = 1 OrElse RowNo = 2 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo + 2).Value = rs.Fields(RowNo + 0).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 3) & RowNo + 2).Value = rs.Fields(RowNo + 2).Value
                        ElseIf RowNo = 3 OrElse RowNo = 4 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo + 2).Value = rs.Fields(RowNo + 2).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 3) & RowNo + 2).Value = rs.Fields(RowNo + 4).Value
                        ElseIf 5 <= RowNo And RowNo <= 28 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo + 2).Value = rs.Fields(RowNo * 2 - 1).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo + 2).Value = rs.Fields(RowNo * 2).Value
                        End If
                    Next

                    'Datagridview2への表示
                    For rowno2 As Integer = 1 To 7
                        oSheet.Range(Cell(ColumnsNo * 4) & rowno2 + 30).Value = rs.Fields(rowno2 + 56).Value
                        oSheet.Range(Cell(ColumnsNo * 4) & "36").Value = ""
                        oSheet.Range(Cell(ColumnsNo * 4) & "37").Value = ""
                    Next

                    rs.MoveNext()

                    ColumnsNo = ColumnsNo + 1
                End While
                cnn.Close()

                '保存
                objExcel.DisplayAlerts = False

                ' エクセル表示
                objExcel.Visible = True

                '印刷
                If TopForm.rbtnPreview.Checked = True Then
                    oSheet.PrintPreview(1)
                ElseIf TopForm.rbtnPrintout.Checked = True Then
                    oSheet.Printout(1)
                End If

                ' EXCEL解放
                objExcel.Quit()
                Marshal.ReleaseComObject(oSheet)
                Marshal.ReleaseComObject(objWorkBook)
                Marshal.ReleaseComObject(objExcel)
                oSheet = Nothing
                objWorkBook = Nothing
                objExcel = Nothing
            Else
                MsgBox("出力するデータがありません")
            End If
        End If

    End Sub

End Class