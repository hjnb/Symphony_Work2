Imports System.Data.OleDb

Public Class 勤務割

    Private cn As OleDbConnection
    Private sqlCm As OleDbCommand
    Private adapter As OleDbDataAdapter
    Private dt As DataTable
    Private disableCellStyle As DataGridViewCellStyle
    Private namColumnCellStyle As DataGridViewCellStyle
    Private sundayColumnCellStyle As DataGridViewCellStyle
    Private sundayCharCellStyle As DataGridViewCellStyle
    Private workChangeCellStyle As DataGridViewCellStyle
    Private editBeforeCellValue As String
    Private Const MAX_ROW_COUNT As Integer = 50

    Private unitDictionary2F As Dictionary(Of String, String)
    Private unitDictionary3F As Dictionary(Of String, String)
    Private wordDictionary As Dictionary(Of String, String)
    Private dayCharArray() As String = {"日", "月", "火", "水", "木", "金", "土"}

    Private Sub workForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        createCellStyles()
        createDictionary()

        '当月のデータを表示
        'とりあえず今は仮でこれ
        displayWorkTable("2018/04", "2")
    End Sub

    Private Sub displayWorkTable(ymStr As String, floar As String)
        dgvWork.Columns.Clear()
        Dim year As Integer = CInt(ymStr.Split("/")(0))
        Dim month As Integer = CInt(ymStr.Split("/")(1))
        dt = New DataTable()

        cn = New OleDbConnection(TopForm.DB_Work)
        sqlCm = cn.CreateCommand
        adapter = New OleDbDataAdapter(sqlCm)
        sqlCm.CommandText = "SELECT * FROM KinD WHERE YM='" & ymStr & "' AND (Seq2='00' OR ('" & floar & "0' <= Seq2 AND Seq2 <= '" & floar & "9')) order by Seq"
        cn.Open()
        adapter.Fill(dt)
        Dim builder As OleDbCommandBuilder = New OleDbCommandBuilder(adapter)
        adapter.SelectCommand.Connection = cn

        addHenkouRow(dt)
        addDayCharRow(dt, year, month)
        addTypeColumn(dt)
        addBlankRow(dt)
        setSeqValue(dt)

        settingDgv(dgvWork)
        dgvWork.DataSource = dt
        settingDgvColumnsAndRows(dgvWork)
        setReadonlyCell(dgvWork)
    End Sub

    Private Sub createDictionary()
        'ﾕﾆｯﾄ(2F)の連想配列作成
        unitDictionary2F = New Dictionary(Of String, String)
        unitDictionary2F.Add("※", "00")
        unitDictionary2F.Add("星", "21")
        unitDictionary2F.Add("森", "22")
        unitDictionary2F.Add("空", "23")

        'ﾕﾆｯﾄ(3F)の連想配列作成
        unitDictionary3F = New Dictionary(Of String, String)
        unitDictionary3F.Add("※", "00")
        unitDictionary3F.Add("月", "31")
        unitDictionary3F.Add("花", "32")
        unitDictionary3F.Add("海", "33")

        'Y1～Y31の列のセルの入力文字変換連想配列
        wordDictionary = New Dictionary(Of String, String)
        wordDictionary.Add("0", "")
        wordDictionary.Add("1", "早")
        wordDictionary.Add("2", "日早")
        wordDictionary.Add("3", "日")
        wordDictionary.Add("4", "日遅")
        wordDictionary.Add("5", "遅")
        wordDictionary.Add("6", "遅々")
        wordDictionary.Add("7", "夜")
        wordDictionary.Add("8", "深")
        wordDictionary.Add("10", "半")
        wordDictionary.Add("11", "半Ａ")
        wordDictionary.Add("12", "半Ｂ")
        wordDictionary.Add("13", "半夜")
        wordDictionary.Add("21", "半行")
        wordDictionary.Add("22", "研")
        wordDictionary.Add("31", "有")
        wordDictionary.Add("32", "公")
        wordDictionary.Add("33", "明")
        wordDictionary.Add("34", "希")
        wordDictionary.Add("35", "産")
        wordDictionary.Add("36", "特")
    End Sub

    Private Sub createCellStyles()
        '曜日の行、(予定or変更)の列のスタイル
        disableCellStyle = New DataGridViewCellStyle()
        disableCellStyle.BackColor = Color.FromKnownColor(KnownColor.Control)
        disableCellStyle.SelectionBackColor = Color.FromKnownColor(KnownColor.Control)
        disableCellStyle.SelectionForeColor = Color.Black
        disableCellStyle.Font = New Font("MS UI Gothic", 9, FontStyle.Bold)

        '氏名の列のスタイル
        namColumnCellStyle = New DataGridViewCellStyle()
        namColumnCellStyle.ForeColor = Color.Blue
        namColumnCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        '日曜日の列のスタイル
        sundayColumnCellStyle = New DataGridViewCellStyle()
        sundayColumnCellStyle.BackColor = Color.FromArgb(255, 200, 200)
        sundayColumnCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        '日の文字のセルのスタイル
        sundayCharCellStyle = New DataGridViewCellStyle()
        sundayCharCellStyle.BackColor = Color.FromArgb(255, 200, 200)
        sundayCharCellStyle.SelectionBackColor = Color.FromArgb(255, 200, 200)
        sundayCharCellStyle.SelectionForeColor = Color.Black
        sundayCharCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        sundayCharCellStyle.Font = New Font("MS UI Gothic", 9, FontStyle.Bold)

        'Y1～Y31列の変更の行のスタイル
        workChangeCellStyle = New DataGridViewCellStyle()
        workChangeCellStyle.ForeColor = Color.Red
        workChangeCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

    End Sub

    Private Sub addHenkouRow(dt As DataTable)
        Dim row As DataRow
        For i As Integer = 0 To dt.Rows.Count * 2 - 1 Step 2
            row = dt.NewRow()
            For j As Integer = 1 To 31
                If (Not IsDBNull(dt.Rows(i)("Y" & j)) AndAlso Not IsDBNull(dt.Rows(i)("J" & j))) AndAlso dt.Rows(i)("Y" & j) <> dt.Rows(i)("J" & j) Then
                    row("Y" & j) = dt.Rows(i)("J" & j)
                End If
            Next
            dt.Rows.InsertAt(row, i + 1)
        Next
    End Sub

    Private Sub addTypeColumn(dt As DataTable)
        dt.Columns.Add("type", Type.GetType("System.String")).SetOrdinal(6)
        For i As Integer = 1 To dt.Rows.Count - 1
            If i Mod 2 = 0 Then
                dt.Rows(i).Item("type") = "変更"
            Else
                dt.Rows(i).Item("type") = "予定"
            End If
        Next
    End Sub

    Private Sub addDayCharRow(dt As DataTable, year As Integer, month As Integer)
        Dim daysInMonth As Integer = DateTime.DaysInMonth(year, month)
        Dim firstDay As DateTime = New DateTime(year, month, 1)
        Dim weekNumber As Integer = CInt(firstDay.DayOfWeek)
        Dim row As DataRow = dt.NewRow()

        For i As Integer = 1 To daysInMonth
            row("Y" & i) = dayCharArray((weekNumber + (i - 1)) Mod 7)
        Next

        dt.Rows.InsertAt(row, 0)
    End Sub

    Private Sub addBlankRow(dt As DataTable)
        Dim rowCount As Integer = dt.Rows.Count
        If rowCount = MAX_ROW_COUNT + 1 Then
            Return
        End If

        For i As Integer = rowCount To MAX_ROW_COUNT
            Dim row As DataRow = dt.NewRow()
            dt.Rows.Add(row)
        Next
    End Sub

    Private Sub settingDgv(dgv As DataGridView)
        Util.EnableDoubleBuffering(dgv)

        With dgv
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False '行削除禁止
            .RowHeadersVisible = False '行ヘッダー非表示
            .SelectionMode = DataGridViewSelectionMode.CellSelect
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
            .RowTemplate.Height = 16
            .ColumnHeadersHeight = 19
        End With
    End Sub

    Private Sub settingDgvColumnsAndRows(dgv As DataGridView)
        With dgv
            '並び替えができないようにする
            For Each c As DataGridViewColumn In dgv.Columns
                c.SortMode = DataGridViewColumnSortMode.NotSortable
            Next

            '非表示列
            .Columns("Ym").Visible = False
            .Columns("Seq").Visible = False
            .Columns("Seq2").Visible = False
            For i As Integer = 1 To 31
                .Columns("J" & i).Visible = False
            Next

            '行固定
            .Rows(0).Frozen = True

            '列固定
            .Columns("type").Frozen = True

            'ユニット列
            With .Columns("Unt")
                .Width = 34
                .HeaderText = "ﾕﾆｯﾄ"
                .HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With

            'R列
            With .Columns("Rdr")
                .Width = 19
                .HeaderText = "R"
                .HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With

            '氏名列
            With .Columns("Nam")
                .Width = 90
                .HeaderText = "氏名"
                .HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                .DefaultCellStyle = namColumnCellStyle
            End With

            '予定or変更列
            With .Columns("type")
                .Width = 32
                .HeaderText = ""
                .DefaultCellStyle = disableCellStyle
            End With

            'Y1～Y31の列
            For i As Integer = 1 To 31
                With .Columns("Y" & i)
                    .Width = 46
                    .HeaderText = i.ToString()
                    .HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                End With
            Next

            'Y1～Y31の列の変更の行
            For i As Integer = 2 To MAX_ROW_COUNT Step 2
                For j As Integer = 1 To 31
                    dgv("Y" & j, i).Style = workChangeCellStyle
                Next
            Next

            '日曜日の列
            For i As Integer = 1 To 31
                If Not IsDBNull(dt.Rows(0).Item("Y" & i)) AndAlso dt.Rows(0).Item("Y" & i) = "日" Then
                    dgv("Y" & i, 0).Style = sundayCharCellStyle
                    dgv.Columns("Y" & i).DefaultCellStyle = sundayColumnCellStyle
                End If
            Next

            '日曜日以外の曜日の行
            For Each cell As DataGridViewCell In dgv.Rows(0).Cells
                If IsDBNull(cell.Value) OrElse cell.Value <> "日" Then
                    cell.Style = disableCellStyle
                End If
            Next

        End With
    End Sub

    Private Sub setReadonlyCell(dgv As DataGridView)
        With dgv
            '曜日の行
            .Rows(0).ReadOnly = True

            '予定or変更の列
            .Columns("type").ReadOnly = True

            '変更の行のﾕﾆｯﾄ、R、Nam列のセル
            For i As Integer = 2 To dgv.Rows.Count - 1 Step 2
                dgv("Unt", i).ReadOnly = True
                dgv("Rdr", i).ReadOnly = True
                dgv("Nam", i).ReadOnly = True
            Next
        End With
    End Sub

    Private Sub setSeqValue(dt As DataTable)
        For i As Integer = 1 To MAX_ROW_COUNT Step 2
            dt.Rows(i).Item("Seq") = i + 1
        Next
    End Sub

    Private Sub setAddState(dt As DataTable)
        For Each row As DataRow In dt.Rows
            If Not IsDBNull(row("Nam")) AndAlso row("Nam") <> "" Then
                row.SetAdded()
            End If
        Next
    End Sub

    Private Sub rbtnF_MouseClick(sender As Object, e As MouseEventArgs) Handles rbtn2F.MouseClick, rbtn3F.MouseClick
        If sender Is rbtn2F Then
            displayWorkTable("2018/04", "2")
        ElseIf sender Is rbtn3F Then
            displayWorkTable("2018/04", "3")
        End If
    End Sub

    Private Sub btnRowAdd_Click(sender As Object, e As EventArgs) Handles btnRowAdd.Click
        Dim selectedRowIndex As Integer = If(IsNothing(dgvWork.CurrentRow), -1, dgvWork.CurrentRow.Index)
        If selectedRowIndex = -1 OrElse selectedRowIndex = 0 Then
            Return
        ElseIf Not IsDBNull(dt.Rows(MAX_ROW_COUNT - 1).Item("Nam")) AndAlso dt.Rows(MAX_ROW_COUNT - 1).Item("Nam") <> "" Then
            MsgBox("行挿入できません。")
            Return
        Else
            '変更の行を選択してる場合は予定の行を選択しているindexとする
            If selectedRowIndex Mod 2 = 0 Then
                selectedRowIndex -= 1
            End If

            Dim row1 As DataRow = dt.NewRow()
            Dim row2 As DataRow = dt.NewRow()
            row2("Seq") = selectedRowIndex + 1

            '行追加
            dt.Rows.InsertAt(row1, selectedRowIndex)
            dt.Rows.InsertAt(row2, selectedRowIndex)
            '追加した行(変更の行)のreadonly設定
            dgvWork("Unt", selectedRowIndex + 1).ReadOnly = True
            dgvWork("Rdr", selectedRowIndex + 1).ReadOnly = True
            dgvWork("Nam", selectedRowIndex + 1).ReadOnly = True

            '追加された行以降のSeqの値を更新
            For i As Integer = selectedRowIndex + 2 To MAX_ROW_COUNT - 1 Step 2
                dt.Rows(i).Item("Seq") += 2
            Next

            '下から２行削除
            dt.Rows.RemoveAt(MAX_ROW_COUNT + 2)
            dt.Rows.RemoveAt(MAX_ROW_COUNT + 1)
        End If
    End Sub

    Private Sub btnRowDelete_Click(sender As Object, e As EventArgs) Handles btnRowDelete.Click
        Dim selectedRowIndex As Integer = If(IsNothing(dgvWork.CurrentRow), -1, dgvWork.CurrentRow.Index)
        If selectedRowIndex = -1 OrElse selectedRowIndex = 0 Then
            Return
        Else
            '変更の行を選択してる場合は予定の行を選択しているindexとする
            If selectedRowIndex Mod 2 = 0 Then
                selectedRowIndex -= 1
            End If

            '行削除
            dt.Rows.RemoveAt(selectedRowIndex)
            dt.Rows.RemoveAt(selectedRowIndex)

            '削除された行以降のSeqの値を更新
            For i As Integer = selectedRowIndex To MAX_ROW_COUNT - 3 Step 2
                dt.Rows(i).Item("Seq") -= 2
            Next

            '下に２行追加
            Dim row As DataRow = dt.NewRow()
            row("Seq") = MAX_ROW_COUNT
            dt.Rows.Add(row)
            dt.Rows.Add(dt.NewRow())
            '追加した行(変更の行)のreadonly設定
            dgvWork("Unt", MAX_ROW_COUNT).ReadOnly = True
            dgvWork("Rdr", MAX_ROW_COUNT).ReadOnly = True
            dgvWork("Nam", MAX_ROW_COUNT).ReadOnly = True
        End If
    End Sub

    Private Sub deleteMonthData(ymStr As String, floar As String)
        cn = New OleDbConnection(TopForm.DB_Work)
        sqlCm = cn.CreateCommand
        sqlCm.CommandText = "delete from KinD where Ym='" & ymStr & "' and (Seq2='00' OR ('" & floar & "0' <= Seq2 AND Seq2 <= '" & floar & "9'))"
        cn.Open()
        sqlCm.ExecuteNonQuery()
        cn.Close()
        sqlCm.Dispose()
        cn.Dispose()
    End Sub

    Private Sub btnRegist_Click(sender As Object, e As EventArgs) Handles btnRegist.Click
        Dim floar As String = If(rbtn2F.Checked = True, "2", "3")
        dt.AcceptChanges()
        setAddState(dt)
        deleteMonthData("2018/04", floar)
        adapter.Update(dt)
        MsgBox("登録しました。")
    End Sub

    Private Sub dgvWork_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dgvWork.CellBeginEdit
        editBeforeCellValue = If(IsDBNull(dgvWork(e.ColumnIndex, e.RowIndex).Value), "", dgvWork(e.ColumnIndex, e.RowIndex).Value)
    End Sub

    Private Sub dgvWork_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgvWork.CellEndEdit
        If dgvWork.Columns(e.ColumnIndex).Name = "Unt" Then
            'ﾕﾆｯﾄ列の編集終了時、Seq2列のセルに対応した値を設定
            Dim inputStr As String = If(IsDBNull(dgvWork(e.ColumnIndex, e.RowIndex).Value), "", dgvWork(e.ColumnIndex, e.RowIndex).Value)
            Try
                If rbtn2F.Checked = True Then
                    dgvWork("Seq2", e.RowIndex).Value = unitDictionary2F(inputStr)
                Else
                    dgvWork("Seq2", e.RowIndex).Value = unitDictionary3F(inputStr)
                End If
            Catch ex As KeyNotFoundException
                dgvWork(e.ColumnIndex, e.RowIndex).Value = editBeforeCellValue
                MsgBox("正しいﾕﾆｯﾄ名を入力してください。")
            End Try
        ElseIf 7 <= e.ColumnIndex AndAlso e.ColumnIndex <= 37 Then
            'Y1～Y31列の編集終了時の処理
            If e.RowIndex Mod 2 = 1 Then
                '予定の行の場合、値の変換処理をする
                Try
                    dgvWork(e.ColumnIndex, e.RowIndex).Value = wordDictionary(dgvWork(e.ColumnIndex, e.RowIndex).Value)
                Catch ex As KeyNotFoundException
                    '何もしない
                End Try
            Else
                '変更の行の場合、値の変換処理をして、入力しているY列に対応する非表示のJ列の値を置き換える
                Try
                    dgvWork(e.ColumnIndex, e.RowIndex).Value = wordDictionary(dgvWork(e.ColumnIndex, e.RowIndex).Value)
                    dgvWork(e.ColumnIndex + 31, e.RowIndex - 1).Value = dgvWork(e.ColumnIndex, e.RowIndex).Value
                Catch ex As KeyNotFoundException
                    dgvWork(e.ColumnIndex + 31, e.RowIndex - 1).Value = dgvWork(e.ColumnIndex, e.RowIndex).Value
                End Try
            End If

        End If
    End Sub

End Class