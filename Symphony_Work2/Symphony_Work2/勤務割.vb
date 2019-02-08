Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop

Public Class 勤務割

    '勤務割データテーブル
    Private workDt As DataTable

    'ユニット名dic
    Private unitDictionary As Dictionary(Of String, String)

    '勤務略名dic
    Private wordDictionary As Dictionary(Of String, String)

    '勤務時間dic
    Private workTimeDictionary As Dictionary(Of String, Double)

    '略語dic
    Private abbreviationDictionary As Dictionary(Of String, String)

    '小計行インデックスdic
    Private subtotalStrIndexDictionary As Dictionary(Of String, Integer)

    '曜日配列
    Private dayCharArray() As String = {"日", "月", "火", "水", "木", "金", "土"}

    'アルファベット配列
    Private NAME_COLUMN_VALUES As Char() = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray

    'アルファベット配列長さ
    Private NAME_COLUMN_VALUES_LENGTH As Integer = NAME_COLUMN_VALUES.Length

    '編集不可セルスタイル
    Private disableCellStyle As DataGridViewCellStyle

    '氏名列セルスタイル
    Private namColumnCellStyle As DataGridViewCellStyle

    '日曜日列セルスタイル
    Private sundayColumnCellStyle As DataGridViewCellStyle

    '"日"の文字セルスタイル
    Private sundayCharCellStyle As DataGridViewCellStyle

    '変更セルスタイル
    Private workChangeCellStyle As DataGridViewCellStyle

    '小計予定セルスタイル
    Private subtotalPlanCellStyle As DataGridViewCellStyle

    '小計変更セルスタイル
    Private subtotalChangeCellStyle As DataGridViewCellStyle

    '入力可能行数（勤務入力部分）
    Private Const INPUT_ROW_COUNT As Integer = 50

    '入力不可行数（小計表示部分）
    Private Const READONLY_ROW_COUNT As Integer = 32

    '同姓略名フォーム
    Private abbreviationNamForm As 同姓略名

    ''' <summary>
    ''' keyDownイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 勤務割_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.Alt AndAlso e.KeyCode = Keys.F12 Then
            '(Alt + F12)キー押下
            btnRowAdd.Visible = Not btnRowAdd.Visible '行挿入ボタン表示、非表示
            btnRowDelete.Visible = Not btnRowDelete.Visible '行削除ボタン表示、非表示
            btnRegist.Visible = Not btnRegist.Visible '登録ボタン表示、非表示
            btnDelete.Visible = Not btnDelete.Visible '削除ボタン表示、非表示
            btnPrint.Visible = Not btnPrint.Visible '印刷ボタン表示、非表示
            wordPanel.Visible = Not wordPanel.Visible '勤務名ラベル表示、非表示
        End If

        If e.Alt AndAlso e.KeyCode = Keys.F11 Then
            '(Alt + F11)キー押下
            If IsNothing(abbreviationNamForm) OrElse abbreviationNamForm.IsDisposed Then
                '同姓略名フォーム表示
                abbreviationNamForm = New 同姓略名(ymBox.getADStr4Ym())
                abbreviationNamForm.Show()
            End If
        End If
    End Sub

    ''' <summary>
    ''' loadイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 勤務割_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.KeyPreview = True

        'dic作成
        createDictionary()

        'セルスタイル作成
        createCellStyles()

        'dgv初期設定
        initDgvWork()

        'ラジオボタンを2階にセット
        rbtn2F.Checked = True
    End Sub

    ''' <summary>
    ''' dic作成
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub createDictionary()
        'ﾕﾆｯﾄの連想配列作成
        unitDictionary = New Dictionary(Of String, String)
        unitDictionary.Add("※", "00")
        unitDictionary.Add("虹", "21")
        unitDictionary.Add("光", "22")
        unitDictionary.Add("丘", "23")
        unitDictionary.Add("風", "31")
        unitDictionary.Add("雪", "32")

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

        '略語
        abbreviationDictionary = New Dictionary(Of String, String)
        abbreviationDictionary.Add("早", "早出")
        abbreviationDictionary.Add("日早", "日早")
        abbreviationDictionary.Add("日", "日勤")
        abbreviationDictionary.Add("日遅", "日遅")
        abbreviationDictionary.Add("遅", "遅出")
        abbreviationDictionary.Add("遅々", "遅々")
        abbreviationDictionary.Add("夜", "夜勤")
        abbreviationDictionary.Add("深", "深夜")
        abbreviationDictionary.Add("半", "半")
        abbreviationDictionary.Add("半Ａ", "半Ａ")
        abbreviationDictionary.Add("半Ｂ", "半Ｂ")
        abbreviationDictionary.Add("半夜", "半夜")
        abbreviationDictionary.Add("半行", "半行")
        abbreviationDictionary.Add("研", "研修")
        abbreviationDictionary.Add("公", "公休")
        abbreviationDictionary.Add("明", "明等")

        '勤務時間
        workTimeDictionary = New Dictionary(Of String, Double)
        workTimeDictionary.Add("早", 7.5)
        workTimeDictionary.Add("日早", 7.5)
        workTimeDictionary.Add("日", 7.5)
        workTimeDictionary.Add("日遅", 7.5)
        workTimeDictionary.Add("遅", 7.5)
        workTimeDictionary.Add("遅々", 7.5)
        workTimeDictionary.Add("夜", 15.0)
        workTimeDictionary.Add("深", 7.5)
        workTimeDictionary.Add("半", 3.5)
        workTimeDictionary.Add("半Ａ", 3.5)
        workTimeDictionary.Add("半Ｂ", 3.5)
        workTimeDictionary.Add("半夜", 3.5)
        workTimeDictionary.Add("半行", 3.5)
        workTimeDictionary.Add("研", 7.5)
        workTimeDictionary.Add("有", 0.0)
        workTimeDictionary.Add("公", 7.5)
        workTimeDictionary.Add("明", 0.0)
        workTimeDictionary.Add("希", 0.0)
        workTimeDictionary.Add("産", 0.0)
        workTimeDictionary.Add("特", 0.0)
        workTimeDictionary.Add("A", 5.0)
        workTimeDictionary.Add("B", 5.5)
        workTimeDictionary.Add("C", 7.0)
        workTimeDictionary.Add("D", 3.5)
        workTimeDictionary.Add("E", 5.0)
        workTimeDictionary.Add("F", 6.0)
        workTimeDictionary.Add("G", 7.0)
        workTimeDictionary.Add("H", 4.0)
        workTimeDictionary.Add("I", 3.0)
        workTimeDictionary.Add("J", 5.5)
        workTimeDictionary.Add("K", 7.0)
        workTimeDictionary.Add("L", 2.5)
        workTimeDictionary.Add("M", 3.5)
        workTimeDictionary.Add("N", 2.0)
        workTimeDictionary.Add("P", 6.5)
        workTimeDictionary.Add("R", 2.5)
        workTimeDictionary.Add("S", 7.5)
        workTimeDictionary.Add("T", 4.5)

        '小計の行インデックス
        subtotalStrIndexDictionary = New Dictionary(Of String, Integer)
        subtotalStrIndexDictionary.Add("早出", 51)
        subtotalStrIndexDictionary.Add("日早", 53)
        subtotalStrIndexDictionary.Add("日勤", 55)
        subtotalStrIndexDictionary.Add("日遅", 57)
        subtotalStrIndexDictionary.Add("遅出", 59)
        subtotalStrIndexDictionary.Add("遅々", 61)
        subtotalStrIndexDictionary.Add("夜勤", 63)
        subtotalStrIndexDictionary.Add("深夜", 65)
        subtotalStrIndexDictionary.Add("半", 67)
        subtotalStrIndexDictionary.Add("半Ａ", 69)
        subtotalStrIndexDictionary.Add("半Ｂ", 71)
        subtotalStrIndexDictionary.Add("半夜", 73)
        subtotalStrIndexDictionary.Add("半行", 75)
        subtotalStrIndexDictionary.Add("研修", 77)
        subtotalStrIndexDictionary.Add("公休", 79)

    End Sub

    ''' <summary>
    ''' セルスタイル作成
    ''' </summary>
    ''' <remarks></remarks>
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
        namColumnCellStyle.SelectionForeColor = Color.Blue
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
        workChangeCellStyle.SelectionForeColor = Color.Red
        workChangeCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        workChangeCellStyle.Font = New Font("MS UI Gothic", 8.5)

        '小計予定セルスタイル
        subtotalPlanCellStyle = New DataGridViewCellStyle()
        subtotalPlanCellStyle.BackColor = Color.FromKnownColor(KnownColor.Control)
        subtotalPlanCellStyle.SelectionBackColor = Color.FromKnownColor(KnownColor.Control)
        subtotalPlanCellStyle.ForeColor = Color.Blue
        subtotalPlanCellStyle.SelectionForeColor = Color.Blue
        subtotalPlanCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        '小計変更セルスタイル
        subtotalChangeCellStyle = New DataGridViewCellStyle()
        subtotalChangeCellStyle.BackColor = Color.FromKnownColor(KnownColor.Control)
        subtotalChangeCellStyle.SelectionBackColor = Color.FromKnownColor(KnownColor.Control)
        subtotalChangeCellStyle.ForeColor = Color.Red
        subtotalChangeCellStyle.SelectionForeColor = Color.Red
        subtotalChangeCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

    End Sub

    ''' <summary>
    ''' dgv初期設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub initDgvWork()
        'dictionary設定
        dgvWork.setUnitDictionary(unitDictionary)
        dgvWork.setWordDictionary(wordDictionary)

        'dgv設定
        With dgvWork
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False '行削除禁止
            .RowHeadersVisible = False '行ヘッダー非表示
            .SelectionMode = DataGridViewSelectionMode.CellSelect 'セル選択
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
            .DefaultCellStyle.SelectionForeColor = Color.Black
            .RowTemplate.Height = 15
            .ColumnHeadersHeight = 19
            .ShowCellToolTips = False
            .EnableHeadersVisualStyles = False
            .DefaultCellStyle.Font = New Font("MS UI Gothic", 8.5)
        End With

    End Sub

    ''' <summary>
    ''' 空行作成
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub setEmptyCell()
        dgvWork.Columns.Clear()

        workDt = New DataTable()

        '列定義
        workDt.Columns.Add("Ym", Type.GetType("System.String"))
        workDt.Columns.Add("Seq", Type.GetType("System.String"))
        workDt.Columns.Add("Seq2", Type.GetType("System.String"))
        workDt.Columns.Add("Unt", Type.GetType("System.String"))
        workDt.Columns.Add("Rdr", Type.GetType("System.String"))
        workDt.Columns.Add("Nam", Type.GetType("System.String"))
        workDt.Columns.Add("Type", Type.GetType("System.String"))
        For i As Integer = 1 To 31
            workDt.Columns.Add("Y" & i, Type.GetType("System.String"))
        Next
        workDt.Columns.Add("月合計", Type.GetType("System.String"))
        workDt.Columns.Add("早出", Type.GetType("System.String"))
        workDt.Columns.Add("日早", Type.GetType("System.String"))
        workDt.Columns.Add("日勤", Type.GetType("System.String"))
        workDt.Columns.Add("日遅", Type.GetType("System.String"))
        workDt.Columns.Add("遅出", Type.GetType("System.String"))
        workDt.Columns.Add("遅々", Type.GetType("System.String"))
        workDt.Columns.Add("夜勤", Type.GetType("System.String"))
        workDt.Columns.Add("深夜", Type.GetType("System.String"))
        workDt.Columns.Add("半", Type.GetType("System.String"))
        workDt.Columns.Add("半Ａ", Type.GetType("System.String"))
        workDt.Columns.Add("半Ｂ", Type.GetType("System.String"))
        workDt.Columns.Add("半夜", Type.GetType("System.String"))
        workDt.Columns.Add("半行", Type.GetType("System.String"))
        workDt.Columns.Add("研修", Type.GetType("System.String"))
        workDt.Columns.Add("公休", Type.GetType("System.String"))
        workDt.Columns.Add("明等", Type.GetType("System.String"))

        '空行追加
        For i = 0 To 1 + INPUT_ROW_COUNT + READONLY_ROW_COUNT
            workDt.Rows.Add(workDt.NewRow())
        Next

        '小計項目名
        Dim itemArray As String() = {"早出", "日早", "日勤", "日遅", "遅出", "遅々", "夜勤", "深夜", "半", "半Ａ", "半Ｂ", "半夜", "半行", "研修", "公休"}
        Dim index As Integer = 0
        For i As Integer = 51 To 79 Step 2
            workDt.Rows(i).Item("Nam") = itemArray(index)
            index += 1
        Next

        '表示
        dgvWork.DataSource = workDt
    End Sub

    ''' <summary>
    ''' dgv列行スタイル設定等
    ''' </summary>
    ''' <param name="year"></param>
    ''' <param name="month"></param>
    ''' <remarks></remarks>
    Private Sub settingDgvWorkColumnsAndRows(year As Integer, month As Integer)
        '空セル表示
        setEmptyCell()

        '曜日設定
        setDayCharRow(year, month)

        '列設定
        With dgvWork
            '並び替えができないようにする
            For Each c As DataGridViewColumn In .Columns
                c.SortMode = DataGridViewColumnSortMode.NotSortable
            Next

            '非表示列
            .Columns("Ym").Visible = False
            .Columns("Seq").Visible = False
            .Columns("Seq2").Visible = False

            '行固定
            .Rows(0).Frozen = True

            '列固定
            .Columns("type").Frozen = True

            'ユニット列
            With .Columns("Unt")
                .Width = 32
                .HeaderText = "ﾕﾆｯﾄ"
                .HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With

            'R列
            With .Columns("Rdr")
                .Width = 20
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
                    .Width = 50
                    .HeaderText = i.ToString()
                    .HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                End With
            Next

            'Y1～Y31,小計列の変更の行
            For i As Integer = 2 To INPUT_ROW_COUNT Step 2
                For j As Integer = 1 To 31
                    dgvWork("Y" & j, i).Style = workChangeCellStyle
                Next
                For k As Integer = 38 To 54
                    dgvWork(k, i).Style = subtotalChangeCellStyle
                Next
            Next

            '日曜日の列
            For i As Integer = 1 To 31
                If Not IsDBNull(workDt.Rows(0).Item("Y" & i)) AndAlso workDt.Rows(0).Item("Y" & i) = "日" Then
                    dgvWork("Y" & i, 0).Style = sundayCharCellStyle
                    dgvWork.Columns("Y" & i).DefaultCellStyle = sundayColumnCellStyle
                End If
            Next

            '日曜日以外の曜日の行
            For Each cell As DataGridViewCell In dgvWork.Rows(0).Cells
                If IsDBNull(cell.Value) OrElse cell.Value <> "日" Then
                    cell.Style = disableCellStyle
                End If
            Next

            '小計記載行
            For i As Integer = 1 + INPUT_ROW_COUNT To 1 + INPUT_ROW_COUNT + READONLY_ROW_COUNT
                If i Mod 2 = 1 Then
                    .Rows(i).DefaultCellStyle = subtotalPlanCellStyle
                Else
                    .Rows(i).DefaultCellStyle = subtotalChangeCellStyle
                End If
            Next

            '小計記載列
            .Columns(38).DefaultCellStyle = subtotalPlanCellStyle
            .Columns(38).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(38).Width = 60
            For i As Integer = 39 To 54
                .Columns(i).DefaultCellStyle = subtotalPlanCellStyle
                .Columns(i).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns(i).Width = 36
            Next

            'ReadOnlyセルの設定
            setReadonlyCell()

        End With
    End Sub

    ''' <summary>
    ''' 曜日行作成
    ''' </summary>
    ''' <param name="year"></param>
    ''' <param name="month"></param>
    ''' <remarks></remarks>
    Private Sub setDayCharRow(year As Integer, month As Integer)
        Dim daysInMonth As Integer = DateTime.DaysInMonth(year, month) '月の日数
        Dim firstDay As DateTime = New DateTime(year, month, 1)
        Dim weekNumber As Integer = CInt(firstDay.DayOfWeek) '月の初日の曜日のindex
        Dim row As DataRow = workDt.Rows(0)

        For i As Integer = 1 To daysInMonth
            row("Y" & i) = dayCharArray((weekNumber + (i - 1)) Mod 7)
        Next
    End Sub

    ''' <summary>
    ''' readonlyセルの設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub setReadonlyCell()
        With dgvWork
            '曜日の行
            .Rows(0).ReadOnly = True

            '予定or変更の列
            .Columns("type").ReadOnly = True

            '変更の行のﾕﾆｯﾄ、R、Nam列のセル
            For i As Integer = 2 To dgvWork.Rows.Count - 1 Step 2
                dgvWork("Unt", i).ReadOnly = True
                dgvWork("Rdr", i).ReadOnly = True
                dgvWork("Nam", i).ReadOnly = True
            Next

            '小計記載行
            For i As Integer = 1 + INPUT_ROW_COUNT To 1 + INPUT_ROW_COUNT + READONLY_ROW_COUNT
                .Rows(i).ReadOnly = True
            Next

            '小計記載列
            For i As Integer = 38 To 54
                dgvWork.Columns(i).ReadOnly = True
            Next
        End With
    End Sub

    ''' <summary>
    ''' 行番号(seq)セット
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub setSeqValue()
        For i As Integer = 1 To INPUT_ROW_COUNT Step 2
            workDt.Rows(i).Item("Seq") = i + 1
        Next
    End Sub

    ''' <summary>
    ''' 勤務割表示
    ''' </summary>
    ''' <param name="ymStr">年月(yyyy/MM)</param>
    ''' <param name="floor">階</param>
    ''' <param name="deleteAfterFlg"></param>
    ''' <remarks></remarks>
    Private Sub displayWork(ymStr As String, floor As String, Optional deleteAfterFlg As Boolean = False)
        Dim year As Integer = CInt(ymStr.Split("/")(0)) '年
        Dim month As Integer = CInt(ymStr.Split("/")(1)) '月

        'dgv列行設定
        settingDgvWorkColumnsAndRows(year, month)
        '行番号設定
        setSeqValue()

        If deleteAfterFlg Then
            Return
        End If

        Dim cnn As New ADODB.Connection
        cnn.Open(TopForm.DB_Work2)
        Dim rs As New ADODB.Recordset
        Dim sql = "SELECT * FROM KinD WHERE YM='" & ymStr & "' AND ((Seq2='00' AND Unt='※') OR ('" & floor & "0' <= Seq2 AND Seq2 <= '" & floor & "9')) order by Seq"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)

        If rs.RecordCount <= 0 Then '当月データが無い場合
            Dim warekiStr As String = Util.convADStrToWarekiStr(ymStr & "/01")
            Dim eraStr As String = warekiStr.Substring(0, 3)
            Dim monthStr As String = warekiStr.Substring(4, 2)
            Dim result As DialogResult = MessageBox.Show(eraStr & "年" & monthStr & "月分は登録されていません" & Environment.NewLine & "登録しますか？", "Work", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
            If result = Windows.Forms.DialogResult.Yes Then
                rs.Close()

                '前月データの名前を表示
                Dim prevMonth As String
                Dim prevYear As String
                If month - 1 = 0 Then
                    prevYear = (year - 1).ToString()
                    prevMonth = "12"
                Else
                    prevYear = year.ToString()
                    prevMonth = If(month - 1 >= 10, (month - 1).ToString(), "0" & (month - 1).ToString())
                End If
                Dim prevYmStr As String = prevYear & "/" & prevMonth

                sql = "SELECT * FROM KinD WHERE YM='" & prevYmStr & "' AND ((Seq2='00' AND Unt='※') OR ('" & floor & "0' <= Seq2 AND Seq2 <= '" & floor & "9')) order by Seq"
                rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)

                Dim rowIndex As Integer = 1
                While Not rs.EOF
                    '予定行の値設定
                    dgvWork("Seq", rowIndex).Value = Util.checkDBNullValue(rs.Fields("Seq").Value)
                    dgvWork("Seq2", rowIndex).Value = Util.checkDBNullValue(rs.Fields("Seq2").Value)
                    dgvWork("Unt", rowIndex).Value = Util.checkDBNullValue(rs.Fields("Unt").Value)
                    dgvWork("Rdr", rowIndex).Value = Util.checkDBNullValue(rs.Fields("Rdr").Value)
                    dgvWork("Nam", rowIndex).Value = Util.checkDBNullValue(rs.Fields("Nam").Value)
                    dgvWork("Type", rowIndex).Value = "予定"

                    '変更行の値設定
                    dgvWork("Type", (rowIndex + 1)).Value = "変更"

                    rowIndex += 2
                    rs.MoveNext()
                End While

                rs.Close()
                cnn.Close()
                Return
            Else
                rs.Close()
                cnn.Close()
                Return
            End If
        Else
            '表示処理
            '現在日付が見えるようにスクロール
            Dim todayDate As Integer = Today.Day
            If todayDate >= 24 Then
                dgvWork.FirstDisplayedScrollingColumnIndex = 21
            ElseIf 10 <= todayDate AndAlso todayDate <= 23 Then
                dgvWork.FirstDisplayedScrollingColumnIndex = todayDate - 2
            Else
                dgvWork.FirstDisplayedScrollingColumnIndex = 7
            End If

            Dim rowIndex As Integer = 1
            While Not rs.EOF
                '予定行の値設定
                dgvWork("Ym", rowIndex).Value = Util.checkDBNullValue(rs.Fields("Ym").Value)
                dgvWork("Seq", rowIndex).Value = Util.checkDBNullValue(rs.Fields("Seq").Value)
                dgvWork("Seq2", rowIndex).Value = Util.checkDBNullValue(rs.Fields("Seq2").Value)
                dgvWork("Unt", rowIndex).Value = Util.checkDBNullValue(rs.Fields("Unt").Value)
                dgvWork("Rdr", rowIndex).Value = Util.checkDBNullValue(rs.Fields("Rdr").Value)
                dgvWork("Nam", rowIndex).Value = Util.checkDBNullValue(rs.Fields("Nam").Value)
                dgvWork("Type", rowIndex).Value = "予定"
                For i As Integer = 1 To 31
                    dgvWork("Y" & i, rowIndex).Value = Util.checkDBNullValue(rs.Fields("Y" & i).Value)
                Next

                '変更行の値設定
                dgvWork("Type", (rowIndex + 1)).Value = "変更"
                For i As Integer = 1 To 31
                    '予定と変更の内容が異なる場合のみ変更を表示
                    dgvWork("Y" & i, (rowIndex + 1)).Value = If(Util.checkDBNullValue(rs.Fields("J" & i).Value) = Util.checkDBNullValue(rs.Fields("Y" & i).Value), "", Util.checkDBNullValue(rs.Fields("J" & i).Value))
                Next

                rowIndex += 2
                rs.MoveNext()
            End While
            rs.Close()
            cnn.Close()
        End If

    End Sub

    ''' <summary>
    ''' 年月ボックス変更イベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ymBox_YmLabelTextChange(sender As Object, e As System.EventArgs) Handles ymBox.YmLabelTextChange
        Dim ym As String = ymBox.getADStr4Ym() '選択年月
        Dim floor As String = If(rbtn2F.Checked, "2", "3") '選択されている階
        displayWork(ym, floor) '表示
    End Sub

    ''' <summary>
    ''' 階ラジオボタン変更イベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub floorRadioButton_CheckedChanged(sender As Object, e As System.EventArgs) Handles rbtn2F.CheckedChanged, rbtn3F.CheckedChanged
        Dim rbtn As RadioButton = CType(sender, RadioButton)
        If rbtn.Checked = True Then
            rbtn.BackColor = Color.FromArgb(255, 255, 0)
            Dim floor As String = rbtn.Name.Substring(4, 1)
            displayWork(ymBox.getADStr4Ym(), floor) '選択年月、階のデータ表示
        Else
            rbtn.BackColor = Color.FromKnownColor(KnownColor.Control)
        End If
    End Sub

    ''' <summary>
    ''' 対象年月階のデータを削除
    ''' </summary>
    ''' <param name="ymStr">年月(yyyy/MM)</param>
    ''' <param name="floor">階</param>
    ''' <param name="cnn">データベースコネクション</param>
    ''' <remarks></remarks>
    Private Sub monthDataDelete(ymStr As String, floor As String, cnn As ADODB.Connection)
        Dim cmd As New ADODB.Command()
        cmd.ActiveConnection = cnn
        cmd.CommandText = "delete from KinD where YM='" & ymStr & "' AND (Seq2='00' OR ('" & floor & "0' <= Seq2 AND Seq2 <= '" & floor & "9'))"
        cmd.Execute()
    End Sub

    ''' <summary>
    ''' 対象行に勤務の入力があるかチェック
    ''' </summary>
    ''' <param name="row">dgv行</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function existsWorkStr(row As DataGridViewRow) As Boolean
        For i As Integer = 1 To 31
            If Util.checkDBNullValue(row.Cells("Y" & i).Value) <> "" Then
                Return True
            End If
        Next
        Return False
    End Function

    ''' <summary>
    ''' 曜日の無い列に対象の行が入力があるかチェック
    ''' </summary>
    ''' <param name="row">dgv行</param>
    ''' <param name="ymStr">年月(yyyy/MM)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function existsNoneDayCell(row As DataGridViewRow, ymStr As String) As Boolean
        Dim year As Integer = CInt(ymStr.Split("/")(0)) '年
        Dim month As Integer = CInt(ymStr.Split("/")(1)) '月
        Dim daysInMonth As Integer = DateTime.DaysInMonth(year, month) '月の日数
        For i As Integer = daysInMonth + 1 To 31
            If Util.checkDBNullValue(row.Cells("Y" & i).Value) <> "" Then
                Return True
            End If
        Next
        Return False
    End Function

    ''' <summary>
    ''' 小計のクリア
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subtotalClear()
        '小計列のクリア
        For i As Integer = 38 To 54
            For j As Integer = 1 To 50
                dgvWork(i, j).Value = ""
            Next
        Next

        '小計行のクリア
        For i As Integer = 51 To 80
            For j As Integer = 1 To 31
                dgvWork("Y" & j, i).Value = ""
            Next
        Next
    End Sub

    ''' <summary>
    ''' 行追加ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRowAdd_Click(sender As System.Object, e As System.EventArgs) Handles btnRowAdd.Click
        Dim selectedRowIndex As Integer = If(IsNothing(dgvWork.CurrentRow), -1, dgvWork.CurrentRow.Index) '選択行index
        If selectedRowIndex = -1 OrElse selectedRowIndex = 0 OrElse selectedRowIndex >= 51 Then
            Return
        ElseIf Not IsDBNull(workDt.Rows(INPUT_ROW_COUNT - 1).Item("Nam")) AndAlso workDt.Rows(INPUT_ROW_COUNT - 1).Item("Nam") <> "" Then
            '一番下の予定行に既に名前が入力されている場合は行挿入禁止
            MsgBox("行挿入できません。")
            Return
        Else
            '変更の行を選択してる場合は予定の行を選択しているindexとする
            If selectedRowIndex Mod 2 = 0 Then
                selectedRowIndex -= 1
            End If

            Dim rowJ As DataRow = workDt.NewRow()
            Dim rowY As DataRow = workDt.NewRow()
            rowY("Seq") = selectedRowIndex + 1

            '行追加
            workDt.Rows.InsertAt(rowJ, selectedRowIndex) '変更行
            workDt.Rows.InsertAt(rowY, selectedRowIndex) '予定行

            '追加した変更行の設定
            dgvWork("Unt", selectedRowIndex + 1).ReadOnly = True
            dgvWork("Rdr", selectedRowIndex + 1).ReadOnly = True
            dgvWork("Nam", selectedRowIndex + 1).ReadOnly = True
            For i As Integer = 1 To 31 'Y1～Y31列のセルスタイル設定
                dgvWork("Y" & i, selectedRowIndex + 1).Style = workChangeCellStyle
            Next
            For i As Integer = 38 To 54 '小計部分のセルスタイル設定
                dgvWork(i, selectedRowIndex + 1).Style = subtotalChangeCellStyle
            Next

            '追加された行以降のSeqの値を更新
            For i As Integer = selectedRowIndex + 2 To INPUT_ROW_COUNT - 1 Step 2
                workDt.Rows(i).Item("Seq") += 2
            Next

            '下から２行削除
            workDt.Rows.RemoveAt(INPUT_ROW_COUNT + 2)
            workDt.Rows.RemoveAt(INPUT_ROW_COUNT + 1)
        End If
    End Sub

    ''' <summary>
    ''' 行削除ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRowDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnRowDelete.Click
        Dim selectedRowIndex As Integer = If(IsNothing(dgvWork.CurrentRow), -1, dgvWork.CurrentRow.Index) '選択行index
        If selectedRowIndex = -1 OrElse selectedRowIndex = 0 OrElse selectedRowIndex >= 51 Then
            Return
        Else
            '変更の行を選択してる場合は予定の行を選択しているindexとする
            If selectedRowIndex Mod 2 = 0 Then
                selectedRowIndex -= 1
            End If

            '行削除
            workDt.Rows.RemoveAt(selectedRowIndex)
            workDt.Rows.RemoveAt(selectedRowIndex)

            '削除された行以降のSeqの値を更新
            For i As Integer = selectedRowIndex To INPUT_ROW_COUNT - 3 Step 2
                workDt.Rows(i).Item("Seq") -= 2
            Next

            '下に２行追加
            Dim row As DataRow = workDt.NewRow()
            row("Seq") = INPUT_ROW_COUNT
            workDt.Rows.InsertAt(workDt.NewRow(), INPUT_ROW_COUNT - 1)
            workDt.Rows.InsertAt(row, INPUT_ROW_COUNT - 1)

            '追加した変更行の設定
            dgvWork("Unt", INPUT_ROW_COUNT).ReadOnly = True
            dgvWork("Rdr", INPUT_ROW_COUNT).ReadOnly = True
            dgvWork("Nam", INPUT_ROW_COUNT).ReadOnly = True
            For i As Integer = 1 To 31 'Y1～Y31列のセルスタイル設定
                dgvWork("Y" & i, INPUT_ROW_COUNT).Style = workChangeCellStyle
            Next
            For i As Integer = 38 To 54 '小計部分のセルスタイル設定
                dgvWork(i, INPUT_ROW_COUNT).Style = subtotalChangeCellStyle
            Next
        End If
    End Sub

    ''' <summary>
    ''' 登録ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRegist_Click(sender As System.Object, e As System.EventArgs) Handles btnRegist.Click
        Dim cnn As New ADODB.Connection
        cnn.Open(TopForm.DB_Work2)
        Dim rs As New ADODB.Recordset
        rs.Open("KinD", cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)

        Dim ymStr As String = ymBox.getADStr4Ym() '選択年月
        Dim floor As String = If(rbtn2F.Checked, "2", "3") '選択階
        Dim seq As Integer = 2
        Dim existsUnt As Boolean
        Dim existsNam As Boolean
        Dim existsWork As Boolean
        Dim existsNoneDay As Boolean

        '登録チェック
        For i As Integer = 1 To 49 Step 2
            existsUnt = If(Util.checkDBNullValue(dgvWork("Unt", i).Value) <> "", True, False) 'ユニット名の入力チェック
            existsNam = If(Util.checkDBNullValue(dgvWork("Nam", i).Value) <> "", True, False) '氏名の入力チェック
            existsWork = existsWorkStr(dgvWork.Rows(i)) '勤務の入力チェック
            existsNoneDay = existsNoneDayCell(dgvWork.Rows(i), ymStr) '曜日の無い列への入力チェック

            If (existsUnt AndAlso Not existsNam AndAlso existsWork) OrElse (Not existsUnt AndAlso Not existsNam AndAlso existsWork) Then
                MsgBox("氏名の無い行に入力しています。", MsgBoxStyle.Exclamation, "Work")
                rs.Close()
                cnn.Close()
                Return
            ElseIf (Not existsUnt AndAlso existsNam AndAlso existsWork) OrElse (Not existsUnt AndAlso existsNam AndAlso Not existsWork) Then
                MsgBox("ﾕﾆｯﾄが空白です。", MsgBoxStyle.Exclamation, "Work")
                rs.Close()
                cnn.Close()
                Return
            Else
                If existsNoneDay Then
                    MsgBox("曜日の無い列に入力しています。", MsgBoxStyle.Exclamation, "Work")
                    rs.Close()
                    cnn.Close()
                    Return
                Else
                    Continue For
                End If
            End If
        Next

        '既存データ削除
        monthDataDelete(ymStr, floor, cnn)

        '登録
        For i As Integer = 1 To 49 Step 2
            existsUnt = If(Util.checkDBNullValue(dgvWork("Unt", i).Value) <> "", True, False)
            existsNam = If(Util.checkDBNullValue(dgvWork("Nam", i).Value) <> "", True, False)
            existsWork = existsWorkStr(dgvWork.Rows(i))

            If (existsUnt AndAlso Not existsNam AndAlso Not existsWork) OrElse (Not existsUnt AndAlso Not existsNam AndAlso Not existsWork) Then
                Continue For
            Else
                If Not unitDictionary.ContainsKey(Util.checkDBNullValue(dgvWork("Unt", i).Value)) Then
                    Continue For
                End If
                With rs
                    .AddNew()
                    .Fields("Ym").Value = ymStr
                    .Fields("Seq").Value = seq
                    .Fields("Seq2").Value = Util.checkDBNullValue(dgvWork("Seq2", i).Value)
                    .Fields("Unt").Value = Util.checkDBNullValue(dgvWork("Unt", i).Value)
                    .Fields("Rdr").Value = Util.checkDBNullValue(dgvWork("Rdr", i).Value)
                    .Fields("Nam").Value = Util.checkDBNullValue(dgvWork("Nam", i).Value)
                    For j As Integer = 1 To 31
                        .Fields("Y" & j).Value = Util.checkDBNullValue(dgvWork("Y" & j, i).Value)
                        .Fields("J" & j).Value = If(Util.checkDBNullValue(dgvWork("Y" & j, i + 1).Value) = "", Util.checkDBNullValue(dgvWork("Y" & j, i).Value), Util.checkDBNullValue(dgvWork("Y" & j, i + 1).Value))
                    Next
                End With
                rs.Update()
                seq += 2
            End If
        Next
        rs.Close()
        cnn.Close()
        MsgBox("登録しました。", , "Work")
    End Sub

    ''' <summary>
    ''' 削除ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click
        Dim ymStr As String = ymBox.getADStr4Ym() '選択年月
        Dim floor As String = If(rbtn2F.Checked, "2", "3") '選択階
        Dim cnn As New ADODB.Connection
        cnn.Open(TopForm.DB_Work2)
        Dim rs As New ADODB.Recordset
        Dim sql = "SELECT * FROM KinD WHERE YM='" & ymStr & "' AND (Seq2='00' OR ('" & floor & "0' <= Seq2 AND Seq2 <= '" & floor & "9')) order by Seq"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)

        If rs.RecordCount <= 0 Then '対象年月のデータが存在しない場合
            MsgBox("登録されていません", , "Work")
            rs.Close()
            cnn.Close()
        Else
            Dim result As DialogResult = MessageBox.Show("削除してよろしいですか？", "Work", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
            If result = Windows.Forms.DialogResult.Yes Then
                monthDataDelete(ymStr, floor, cnn) '削除処理
                rs.Close()
                cnn.Close()

                '再表示
                displayWork(ymStr, floor, True)
                MsgBox("削除しました", , "Work")
            End If
        End If
    End Sub

    ''' <summary>
    ''' 印刷ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnPrint_Click(sender As System.Object, e As System.EventArgs) Handles btnPrint.Click
        'パスワードフォーム表示
        Dim passForm As Form = New passwordForm(TopForm.iniFilePath, 2)
        If passForm.ShowDialog() <> Windows.Forms.DialogResult.OK Then
            Return
        End If

        Dim ymStr As String = ymBox.getADStr4Ym() '選択年月
        Dim floor As String = If(rbtn2F.Checked, "2", "3") '選択階
        Dim cnn As New ADODB.Connection
        cnn.Open(TopForm.DB_Work2)
        Dim rs As New ADODB.Recordset
        Dim sql = "SELECT * FROM KinD WHERE YM='" & ymStr & "' AND ((Seq2='00' AND Unt='※') OR ('20' <= Seq2 AND Seq2 <= '39')) order by Seq2" '選択年月の全てのデータ(2階、3階共に)抽出
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)

        If rs.RecordCount <= 0 Then
            MsgBox("該当がありません。", MsgBoxStyle.Exclamation, "Work")
            rs.Close()
            cnn.Close()
            Return
        Else
            rs.Close()
            '小計表示部分クリア
            subtotalClear()

            '予定の小計表示
            sql = "SELECT * FROM KinD WHERE YM='" & ymStr & "' AND ((Seq2='00' AND Unt='※') OR ('" & floor & "0' <= Seq2 AND Seq2 <= '" & floor & "9')) order by Seq"
            rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
            Dim rowIndex As Integer = 1
            Dim totalTime As Double
            While Not rs.EOF
                totalTime = 0.0
                For i As Integer = 1 To 31
                    Dim inputPlan As String = Util.checkDBNullValue(rs.Fields("Y" & i).Value) '予定勤務
                    If workTimeDictionary.ContainsKey(inputPlan) Then
                        '勤務名に対応する時間を加算
                        totalTime = totalTime + workTimeDictionary(inputPlan)
                    End If
                    If Not abbreviationDictionary.ContainsKey(inputPlan) AndAlso inputPlan <> "" Then
                        '空ではなく対応する勤務名が無い場合
                        inputPlan = "明"
                    End If
                    If abbreviationDictionary.ContainsKey(inputPlan) Then
                        Dim columnStr As String = abbreviationDictionary(inputPlan)
                        '小計（右部）
                        dgvWork(columnStr, rowIndex).Value = If(IsNumeric(dgvWork(columnStr, rowIndex).Value), CInt(dgvWork(columnStr, rowIndex).Value), 0) + 1
                        If columnStr <> "明等" Then
                            '小計（下部）
                            dgvWork("Y" & i, subtotalStrIndexDictionary(columnStr)).Value = If(IsNumeric(dgvWork("Y" & i, subtotalStrIndexDictionary(columnStr)).Value), CInt(dgvWork("Y" & i, subtotalStrIndexDictionary(columnStr)).Value), 0) + 1
                        End If
                    End If
                Next
                If totalTime <> 0.0 Then
                    '合計時間を小数第一位まで表示
                    dgvWork("月合計", rowIndex).Value = totalTime.ToString("f1")
                End If
                rowIndex += 2
                rs.MoveNext()
            End While

            '小計部分が見えるようにスクロール
            dgvWork.FirstDisplayedScrollingColumnIndex = 33

            '変更の小計表示
            Dim changeRowResult As DialogResult = MessageBox.Show("縦/横計の変更分も表示しますか？", "Work", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
            If changeRowResult = Windows.Forms.DialogResult.Yes Then
                rs.MoveFirst() 'レコードセットを先頭へ
                rowIndex = 1
                While Not rs.EOF
                    totalTime = 0.0
                    For i As Integer = 1 To 31
                        Dim inputChange As String = Util.checkDBNullValue(rs.Fields("J" & i).Value) '変更勤務
                        If workTimeDictionary.ContainsKey(inputChange) Then
                            '勤務名に対応する時間を加算
                            totalTime = totalTime + workTimeDictionary(inputChange)
                        End If
                        If Not abbreviationDictionary.ContainsKey(inputChange) AndAlso inputChange <> "" Then
                            '空ではなく対応する勤務名が無い場合
                            inputChange = "明"
                        End If
                        If abbreviationDictionary.ContainsKey(inputChange) Then
                            Dim columnStr As String = abbreviationDictionary(inputChange)
                            '小計（右部）
                            dgvWork(columnStr, rowIndex + 1).Value = If(IsNumeric(dgvWork(columnStr, rowIndex + 1).Value), CInt(dgvWork(columnStr, rowIndex + 1).Value), 0) + 1
                            If columnStr <> "明等" Then
                                '小計（下部）
                                dgvWork("Y" & i, subtotalStrIndexDictionary(columnStr) + 1).Value = If(IsNumeric(dgvWork("Y" & i, subtotalStrIndexDictionary(columnStr) + 1).Value), CInt(dgvWork("Y" & i, subtotalStrIndexDictionary(columnStr) + 1).Value), 0) + 1
                            End If
                        End If
                    Next
                    If totalTime <> 0.0 Then
                        '合計時間を小数第一位まで表示
                        dgvWork("月合計", rowIndex + 1).Value = totalTime.ToString("f1")
                    End If
                    rowIndex += 2
                    rs.MoveNext()
                End While
            End If
            rs.Close()

            '勤務割表印刷
            Dim workPrintResult As DialogResult = MessageBox.Show("勤務割表を印刷しますか？", "Work", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
            If workPrintResult = Windows.Forms.DialogResult.Yes Then
                Dim objExcel As Object = CreateObject("Excel.Application")
                Dim objWorkBooks As Object = objExcel.Workbooks

                'エクセルに書き込み
                Dim count As Integer = 0
                For Each type As String In {"２階", "３階", "常勤", "２階", "３階", "非常勤", "常勤"}
                    Dim objWorkBook As Object = objWorkBooks.Open(TopForm.excelFilePass)
                    Dim oSheet As Object = If(count <= 2, objWorkBook.Worksheets("勤務横計表改"), objWorkBook.Worksheets("勤務表改"))
                    Dim writeFlg As Boolean = If(count <= 2, writeWorkTotalTable(oSheet, cnn, type), writeWorkTable(oSheet, cnn, type))
                    If writeFlg Then
                        objExcel.DisplayAlerts = False '変更保存確認ダイアログ非表示
                        If TopForm.rbtnPrintout.Checked = True Then
                            '印刷
                            oSheet.printOut()
                        ElseIf TopForm.rbtnPreview.Checked = True Then
                            '印刷プレビュー
                            objExcel.Visible = True
                            oSheet.PrintPreview(1)
                        End If
                    End If
                    Marshal.ReleaseComObject(objWorkBook)
                    count += 1
                Next

                ' EXCEL解放
                objExcel.Quit()
                Marshal.ReleaseComObject(objExcel)
                objExcel = Nothing
            End If

            '個人別勤務割の印刷
            Dim personalPrintResult As DialogResult = MessageBox.Show("個人別勤務割を印刷しますか？", "Work", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
            If personalPrintResult = Windows.Forms.DialogResult.Yes Then
                Dim objExcel As Object = CreateObject("Excel.Application")
                Dim objWorkBooks As Object = objExcel.Workbooks
                Dim objWorkBook As Object = objWorkBooks.Open(TopForm.excelFilePass)
                Dim oSheet As Object = objWorkBook.Worksheets("ｶﾚﾝﾀﾞｰ改")

                'エクセルに書き込み
                writePersonalCalendar(oSheet, cnn)

                objExcel.DisplayAlerts = False '変更保存確認ダイアログ非表示
                If TopForm.rbtnPrintout.Checked = True Then
                    '印刷
                    oSheet.printOut()
                ElseIf TopForm.rbtnPreview.Checked = True Then
                    '印刷プレビュー
                    objExcel.Visible = True
                    oSheet.PrintPreview(1)
                End If

                ' EXCEL解放
                objExcel.Quit()
                Marshal.ReleaseComObject(objWorkBook)
                Marshal.ReleaseComObject(objExcel)
                objWorkBook = Nothing
                objExcel = Nothing
            End If
            cnn.Close()
        End If

    End Sub

    ''' <summary>
    ''' 勤務割横計表書き込み
    ''' </summary>
    ''' <param name="osheet">書き込み対象シート</param>
    ''' <param name="cnn">データベースコネクション</param>
    ''' <param name="type">勤務種類</param>
    ''' <returns>シートへの書き込みの有無</returns>
    ''' <remarks></remarks>
    Private Function writeWorkTotalTable(osheet As Object, cnn As ADODB.Connection, type As String) As Boolean
        '共通部分
        Dim ymStr As String = ymBox.getADStr4Ym() '選択年月
        Dim year As Integer = CInt(ymStr.Split("/")(0))
        Dim month As Integer = CInt(ymStr.Split("/")(1))
        osheet.Range("E2").value = ymBox.EraLabelText & " 年 " & month & " 月" '年月
        Dim daysInMonth As Integer = DateTime.DaysInMonth(year, month) '月の日数
        Dim firstDay As DateTime = New DateTime(year, month, 1)
        Dim weekNumber As Integer = CInt(firstDay.DayOfWeek) '初日の曜日のindex
        Dim sundayColumnAlphabetList As New List(Of String)
        Dim columnAlphabet As String
        Dim dayChar As String
        For i As Integer = 1 To daysInMonth '曜日書き込み
            columnAlphabet = getColumnAlphabet(4 + i)
            dayChar = dayCharArray((weekNumber + (i - 1)) Mod 7)
            If dayChar = "日" Then
                sundayColumnAlphabetList.Add(columnAlphabet)
            End If
            osheet.range(columnAlphabet & "5").value = dayChar
        Next
        For Each alphabet As String In sundayColumnAlphabetList '日曜日の列の色設定
            For i As Integer = 4 To 45
                osheet.range(alphabet & i).Interior.ColorIndex = 27
            Next
        Next

        '書き込み処理
        osheet.Range("I2").value = type
        Dim sql As String
        If type = "２階" Then
            '2階
            sql = "SELECT * FROM KinD WHERE YM='" & ymStr & "' AND ((Seq2='00' AND Unt='※') OR ('20' <= Seq2 AND Seq2 <= '29')) order by Seq"
        ElseIf type = "３階" Then
            '3階
            sql = "SELECT * FROM KinD WHERE YM='" & ymStr & "' AND ((Seq2='00' AND Unt='※') OR ('30' <= Seq2 AND Seq2 <= '39')) order by Seq"
        Else
            '常勤
            sql = "SELECT * FROM KinD WHERE YM='" & ymStr & "' AND Rdr<>'' order by Seq2, Seq"
        End If
        Dim rs As New ADODB.Recordset
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
        If rs.RecordCount <= 0 Then
            '該当データがなくエクセルへ書き込み処理が無い場合Falseを返す
            rs.Close()
            Return False
        Else
            If rs.RecordCount >= 21 Then
                '２枚目枠作成
                Dim xlPasteRange As Excel.Range = osheet.Range("A47") 'ペースト先
                osheet.rows("1:46").copy(xlPasteRange)
            End If

            Dim unit As String
            Dim unitTmp As String = ""
            Dim index As Integer = 6
            Dim yjVal1(39, 36) As String '1枚目データ用配列
            Dim yjVal2(39, 36) As String '2枚目データ用配列
            Dim yVal, jVal As String
            Dim workTypeIndexDic As New Dictionary(Of String, Integer) From {{"公", 31}, {"夜", 32}, {"深", 33}, {"遅", 34}, {"遅々", 35}, {"明", 36}}

            '小計の変更行に0をセット
            For i As Integer = 1 To 39 Step 2
                For j As Integer = 31 To 36
                    yjVal1(i, j) = "0"
                    yjVal2(i, j) = "0"
                Next
            Next

            Dim border As Excel.Border
            While Not rs.EOF
                'ユニット
                unit = Util.checkDBNullValue(rs.Fields("Unt").Value)
                If unit = unitTmp Then
                    If unit = "※" Then
                        osheet.range("B" & index).value = unit
                    Else
                        osheet.range("B" & index).value = ""
                    End If
                Else
                    osheet.range("B" & index).value = unit
                    unitTmp = unit
                    '罫線
                    border = osheet.Range("B" & index, "AO" & index).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlThin
                End If
                'Rdr
                osheet.range("C" & index).value = Util.checkDBNullValue(rs.Fields("Rdr").Value)
                '氏名
                osheet.range("D" & index).value = Util.checkDBNullValue(rs.Fields("Nam").Value)
                '予定と変更
                If index <= 46 Then '1枚目データ作成
                    For i As Integer = 1 To 31
                        yVal = Util.checkDBNullValue(rs.Fields("Y" & i).Value)
                        jVal = Util.checkDBNullValue(rs.Fields("J" & i).Value)
                        yjVal1(index - 6, i - 1) = yVal
                        If yVal <> jVal Then
                            yjVal1((index + 1) - 6, i - 1) = jVal
                        End If
                        If workTypeIndexDic.ContainsKey(yVal) Then
                            yjVal1(index - 6, workTypeIndexDic(yVal)) = CInt(yjVal1(index - 6, workTypeIndexDic(yVal))) + 1
                        End If
                        If workTypeIndexDic.ContainsKey(jVal) Then
                            yjVal1((index + 1) - 6, workTypeIndexDic(jVal)) = CInt(yjVal1((index + 1) - 6, workTypeIndexDic(jVal))) + 1
                        End If
                    Next
                Else '2枚目データ作成
                    For i As Integer = 1 To 31
                        yVal = Util.checkDBNullValue(rs.Fields("Y" & i).Value)
                        jVal = Util.checkDBNullValue(rs.Fields("J" & i).Value)
                        yjVal2(index - 52, i - 1) = yVal
                        If yVal <> jVal Then
                            yjVal2((index + 1) - 52, i - 1) = jVal
                        End If
                        If workTypeIndexDic.ContainsKey(yVal) Then
                            yjVal2(index - 52, workTypeIndexDic(yVal)) = CInt(yjVal2(index - 52, workTypeIndexDic(yVal))) + 1
                        End If
                        If workTypeIndexDic.ContainsKey(jVal) Then
                            yjVal2((index + 1) - 52, workTypeIndexDic(jVal)) = CInt(yjVal2((index + 1) - 52, workTypeIndexDic(jVal))) + 1
                        End If
                    Next
                End If

                rs.MoveNext()
                index += 2
                If index = 46 Then
                    index = 52
                End If
            End While

            '小計部分
            For i As Integer = 1 To 39 Step 2
                For j As Integer = 31 To 36
                    '予定が空、または予定と変更が同じならば変更を空にする
                    If yjVal1(i - 1, j) = "" OrElse yjVal1(i - 1, j) = yjVal1(i, j) Then
                        yjVal1(i, j) = ""
                    End If
                    If yjVal2(i - 1, j) = "" OrElse yjVal2(i - 1, j) = yjVal2(i, j) Then
                        yjVal2(i, j) = ""
                    End If
                Next
            Next

            'シートの対象範囲に作成データをセット
            osheet.range("E6", "AO45").value = yjVal1 '1枚目
            If rs.RecordCount >= 21 Then
                osheet.range("E52", "AO91").value = yjVal2 '2枚目
            End If

            rs.Close()
            Return True
        End If
    End Function

    ''' <summary>
    ''' 勤務割表書き込み
    ''' </summary>
    ''' <param name="osheet">書き込み対象シート</param>
    ''' <param name="cnn">データベースコネクション</param>
    ''' <param name="type">勤務種類</param>
    ''' <returns>シートへの書き込みの有無</returns>
    ''' <remarks></remarks>
    Private Function writeWorkTable(osheet As Object, cnn As ADODB.Connection, type As String) As Boolean
        '共通部分
        Dim ymStr As String = ymBox.getADStr4Ym() '選択年月
        Dim year As Integer = CInt(ymStr.Split("/")(0))
        Dim month As Integer = CInt(ymStr.Split("/")(1))
        osheet.Range("E2").value = ymBox.EraLabelText & " 年 " & month & " 月" '年月
        Dim daysInMonth As Integer = DateTime.DaysInMonth(year, month) '月の日数
        Dim firstDay As DateTime = New DateTime(year, month, 1)
        Dim weekNumber As Integer = CInt(firstDay.DayOfWeek) '初日の曜日のindex
        Dim sundayColumnAlphabetList As New List(Of String)
        Dim columnAlphabet As String
        Dim dayChar As String
        For i As Integer = 1 To daysInMonth '曜日書き込み
            columnAlphabet = getColumnAlphabet(4 + i)
            dayChar = dayCharArray((weekNumber + (i - 1)) Mod 7)
            If dayChar = "日" Then
                sundayColumnAlphabetList.Add(columnAlphabet)
            End If
            osheet.range(columnAlphabet & "5").value = dayChar
        Next
        For Each alphabet As String In sundayColumnAlphabetList '日曜日の列の色設定
            For i As Integer = 4 To 45
                osheet.range(alphabet & i).Interior.ColorIndex = 27
            Next
        Next

        '書き込み処理
        osheet.Range("I2").value = type
        Dim sql As String
        If type = "２階" Then
            '2階
            sql = "SELECT * FROM KinD WHERE YM='" & ymStr & "' AND ((Seq2='00' AND Unt='※') OR ('20' <= Seq2 AND Seq2 <= '29')) order by Seq"
        ElseIf type = "３階" Then
            '3階
            sql = "SELECT * FROM KinD WHERE YM='" & ymStr & "' AND ((Seq2='00' AND Unt='※') OR ('30' <= Seq2 AND Seq2 <= '39')) order by Seq"
        ElseIf type = "非常勤" Then
            '非常勤
            sql = "SELECT * FROM KinD WHERE YM='" & ymStr & "' AND Rdr='' order by Seq2, Seq"
        Else
            '常勤
            sql = "SELECT * FROM KinD WHERE YM='" & ymStr & "' AND Rdr<>'' order by Seq2, Seq"
        End If
        Dim rs As New ADODB.Recordset
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
        If rs.RecordCount <= 0 Then
            '該当データがなくエクセルへ書き込み処理が無い場合Falseを返す
            rs.Close()
            Return False
        Else
            If rs.RecordCount >= 21 Then
                '２枚目枠作成
                Dim xlPasteRange As Excel.Range = osheet.Range("A47") 'ペースト先
                osheet.rows("1:46").copy(xlPasteRange)
            End If

            Dim unit As String
            Dim unitTmp As String = ""
            Dim index As Integer = 6
            Dim yjVal1(39, 31) As String '1枚目データ用配列
            Dim yjVal2(39, 31) As String '2枚目データ用配列
            Dim yVal, jVal As String
            Dim workTypeIndexDic As New Dictionary(Of String, Integer) From {{"公", 31}}

            '小計の変更行に0をセット
            For i As Integer = 1 To 39 Step 2
                For j As Integer = 31 To 31
                    yjVal1(i, j) = "0"
                    yjVal2(i, j) = "0"
                Next
            Next

            Dim border As Excel.Border
            While Not rs.EOF
                'ユニット
                unit = Util.checkDBNullValue(rs.Fields("Unt").Value)
                If type = "非常勤" Then
                    osheet.range("B" & index).value = unit
                ElseIf unit = unitTmp Then
                    If unit = "※" Then
                        osheet.range("B" & index).value = unit
                    Else
                        osheet.range("B" & index).value = ""
                    End If
                Else
                    osheet.range("B" & index).value = unit
                    unitTmp = unit
                    border = osheet.Range("B" & index, "AJ" & index).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlThin
                End If
                'Rdr
                osheet.range("C" & index).value = Util.checkDBNullValue(rs.Fields("Rdr").Value)
                '氏名
                osheet.range("D" & index).value = Util.checkDBNullValue(rs.Fields("Nam").Value)
                '予定と変更
                If index <= 46 Then '1枚目データ作成
                    For i As Integer = 1 To 31
                        yVal = Util.checkDBNullValue(rs.Fields("Y" & i).Value)
                        jVal = Util.checkDBNullValue(rs.Fields("J" & i).Value)
                        yjVal1(index - 6, i - 1) = yVal
                        If yVal <> jVal Then
                            yjVal1((index + 1) - 6, i - 1) = jVal
                        End If
                        If workTypeIndexDic.ContainsKey(yVal) Then
                            yjVal1(index - 6, workTypeIndexDic(yVal)) = CInt(yjVal1(index - 6, workTypeIndexDic(yVal))) + 1
                        End If
                        If workTypeIndexDic.ContainsKey(jVal) Then
                            yjVal1((index + 1) - 6, workTypeIndexDic(jVal)) = CInt(yjVal1((index + 1) - 6, workTypeIndexDic(jVal))) + 1
                        End If
                    Next
                Else '2枚目データ作成
                    For i As Integer = 1 To 31
                        yVal = Util.checkDBNullValue(rs.Fields("Y" & i).Value)
                        jVal = Util.checkDBNullValue(rs.Fields("J" & i).Value)
                        yjVal2(index - 52, i - 1) = yVal
                        If yVal <> jVal Then
                            yjVal2((index + 1) - 52, i - 1) = jVal
                        End If
                        If workTypeIndexDic.ContainsKey(yVal) Then
                            yjVal2(index - 52, workTypeIndexDic(yVal)) = CInt(yjVal2(index - 52, workTypeIndexDic(yVal))) + 1
                        End If
                        If workTypeIndexDic.ContainsKey(jVal) Then
                            yjVal2((index + 1) - 52, workTypeIndexDic(jVal)) = CInt(yjVal2((index + 1) - 52, workTypeIndexDic(jVal))) + 1
                        End If
                    Next
                End If

                rs.MoveNext()
                index += 2
                If index = 46 Then
                    index = 52
                End If
            End While

            '小計部分
            For i As Integer = 1 To 39 Step 2
                '予定が空、または予定と変更が同じならば変更を空にする
                For j As Integer = 31 To 31
                    If yjVal1(i - 1, j) = "" OrElse yjVal1(i - 1, j) = yjVal1(i, j) Then
                        yjVal1(i, j) = ""
                    End If
                    If yjVal2(i - 1, j) = "" OrElse yjVal2(i - 1, j) = yjVal2(i, j) Then
                        yjVal2(i, j) = ""
                    End If
                Next
            Next

            'シートの対象範囲に作成データをセット
            osheet.range("E6", "AJ45").value = yjVal1 '1枚目
            If rs.RecordCount >= 21 Then
                osheet.range("E52", "AJ91").value = yjVal2 '2枚目
            End If

            rs.Close()
            Return True
        End If
    End Function

    ''' <summary>
    ''' 個人別勤務割書き込み
    ''' </summary>
    ''' <param name="oSheet">書き込み対象シート</param>
    ''' <param name="cnn">データベースコネクション</param>
    ''' <remarks></remarks>
    Private Sub writePersonalCalendar(oSheet As Object, cnn As ADODB.Connection)
        Dim ymStr As String = ymBox.getADStr4Ym() '選択年月
        Dim year As Integer = CInt(ymStr.Split("/")(0))
        Dim month As Integer = CInt(ymStr.Split("/")(1))
        oSheet.Range("C1").value = ymBox.EraLabelText & " 年 " & month & " 月" '年月
        oSheet.Range("C31").value = ymBox.EraLabelText & " 年 " & month & " 月" '年月

        Dim sql As String = "SELECT * FROM KinD WHERE YM='" & ymStr & "' AND ((Seq2='00' AND Unt='※') OR ('20' <= Seq2 AND Seq2 <= '39')) order by Seq2, Seq" '選択年月の全てのデータ(2階、3階共に)抽出
        Dim rs As New ADODB.Recordset
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
        Dim personCount As Integer = rs.RecordCount '人数

        '人数分の枠準備
        Dim forCount As Integer
        For forCount = 1 To ((personCount - 1) \ 2)
            'コピペ処理
            Dim xlPasteRange As Excel.Range = oSheet.Range("A" & ((forCount * 53) + 1)) 'ペースト先
            oSheet.rows("1:53").copy(xlPasteRange)
        Next
        If (personCount Mod 2) = 1 Then
            oSheet.Range("C" & (((forCount - 1) * 53) + 1 + 30)).value = ""
        End If

        '勤務データ作成
        Dim daysInMonth As Integer = DateTime.DaysInMonth(year, month) '月の日数
        Dim firstDay As DateTime = New DateTime(year, month, 1)
        Dim weekNumber As Integer = CInt(firstDay.DayOfWeek) '初日の曜日のindex
        Dim count As Integer = 1
        While Not rs.EOF
            'データ作成
            Dim yVal, jVal As String
            Dim numIndex As Integer = weekNumber
            Dim workData(17, 6) As String
            For i As Integer = 1 To daysInMonth
                workData((numIndex \ 7) * 3, numIndex Mod 7) = i '日にち
                yVal = Util.checkDBNullValue(rs.Fields("Y" & i).Value)
                jVal = Util.checkDBNullValue(rs.Fields("J" & i).Value)
                workData((numIndex \ 7) * 3 + 1, numIndex Mod 7) = yVal '予定
                If yVal <> jVal Then
                    workData((numIndex \ 7) * 3 + 2, numIndex Mod 7) = jVal '変更
                End If
                numIndex += 1
            Next

            'エクセルにデータ貼り付け
            If (count Mod 2) = 1 Then
                'ページ上部
                oSheet.range("E" & (53 * (count \ 2) + 1)).value = Util.checkDBNullValue(rs.Fields("Nam").Value) '氏名
                oSheet.range("B" & (53 * (count \ 2) + 4), "H" & (53 * (count \ 2) + 21)).value = workData '勤務データ
            Else
                'ページ下部
                oSheet.range("E" & (53 * ((count - 1) \ 2) + 31)).value = Util.checkDBNullValue(rs.Fields("Nam").Value) '氏名
                oSheet.range("B" & (53 * ((count - 1) \ 2) + 34), "H" & (53 * ((count - 1) \ 2) + 51)).value = workData '勤務データ
            End If

            rs.MoveNext()
            count += 1
        End While
    End Sub

    ''' <summary>
    ''' エクセル列番号文字列を取得
    ''' </summary>
    ''' <param name="num">列番号数値</param>
    ''' <returns>エクセル列番号文字</returns>
    ''' <remarks></remarks>
    Private Function getColumnAlphabet(num As Integer) As String
        Dim s As String = ""
        Do While num > 0
            num -= 1
            Dim m As Integer = num Mod NAME_COLUMN_VALUES_LENGTH
            s = NAME_COLUMN_VALUES(m) & s
            num = Math.Floor(num / NAME_COLUMN_VALUES_LENGTH)
        Loop
        Return s
    End Function

End Class