Public Class CSV作成

    '保存ファイル名定型部
    Private Const DEFAULT_SAVE_NAME As String = "勤務割Annex"

    'ヘッダー文字列
    Private columnCaption() As String = {"表示順", "対象年月", "勤務表", "職員№", "氏名", "予形態", "予職種", "予1", "予2", "予3", "予4", "予5", "予6", "予7", "予8", "予9", "予10", "予11", "予12", "予13", "予14", "予15", "予16", "予17", "予18", "予19", "予20", "予21", "予22", "予23", "予24", "予25", "予26", "予27", "予28", "予29", "予30", "予31", "予換算", "実形態", "実職種", "実1", "実2", "実3", "実4", "実5", "実6", "実7", "実8", "実9", "実10", "実11", "実12", "実13", "実14", "実15", "実16", "実17", "実18", "実19", "実20", "実21", "実22", "実23", "実24", "実25", "実26", "実27", "実28", "実29", "実30", "実31", "実換算"}

    '勤務名対応dic
    Private workDictionary As Dictionary(Of String, String)

    ''' <summary>
    ''' loadイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub CSV作成_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        'dic作成
        createDictionary()

        '年月ボックスを現在年月に設定
        ymBox.setADStr(Today.ToString("yyyy/MM/dd"))
    End Sub

    ''' <summary>
    ''' 勤務対応dic作成
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub createDictionary()
        workDictionary = New Dictionary(Of String, String)
        workDictionary.Add("日", "日勤")
        workDictionary.Add("遅々", "遅々")
        workDictionary.Add("夜", "夜勤")
        workDictionary.Add("深", "深夜")
        workDictionary.Add("半", "半")
        workDictionary.Add("半Ａ", "半Ａ")
        workDictionary.Add("半Ｂ", "半Ｂ")
        workDictionary.Add("半夜", "半夜")
        workDictionary.Add("半行", "半行")
        workDictionary.Add("研", "研修")
        workDictionary.Add("有", "有休")
        workDictionary.Add("公", "公休")
        workDictionary.Add("明", "明け")
        workDictionary.Add("希", "")
        workDictionary.Add("A", "5.0")
        workDictionary.Add("B", "5.5")
        workDictionary.Add("C", "7.0")
        workDictionary.Add("D", "3.5")
        workDictionary.Add("E", "5.0")
        workDictionary.Add("F", "6.0")
        workDictionary.Add("G", "7.0")
        workDictionary.Add("H", "4.0")
        workDictionary.Add("I", "3.0")
        workDictionary.Add("J", "5.5")
        workDictionary.Add("K", "7.0")
        workDictionary.Add("L", "2.5")
        workDictionary.Add("M", "3.5")
        workDictionary.Add("N", "2.0")
        'workDictionary.Add("P", "6.5")
        workDictionary.Add("R", "2.5")
        workDictionary.Add("S", "7.5")
        'workDictionary.Add("T", "4.5")
    End Sub

    ''' <summary>
    ''' 実行ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnExecution_Click(sender As System.Object, e As System.EventArgs) Handles btnExecution.Click
        'パスワードフォーム表示
        Dim passForm As Form = New passwordForm(TopForm.iniFilePath, 2)
        If passForm.ShowDialog() <> Windows.Forms.DialogResult.OK Then
            Return
        End If

        '保存ファイル名等初期値
        Dim ymStr As String = ymBox.getADYmStr() '年月
        Me.saveCsvFileDialog.FileName = DEFAULT_SAVE_NAME & ymStr.Replace("/", "") & ".csv"
        Me.saveCsvFileDialog.Filter = "Csv|"

        '保存ダイアログでファイル名を設定した場合に処理を実行します。
        If Me.saveCsvFileDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim strResult As New System.Text.StringBuilder

            'ヘッダー部分作成
            Dim columnCount As Integer = columnCaption.Length - 1
            For i As Integer = 0 To columnCount
                Dim s As String = EncloseDoubleQuotesIfNeed(columnCaption(i)) '"で囲む
                strResult.Append(s)
                'カンマ追加
                If i < columnCount Then
                    strResult.Append(",")
                End If
            Next
            strResult.Append(vbCrLf) '改行

            'レコード作成
            Dim count As Integer = 1
            Dim cnn As New ADODB.Connection
            cnn.Open(TopForm.DB_Work)
            Dim rs As New ADODB.Recordset
            '常勤部分
            Dim sql = "SELECT * FROM KinD WHERE YM='" & ymStr & "' AND Rdr<>'' AND ((Seq2='00' AND Unt='※') OR ('20' <= Seq2 AND Seq2 <= '39')) order by Rdr, Seq, Seq2"
            rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
            While Not rs.EOF
                writeWorkData(rs, strResult, count)
                rs.MoveNext()
                count += 1
            End While
            rs.Close()
            '非常勤部分
            sql = "SELECT * FROM KinD WHERE YM='" & ymStr & "' AND Rdr='' AND ((Seq2='00' AND Unt='※') OR ('20' <= Seq2 AND Seq2 <= '39')) order by Seq, Seq2"
            rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
            While Not rs.EOF
                writeWorkData(rs, strResult, count)
                rs.MoveNext()
                count += 1
            End While

            '保存処理等
            Dim fileName As String = If(Me.saveCsvFileDialog.FileName.EndsWith(".csv"), Me.saveCsvFileDialog.FileName, Me.saveCsvFileDialog.FileName & ".csv") 'ファイル名
            Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS") 'エンコードをShift_JISに
            Dim sw As New System.IO.StreamWriter(fileName, False, enc)
            sw.Write(strResult.ToString)
            sw.Close()
            MsgBox("勤務割ＦＤの書き出しが終了しました。", MsgBoxStyle.Information, "Work")
        End If
    End Sub

    ''' <summary>
    ''' レコードセットのデータを整形してStringBuilderに追加
    ''' </summary>
    ''' <param name="rs">追加するレコードセット</param>
    ''' <param name="sb">追加されるStringBuilder</param>
    ''' <param name="count">表示順番号</param>
    ''' <remarks></remarks>
    Private Sub writeWorkData(rs As ADODB.Recordset, sb As System.Text.StringBuilder, count As Integer)
        '表示順
        sb.Append(EncloseDoubleQuotesIfNeed(count.ToString()) & ",")
        '対象年月
        sb.Append(EncloseDoubleQuotesIfNeed(Util.checkDBNullValue(rs.Fields("Ym").Value)) & ",")
        '勤務表
        sb.Append(EncloseDoubleQuotesIfNeed("特養") & ",")
        '職員№
        sb.Append(EncloseDoubleQuotesIfNeed("0") & ",")
        '氏名
        sb.Append(EncloseDoubleQuotesIfNeed(Util.checkDBNullValue(rs.Fields("Nam").Value)) & ",")
        '予形態,予職種
        If Util.checkDBNullValue(rs.Fields("Rdr").Value) <> "" Then
            sb.Append(EncloseDoubleQuotesIfNeed("常勤専従") & ",")
            sb.Append(EncloseDoubleQuotesIfNeed("介護職") & ",")
        Else
            sb.Append(EncloseDoubleQuotesIfNeed("常勤以外専従") & ",")
            sb.Append(EncloseDoubleQuotesIfNeed("介護職ﾊﾟｰﾄ") & ",")
        End If
        '予1～予31
        For i As Integer = 1 To 31
            Dim yVal As String = Util.checkDBNullValue(rs.Fields("Y" & i).Value)
            If workDictionary.ContainsKey(yVal) Then
                sb.Append(EncloseDoubleQuotesIfNeed(workDictionary(yVal)) & ",")
            Else
                sb.Append(EncloseDoubleQuotesIfNeed(yVal) & ",")
            End If
        Next
        '予換算
        sb.Append("" & ",")
        '実形態
        sb.Append("" & ",")
        '実職種
        sb.Append("" & ",")
        '実1～実31
        For i As Integer = 1 To 31
            Dim jVal As String = Util.checkDBNullValue(rs.Fields("J" & i).Value)
            If workDictionary.ContainsKey(jVal) Then
                sb.Append(EncloseDoubleQuotesIfNeed(workDictionary(jVal)) & ",")
            Else
                sb.Append(EncloseDoubleQuotesIfNeed(jVal) & ",")
            End If
        Next
        '実換算
        sb.Append("")
        '改行
        sb.Append(vbCrLf)
    End Sub

    ''' <summary>
    ''' 必要ならば、文字列をダブルクォートで囲む
    ''' </summary>
    Private Function EncloseDoubleQuotesIfNeed(field As String) As String
        If NeedEncloseDoubleQuotes(field) Then
            Return EncloseDoubleQuotes(field)
        End If
        Return field
    End Function

    ''' <summary>
    ''' 文字列をダブルクォートで囲む
    ''' </summary>
    Private Function EncloseDoubleQuotes(field As String) As String
        If field.IndexOf(""""c) > -1 Then
            '"を""とする
            field = field.Replace("""", """""")
        End If
        Return """" & field & """"
    End Function

    ''' <summary>
    ''' 文字列をダブルクォートで囲む必要があるか調べる
    ''' </summary>
    Private Function NeedEncloseDoubleQuotes(field As String) As Boolean
        Return field.IndexOf(""""c) > -1 OrElse _
            field.IndexOf(","c) > -1 OrElse _
            field.IndexOf(ControlChars.Cr) > -1 OrElse _
            field.IndexOf(ControlChars.Lf) > -1 OrElse _
            field.StartsWith(" ") OrElse _
            field.StartsWith(vbTab) OrElse _
            field.EndsWith(" ") OrElse _
            field.EndsWith(vbTab)
    End Function
End Class