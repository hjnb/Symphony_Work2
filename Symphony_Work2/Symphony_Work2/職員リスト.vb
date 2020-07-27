Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Public Class 職員リスト
    Public CellStylepink As DataGridViewCellStyle
    Public CellStyleWhite As DataGridViewCellStyle

    Private Sub madestyle()
        CellStylepink = New DataGridViewCellStyle()
        CellStylepink.BackColor = Color.FromArgb(255, 192, 255)
        CellStylepink.Alignment = DataGridViewContentAlignment.MiddleCenter

        CellStyleWhite = New DataGridViewCellStyle()
        CellStyleWhite.BackColor = Color.FromArgb(255, 255, 255)
        CellStyleWhite.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Private Sub 職員リスト_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        btnPaint.BackColor = Color.FromArgb(255, 192, 255)

        Dim Ym As String = Strings.Left(CType(Me.Owner, 週間表).lblYmd.Text, 10)
        Dim reader As System.Data.OleDb.OleDbDataReader
        Dim Cn As New OleDbConnection(TopForm.DB_Work2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        SQLCm.CommandText = "select Nam from KinD WHERE Ym = '" & Ym & "' order by seq2, seq"
        Cn.Open()
        reader = SQLCm.ExecuteReader()
        While reader.Read() = True
            lstName.Items.Add(reader("Nam"))
        End While
        reader.Close()

        Cn.Close()

        Dim SQLCm1 As OleDbCommand = Cn.CreateCommand
        Dim Adapter1 As New OleDbDataAdapter(SQLCm1)
        Dim Table1 As New DataTable
        SQLCm1.CommandText = "SELECT * FROM SNam"
        Adapter1.Fill(Table1)
        DataGridView1.DataSource = Table1

        madestyle()

    End Sub

    Private Sub btnPaint_Click(sender As System.Object, e As System.EventArgs) Handles btnPaint.Click
        If btnPaint.Text = "PAINT" Then
            lstName.BackColor = Color.FromArgb(255, 192, 255)
            btnPaint.Text = "RESET"
            btnPaint.BackColor = Color.FromArgb(234, 234, 234)
        Else
            lstName.BackColor = Color.FromArgb(234, 234, 234)
            btnPaint.Text = "PAINT"
            btnPaint.BackColor = Color.FromArgb(255, 192, 255)
        End If
    End Sub

    Private Sub lstName_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles lstName.SelectedIndexChanged
        Dim DGV1rowcount As Integer = DataGridView1.Rows.Count
        Dim cell As DataGridViewCell = CType(Me.Owner, 週間表).DataGridView1.CurrentCell

        If cell.RowIndex < 1 OrElse cell.ColumnIndex = 0 OrElse (cell.ColumnIndex Mod 2) = 1 Then
            Return
        End If

        If btnPaint.Text = "PAINT" Then
            cell.Value = Strings.Left(lstName.Text, If(lstName.Text.IndexOf("　") >= 0, lstName.Text.IndexOf("　"), 3))
            cell.Style = CellStyleWhite
            For i As Integer = 0 To DGV1rowcount - 1
                If lstName.Text = DataGridView1(0, i).Value Then
                    cell.Value = DataGridView1(1, i).Value
                End If
            Next
        Else
            cell.Value = Strings.Left(lstName.Text, If(lstName.Text.IndexOf("　") >= 0, lstName.Text.IndexOf("　"), 3))
            cell.Style = CellStylepink
            For i As Integer = 0 To DGV1rowcount - 1
                If lstName.Text = DataGridView1(0, i).Value Then
                    cell.Value = DataGridView1(1, i).Value
                End If
            Next
        End If

    End Sub
End Class