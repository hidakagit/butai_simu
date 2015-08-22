Imports Microsoft.VisualBasic.FileIO
Imports System.Text

Public Class Form12
    Private Sub Form_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' データをすべてクリア
        DataGridView1.Rows.Clear()
        Call CsvLoad(fspath, DataGridView1)
    End Sub

    '閉じる時に設定を保存する
    Private Sub Form1_Closed(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.FormClosing
        SaveToCsv(Me.DataGridView1, fspath)
    End Sub
End Class