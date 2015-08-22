Public Class Form14
    Private Sub Form_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' データをすべてクリア
        DataGridView1.Rows.Clear()
        Call CsvLoad(fcpath, DataGridView1)
    End Sub

    Private Sub 設定適用再起動(sender As Object, e As EventArgs) Handles Button1.Click
        SaveToCsv(Me.DataGridView1, fcpath)
        Application.Restart()
    End Sub
End Class