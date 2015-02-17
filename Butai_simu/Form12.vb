Imports Microsoft.VisualBasic.FileIO
Imports System.Text

Public Class Form12
    Private Sub Form_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' データをすべてクリア
        DataGridView1.Rows.Clear()
        Call CsvLoad(fspath, DataGridView1)
    End Sub

    Public Sub SaveToCsv(ByVal tempDgv As DataGridView)
        '変数を定義します。
        Dim i As Integer
        Dim j As Integer
        Dim strResult As New System.Text.StringBuilder
        For i = 0 To tempDgv.Rows.Count - 2
            For j = 0 To tempDgv.Columns.Count - 1
                Select Case j
                    Case 0
                        strResult.Append("""" & _
                        tempDgv.Rows(i).Cells(j).Value.ToString & _
                        """")

                    Case tempDgv.Columns.Count - 1
                        strResult.Append("," & """" & _
                        tempDgv.Rows(i).Cells(j).Value.ToString & _
                        """" & vbCrLf)

                    Case Else
                        strResult.Append("," & """" & _
                        tempDgv.Rows(i).Cells(j).Value.ToString & _
                        """")
                End Select
            Next
        Next
        'Shift-JISで保存します。
        Dim swText As New System.IO.StreamWriter(fspath, False, System.Text.Encoding.GetEncoding(932))
        swText.Write(strResult.ToString)
        swText.Dispose()
    End Sub

    '閉じる時に設定を保存する
    Private Sub Form1_Closed(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.FormClosing
        SaveToCsv(DataGridView1)
    End Sub
End Class