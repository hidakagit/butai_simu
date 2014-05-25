Public Class Form11

    Private Sub Form11_load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            With Form10
                If .syosklflg = True Then
                    CheckBox1.Checked = True
                End If
                For i As Integer = 0 To 2 '初期スキル
                    If Not .simu_bs(i).skill(0).name = "" Then
                        Label(Me, "00" & Val(i + 1)).Text = .simu_bs(i).skill(0).name
                    End If
                Next
                If Not .cus_addskl Is Nothing Then '空じゃなければ
                    For i As Integer = 0 To 3
                        ComboBox(Me, Val(i) & "09").Text = .cus_addskl(i)(1)
                        ComboBox(Me, Val(i) & "10").Text = .cus_addskl(i)(3)
                        ComboBox(Me, Val(i) & "11").Text = .cus_addskl(i)(0)
                        ComboBox(Me, Val(i) & "12").Text = .cus_addskl(i)(2)
                    Next
                End If
            End With
        Catch ex As Exception
        End Try
    End Sub

    Private Sub 詳細設定ON(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Form10.syosklflg = True
            スキル詳細設定関連ONOFF(False)
        Else
            Form10.syosklflg = False
            スキル詳細設定関連ONOFF(True)
        End If
    End Sub

    Public Sub スキル詳細設定関連ONOFF(ByVal TorF As Boolean)
        With Form10
            .ComboBox01.Enabled = TorF
            .ComboBox02.Enabled = TorF
            .ComboBox011.Enabled = TorF
            .ComboBox012.Enabled = TorF

            If TorF Then
                .GroupBox3.BackColor = SystemColors.Control
                .GroupBox5.BackColor = SystemColors.Control
                With .Label14
                    .ForeColor = Color.Blue
                    .Text = "OFF"
                End With
            Else
                .GroupBox3.BackColor = SystemColors.ControlLight
                .GroupBox5.BackColor = SystemColors.ControlLight
                With .Label14
                    .ForeColor = Color.Red
                    .Text = "ON"
                End With
            End If

            '.ComboBox41.Enabled = TorF
            '.ComboBox42.Enabled = TorF
            '.ComboBox041.Enabled = TorF
            '.ComboBox042.Enabled = TorF

            .CheckBox1.Enabled = TorF
        End With
    End Sub

    Private Sub 追加スキル表示(sender As Object, e As EventArgs) _
       Handles ComboBox009.SelectedIndexChanged, ComboBox010.SelectedIndexChanged, _
               ComboBox109.SelectedIndexChanged, ComboBox110.SelectedIndexChanged, _
               ComboBox209.SelectedIndexChanged, ComboBox210.SelectedIndexChanged, _
               ComboBox309.SelectedIndexChanged, ComboBox310.SelectedIndexChanged
        Dim p As DataSet
        Dim c As ComboBox = Nothing
        Select Case (String_onlyNumber(sender.name).Remove(0, 1))
            Case "09"
                c = ComboBox(Me, String_onlyNumber(sender.name).Substring(0, 1) & "11")
            Case "10"
                c = ComboBox(Me, String_onlyNumber(sender.name).Substring(0, 1) & "12")
        End Select
        p = DB_TableOUT(con, cmd, "SELECT Index,分類,名前,LV FROM Skill WHERE 分類 = """ & sender.text & """ AND LV = 1 ORDER BY Index", "Skill")
        With c
            .ValueMember = "Index"
            .DisplayMember = "名前"
            .DataSource = p.Tables("Skill")
            .SelectedIndex = -1
        End With
    End Sub

    Private Sub 確定(sender As Object, e As EventArgs) Handles Button1.Click
        For i As Integer = 0 To 3
            Form10.cus_addskl(i) = {ComboBox(Me, Val(i) & "11").Text, ComboBox(Me, Val(i) & "09").Text, _
                                    ComboBox(Me, Val(i) & "12").Text, ComboBox(Me, Val(i) & "10").Text}
        Next
        Me.Close()
    End Sub

    Private Sub 消去(sender As Object, e As EventArgs) Handles Button2.Click
        Form10.syosklflg = False
        スキル詳細設定関連ONOFF(True)
        Me.Close()
    End Sub
End Class