Public Class Form4
    Private Sub Form4_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Form1.ToolStripComboBox1.Text = "攻撃" Then
            ComboBox1.Text = "攻撃"
        Else
            ComboBox1.Text = "防御"
        End If
        With bskill
            If Not .flg = False Then
                CheckBox1.Checked = True
                If .koubou = "攻" Then
                    ComboBox1.Text = "攻撃"
                Else
                    ComboBox1.Text = "防御"
                End If
                TextBox1.Text = .kouka_p * 100
                TextBox2.Text = .kouka_f * 100
                ComboBox2.Text = .taisyo
                If Not .speed = 0 Then '加速スキルが入っていれば
                    CheckBox2.Checked = True
                    TextBox3.Text = .speed * 100
                End If
            End If
        End With
    End Sub

    Private Sub 部隊スキル情報入力(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If CheckBox1.Checked = False Or Val(TextBox1.Text) = -1 Or Val(TextBox2.Text) = -1 Then '部隊スキルそのもの、もしくは設定が無効ならば
            Me.Close()
            Exit Sub
        End If
        With bskill
            If ComboBox1.Text = "攻撃" Then
                bskill.koubou = "攻"
            Else
                bskill.koubou = "防"
            End If
            .kouka_p = Val(TextBox1.Text) * 0.01
            .kouka_f = Val(TextBox2.Text) * 0.01
            .taisyo = ComboBox2.Text
            .ONOFF = True
            If CheckBox3.Checked = True Then '将攻二乗モードON
                .qq = True
            Else
                .qq = False
            End If
        End With
        Me.Close()
    End Sub

    Private Sub 部隊スキル初期化()
        bskill.ONOFF = False
        bskill = Nothing
    End Sub

    Private Sub 部隊スキル削除(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If CheckBox1.Checked = True Then '削除ボタンを押しているのにチェックが有効
            CheckBox1.Checked = False
        End If
        Call 部隊スキル初期化()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        CheckBox2.Checked = False
        Me.Close()
    End Sub

    Private Sub 加速の有無(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged, CheckBox2.CheckedChanged
        If CheckBox1.Checked = True And CheckBox2.Checked = True Then '加速要素有ならば
            TextBox3.Enabled = True
        Else
            TextBox3.Text = ""
            TextBox3.Enabled = False
        End If
    End Sub

    Private Sub 加速率入力(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        If sender.text = "" Then
            bskill.speed = 0
        Else
            bskill.speed = Val(sender.text) * 0.01
        End If
    End Sub
End Class