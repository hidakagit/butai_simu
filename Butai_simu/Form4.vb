Public Class Form4
    Dim selectedi As Integer = -1 '現在編集中の部隊スキル位置。-1:新規追加か、選択無

    Private Sub Form4_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Form1.ToolStripComboBox1.Text = "攻撃" Then
            ComboBox1.SelectedIndex = 0
        Else
            ComboBox1.SelectedIndex = 1
        End If
        If bskill.flg = False Then CheckBox1.Checked = True
        If bskill.bsk Is Nothing Then Exit Sub
        Call 部隊スキル情報読み込み()
    End Sub

    Private Sub 部隊スキル情報読み込み()
        DataGridView1.Rows.Clear() '表をクリア
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        For i As Integer = 0 To bskill.bsk.Length - 1
            ' bskillを読み込んでいく
            With bskill.bsk(i)
                Dim row As DataGridViewRow = New DataGridViewRow()
                row.CreateCells(DataGridView1)
                Dim kb As String = ""
                If InStr(.koubou, "攻") Then
                    kb = "攻"
                Else
                    kb = "防"
                End If
                row.SetValues(New Object() {kb, .kouka_p, .kouka_f, .taisyo, .speed})
                DataGridView1.Rows.AddRange(row)
            End With
        Next
    End Sub

    Private Sub 登録部隊スキル削除(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) _
        Handles DataGridView1.CellClick
        '"Button"列ならば、ボタンがクリックされた
        If e.ColumnIndex < 0 Then Exit Sub
        If DataGridView1.Columns(e.ColumnIndex).Name = "Column6" Then
            'MessageBox.Show((e.RowIndex.ToString() + "行のボタンがクリックされました。"))
            selectedi = e.RowIndex
            'DataGridView1.CurrentCell.Value = "編集中"
            'DataGridView1.CurrentCell.Style.ForeColor = Color.Orange
        Else
            Exit Sub
        End If
        'データを埋めていく
        With DataGridView1.Rows(selectedi)
            If .Cells(0).Value = Nothing Then Exit Sub '登録されていなければ削除は出来ない
            If CStr(.Cells(0).Value) = "攻" Then
                ComboBox1.Text = "攻撃"
            Else
                ComboBox1.Text = "防御"
            End If
            TextBox1.Text = Val(.Cells(1).Value) * 100
            TextBox2.Text = Val(.Cells(2).Value) * 100
            ComboBox2.Text = .Cells(3).Value
            If Not Val(.Cells(4).Value) = 0 Then
                CheckBox2.Checked = True
                TextBox3.Text = Val(.Cells(4).Value) * 100
            Else
                CheckBox2.Checked = False
            End If
        End With
        DataGridView1.Rows.RemoveAt(selectedi)
    End Sub

    Private Sub 部隊スキル追加(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '入力漏れチェック
        If Not スキル入力チェック() Then Exit Sub
        Dim row As DataGridViewRow = New DataGridViewRow()
        row.CreateCells(DataGridView1)
        Dim kb As String = ""
        If ComboBox1.Text = "攻撃" Then
            kb = "攻"
        Else
            kb = "防"
        End If
        row.SetValues(New Object() {kb, Val(TextBox1.Text) * 0.01, Val(TextBox2.Text) * 0.01, ComboBox2.Text, Val(TextBox3.Text) * 0.01, "編集"})
        DataGridView1.Rows.Add(row)
    End Sub

    Private Function スキル入力チェック() As Boolean
        If Val(TextBox1.Text) = 0 Then Return False
        If ComboBox2.Text = "" Then Return False
        If CheckBox2.Checked Then '加速ON
            If Val(TextBox3.Text) = 0 Then
                MsgBox("加速率ゼロの速度スキル")
                Return False
            End If
        Else '加速OFFなのに威力がゼロのスキル
            If Val(TextBox2.Text) = 0 Then
                MsgBox("上昇率か加速率（加速スキルの場合）のどちらかを入力して下さい")
                Return False
            End If
        End If
        Return True
    End Function

    Private Sub 部隊スキル初期化()
        '残ったフォームをクリア
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        CheckBox2.Checked = False
        bskill = Nothing
    End Sub

    Private Sub 部隊スキル削除(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If CheckBox1.Checked = True Then '削除ボタンを押しているのにチェックが有効
            Dim result As DialogResult = MessageBox.Show("部隊スキルを削除しますか？", "削除確認", MessageBoxButtons.OKCancel)
            If result = DialogResult.Cancel Then Exit Sub
            CheckBox1.Checked = False
        End If
        Call 部隊スキル初期化()
        Call 部隊スキルONOFF(False)
        Me.Close()
    End Sub

    Private Sub 加速の有無(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then '加速要素有ならば
            TextBox3.Enabled = True
        Else
            TextBox3.Text = ""
            TextBox3.Enabled = False
        End If
    End Sub

    Public Sub 部隊スキルONOFF(ByVal onoff As Boolean)
        With Form1.ToolStripButton4
            If onoff = True Then
                bskill.flg = True '部隊スキルボタンの画像を変更
                .Image = Bitmap.FromFile(My.Application.Info.DirectoryPath & "\settings\ico\prettyxstickxstripe_p24_rd_nl_l.png")
                'Dim bkb As String
                'If koubou = "攻" Then
                '    bkb = "攻撃"
                'Else
                '    bkb = "防衛"
                'End If
                .ToolTipText = "部隊スキル : 有効" _
                & vbCrLf & "（" & bskill.bsk.Length & "個の部隊スキルが登録されています）"
                '& " [" & bkb & "時発動]" & vbCrLf & _
                '"発動率: " & kouka_p & "/ 上昇率: " & kouka_f
                If Not bskill.speed = 0 Then '加速有効
                    .ToolTipText = .ToolTipText & vbCrLf & "最大加速: +" & bskill.speed * 100 & "%"
                End If
            Else
                bskill.flg = False
                .Image = Bitmap.FromFile(My.Application.Info.DirectoryPath & "\settings\ico\prettyxstickxstripe_p24_bk_nl_l.png")
                .ToolTipText = "部隊スキル : 無効"
            End If
        End With
    End Sub

    Private Sub 部隊スキル確定(sender As Object, e As EventArgs) Handles Button3.Click
        If CheckBox1.Checked = False Then '確定ボタンを押しているのにチェックが無効
            CheckBox1.Checked = True
        End If
        Call 部隊スキル初期化()
        For i As Integer = 0 To DataGridView1.Rows.Count - 2
            Dim row As DataGridViewRow = DataGridView1.Rows(i)
            If i = 0 Then
                ReDim bskill.bsk(0)
            Else
                ReDim Preserve bskill.bsk(i)
            End If
            With bskill.bsk(i)
                .koubou = CStr(row.Cells(0).Value)
                .type = .koubou
                .kouka_p = Val(row.Cells(1).Value)
                .kouka_f = Val(row.Cells(2).Value)
                .taisyo = CStr(row.Cells(3).Value)
                .speed = Val(row.Cells(4).Value)
                If .speed > 0 And .type = "攻" Then '加速率があれば「速度」カテゴリー。「加速があって、防衛スキル」はまだ存在しないので未考慮
                    .type = "速"
                End If
            End With
        Next
        If DataGridView1.Rows.Count = 1 Then 'そもそも部隊スキルが一つもセットされていない
            MsgBox("部隊スキルが1つも設定されていません")
            Call 部隊スキル削除(sender, e)
            Exit Sub
        End If
        'Form1の部隊スキルのTooltips
        Call 部隊スキルONOFF(True)
        Me.Close()
    End Sub
End Class