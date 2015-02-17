Public Class Form3

    Public bd As Form5.bData
    Public selectbs As Integer = 0

    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Opacity = 0.8 '初期の透過度は80%
        'TrackBar1.Value = 8 '初期位置8
        'Form1.Hide()
        Me.TopMost = True
        ComboBox1.SelectedIndex = 0
    End Sub

    'Private Sub Form3_Closing(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.FormClosing
    '    Form1.Show()
    'End Sub
    'Private Sub 透過度変更(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Me.Opacity = 0.1 * Val(sender.value)
    'End Sub
    'Private Sub 常に手前に表示(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If CheckBox1.Checked = True Then
    '        Me.TopMost = True
    '    Else
    '        Me.TopMost = False
    '    End If
    'End Sub

    Private Sub データ読取(sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles RichTextBox1.DragDrop
        'エラータグをクリア
        RichTextBox1.Tag = Nothing

        Dim stmp() = Split(e.Data.GetData(GetType(String)), vbCrLf)
        Dim tmp() As String = Nothing
        Dim k As Integer = 0
        For i As Integer = 0 To stmp.Length - 1
            If Not ((stmp(i) = vbNullString) Or (stmp(i) = "ステータス強化") Or (stmp(i) = "指揮力強化")) Then
                ReDim Preserve tmp(k)
                If InStr(stmp(i), "LV") = 0 Then 'スキル名の空白を消すとマズイ場合がある
                    tmp(k) = Replace(stmp(i), " ", "")
                Else
                    tmp(k) = stmp(i)
                End If
                k = k + 1
            End If
        Next
        With bd
            Try
                .rare = Mid(tmp(0), tmp(0).Length, 1)
                If InStr(tmp(2), "限界突破") Then '限界突破時
                    .rank = 6
                    .level = 20
                Else
                    .rank = Val(Mid(tmp(2), InStr(tmp(2), "★") + 1, 1))
                    .level = Val(Mid(tmp(2), InStr(tmp(2), "｜") + 1))
                End If
                .hei_sum = Val(Mid(tmp(4), 5))
                .heisyu_name = Mid(tmp(5), 4)
                ReDim .st(2), .tou_a(3)
                .st(0) = Val(Mid(tmp(6), 3))
                .st(1) = Val(Mid(tmp(7), 3))
                .st(2) = Val(Mid(tmp(6), InStr(tmp(6), "兵") + 2))
                .tou_a(0) = Mid(tmp(8), 3, InStr(tmp(8), "馬") - 4)
                .tou_a(1) = Mid(tmp(9), 3, InStr(tmp(9), "器") - 4)
                .tou_a(2) = Mid(tmp(8), InStr(tmp(8), "馬") + 2)
                .tou_a(3) = Mid(tmp(9), InStr(tmp(9), "器") + 2)
                For j As Integer = 0 To 3
                    If .tou_a(j) = "S" Then
                        .tou_a(j) = ".S"
                    End If
                Next
                .skill_no = tmp.Length - 10
                ReDim .skill_name(.skill_no - 1), .skill_lv(.skill_no - 1)
                Dim syokisk As String = Nothing '初期スキル名
                For j As Integer = 0 To bd.skill_no - 1
                    Dim ttmp As String = Replace(tmp(10 + j), "技" & (j + 1) & vbTab, "") '"技1"みたいなのが付いてる場合、外す
                    If j = 0 Then
                        syokisk = Mid(ttmp, 1, InStr(ttmp, "LV") - 1)
                    End If
                    .skill_name(j) = Mid(ttmp, 1, InStr(ttmp, "LV") - 1)
                    .skill_lv(j) = Val(Mid(ttmp, InStr(ttmp, "LV") + 2))
                Next
                'Dim ntmp As String = GetINIValue(syokisk, Replace(Replace(tmp(0), "名", ""), "レア", "・"), bnpath)
                .name = Mid(tmp(0), 2, InStr(tmp(0), "レア") - 2)
                Dim repstr As String = "replace(replace(初期スキル名, " & """ "" , """"), " & """　"" , """") = "
                'Dim repstr As String = "replace(初期スキル名, " & """ "" , """") = "
                Dim s() As String = DB_DirectOUT("SELECT 武将名, 初期スキル名 FROM BData WHERE 武将R = " _
                                & ダブルクオート(.rare) & " AND 武将名 LIKE """ & .name & "%""" & " AND " & repstr & ダブルクオート(TrimJ(.skill_name(0))), _
                                {"武将名", "初期スキル名"})
                .name = s(0)
                .skill_name(0) = s(1)
                'If Not ntmp = "－" Then
                '    .name = ntmp
                'Else
                '    .name = Mid(tmp(0), 2, InStr(tmp(0), "レア") - 2)
                'End If
            Catch ex As Exception
                'Me.Focus()
                'MsgBox("データの読み込みが正常に行われなかった可能性があります")
                RichTextBox1.Tag = "E" 'エラータグ
                'Call テキスト確定(RichTextBox(Me, Val(cc) + 1), Nothing)
            End Try
        End With
        Call Tree登録()
    End Sub

    Private Sub Tree登録()
        TreeView1.Nodes.Clear()
        Label3.Text = ""
        Try
            With bd
                Dim treeNode_name As New TreeNode("武将名: " & .name)
                Dim treeNode_rare As New TreeNode("レア: " & .rare)
                Dim treeNode_rank As TreeNode
                If .rank = 6 Then '限界突破時
                    treeNode_rank = New TreeNode("☆限界突破☆")
                Else
                    treeNode_rank = New TreeNode("★" & .rank & "/LV" & .level)
                End If
                Dim treeNode_heisyu As New TreeNode("指揮兵種: " & .heisyu_name)
                Dim treeNode_heisum As New TreeNode("指揮兵数: " & .hei_sum)

                Dim treeNode_s_atk As New TreeNode("攻: " & .st(0))
                Dim treeNode_s_def As New TreeNode("防: " & .st(1))
                Dim treeNode_s_hei As New TreeNode("兵: " & .st(2))
                Dim treeNode_stsub() As TreeNode = {treeNode_s_atk, treeNode_s_def, treeNode_s_hei}
                Dim treeNode_st As New TreeNode("ステータス", treeNode_stsub)

                Dim treeNode_t_1 As New TreeNode("槍: " & .tou_a(0))
                Dim treeNode_t_2 As New TreeNode("弓: " & .tou_a(1))
                Dim treeNode_t_3 As New TreeNode("馬: " & .tou_a(2))
                Dim treeNode_t_4 As New TreeNode("器: " & .tou_a(3))
                Dim treeNode_tousub() As TreeNode = {treeNode_t_1, treeNode_t_2, treeNode_t_3, treeNode_t_4}
                Dim treeNode_tou As New TreeNode("統率", treeNode_tousub)

                Dim treeNode_ssub() As TreeNode
                Dim treeNode_s1 As New TreeNode("初期スキル: " & .skill_name(0) & .skill_lv(0))
                treeNode_ssub = {treeNode_s1}
                If .skill_no >= 2 Then
                    Dim treeNode_s2 As New TreeNode("スロ2: " & .skill_name(1) & .skill_lv(1))
                    treeNode_ssub = {treeNode_s1, treeNode_s2}
                    If .skill_no >= 3 Then
                        Dim treeNode_s3 As New TreeNode("スロ3: " & .skill_name(2) & .skill_lv(2))
                        treeNode_ssub = {treeNode_s1, treeNode_s2, treeNode_s3}
                    End If
                End If
                Dim treeNode_skill As New TreeNode("スキル", treeNode_ssub)

                'TreeViewに追加
                Dim TreeNode_sub() As TreeNode = _
                {treeNode_name, treeNode_rare, treeNode_rank, treeNode_heisyu, treeNode_heisum, treeNode_st, treeNode_tou, treeNode_skill}
                TreeView1.Nodes.AddRange(TreeNode_sub)
            End With
        Catch ex As Exception
            Label3.Text = "エラー検出"
        End Try
        If RichTextBox1.Tag = "E" Then
            Label3.Text = "エラー検出"
        End If
        TreeView1.TopNode.Expand()
    End Sub

    Private Sub データクリア(sender As Object, e As EventArgs) Handles Button2.Click
        RichTextBox1.Clear()
        RichTextBox1.Tag = Nothing
        TreeView1.Nodes.Clear()
        Label3.Text = ""
    End Sub

    Private Sub セット位置設定(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Select Case sender.text
            Case "部隊長"
                selectbs = 0
            Case "小隊長A"
                selectbs = 1
            Case "小隊長B"
                selectbs = 2
            Case "小隊長C"
                selectbs = 3
        End Select
    End Sub

    Private Sub 武将登録(sender As Object, e As EventArgs) Handles Button1.Click
        Dim tbno As Integer = Count_Busho()
        If selectbs > tbno + 1 Then '飛んでいる枠に武将を登録しようとしている
            MsgBox("武将を登録する枠は詰めてください")
            Exit Sub
        End If
        If busho_counter < selectbs Then '元の武将数 < 登録武将位置
            ReDim Preserve bs(selectbs)
        End If
        If Not ComboBox(Form1, CStr(selectbs) & "02").Text = "" Then
            Call Form1.武将データ消去(selectbs, False)
        End If
        Call Form1.武将データ手入力用(selectbs, False, True)
        With bs(selectbs)
            .rare = bd.rare
            .name = bd.name
            '.SelectedIndex = ComboBox(Me, CStr(i) & "01").FindString(.rare)
            ComboBox(Form1, CStr(selectbs) & "01").SelectedIndex = ComboBox(Form1, CStr(selectbs) & "01").FindString(bd.rare) '（強制的に）R選択
            'Form1.R選択(ComboBox(Form1, CStr(selectbs) & "01"), Nothing)
            ComboBox(Form1, CStr(selectbs) & "02").SelectedText = .name '（強制的に）武将名選択
            Form1.武将名選択(ComboBox(Form1, CStr(selectbs) & "02"), Nothing)
            .heisyu.name = bd.heisyu_name
            .hei_sum = bd.hei_sum
            .rank = bd.rank
            If .rank = 6 Then '限界突破時
                CheckBox(Form1, CStr(selectbs) & "1").Checked = True
            End If
            If Not .job = "剣" Then
                If .job = "覇" Then
                    .hei_max = .hei_max_d + .rank * 200 'ランクアップで兵数一律+200
                Else
                    .hei_max = .hei_max_d + .rank * 100 'ランクアップで兵数一律+100
                End If
            End If
            .level = bd.level
            .skill_no = bd.skill_no
            For j As Integer = 0 To 2
                .st(j) = bd.st(j)
            Next
            For j As Integer = 0 To 3
                .tou_a(j) = bd.tou_a(j)
            Next
            'この2つは面倒なので入力→更新する
            Dim tmpskill_no As Integer = bd.skill_no '変わっていく値なので一旦移し替え
            '----------
            For j As Integer = 0 To tmpskill_no - 1
                ReDim Preserve .skill(tmpskill_no - 1) '途中、追加スキル追加の部分で空白部分を削られてしまうので逐一Redimする必要がある
                .skill(j).name = bd.skill_name(j)
                .skill(j).lv = bd.skill_lv(j)
                If Not j = 0 Then '初期スキルはスルー
                    .skill(j).kanren = スキル関連推定(bd.skill_name(j))
                End If
                If Not .skill(j).kanren = "" And Not j = 0 Then '関連スキルが空ではない＝追加スキルがある
                    Select Case j
                        Case 1 'スロ2
                            ComboBox(Form1, CStr(selectbs) & "09").Focus()
                            ComboBox(Form1, CStr(selectbs) & "09").SelectedIndex = ComboBox(Form1, CStr(selectbs) & "09").FindString(.skill(j).kanren)
                            ComboBox(Form1, CStr(selectbs) & "11").SelectedText = .skill(j).name
                            Form1.スキル名入力(ComboBox(Form1, CStr(selectbs) & "11"), Nothing)
                            ComboBox(Form1, CStr(selectbs) & "15").Text = .skill(j).lv
                            Form1.追加スキル追加(ComboBox(Form1, CStr(selectbs) & "15"), Nothing)
                        Case 2 'スロ3
                            ComboBox(Form1, CStr(selectbs) & "10").Focus()
                            ComboBox(Form1, CStr(selectbs) & "10").SelectedIndex = ComboBox(Form1, CStr(selectbs) & "10").FindString(.skill(j).kanren)
                            ComboBox(Form1, CStr(selectbs) & "12").SelectedText = .skill(j).name
                            Form1.スキル名入力(ComboBox(Form1, CStr(selectbs) & "12"), Nothing)
                            ComboBox(Form1, CStr(selectbs) & "16").Text = .skill(j).lv
                            Form1.追加スキル追加(ComboBox(Form1, CStr(selectbs) & "16"), Nothing)
                    End Select
                ElseIf j = 0 Then '初期スキルの場合は
                    ComboBox(Form1, CStr(selectbs) & "14").Text = .skill(0).lv
                    Form1.追加スキル追加(ComboBox(Form1, CStr(selectbs) & "14"), Nothing)
                End If
            Next
            '----------
            '他は“表面上フォームを埋めていくだけ”
            Button(Form1, CStr(selectbs) & "001").Text = "/ " & .hei_max
            ComboBox(Form1, CStr(selectbs) & "03").Text = bd.heisyu_name
            ComboBox(Form1, CStr(selectbs) & "04").Text = bd.rank
            For j As Integer = 5 To 8
                ComboBox(Form1, CStr(selectbs) & "0" & CStr(j)).Text = bd.tou_a(j - 5)
            Next
            .Tousotu(True) = .tou_a '統率値を変換
            .兵科情報取得(bs(selectbs).heisyu.name)
            TextBox(Form1, CStr(selectbs) & "01").Text = bd.hei_sum
            TextBox(Form1, CStr(selectbs) & "02").Text = bd.level
            For j As Integer = 3 To 5 'ステ選択
                TextBox(Form1, CStr(selectbs) & "0" & CStr(j)).Text = bd.st(j - 3)
            Next
            .rankup_r = 統率未振り数推定(.tou_d_a, .tou_a, .rank)
            .huri = ステ振り推定(.st_d, .st, .sta_g, .rank, .level)
            ComboBox(Form1, CStr(selectbs) & "13").Text = .huri
            .残りランクアップ可能回数表示(GroupBox(Form1, 4 * (selectbs + 1)))
            '---------
            Form1.武将データ手入力用(selectbs, True, True)
            bs(selectbs).No = selectbs '武将No付け替え必要
        End With
    End Sub

End Class