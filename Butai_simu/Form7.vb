Public Class Form7
    Public kobo As String '攻防
    Public skill_lv As Integer 'スキルLV
    Public busho_cost As Decimal 'コスト
    Public busho_heika As String '兵科
    Public butai_heiho As Decimal '部隊兵法補正
    Public syo_skill As Boolean = False '将スキルフラグ
    Public hakai_onoff As Boolean = False '破壊期待値表示非表示フラグ

    Private Function 条件入力() As Boolean
        Dim flg As Boolean = True
        If ToolStripComboBox1.Text = "" Then
            MsgBox("攻防を設定して下さい")
            flg = False
        Else
            If ToolStripComboBox1.Text = "攻撃" Then
                kobo = "攻"
            Else
                kobo = "防"
            End If
        End If
        If ToolStripComboBox2.Text = "" Or ToolStripComboBox2.Text = "-----" Then
            MsgBox("兵科が未設定です")
            flg = False
        ElseIf ToolStripComboBox2.Text = "[将スキル]" Then '将スキルの場合
            busho_heika = "将"
            syo_skill = True
        Else
            busho_heika = ToolStripComboBox2.Text
        End If
        If ToolStripComboBox3.Text = "" Then
            MsgBox("武将コストが未設定です")
            flg = False
        Else
            busho_cost = Val(ToolStripComboBox3.Text)
        End If
        If ToolStripComboBox4.Text = "" Then
            MsgBox("スキルLVが未設定です")
            flg = False
        Else
            skill_lv = Val(ToolStripComboBox4.Text)
        End If
        If ToolStripTextBox1.Text = "" Then
            MsgBox("部隊兵法値が未設定です")
            flg = False
        Else
            butai_heiho = Val(ToolStripTextBox1.Text)
            If butai_heiho = 0 Then
                MsgBox("部隊兵法値の設定が不正です")
                flg = False
            End If
        End If
        Return flg
    End Function

    Private Sub 表更新(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim jf As Boolean = 条件入力()
        If jf = False Then '条件に不備があれば
            Exit Sub
        End If
        DataGridView1.Rows.Clear() '表をクリア
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        Dim sl()() As String = Nothing
        Dim hp(), kp(), dp(), kitai(), dkitai() As Decimal
        Dim costd() As Boolean
        Dim slid() As Decimal
        Dim de() As Integer
        Dim tk(), erck() As String
        Dim slbl() As String = {"名前", "対象", "確率", "基本効果", "付加効果"}
        Dim target_u() As String = {"武士", "弓騎馬", "赤備え", "騎馬鉄砲", "鉄砲足軽", "焙烙火矢", "大筒兵", "破城鎚", "攻城櫓"}
        Dim target_h() As String = {"武士", "弓騎馬", "赤備え", "騎馬鉄砲", "鉄砲足軽", "焙烙火矢", "大筒兵", "雑賀衆"}
        Dim target_hi() As String = {"国人衆", "雑賀衆", "海賊衆", "母衣衆"}
        Dim jkf_u As Boolean = False '上級器フラグ
        Dim jkf_h As Boolean = False '上級砲フラグ
        Dim hik_f As Boolean = False '秘境兵フラグ

        Dim h() As String
        If syo_skill = False Then '将スキルじゃない場合
            h = DB_DirectOUT(con, cmd, "SELECT 兵種名,兵科 FROM Heika WHERE 兵種名=""" & busho_heika & """", {"兵科"})
            For i As Integer = 0 To target_u.Length - 1 '上級器対応かどうか
                If InStr(busho_heika, target_u(i)) Then
                    jkf_u = True
                End If
            Next
            For i As Integer = 0 To target_h.Length - 1 '上級砲対応かどうか
                If InStr(busho_heika, target_h(i)) Then
                    jkf_h = True
                End If
            Next
            For i As Integer = 0 To target_hi.Length - 1 '秘境兵対応かどうか
                If InStr(busho_heika, target_hi(i)) Then
                    hik_f = True
                End If
            Next
        Else '将スキルの場合
            h = {"将"}
        End If
        For i As Integer = 0 To slbl.Length - 1 '表示適応データを取得
            ReDim Preserve sl(i)
            '特殊スキルではなく、かつ指定LVで、かつ対象に兵科文字列もしくは全を含む、かつ攻防一致
            Dim sqlstr As String
            Dim kobos As String
            If kobo = "攻" Then
                kobos = "( 基本効果 LIKE """ & "%攻%" & """ OR 基本効果" & " LIKE """ & "%破%" & """ )"
            Else
                kobos = "基本効果 LIKE """ & "%防%" & """"
            End If
            If syo_skill Then '将スキルの場合は全スキルを除く
                'sqlstr = "SELECT * FROM Skill WHERE 基本効果 LIKE """ & "%" & kobo & "%" & """ AND NOT 分類 =""" & "特殊""" & " AND LV =" & skill_lv & " AND ( 対象 LIKE """ & "%" & h(0) & "%" & """ )"
                sqlstr = "SELECT * FROM Skill WHERE " & kobos & " AND NOT 分類 =""" & "特殊""" & " AND LV =" & skill_lv & " AND ( 対象 LIKE """ & "%" & h(0) & "%" & ""
            Else
                sqlstr = "SELECT * FROM Skill WHERE " & kobos & " AND NOT 分類 =""" & "特殊""" & " AND LV =" & skill_lv & " AND ( 対象 LIKE """ & "%" & h(0) & "%" & """ OR 対象 =" & """全" & ""
            End If
            If jkf_u Then
                sqlstr = sqlstr & """ OR 対象 =" & """上級器*" & ""
            End If
            If jkf_u Then
                sqlstr = sqlstr & """ OR 対象 =" & """上級砲*" & ""
            End If
            If hik_f Then
                sqlstr = sqlstr & """ OR 対象 =" & """秘境兵*" & ""
            End If
            sqlstr = sqlstr & """ )"
            sl(i) = DB_DirectOUT2(con, cmd, sqlstr, slbl(i))
        Next
        Dim sc As Integer = sl(0).Length - 1 '合致したスキルの個数
        ReDim hp(sc), kp(sc), dp(sc), kitai(sc), dkitai(sc), tk(sc), de(sc), erck(sc), costd(sc)

        For i = 2 To 4 '取得してきたデータを加工
            For j As Integer = 0 To sc
                If i = 2 Then '発動率
                    If sl(i)(j) = "" Or sl(i)(j) = "+%" Then 'データなし
                        erck(j) = "D無"
                        hp(j) = 0
                    ElseIf Val(sl(i)(j)) <> 1 Then
                        hp(j) = Val(sl(i)(j)) + butai_heiho / 100
                    Else
                        hp(j) = 1
                    End If
                End If
                If i = 3 Then '上昇率
                    If kobo = "攻" And InStr(sl(i)(j), "破") Then
                        de(j) = 1 '破壊のみのスキル
                        sl(4)(j) = sl(3)(j) '付加効果へコピー
                    End If
                    Dim rk As String = sl(1)(j) & kobo & "："
                    sl(i)(j) = Replace(sl(i)(j), rk, "")
                    If jkf_u = True Then
                        rk = "上級器" & kobo & "："
                        sl(i)(j) = Replace(sl(i)(j), rk, "")
                    End If
                    If jkf_h = True Then
                        rk = "上級砲" & kobo & "："
                        sl(i)(j) = Replace(sl(i)(j), rk, "")
                    End If
                    If hik_f = True Then
                        rk = "秘境兵" & kobo & "："
                        sl(i)(j) = Replace(sl(i)(j), rk, "")
                    End If

                    If erck(j) <> "" Or sl(i)(j) = "%上昇" Then 'データなし
                        kp(j) = 0
                        erck(j) = "D無"
                    Else
                        sl(i)(j) = Replace(sl(i)(j), "上昇", "")
                        If InStr(sl(i)(j), "コスト") Then
                            sl(i)(j) = Replace(sl(i)(j), "コスト", busho_cost)
                            costd(j) = True
                        Else
                            costd(j) = False
                        End If
                        kp(j) = Decimal.Parse(文字列計算(sl(i)(j), False)) 'ここでのエラーはwikiがおかしい場合が多い、うるさいから切る
                        If kp(j) = 0 And de(j) <> 1 Then '文字列計算をした結果、ゼロ
                            erck(j) = "D異常"
                        End If
                        '期待値計算
                        kitai(j) = hp(j) * kp(j)
                    End If
                End If
                If i = 4 Then '付加効果
                    If InStr(sl(i)(j), "破壊") Then
                        de(j) = 1

                        Dim rk As String = sl(1)(j) & "破壊" & "："
                        sl(i)(j) = Replace(sl(i)(j), rk, "")
                        If jkf_u = True Then
                            rk = "上級器" & "破壊" & "："
                            sl(i)(j) = Replace(sl(i)(j), rk, "")
                        End If
                        If jkf_h = True Then
                            rk = "上級砲" & "破壊" & "："
                            sl(i)(j) = Replace(sl(i)(j), rk, "")
                        End If
                        If hik_f = True Then
                            rk = "秘境兵" & "破壊" & "："
                            sl(i)(j) = Replace(sl(i)(j), rk, "")
                        End If

                        If erck(j) <> "" Or sl(i)(j) = "%上昇" Then 'データなし
                            dp(j) = 0
                            erck(j) = "D無"
                        Else
                            sl(i)(j) = Replace(sl(i)(j), "上昇", "")
                        End If
                        If InStr(sl(i)(j), "コスト") Then
                            sl(i)(j) = Replace(sl(i)(j), "コスト", busho_cost)
                        End If
                        dp(j) = Decimal.Parse(文字列計算(sl(i)(j), False)) '同上
                        If dp(j) = 0 Then '文字列計算をした結果、ゼロ
                            erck(j) = "D異常"
                        End If

                        If InStr(sl(i)(j), "速") Then
                            tk(j) = "速度＋破壊"
                        Else
                            tk(j) = "破壊"
                        End If
                    ElseIf InStr(sl(i)(j), "速") Then
                        tk(j) = "速度"
                    End If
                End If
                If de(j) = 1 Then '破壊が絡むスキルは破壊期待値を計算
                    dkitai(j) = hp(j) * dp(j)
                End If
            Next
        Next
        ReDim slid(sc) 'インデックス配列
        For i As Integer = 0 To sc
            slid(i) = i
        Next
        Array.Sort(kitai, slid) 'ソート
        'Datagridに追加
        For i As Integer = sc To 0 Step -1
            Dim row As DataGridViewRow = New DataGridViewRow()
            Dim cp As Integer = CInt(slid(i))
            row.CreateCells(DataGridView1)
            row.SetValues(New Object() {sl(0)(cp), sl(1)(cp), kitai(i), hp(cp), kp(cp), tk(cp), dp(cp), dkitai(cp), erck(cp)})
            DataGridView1.Rows.AddRange(row)

            Dim cell As DataGridViewCell = DataGridView1.Rows(DataGridView1.Rows.Count - 2).Cells(4) 'コスト依存ならば色づけ
            If costd(cp) Then 'コスト依存
                cell.Style.ForeColor = Color.DeepPink
            End If
            cell = DataGridView1.Rows(DataGridView1.Rows.Count - 2).Cells(5) '破壊等の色付けのため
            If cell.Value = "破壊" Then
                cell.Style.ForeColor = Color.DarkOliveGreen
            ElseIf cell.Value = "速度" Then
                cell.Style.ForeColor = Color.DodgerBlue
            ElseIf cell.Value = "速度＋破壊" Then
                cell.Style.ForeColor = Color.Firebrick
            End If
        Next
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        DataGridView1.Columns(6).Visible = hakai_onoff '破壊期待値表示切替
        DataGridView1.Columns(7).Visible = hakai_onoff
        If hakai_onoff Then '画面サイズ調整
            Me.Width = 670
        Else
            Me.Width = 540
        End If
    End Sub

    Private Sub 破壊期待値表示切替(sender As Object, e As EventArgs) Handles ToolStripComboBox5.TextChanged
        If InStr(ToolStripComboBox5.Text, "非表示") Then
            hakai_onoff = False
        Else
            hakai_onoff = True
        End If
    End Sub
End Class