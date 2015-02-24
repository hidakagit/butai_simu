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
            If butai_heiho < 0 Then
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

        Call フラグ付きスキル読み込み() '読込（更新）

        DataGridView1.Rows.Clear() '表をクリア
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        Dim sl()() As String = Nothing
        Dim hp(), kp(), dp(), sp(), kitai(), dkitai() As Decimal
        Dim costd() As Boolean
        Dim slid() As Decimal
        Dim tk(), erck() As String
        Dim slbl() As String = {"スキル名", "分類", "攻防", "対象", "発動率", "上昇率", "付与効果", "付与率", "Sunf"}
        Dim target_u() As String = {"武士", "弓騎馬", "赤備え", "騎馬鉄砲", "鉄砲足軽", "焙烙火矢", "大筒兵", "破城鎚", "攻城櫓"}
        Dim target_h() As String = {"武士", "弓騎馬", "赤備え", "騎馬鉄砲", "鉄砲足軽", "焙烙火矢", "大筒兵", "雑賀衆"}
        Dim target_hi() As String = {"国人衆", "雑賀衆", "海賊衆", "母衣衆"}
        Dim jkf_u As Boolean = False '上級器フラグ
        Dim jkf_h As Boolean = False '上級砲フラグ
        Dim hik_f As Boolean = False '秘境兵フラグ

        Dim h() As String
        If syo_skill = False Then '将スキルじゃない場合
            h = DB_DirectOUT("SELECT 兵種名, 兵科 FROM HData WHERE 兵種名 = " & ダブルクオート(busho_heika) & "", {"兵科"})
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
                kobos = "( 攻防 = " & ダブルクオート("攻") & " OR 攻防 = " & ダブルクオート("破壊") & " )"
            Else
                kobos = " 攻防 = " & ダブルクオート("防")
            End If
            If syo_skill Then '将スキルの場合は全スキルを除く
                'sqlstr = "SELECT * FROM Skill WHERE 基本効果 LIKE """ & "%" & kobo & "%" & """ AND NOT 分類 =""" & "特殊""" & " AND LV =" & skill_lv & " AND ( 対象 LIKE """ & "%" & h(0) & "%" & """ )"
                sqlstr = "SELECT * FROM SData INNER JOIN SName ON SData.スキル名 = SName.スキル名 WHERE " & kobos _
                    & " AND NOT 分類 = " & ダブルクオート("特殊") & " AND NOT 分類 = " & ダブルクオート("不可") & " AND スキルLV =" & skill_lv & " AND ( 対象 LIKE " & ダブルクオート("%" & h(0) & "%")
            Else
                sqlstr = "SELECT * FROM SData INNER JOIN SName ON SData.スキル名 = SName.スキル名 WHERE " & kobos _
                    & " AND NOT 分類 = " & ダブルクオート("特殊") & " AND NOT 分類 = " & ダブルクオート("不可") & " AND スキルLV =" & skill_lv & " AND ( 対象 LIKE " & ダブルクオート("%" & h(0) & "%") & " OR 対象 = " & ダブルクオート("全")
            End If
            If jkf_u Then
                sqlstr = sqlstr & " OR 対象 = " & ダブルクオート("上級器")
            End If
            If jkf_u Then
                sqlstr = sqlstr & " OR 対象 = " & ダブルクオート("上級砲")
            End If
            If hik_f Then
                sqlstr = sqlstr & " OR 対象 = " & ダブルクオート("秘境兵")
            End If
            sqlstr = sqlstr & " )"
            sl(i) = DB_DirectOUT2(sqlstr, slbl(i))
        Next
        Dim sc As Integer = sl(0).Length - 1 '合致したスキルの個数
        ReDim hp(sc), kp(sc), dp(sc), kitai(sc), dkitai(sc), tk(sc), erck(sc), costd(sc), sp(sc)

        For j As Integer = 0 To sc '取得してきたデータを加工
            Dim dflg As Boolean = False
            Dim sflg As Boolean = False
            costd(j) = False
            If sl(8)(j) = "U" Then 'データなし
                erck(j) = "D無"
                hp(j) = 0
                kp(j) = 0
                dp(j) = 0
                Continue For
            End If
            hp(j) = Val(sl(4)(j)) '発動率
            '特殊スキル
            If sl(1)(j) = "条件" Then
                Dim tmpskl As Busho.skl = Nothing
                With tmpskl
                    .name = sl(0)(j)
                    .lv = skill_lv
                    .kouka_f = Val(sl(5)(j))
                    .t_flg = フラグ付きスキル参照(tmpskl) '条件付きスキルの場合
                    If InStr(.heika, "槍弓馬砲器") Then
                        .heika = "全"
                    End If
                    If .t_flg Then
                        sl(3)(j) = .heika
                        sl(4)(j) = .kouka_p
                        sl(5)(j) = .kouka_f
                    End If
                End With
            End If
            Select Case (sl(2)(j))
                Case "速" '速度オンリー
                    sflg = True
                    sp(j) = Val(sl(5)(j))
                    sl(7)(j) = sl(5)(j) '速度のみのスキル。付加効果にコピー
                    sl(5)(j) = 0
                Case "破壊" '破壊オンリー
                    dflg = True
                    sl(7)(j) = sl(5)(j) '破壊のみのスキル。付加効果にコピー
                    sl(5)(j) = 0
                Case Else '通常スキルの場合
                    If sl(6)(j) = "速" Then '速度を含むスキル
                        sflg = True
                        sp(j) = Val(sl(7)(j))
                    End If
                    If sl(6)(j) = "破壊" Then '破壊を含むスキル
                        dflg = True
                    End If
                    If InStr(sl(5)(j), "C") Then 'コスト依存スキル
                        sl(5)(j) = Replace(sl(5)(j), "C", busho_cost)
                        costd(j) = True
                        kp(j) = Decimal.Parse(文字列計算(sl(5)(j), False)) 'ここでのエラーはwikiがおかしい場合が多い、うるさいから切る
                        If kp(j) = 0 And dflg = False Then '文字列計算をした結果、ゼロ
                            erck(j) = "D異常"
                        End If
                    Else
                        kp(j) = Decimal.Parse(sl(5)(j))
                    End If
            End Select
            If Val(sl(4)(j)) + (butai_heiho / 100) < 1 Then '発動率
                hp(j) = Val(sl(4)(j)) + butai_heiho / 100
            Else
                hp(j) = 1
            End If
            '期待値計算
            kitai(j) = hp(j) * kp(j)
            If dflg Then '付与効果に破壊を含むスキル
                If InStr(sl(7)(j), "C") Then 'コスト依存スキル
                    sl(7)(j) = Replace(sl(7)(j), "C", busho_cost)
                    costd(j) = True
                    'costd(j) = True
                    dp(j) = Decimal.Parse(文字列計算(sl(7)(j), False)) 'ここでのエラーはwikiがおかしい場合が多い、うるさいから切る
                    If dp(j) = 0 Then '文字列計算をした結果、ゼロ
                        erck(j) = "D異常"
                    End If
                Else
                    dp(j) = Decimal.Parse(文字列計算(sl(7)(j)))
                End If
                dkitai(j) = hp(j) * dp(j) '破壊が絡むスキルは破壊期待値を計算
            End If
            If sflg Then tk(j) = "速度"
            If dflg Then tk(j) = "破壊"
            If sflg And dflg Then tk(j) = "速度＋破壊"
        Next
        ReDim slid(sc) 'インデックス配列
        For i As Integer = 0 To sc
            slid(i) = i
        Next
        Array.Sort(kitai, slid) 'ソート
        Dim rows As DataGridViewRow()
        Dim rno As Integer = 0
        ReDim rows(sc)
        'Datagridに追加
        For i As Integer = sc To 0 Step -1
            Dim row As DataGridViewRow = New DataGridViewRow()
            Dim cp As Integer = CInt(slid(i))
            row.CreateCells(DataGridView1)
            row.SetValues(New Object() {sl(0)(cp), sl(3)(cp), kitai(i), hp(cp), kp(cp), tk(cp), dp(cp), dkitai(cp), erck(cp)})

            Dim cell As DataGridViewCell = row.Cells(0) '条件付きスキルならば色づけ
            If sl(1)(cp) = "条件" Then cell.Style.BackColor = Color.LightGoldenrodYellow
            cell = row.Cells(4) 'コスト依存ならば色づけ
            If costd(cp) Then 'コスト依存
                cell.Style.ForeColor = Color.DeepPink
            End If
            cell = row.Cells(5) '破壊等の色付けのため
            If cell.Value = "破壊" Then
                cell.Style.ForeColor = Color.DarkOliveGreen
            ElseIf cell.Value = "速度" Then
                If sp(cp) > 0 Then
                    cell.Style.ForeColor = Color.DodgerBlue
                Else
                    cell.Style.ForeColor = Color.IndianRed
                End If
            ElseIf cell.Value = "速度＋破壊" Then
                If sp(cp) > 0 Then
                    cell.Style.ForeColor = Color.Firebrick
                Else
                    cell.Style.ForeColor = Color.Maroon
                End If
            End If
            rows(rno) = row
            rno = rno + 1
        Next
        DataGridView1.Rows.AddRange(rows)
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