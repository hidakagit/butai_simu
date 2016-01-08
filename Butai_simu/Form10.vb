Imports System.IO

Public Class Form10

    Public kobo As String '攻防
    Public busho_heika As String '兵科
    Public busho_rank As Integer '武将ランク
    Public simu_bs(3) As Busho 'シミュレーションのためのBusho
    Public simu_lv As Integer = 10
    Public simu_skeqflg As Boolean = True
    Public heicount As Integer = 4 '兵法振りの人数の数
    Public heino() As Integer '兵法振りの武将No
    Public heihousum As Decimal '部隊兵法値
    Public butaicostsum As Decimal '部隊コスト合計
    Public butairanksum As Decimal '部隊ランクボーナス（☆合計）
    Public butairankbonus As Decimal '部隊ランクボーナス値
    Public stpt As Integer 'ステ振りポイント数
    Public simu_ss(3)() As Integer 'Skill Status（各武将のスキル状態）
    Public kitai_val() As Decimal '各武将の期待値
    Public kitai_butai() As Decimal '部隊期待値
    Public kitai_max() As Decimal '部隊MAX
    Public taisyo_rare(4) As Boolean '対象にするランキング武将(0:対象外 1:対象)
    Public add_skl() As Busho.skl '追加スキルはキャッシュする
    Public zero_skl()() As String '初期スキル情報
    Public syosklflg As Boolean = False '追加スキル詳細設定フラグ
    Public statusflg As Boolean = False 'ステ振り詳細設定フラウ
    Public cus_addskl(3)() As String '追加スキル詳細設定(第二要素:0,2->スキル名, 1,3->スキル関連)
    Public cus_status(3) As Integer 'ステ振り詳細設定(0:攻, 1:防, 2:兵, 3:適正お任せ)
    Public cus_status_kb As Decimal 'ステ振り詳細設定にて、攻防極閾値
    Public cus_status_hei As Decimal 'ステ振り詳細設定にて、兵法極閾値

    Private Sub Form10_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '武将設定初期化
        For i As Integer = 0 To 3
            simu_bs(i).武将設定初期化()
            If Not i = 0 Then '追加スキル詳細設定部分の初期スキルをクリア
                Label(Form11, "00" & Val(i)).Text = "------------"
            End If
        Next
        ComboBox9.SelectedIndex = 9
        ComboBox10.SelectedIndex = 0
    End Sub

    Public Sub R選択(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles ComboBox1.SelectedIndexChanged, ComboBox2.SelectedIndexChanged, ComboBox3.SelectedIndexChanged
        Dim cc As ComboBox = Nothing
        Select Case (String_onlyNumber(sender.name))
            Case "1"
                cc = ComboBox001
            Case "2"
                cc = ComboBox002
            Case "3"
                cc = ComboBox003
        End Select
        RemoveHandler cc.SelectedValueChanged, AddressOf Me.武将名選択 'これが無いと武将名を選べなくなる

        Dim p As DataSet
        p = DB_TableOUT("SELECT id, 武将R, 武将名 FROM BData WHERE 武将R = " & ダブルクオート(sender.SelectedItem) & " ORDER BY Bid ASC", "BData")
        With cc
            .BeginUpdate()
            .DisplayMember = "武将名"
            .ValueMember = "id"
            .DataSource = p.Tables("BData")
            .SelectedIndex = -1
            .EndUpdate()
        End With
        AddHandler cc.SelectedValueChanged, AddressOf Me.武将名選択
    End Sub
    Public Sub 武将名選択(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles ComboBox001.SelectedValueChanged, ComboBox002.SelectedValueChanged, ComboBox003.SelectedValueChanged
        Dim selbsho As Integer = 0
        Select Case (String_onlyNumber(sender.name))
            Case "001"
                selbsho = 0
            Case "002"
                selbsho = 1
            Case "003"
                selbsho = 2
        End Select
        simu_bs(selbsho) = Nothing '選択した武将をクリア
        simu_bs(selbsho).武将設定初期化()
        Label(Form11, "00" & Val(selbsho + 1)).Text = "------------"
        Dim r() As String = _
         {"Bid", "武将R", "Cost", "指揮兵数", "槍統率", "弓統率", "馬統率", "器統率", "初期攻撃", "初期防御", "初期兵法", "攻成長", "防成長", "兵成長", "初期スキル名", "職"}
        Dim s() As String = _
        DB_DirectOUT("SELECT * FROM BData WHERE 武将R = " & ダブルクオート(ComboBox(Me, CStr(selbsho + 1)).Text) & _
                     " AND 武将名 = " & ダブルクオート(sender.Text) & " AND Bunf = 'F'", r)
        'ここから武将初期化
        With simu_bs(selbsho)
            .No = selbsho
            .name = sender.Text
            .id = s(0)
            .rare = s(1)
            .cost = s(2)
            .hei_max_d = s(3)
            .hei_max = .hei_max_d '初期値設定
            .hei_sum = .hei_max
            .Tousotu(False, True) = {s(4), s(5), s(6), s(7)}
            .Sta(False) = {s(8), s(9), s(10)}
            ReDim .sta_g(2)
            .sta_g = {s(11), s(12), s(13)}
            .skill(0).name = Replace(s(14), Mid(s(14), 1, InStr(s(14), "：")), "")
            .skill(0).bno = selbsho
            .job = s(15)
            .スキル取得(0, selbsho, .skill(0).name, .skill(0).lv, {0})
            Label(Form11, "00" & Val(selbsho + 1)).Text = .skill(0).name
        End With
    End Sub

    Private Sub 追加スキル表示(sender As Object, e As EventArgs) _
        Handles ComboBox01.SelectedIndexChanged, ComboBox02.SelectedIndexChanged, _
                ComboBox41.SelectedIndexChanged, ComboBox42.SelectedIndexChanged
        Dim p As DataSet
        Dim cc As ComboBox = Nothing
        Select Case (String_onlyNumber(sender.name))
            Case "01"
                cc = ComboBox011
            Case "02"
                cc = ComboBox012
            Case "41"
                cc = ComboBox041
            Case "42"
                cc = ComboBox042
        End Select
        Dim sqlwhere As String = ダブルクオート(sender.text)
        If sqlwhere = ダブルクオート("特殊") Then '特殊項目には、条件付きスキルも含む
            sqlwhere = sqlwhere & " OR 分類 = " & ダブルクオート("条件")
        End If
        p = DB_TableOUT("SELECT id, 分類, スキル名 FROM SName WHERE 分類 = " & sqlwhere & " ORDER BY id", "SName")
        With cc
            .BeginUpdate()
            .ValueMember = "id"
            .DisplayMember = "スキル名"
            .DataSource = p.Tables("SName")
            .SelectedIndex = -1
            .EndUpdate()
        End With
    End Sub

    Private Sub スキル設定EQUAL(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then 'ON
            simu_skeqflg = True
            ComboBox41.Enabled = False
            ComboBox42.Enabled = False
            ComboBox041.Enabled = False
            ComboBox042.Enabled = False
        Else
            simu_skeqflg = False
            ComboBox41.Enabled = True
            ComboBox42.Enabled = True
            ComboBox041.Enabled = True
            ComboBox042.Enabled = True
        End If
    End Sub
    Private Sub 攻防変更(sender As Object, e As EventArgs) Handles ToolStripComboBox2.SelectedIndexChanged
        If sender.Text = "攻撃" Then
            kobo = "攻撃"
        Else
            kobo = "防御"
        End If
    End Sub
    Private Sub 兵科変更(sender As Object, e As EventArgs) Handles ToolStripComboBox2.SelectedIndexChanged
        busho_heika = ToolStripComboBox2.Text
    End Sub
    Private Sub ランク変更(sender As Object, e As EventArgs) Handles ToolStripComboBox3.SelectedIndexChanged
        Select Case sender.Text
            Case "☆☆☆☆☆"
                busho_rank = 0
            Case "★☆☆☆☆"
                busho_rank = 1
            Case "★★☆☆☆"
                busho_rank = 2
            Case "★★★☆☆"
                busho_rank = 3
            Case "★★★★☆"
                busho_rank = 4
            Case "★★★★★"
                busho_rank = 5
            Case "限界突破"
                busho_rank = 6
        End Select
    End Sub
    Private Sub スキルレベル変更(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged
        simu_lv = Val(sender.text)
    End Sub
    Private Sub ステ振り設定(sender As Object, e As EventArgs) Handles ComboBox10.SelectedIndexChanged
        heicount = Val(Mid(sender.Text, 4, 1))
        If sender.text = "お好み設定" Then
            statusflg = True
            Call Form15.ステ振り詳細設定関連ONOFF(False)
        Else
            statusflg = False
            Call Form15.ステ振り詳細設定関連ONOFF(True)
        End If
    End Sub
    Private Sub プログレスバー初期化()
        ToolStripProgressBar1.Value = 0
    End Sub
    Private Sub プログレスバー変化(ByVal val As Decimal)
        ToolStripProgressBar1.Value = val
    End Sub

    Private Function 条件入力() As Boolean
        Dim flg As Boolean = True
        '攻防
        If ToolStripComboBox1.Text = "" Then
            MsgBox("攻防を設定して下さい")
            flg = False
        Else
            If ToolStripComboBox1.Text = "攻撃" Then
                kobo = "攻撃"
            Else
                kobo = "防御"
            End If
        End If
        '兵科
        If ToolStripComboBox2.Text = "" Or ToolStripComboBox2.Text = "-----" Then
            MsgBox("兵科が未設定です")
            flg = False
        Else
            busho_heika = ToolStripComboBox2.Text
        End If
        '武将ランク
        If ToolStripComboBox3.Text = "" Then
            MsgBox("武将ランクが未設定です")
            flg = False
        End if
        '武将チェック（ちゃんと3将登録されていないとダメ）
        For i As Integer = 1 To 3
            If ComboBox(Me, "00" & CStr(i)).Text = "" Then
                flg = False
            End If
        Next
        '対象武将範囲
        For i As Integer = 0 To 4
            If taisyo_rare(i) = True Then
                Exit For
            End If
            If i = 4 Then 'ここまで来てもループを抜けていない→一個も対象レアが無い
                flg = False
            End If
        Next
        'ステ振りポイントの合計
        stpt = 0
        If busho_rank < 2 Then
            stpt = (busho_rank * 20 + 20) * 4
        ElseIf busho_rank < 4 Then
            stpt = 160 + ((busho_rank - 2) * 20 + 20) * 5
        ElseIf busho_rank < 6 Then
            stpt = 360 + ((busho_rank - 4) * 20 + 20) * 6
        ElseIf busho_rank = 6 Then
            stpt = 630
        End If
        Return flg
    End Function

    Private Sub 対象範囲ALL(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then 'ON
            CheckBox001.Checked = True
            CheckBox002.Checked = True
            CheckBox003.Checked = True
            CheckBox004.Checked = True
            CheckBox005.Checked = True
            For i As Integer = 0 To 4
                taisyo_rare(i) = True
            Next
        Else
            CheckBox001.Checked = False
            CheckBox002.Checked = False
            CheckBox003.Checked = False
            CheckBox004.Checked = False
            CheckBox005.Checked = False
            For i As Integer = 0 To 4
                taisyo_rare(i) = False
            Next
        End If
    End Sub
    Private Sub 対象範囲レア設定(sender As Object, e As EventArgs) _
        Handles CheckBox001.CheckedChanged, CheckBox002.CheckedChanged, CheckBox003.CheckedChanged, _
                CheckBox004.CheckedChanged, CheckBox005.CheckedChanged
        If sender.checked = True Then 'ON
            taisyo_rare(Val(Mid((sender.name), 11, 1)) - 1) = True
        Else
            taisyo_rare(Val(Mid((sender.name), 11, 1)) - 1) = False
        End If
    End Sub
    Private Function 対象範囲文字列出力() As String
        Dim target_rare As String = ""
        For i As Integer = 0 To 4
            If taisyo_rare(i) Then
                Select Case i
                    Case 0
                        target_rare = target_rare & "武将R = " & """天""" & " OR "
                    Case 1
                        target_rare = target_rare & "武将R = " & """極""" & " OR "
                    Case 2
                        target_rare = target_rare & "武将R = " & """特""" & " OR "
                    Case 3
                        target_rare = target_rare & "武将R = " & """上""" & " OR "
                    Case 4
                        target_rare = target_rare & "武将R = " & """序""" & " OR "
                End Select
            End If
        Next
        If Not target_rare = "" Then
            target_rare = target_rare.Remove(target_rare.Length - 4, 4)
        End If
        Return target_rare
    End Function

    'ごちゃごちゃしたので関数に押し込んだ・・・
    Private Sub 追加スキルセット_def()
        'simu_ss = Nothing '初期化
        '3将スキル
        If Not ComboBox011.Text = "" Then '2スロ目埋まっている
            If Not ComboBox012.Text = "" Then '3スロ目埋まっている
                simu_ss(0) = {0, 1, 2}
            Else
                simu_ss(0) = {0, 1}
            End If
        Else
            If Not ComboBox012.Text = "" Then
                simu_ss(0) = {0, 2}
            Else
                simu_ss(0) = {0}
            End If
        End If

        'ランキング武将スキル
        If simu_skeqflg Then '上と同じならば
            simu_ss(3) = simu_ss(0)
        Else
            If Not ComboBox041.Text = "" Then '2スロ目埋まっている
                If Not ComboBox042.Text = "" Then '3スロ目埋まっている
                    simu_ss(3) = {0, 1, 2}
                Else
                    simu_ss(3) = {0, 1}
                End If
            Else
                If Not ComboBox042.Text = "" Then
                    simu_ss(3) = {0, 2}
                Else
                    simu_ss(3) = {0}
                End If
            End If
        End If

        'スキルセット
        For i As Integer = 0 To 3
            If i = 3 Then 'ランキング武将
                simu_bs(i).skill_no = simu_ss(3).Length
            Else
                simu_bs(i).skill_no = simu_ss(0).Length
            End If
        Next
        '3武将スキルセット
        For i As Integer = 0 To 2
            For j As Integer = 0 To simu_ss(0).Length - 1
                Select Case (simu_ss(0)(j))
                    Case 0
                        simu_bs(i).スキル取得(j, i, simu_bs(i).skill(0).name, simu_lv, simu_ss(0))
                    Case 1
                        simu_bs(i).スキル取得(j, i, ComboBox011.Text, simu_lv, simu_ss(0))
                    Case 2
                        simu_bs(i).スキル取得(j, i, ComboBox012.Text, simu_lv, simu_ss(0))
                End Select
            Next
        Next
    End Sub

    Private Sub 追加スキルセット_Cus()
        '各武将毎にスキルを詳細設定モード
        For i As Integer = 0 To 3
            If Not cus_addskl(i)(0) = "" Then '2スロ目埋まっている
                If Not cus_addskl(i)(2) = "" Then '3スロ目埋まっている
                    simu_ss(i) = {0, 1, 2}
                Else
                    simu_ss(i) = {0, 1}
                End If
            Else
                If Not cus_addskl(i)(2) = "" Then
                    simu_ss(i) = {0, 2}
                Else
                    simu_ss(i) = {0}
                End If
            End If
            'スキルセット
            simu_bs(i).skill_no = simu_ss(i).Length
        Next

        '3武将スキルセット
        For i As Integer = 0 To 2
            For j As Integer = 0 To simu_ss(i).Length - 1
                Select Case (simu_ss(i)(j))
                    Case 0
                        simu_bs(i).スキル取得(j, i, simu_bs(i).skill(0).name, simu_lv, simu_ss(i))
                    Case 1
                        simu_bs(i).スキル取得(j, i, cus_addskl(i)(0), simu_lv, simu_ss(i))
                    Case 2
                        simu_bs(i).スキル取得(j, i, cus_addskl(i)(2), simu_lv, simu_ss(i))
                End Select
            Next
        Next
    End Sub

    'コスト依存スキルでなければ全武将に追加スキルをコピーするだけで回る・・・
    Private Sub 追加スキルキャッシュ(ByVal sno As Integer, ByVal skillname As String)
        If sno = 0 Then 'add_sklはnothing
            ReDim add_skl(0)
        Else
            ReDim Preserve add_skl(sno)
        End If
        Dim tmp() As String = Skill_ref(skillname, simu_lv)
        With add_skl(sno)
            If tmp(7) Is Nothing Then tmp(7) = "U"
            .name = skillname
            .bno = 3 '追加武将は4人目
            .lv = simu_lv
            .kanren = tmp(1)
            .heika = tmp(2)
            If InStr(.heika, "全") Then
                .heika = "槍弓馬砲器"
            End If
            .koubou = tmp(3)
            'データ不足の場合を除く
            If tmp(7) = "U" Then
                .tokusyu = 5
                .kouka_p = 0
                .kouka_f = 0
                .kanren = 0
                Exit Sub
            End If
            .kouka_p = Decimal.Parse(tmp(0))
            Select Case (.kanren)
                '特殊スキル
                Case "特殊"
                    .tokusyu = 9
                    .kouka_p = 0
                    .kouka_f = 0
                    .kanren = 0
                    Exit Sub
                Case "条件" ', "童"
                    .tokusyu = 9
                    .t_flg = フラグ付きスキル参照(add_skl(sno)) '条件付きスキルの場合
                    If .koubou = "速" Then
                        .speed = Decimal.Parse(tmp(4)) '速度はコスト依存・・・しない・・・（現状
                    ElseIf tmp(5) = "速" Then
                        .speed = Decimal.Parse(tmp(6)) '付与効果に速度がある
                    End If
                    .kanren = tmp(4)
                Case Else
                    Select Case (.koubou)
                        Case "速" '速度オンリー
                            .tokusyu = 1
                            .speed = Decimal.Parse(tmp(4)) '速度はコスト依存・・・しない・・・（現状
                        Case "破壊" '破壊オンリー
                            .tokusyu = 2
                        Case Else '通常スキルの場合
                            .tokusyu = 0
                            'If InStr(tmp(4), "C") Then tmp(4) = 文字列計算(Replace(tmp(4), "C", CStr(cost))) 'コスト依存スキルの扱い。「スキル所持武将の」コストで一括適用
                            '.kouka_f = Decimal.Parse(tmp(4))
                            .kanren = tmp(4)
                            If InStr(tmp(4), "C") = 0 Then .kouka_f = Decimal.Parse(tmp(4))
                            If tmp(5) = "速" Then .speed = Decimal.Parse(tmp(6)) '付与効果に速度がある
                    End Select
            End Select
        End With
    End Sub
    '初期スキルをキャッシュから生成
    Private Sub 初期スキル読み込み(ByVal skillname As String)
        Dim zero_no As Integer = -1 '対応するスキルのIndex
        For i As Integer = 0 To zero_skl.GetLength(0) - 1
            If zero_skl(i)(0) = skillname Then
                zero_no = i
                Exit For
            End If
        Next
        With simu_bs(3).skill(0)
            'If zero_skl(zero_no)(8) Is Nothing Then Exit Sub
            .name = skillname
            .bno = 3
            .lv = simu_lv
            If zero_no = -1 Then '正常に登録されていない場合
                .tokusyu = 5
                .kouka_p = 0
                .kouka_f = 0
                .kanren = 0
                Exit Sub
            End If
            .kanren = zero_skl(zero_no)(2)
            .heika = zero_skl(zero_no)(3)
            If InStr(.heika, "全") Then
                .heika = "槍弓馬砲器"
            End If
            .koubou = zero_skl(zero_no)(4)
            'データ不足の場合を除く
            If zero_skl(zero_no)(8) = "U" Then
                .tokusyu = 5
                .kouka_p = 0
                .kouka_f = 0
                .kanren = 0
                Exit Sub
            End If
            .kouka_p = Decimal.Parse(zero_skl(zero_no)(1))
            Select Case (.kanren)
                '特殊スキル
                Case "特殊"
                    .tokusyu = 9
                    .kouka_p = 0
                    .kouka_f = 0
                    .kanren = 0
                    Exit Sub
                Case "条件" ', "童"
                    .tokusyu = 9
                    .t_flg = フラグ付きスキル参照(simu_bs(3).skill(0)) '条件付きスキルの場合
                    If .koubou = "速" Then
                        .speed = Decimal.Parse(zero_skl(zero_no)(5)) '速度はコスト依存・・・しない・・・（現状
                    ElseIf zero_skl(zero_no)(6) = "速" Then
                        .speed = Decimal.Parse(zero_skl(zero_no)(7)) '付与効果に速度がある
                    End If
                    .kanren = zero_skl(zero_no)(5)
                Case Else
                    Select Case (.koubou)
                        Case "速" '速度オンリー
                            .tokusyu = 1
                            .speed = Decimal.Parse(zero_skl(zero_no)(5)) '速度はコスト依存・・・しない・・・（現状
                        Case "破壊" '破壊オンリー
                            .tokusyu = 2
                        Case Else '通常スキルの場合
                            .tokusyu = 0
                            'If InStr(tmp(4), "C") Then tmp(4) = 文字列計算(Replace(tmp(4), "C", CStr(cost))) 'コスト依存スキルの扱い。「スキル所持武将の」コストで一括適用
                            '.kouka_f = Decimal.Parse(tmp(4))
                            .kanren = zero_skl(zero_no)(5)
                            If InStr(zero_skl(zero_no)(5), "C") Then zero_skl(zero_no)(5) = _
                                文字列計算(Replace(zero_skl(zero_no)(5), "C", CStr(simu_bs(3).cost)), False) 'コスト依存スキルの扱い。「スキル所持武将の」コストで一括適用
                            .kouka_f = Decimal.Parse(zero_skl(zero_no)(5))
                            If zero_skl(zero_no)(6) = "速" Then .speed = Decimal.Parse(zero_skl(zero_no)(7)) '付与効果に速度がある
                    End Select
            End Select
        End With
    End Sub

    Private Sub ランキング武将読み込み(ByVal bstr() As String)
        With simu_bs(3)
            .No = 3
            .id = bstr(0)
            .rare = bstr(1)
            .name = bstr(2)
            .cost = bstr(3) 'ここで4武将の合計コストが確定
            butaicostsum = 0
            For i As Integer = 0 To 3
                butaicostsum = butaicostsum + simu_bs(i).cost
            Next
            .heisyu = simu_bs(0).heisyu.Clone
            .job = bstr(16)
            .hei_max_d = bstr(4)
            .rank = busho_rank
            If Not .job = "剣" Then
                If .job = "覇" Then
                    .hei_max = .hei_max_d + .rank * 200 'ランクアップで兵数一律+200
                Else
                    .hei_max = .hei_max_d + .rank * 100 'ランクアップで兵数一律+100
                End If
            Else
                .hei_max = .hei_max_d
            End If
            .hei_sum = .hei_max
            .Tousotu(False, True) = {bstr(5), bstr(6), bstr(7), bstr(8)}
            .Sta(False) = {bstr(9), bstr(10), bstr(11)}
            ReDim .sta_g(2)
            .sta_g = {Val(bstr(12)), Val(bstr(13)), Val(bstr(14))}
            heino = 兵振り武将決定() '誰を兵法振りにするか決める
            Call ステ振りセット() '全武将のステ振りが確定できる
            .heisyu.ts = 実質統率計算(.heisyu.tousotu, .No)
            .skill(0).name = bstr(15)
            Call ランキング武将スキル()
        End With
    End Sub
    Private Sub ランキング武将スキル()
        If add_skl Is Nothing Then 'まだ追加スキルがキャッシュされていない(一番最初の武将のみ)
            Dim rec() As String '参照するスキル

            If syosklflg Then '追加スキル詳細設定ONならば
                rec = {cus_addskl(3)(0), cus_addskl(3)(2)}
            Else
                If simu_skeqflg Then '3武将と同じならば
                    rec = {ComboBox011.Text, ComboBox012.Text}
                Else
                    rec = {ComboBox041.Text, ComboBox042.Text}
                End If
            End If

            For i As Integer = 1 To simu_ss(3).Length - 1
                Call 追加スキルキャッシュ(i - 1, rec(simu_ss(3)(i) - 1))
            Next
        End If
        ReDim Preserve simu_bs(3).skill(simu_bs(3).skill_no - 1)
        For i As Integer = 0 To simu_ss(3).Length - 1
            Select Case (simu_ss(3)(i))
                Case 0
                    Call 初期スキル読み込み(simu_bs(3).skill(0).name)
                Case Else
                    simu_bs(3).skill(i) = add_skl(i - 1).Clone
                    With simu_bs(3).skill(i)
                        If InStr(.kanren, "C") Then 'コスト依存スキルの扱い
                            .kanren = Replace(.kanren, "C", CStr(simu_bs(3).cost)) '変更。「スキル所持武将の」コストで一括適用
                            .kanren = 文字列計算(.kanren, False)
                            .kouka_f = Decimal.Parse(.kanren)
                        Else
                            .kouka_f = 文字列計算(.kanren, False)
                        End If
                        'If .kanren Then 'データが無いスキルの場合は計算しない
                        '    .kanren = 文字列計算(.kanren)
                        '    .kouka_f = Decimal.Parse(.kanren)
                        'Else
                        '    .kouka_f = 0
                        'End If
                    End With
            End Select
        Next
    End Sub

    '*** 全武将セットでの対象の関数 ***
    '兵法振りにする武将Noセットを返す
    Private Function 兵振り武将決定() As Integer()
        Dim heiho(3) As Decimal
        Dim pwr(3) As Decimal
        Dim ret_int() As Integer
        'ステ振り詳細設定がONの時
        If statusflg Then
            Return {-1}
        End If
        '全員兵法じゃない時
        If heicount = 0 Then
            ReDim ret_int(0)
            ret_int(0) = 0
            Return ret_int
            Exit Function
        End If
        ReDim ret_int(heicount - 1)
        '兵法値計算
        For i As Integer = 0 To 3
            heiho(i) = simu_bs(i).st_d(2) + simu_bs(i).sta_g(2) * stpt
            If InStr(kobo, "攻") Then
                pwr(i) = simu_bs(i).st_d(0) + simu_bs(i).sta_g(0) * stpt
            Else
                pwr(i) = simu_bs(i).st_d(1) + simu_bs(i).sta_g(1) * stpt
            End If
        Next
        For i As Integer = 0 To heicount - 1
            ret_int(i) = Array_MAX(heiho, pwr)
            heiho(ret_int(i)) = -1
        Next
        Return ret_int
    End Function

    '全武将のステータスセット
    Private Sub ステ振りセット()
        For i As Integer = 0 To 3
            Dim heiflg As Boolean = False
            With simu_bs(i)
                If statusflg Then 'ステ振り詳細設定がONならば
                    Select Case cus_status(i)
                        Case 0 '攻極振
                            .st(0) = .st_d(0) + stpt * .sta_g(0)
                            .st(1) = .st_d(1)
                            .st(2) = .st_d(2)
                        Case 1 '防極振
                            .st(0) = .st_d(0)
                            .st(1) = .st_d(1) + stpt * .sta_g(1)
                            .st(2) = .st_d(2)
                        Case 2 '兵極振
                            .st(0) = .st_d(0)
                            .st(1) = .st_d(1)
                            .st(2) = .st_d(2) + stpt * .sta_g(2)
                        Case 3 '適正お任せ
                            'まずは攻防成長値で決定
                            If InStr(kobo, "攻") Then
                                If .sta_g(0) >= cus_status_kb Then
                                    .st(0) = .st_d(0) + stpt * .sta_g(0)
                                    .st(1) = .st_d(1)
                                    .st(2) = .st_d(2)
                                Else
                                    heiflg = True
                                End If
                            Else
                                If .sta_g(1) >= cus_status_kb Then
                                    .st(0) = .st_d(0)
                                    .st(1) = .st_d(1) + stpt * .sta_g(1)
                                    .st(2) = .st_d(2)
                                Else
                                    heiflg = True
                                End If
                            End If
                            '次に兵法成長値で決定
                            If heiflg Then
                                If .sta_g(2) >= cus_status_hei Then
                                    .st(0) = .st_d(0)
                                    .st(1) = .st_d(1)
                                    .st(2) = .st_d(2) + stpt * .sta_g(2)
                                Else
                                    If InStr(kobo, "攻") Then
                                        .st(0) = .st_d(0) + stpt * .sta_g(0)
                                        .st(1) = .st_d(1)
                                    Else
                                        .st(0) = .st_d(0)
                                        .st(1) = .st_d(1) + stpt * .sta_g(1)
                                    End If
                                End If
                            End If
                    End Select
                Else
                    For j As Integer = 0 To heicount - 1
                        If .No = heino(j) Then '兵法振りにする武将リストと合致
                            heiflg = True
                        End If
                    Next
                    If heiflg Then
                        .st(0) = .st_d(0)
                        .st(1) = .st_d(1)
                        .st(2) = .st_d(2) + stpt * .sta_g(2)
                    Else
                        If InStr(kobo, "攻") Then
                            .st(0) = .st_d(0) + stpt * .sta_g(0)
                            .st(1) = .st_d(1)
                        Else
                            .st(0) = .st_d(0)
                            .st(1) = .st_d(1) + stpt * .sta_g(1)
                        End If
                        .st(2) = .st_d(2)
                    End If
                End If
            End With
        Next
    End Sub

    '部隊兵法値や各小隊の戦闘力を計算
    Private Sub 部隊兵法値計算() 'いわば前段階
        '部隊兵法値計算
        Dim maxheihou As Decimal = 0
        Dim heihou_kei As Decimal = 0
        For i As Integer = 0 To 3
            If (maxheihou < simu_bs(i).st(2)) Then
                maxheihou = simu_bs(i).st(2)
            End If
            heihou_kei = heihou_kei + simu_bs(i).st(2)
        Next
        heihousum = (maxheihou + (heihou_kei - maxheihou) / 6) / 100
        '小隊戦闘力計算
        For i As Integer = 0 To 3
            With simu_bs(i)
                .小隊攻撃力計算(kobo)
                .スキル期待値計算(heihousum, butairankbonus, butaicostsum)
            End With
        Next
        '海野六郎のようなスキル（他武将のスキル条件に影響を与えるスキル）がある。その部分を補正
        For i As Integer = 0 To 3
            For j As Integer = 0 To simu_bs(i).skill.Length - 1
                With simu_bs(i).skill(j)
                    '発動率補正(up_kouka_p)
                    If (.kouka_p_b + .up_kouka_p) > 1 Then '合計100%を超えるならば
                        .kouka_p_b = 1
                    Else
                        .kouka_p_b = .kouka_p_b + .up_kouka_p
                    End If
                    .exp_kouka_b = .kouka_p_b * .kouka_f
                    '上昇率補正(up_kouka_f)
                    .kouka_f_b = .kouka_f_b + .up_kouka_f
                    .exp_kouka_b = .kouka_p_b * .kouka_f_b
                    '兵科追加補正(heika_add)
                    .heika_b = .heika_b + .heika
                End With
            Next
        Next
    End Sub
    '**********************************

    'tc:統率に関係する兵科, bn:武将No
    Private Function 実質統率計算(ByVal ts As String, ByVal bn As Integer) As Decimal
        Dim t As Integer = ts.Length
        Dim td, v() As Decimal
        ReDim v(t - 1)
        For j As Integer = 1 To t
            Select Case Mid(ts, j, 1)
                Case "槍"
                    v(j - 1) = 統率_数値変換(simu_bs(bn).tou_a(0))
                Case "弓"
                    v(j - 1) = 統率_数値変換(simu_bs(bn).tou_a(1))
                Case "馬"
                    v(j - 1) = 統率_数値変換(simu_bs(bn).tou_a(2))
                Case "器"
                    v(j - 1) = 統率_数値変換(simu_bs(bn).tou_a(3))
            End Select
            td = td + v(j - 1)
        Next
        td = td / t '該当兵科の実質統率値
        If td + busho_rank * (0.05 / t) < 1.2 Then 'ランクアップ時の統率変化
            td = td + busho_rank * (0.05 / t)
        Else
            td = 1.2
        End If
        Return td
    End Function

    '*** スキル期待値計算に関わる関数 ***
    'can_skillとcan_skillpを計算
    Private Sub スキル状態数計算(ByRef c_skill() As Busho.skl, ByRef c_skillp() As String)
        c_skill = Nothing
        c_skillp = Nothing
        Dim c As Integer = 0 'カウンター
        '全有効スキル数カウント
        For i As Integer = 0 To 3
            With simu_bs(i)
                For j As Integer = 0 To .skill_no - 1 '攻防一致、特殊スキル排除
                    If InStr(kobo, .skill(j).koubou) Or .skill(j).koubou = "攻防" Then
                        If .skill(j).tokusyu = 9 Then
                            .skill(j).t_flg = 条件依存スキル・フラグスキル判定(.skill(j), butaicostsum, butairanksum) '怪しいスキルを疑う
                            .skill(j).t_flg = フラグ付きスキル参照(.skill(j))
                        End If
                        If .skill(j).tokusyu = 0 Or .skill(j).t_flg Then '通常スキル
                            ReDim Preserve c_skill(c)
                            c_skill(c) = .skill(j).Clone
                            c = c + 1
                        End If
                    ElseIf InStr(kobo, "攻") And InStr(.skill(j).koubou, "上級器") Then '攻撃かつ上級器→上級器攻スキルの条件を満たす
                        ReDim Preserve c_skill(c)
                        c_skill(c) = .skill(j).Clone
                        c = c + 1
                    ElseIf InStr(kobo, "防") And InStr(.skill(j).koubou, "上級砲") Then '防御かつ上級砲→上級砲防スキルの条件を満たす
                        ReDim Preserve c_skill(c)
                        c_skill(c) = .skill(j).Clone
                        c = c + 1
                    ElseIf InStr(kobo, "攻") And InStr(.skill(j).koubou, "秘境兵") Then '攻撃かつ秘境兵→秘境兵攻スキルの条件を満たす
                        ReDim Preserve c_skill(c)
                        c_skill(c) = .skill(j).Clone
                        c = c + 1
                    End If
                Next
            End With
        Next
        '条件依存のある遅延スキル
        For i As Integer = 0 To 3
            With simu_bs(i)
                For j As Integer = 0 To .skill_no - 1 '攻防一致、特殊スキル排除
                    If InStr(kobo, .skill(j).koubou) Or .skill(j).koubou = "攻防" Then
                        '上で引っかからなかった特殊スキルのみ
                        If .skill(j).tokusyu = 9 And Not (.skill(j).t_flg) Then
                            .skill(j).t_flg = 遅延条件依存スキル(.skill(j), butaicostsum, butairanksum) '怪しいスキルを疑う
                            '.skill(j).t_flg = フラグ付きスキル参照(.skill(j))
                            If .skill(j).t_flg Then '有効ならば
                                ReDim Preserve c_skill(c)
                                c_skill(c) = .skill(j).Clone
                                c = c + 1
                            End If
                        End If
                    End If
                Next
            End With
        Next

        c = 2 ^ (c) '全スキル有効状態数
        ReDim c_skillp(c - 1)
        For i As Integer = 0 To c - 1
            c_skillp(i) = Convert10to2(i, Math.Ceiling(Math.Log10(c) / Math.Log10(2)))
        Next
    End Sub
    
    '期待値計算(0:ランキング武将の期待値, 1:全体期待値, 2:MAX発動時)
    Private Function 期待値計算() As Decimal()
        Dim c_skill() As Busho.skl = Nothing '発動するスキル
        Dim c_skillp() As String = Nothing 'スキル状態文字列
        Call スキル状態数計算(c_skill, c_skillp)
        Dim s_exk, s_ex(3) As Decimal
        Dim s_x() As Decimal = Nothing 'x->生起確率
        Dim s_y(,) As Decimal = Nothing '総戦闘力f(x)(k) k:k番目の武将の戦闘力
        Dim s_yk() As Decimal = Nothing 'sum(s_y)
        Dim s_yk_max As Decimal = Nothing 'MAX
        '童関係
        Dim harr() As String = {"槍", "弓", "馬", "砲", "器"}
        Dim warr() As Decimal = warabe.warabe_gets(kobo)

        ReDim s_yk(c_skillp.Length - 1)
        ReDim s_x(c_skillp.Length - 1), s_y(c_skillp.Length - 1, 3)

        'スキル状態計算
        For i As Integer = 0 To c_skillp.Length - 1
            Dim syoplus(3) As Decimal '将UP率
            Dim heiplus(3) As Decimal '一般UP率
            s_x(i) = 1
            For k As Integer = 0 To 3
                s_y(i, k) = simu_bs(k).attack
            Next
            If Not c_skill Is Nothing Then 'どれか意味のあるスキルが存在する
                For j As Integer = 1 To c_skill.Length
                    If Mid(c_skillp(i), j, 1) = 1 Then '発動ならば
                        Dim ttmp(,) As Decimal = Form1.スキル実質上昇率(c_skill(j - 1))
                        s_x(i) = s_x(i) * c_skill(j - 1).kouka_p_b
                        For k As Integer = 0 To 3
                            syoplus(k) = syoplus(k) + ttmp(k, 0)
                            heiplus(k) = heiplus(k) + ttmp(k, 1)
                        Next
                    Else
                        s_x(i) = s_x(i) * (1 - c_skill(j - 1).kouka_p_b)
                    End If
                Next
            Else
                For k As Integer = 0 To 3
                    syoplus(k) = 0
                    heiplus(k) = 0
                Next
            End If
            '実際のスキル発動時の戦闘力を計算
            For k As Integer = 0 To 3
                Dim ds, dk As Integer
                With simu_bs(k)
                    If InStr(kobo, "攻") Then
                        ds = .st(0)
                        dk = .heisyu.atk
                    Else
                        ds = .st(1)
                        dk = .heisyu.def
                    End If
                    s_y(i, k) = (ds * .heisyu.ts * (1 + syoplus(k)) + .hei_max * dk * .heisyu.ts) * (1 + heiplus(k))
                    '童適用(今のところ単科防スキルのみ対応)
                    For l As Integer = 0 To harr.Length - 1
                        If InStr(.heisyu.bunrui, harr(l)) Then
                            s_y(i, k) = s_y(i, k) * (1 + 0.01 * warr(l))
                        End If
                    Next
                End With
            Next
        Next
        s_yk = Array_to_Arrayk(s_y)
        s_yk_max = s_yk(s_yk.Length - 1) 'MAX

        '母集団の期待値、分散
        For i As Integer = 0 To s_x.Length - 1
            For k As Integer = 0 To 3
                s_ex(k) = s_ex(k) + s_x(i) * s_y(i, k)
            Next
        Next
        For k As Integer = 0 To 3
            s_exk = s_exk + s_ex(k)
        Next
        'For i As Integer = 0 To s_x.Length - 1
        '    s_ax = s_ax + s_x(i) * ((s_yk(i) - s_exk) ^ 2) '分散
        'Next
        Return {s_ex(3), s_exk, s_yk_max}
    End Function
    'グローバル変数初期化
    Private Sub グローバル変数初期化()
        kitai_val = Nothing
        kitai_butai = Nothing
        add_skl = Nothing
        butairanksum = 0
    End Sub

    '各ランキング毎に回す変数初期化
    Private Sub 変数初期化()
        heihousum = 0
        'スキル計算に使う変数をクリア
        For i As Integer = 0 To 3
            For j As Integer = 0 To simu_bs(i).skill_no - 1
                simu_bs(i).skill(j).スキル計算状態初期化()
            Next
        Next
    End Sub

    Private Sub 表更新(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Dim jf As Boolean = 条件入力()
        If jf = False Then '条件に不備があれば
            MsgBox("必要な条件が設定されていない個所があります")
            Exit Sub
        End If

        simu_execno = 1

        Call プログレスバー初期化()
        Call グローバル変数初期化()
        DataGridView1.Rows.Clear() '表をクリア
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        '全武将兵科セット（1武将にセットして他はコピー方式）
        simu_bs(0).兵科情報取得(busho_heika, kobo)
        For i As Integer = 1 To 3
            simu_bs(i).heisyu = simu_bs(0).heisyu.Clone
        Next
        '3武将ランクアップ、統率計算
        For i As Integer = 0 To 2
            With simu_bs(i)
                .rank = busho_rank
                If Not .job = "剣" Then
                    If .job = "覇" Then
                        .hei_max = .hei_max_d + .rank * 200 'ランクアップで兵数一律+200
                    Else
                        .hei_max = .hei_max_d + .rank * 100 'ランクアップで兵数一律+100
                    End If
                End If
                .hei_sum = .hei_max
                .heisyu.ts = 実質統率計算(.heisyu.tousotu, i)
            End With
        Next
        '部隊ランクボーナス計算
        butairanksum = simu_bs(0).rank * 4
        butairankbonus = 部隊ランクボーナス計算(butairanksum)

        'スキル状態、スキル数をセット
        If Not syosklflg Then
            Call 追加スキルセット_def()
        Else
            Call 追加スキルセット_Cus()
        End If

        Call フラグ付きスキル読み込み() '読込（更新）
        Call 童ボーナス加算()

        'ランキング対象武将をDBから取得
        Dim sl()() As String = Nothing
        Dim slbl() As String = _
        {"Bid", "武将R", "武将名", "Cost", "指揮兵数", "槍統率", "弓統率", "馬統率", "器統率", "初期攻撃", "初期防御", "初期兵法", "攻成長", "防成長", "兵成長", "初期スキル名", "職"}
        Dim target_r As String = 対象範囲文字列出力()
        Dim sqlstr As String = "SELECT * FROM BData WHERE " & target_r & " AND Bunf = 'F'"
        sl = DB_DirectOUT3(sqlstr, slbl)

        '各武将の初期スキルの性能をまとめて取ってくる
        Dim syoki_skl() As String = Nothing
        Dim syoki_skinm As String = ""
        Dim syoki_c As Integer = 0
        Dim sdflg As Boolean = False
        For i As Integer = 0 To sl.GetLength(0) - 1
            'スキル重複登録を防ぐ
            sdflg = False
            syoki_skinm = sl(i)(15)
            If Not syoki_skl Is Nothing Then '空でないならば
                For j As Integer = 0 To syoki_skl.Length - 1
                    If syoki_skinm = syoki_skl(j) Then '既にsyoki_sklに登録済
                        sdflg = True '重複フラグON
                    End If
                Next
            End If
            If Not sdflg Then
                ReDim Preserve syoki_skl(syoki_c)
                syoki_skl(syoki_c) = syoki_skinm
                syoki_c = syoki_c + 1
            End If
        Next
        zero_skl = Skill_ref_list(syoki_skl, simu_lv)

        Dim sc As Integer = sl.GetLength(0) - 1 '合致したデータ個数
        ReDim kitai_val(sc), kitai_butai(sc), kitai_max(sc)
        Dim rows As DataGridViewRow()
        ReDim rows(sc)

        For i As Integer = 0 To sc
            Dim tmprec(2) As Decimal
            Call ランキング武将読み込み(sl(i))
            Call 変数初期化()
            Call 部隊兵法値計算()
            tmprec = 期待値計算()
            kitai_val(i) = tmprec(0)
            kitai_butai(i) = tmprec(1)
            kitai_max(i) = tmprec(2)
            Dim bcost As Decimal = simu_bs(0).cost + simu_bs(1).cost + simu_bs(2).cost '3将コスト
            Dim batk As Decimal = simu_bs(0).attack + simu_bs(1).attack + simu_bs(2).attack '3将素攻/防
            '*** 表示 ***
            Call プログレスバー変化(200 + ((i + 1) / (sc + 1)) * 800)
            Dim row As DataGridViewRow = New DataGridViewRow()
            row.CreateCells(DataGridView1)
            With simu_bs(3)
                Dim max_hei, max_syo As Decimal 'MAX兵法値, MAX将攻/防
                max_hei = .st_d(2) + .sta_g(2) * stpt
                If InStr(kobo, "攻") Then
                    max_syo = .st_d(0) + .sta_g(0) * stpt
                Else
                    max_syo = .st_d(1) + .sta_g(1) * stpt
                End If
                Dim kitai_cost As String = Math.Floor(kitai_butai(i) / (.cost + bcost)).ToString("#,#") 'C1コス比
                If .cost = 193 Then Continue For 'デッキセットできないカードは飛ばす
                row.SetValues(New Object() {.id, .rare, .name, .cost, .hei_max, .heisyu.ts, max_hei, max_syo, _
                                            Math.Floor(kitai_butai(i)).ToString("#,#"), _
                                            "+" & (Math.Ceiling(((kitai_butai(i) / (.attack + batk)) - 1) * 10000) / 10000) * 100.ToString & "%", _
                                            Math.Floor(kitai_val(i)).ToString("#,#"), _
                                            kitai_cost, Math.Floor(kitai_max(i)).ToString("#,#")})
                'DataGridView1.Rows.AddRange(row)
                Call 表色付け(row)
            End With
            rows(i) = row
        Next
        DataGridView1.Rows.AddRange(rows)
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Call プログレスバー変化(0)
    End Sub

    Private Sub 表色付け(ByRef row As DataGridViewRow)
        Dim cell As DataGridViewCell = Nothing
        With simu_bs(3)
            cell = row.Cells(6)
            If .sta_g(2) >= 3.0 Then '兵法成長で色づけ(3.0->紫, 2.5->赤)
                cell.Style.ForeColor = Color.DarkMagenta
                cell.Style.Font = New Font("Consolas", 10, FontStyle.Bold)
            ElseIf .sta_g(2) >= 2.5 Then
                cell.Style.ForeColor = Color.DarkRed
                cell.Style.Font = New Font("Consolas", 10, FontStyle.Bold)
            End If

            Dim syflg As Boolean = False
            '初期スキルが兵科対応しているかどうか
            If InStr(kobo, .skill(0).koubou) Or .skill(0).koubou = "攻防" Then
                If InStr(.skill(0).heika, .heisyu.bunrui) Then
                    syflg = True
                ElseIf InStr(.skill(0).heika, "将") Then
                    syflg = True
                End If
            Else
                If InStr(.skill(0).koubou, "上級器") And simu_bs(3).heisyu.jyk_utuwa = True And InStr(kobo, "攻") Then
                    syflg = True '上級器攻対応
                ElseIf InStr(.skill(0).koubou, "上級砲") And simu_bs(3).heisyu.jyk_hou = True And InStr(kobo, "防") Then
                    syflg = True '上級砲防対応
                ElseIf InStr(.skill(0).koubou, "秘境兵") And simu_bs(3).heisyu.tok_hikyo = True And InStr(kobo, "攻") Then
                    syflg = True '秘境兵攻対応
                End If
            End If
            cell = row.Cells(8)
            If syflg Then '兵科対応スキル持ち色付け
                cell.Style.ForeColor = Color.DarkGreen
                cell.Style.Font = New Font("Consolas", 10, FontStyle.Bold)
            End If
            If .skill(0).tokusyu = 5 Then '初期スキルがDBにデータ×
                cell.Style.ForeColor = Color.RosyBrown
            End If
            cell = row.Cells(1) '武将レアリティによる色分け
            Dim cell2 As DataGridViewCell = row.Cells(2)
            Select Case Mid(CStr(row.Cells(0).Value), 1, 2)
                Case "10" '天
                    cell.Style.ForeColor = Color.Goldenrod
                    cell2.Style.ForeColor = Color.Goldenrod
                Case "20" To "25" '極
                    cell.Style.ForeColor = Color.DimGray
                    cell2.Style.ForeColor = Color.DimGray
                Case "27" 'シクレ極
                    cell.Style.ForeColor = Color.Purple
                    cell2.Style.ForeColor = Color.Purple
                Case "29" 'プラチナ極
                    cell.Style.ForeColor = Color.Black
                    cell2.Style.ForeColor = Color.Black
                Case "30" To "35" '特
                    cell.Style.ForeColor = Color.Firebrick
                    cell2.Style.ForeColor = Color.Firebrick
                Case "37" 'シクレ特
                    cell.Style.ForeColor = Color.DarkOliveGreen
                    cell2.Style.ForeColor = Color.DarkOliveGreen
                Case "40" To "45" '上
                    cell.Style.ForeColor = Color.Orange
                    cell2.Style.ForeColor = Color.Orange
                Case "50" To "55" '序
                    cell.Style.ForeColor = Color.DarkCyan
                    cell2.Style.ForeColor = Color.DarkCyan
                Case "57" '輝く平手
                    cell.Style.ForeColor = Color.Turquoise
                    cell2.Style.ForeColor = Color.Turquoise
            End Select
        End With
    End Sub

    'カンマ区切りの部分は通常ソートでは正しくソートできない
    Private Sub DataGridView1_SortCompare(ByVal sender As Object, ByVal e As DataGridViewSortCompareEventArgs) _
        Handles DataGridView1.SortCompare
        If e.Column.Index = 8 Or e.Column.Index = 9 Or e.Column.Index = 10 Or e.Column.Index = 11 Or e.Column.Index = 12 Then
            '指定されたセルの値を文字列として取得する
            Dim str1, str2 As String
            If e.CellValue1 Is Nothing Then
                str1 = ""
            Else
                str1 = e.CellValue1.ToString
            End If
            If e.CellValue2 Is Nothing Then
                str2 = ""
            Else
                str2 = e.CellValue2.ToString
            End If
            Dim dec1, dec2 As Decimal
            If Not e.Column.Index = 9 Then '上昇率（+が付いていない通常の数字）
                dec1 = Val(Replace(str1, ",", ""))
                dec2 = Val(Replace(str2, ",", ""))
            Else
                dec1 = Val(Replace(str1.Remove(0, 1), "%", ""))
                dec2 = Val(Replace(str2.Remove(0, 1), "%", ""))
            End If
            '結果を代入
            If dec1 > dec2 Then
                e.SortResult = 1
            Else
                e.SortResult = -1
            End If
            '処理したことを知らせる
            e.Handled = True
        End If
    End Sub
    'csv出力
    Private Sub CSV出力(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        '1行もデータが無い場合は、保存を中止します。
        If DataGridView1.Rows.Count = 0 Then
            Exit Sub
        End If

        Dim strResult As New System.Text.StringBuilder
        Dim fname As String = "" 'csvファイル名
        While 1
            fname = InputBox("保存csvファイル名", "現在のランキング結果をcsvファイルに出力しますか？")
            FILENAME_csv = My.Application.Info.DirectoryPath & "\ranking\" & fname & ".csv" 'csvファイルの保存場所
            If fname = "" Then '無名の場合は抜ける
                Exit Sub
            ElseIf IO.File.Exists(FILENAME_csv) Then
                Dim yn As Integer = MsgBox("既に同名の部隊が保存されています。上書きしますか？", vbYesNo)
                If yn = vbYes Then
                    Exit While
                End If
            Else
                Exit While
            End If
        End While

        '条件列挙
        strResult.Append("******** ウェーバーモード：武将ランキング出力 ********" & vbCrLf)
        If busho_rank = 6 Then
            strResult.Append("ランク/LV" & "," & "☆限界突破☆" & vbCrLf)
        Else
            strResult.Append("ランク/LV" & "," & "★" & simu_bs(0).rank & "/LV20" & vbCrLf)
        End If
        strResult.Append("スキルLV" & "," & "ALL" & simu_bs(0).skill(0).lv & vbCrLf)
        strResult.Append("ステ振り" & "," & ComboBox10.Text & vbCrLf)
        strResult.Append("******* 固定3武将 *******" & vbCrLf)
        strResult.Append("No" & "," & "武将名" & "," & "R" & "," & "コスト" & "," & "指揮兵数" & "," & "初期スキル" & "," & "効果範囲" & "," & "発動率(%)" & "," & "上昇率(%)" & vbCrLf)
        For i As Integer = 0 To 2
            With simu_bs(i)
                Dim kskb As String = .skill(0).koubou
                If Not (kskb = "攻" Or kskb = "防") Then
                    kskb = ""
                End If
                Dim kh As String = .skill(0).heika
                If kh = "槍弓馬砲器" Then
                    kh = "全"
                End If
                strResult.Append(.id & "," & .name & "," & .rare & "," & .cost & "," & .hei_sum & "," & _
                                 .skill(0).name & "," & kh & kskb & "," & .skill(0).kouka_p * 100 & "," & .skill(0).kouka_f * 100 & vbCrLf)
            End With
        Next
        strResult.Append("******* 追加スキル *******" & vbCrLf)
        strResult.Append("追加スキル名" & "," & "効果範囲" & "," & "発動率(%)" & "," & "上昇率(%)" & vbCrLf)
        For i As Integer = 1 To simu_bs(0).skill_no - 1
            With simu_bs(0).skill(i)
                Dim kskb As String = .koubou
                If Not (kskb = "攻" Or kskb = "防") Then
                    kskb = ""
                End If
                Dim kh As String = .heika
                If kh = "槍弓馬砲器" Then
                    kh = "全"
                End If
                Dim koukaft As String = (.kouka_f * 100).ToString
                If InStr(.kanren, "C") Then 'コストが含まれている＝コスト依存
                    koukaft = .kanren
                End If
                strResult.Append(.name & "," & kh & kskb & "," & .kouka_p * 100 & "," & koukaft & vbCrLf)
            End With
        Next
        If Not simu_skeqflg Then 'ランキング武将が独自の追加スキル
            strResult.Append("+++++++ 追加スキル（ランキング武将） +++++++" & vbCrLf)
            For i As Integer = 0 To add_skl.Length - 1
                With add_skl(i)
                    Dim kskb As String = .koubou
                    If Not (kskb = "攻" Or kskb = "防") Then
                        kskb = ""
                    End If
                    Dim kh As String = .heika
                    If kh = "槍弓馬砲器" Then
                        kh = "全"
                    End If
                    Dim koukaft As String = (.kouka_f * 100).ToString
                    If koukaft = 0 Then 'セットされていない＝コスト依存
                        koukaft = .kanren
                    End If
                    strResult.Append(.name & "," & kh & kskb & "," & .kouka_p * 100 & "," & koukaft & vbCrLf)
                End With
            Next
        End If
        strResult.Append("***********************************************" & vbCrLf)
        'コラムヘッダを1行目に列記
        For i As Integer = 0 To DataGridView1.Columns.Count - 1
            Select Case i
                Case 0
                    strResult.Append("""" & DataGridView1.Columns(i).HeaderText.ToString & """")
                Case (DataGridView1.Columns.Count - 1)
                    strResult.Append("," & """" & DataGridView1.Columns(i).HeaderText.ToString & """" & vbCrLf)
                Case Else
                    strResult.Append("," & """" & DataGridView1.Columns(i).HeaderText.ToString & """")
            End Select
        Next
        'データを保存
        For i As Integer = 0 To DataGridView1.Rows.Count - 2
            For j As Integer = 0 To DataGridView1.Columns.Count - 1
                Select Case j
                    Case 0
                        strResult.Append("""" & DataGridView1.Rows(i).Cells(j).Value.ToString & """")
                    Case (DataGridView1.Columns.Count - 1)
                        strResult.Append("," & """" & DataGridView1.Rows(i).Cells(j).Value.ToString & """" & vbCrLf)
                    Case Else
                        strResult.Append("," & """" & DataGridView1.Rows(i).Cells(j).Value.ToString & """")
                End Select
            Next
        Next
        'Shift-JISで保存
        Dim swText As New System.IO.StreamWriter(FILENAME_csv, False, System.Text.Encoding.GetEncoding(932))
        swText.Write(strResult.ToString)
        swText.Dispose()
    End Sub

    Private Sub 追加スキル詳細設定(sender As Object, e As EventArgs) Handles Button1.Click
        Form11.Show()
    End Sub

    Private Sub 設定クリア()
        For i As Integer = 1 To 3
            ComboBox(Me, CStr(i)).Text = ""
            ComboBox(Me, "00" & CStr(i)).Text = ""
        Next
        ComboBox011.Text = ""
        ComboBox012.Text = ""
        ComboBox01.Text = ""
        ComboBox02.Text = ""
        CheckBox1.Checked = True
        ComboBox41.Text = ""
        ComboBox42.Text = ""
        ComboBox041.Text = ""
        ComboBox042.Text = ""
        CheckBox2.Checked = True
        CheckBox2.Checked = False
        syosklflg = False
        statusflg = False
        Call Form11.スキル詳細設定関連ONOFF(True)
        Call Form15.ステ振り詳細設定関連ONOFF(True)
    End Sub

    Private Sub お気に入り設定を開く(sender As Object, e As EventArgs) Handles お気に入り設定を開くToolStripMenuItem.Click
        Dim bini As String '適用するINIファイル
        'OpenFileDialogクラスのインスタンスを作成
        Dim ofd As New OpenFileDialog()

        'はじめに表示されるフォルダを指定する
        '指定しない（空の文字列）の時は、現在のディレクトリが表示される
        ofd.InitialDirectory = My.Application.Info.DirectoryPath & "\settings"
        'タイトルを設定する
        ofd.Title = "設定ファイルを選択してください"
        If ofd.ShowDialog() = DialogResult.OK Then
            'OKボタンがクリックされたとき
            '選択されたファイル名を表示する
            bini = ofd.FileName
            ofd.Dispose() '要らなくなれば破棄
            Call INIファイルから読み込み(bini)
        Else
            ofd.Dispose()
            Exit Sub
        End If
    End Sub

    Private Sub INIファイルから読み込み(ByVal bini As String)
        Cursor.Current = Cursors.WaitCursor
        Call 設定クリア()
        '全体設定
        ToolStripComboBox1.Text = GetINIValue("kobo", "設定", bini)
        ToolStripComboBox2.Text = GetINIValue("heika", "設定", bini)
        ToolStripComboBox3.SelectedIndex = CInt(GetINIValue("rank", "設定", bini))
        ComboBox9.Text = GetINIValue("skilllv", "設定", bini)
        ComboBox10.SelectedIndex = CInt(GetINIValue("heicount", "設定", bini))
        '対象武将設定
        Dim rarearr() = {"天", "極", "特", "上", "序"}
        For i As Integer = 0 To rarearr.Length - 1
            CheckBox(Me, "00" & CStr(i + 1)).Checked = GetINIValue(rarearr(i), "対象武将", bini)
        Next
        '各種フラグ
        syosklflg = GetINIValue("detail_skillflg", "設定", bini)
        statusflg = GetINIValue("detail_statusflg", "設定", bini)
        simu_skeqflg = GetINIValue("rankbusho_skillflg", "設定", bini)
        '武将、スキル
        For i As Integer = 0 To 2
            Dim bsho As String = Nothing
            Select Case i
                Case 0
                    bsho = "武将A"
                Case 1
                    bsho = "武将B"
                Case 2
                    bsho = "武将C"
                Case 3
                    bsho = "ランキング武将"
            End Select
            ComboBox(Me, CStr(i + 1)).SelectedIndex = ComboBox(Me, CStr(i + 1)).FindString(GetINIValue("rare", bsho, bini)) '（強制的に）R選択
            'R選択(ComboBox(Me, CStr(i + 1)), Nothing)
            ComboBox(Me, "00" & CStr(i + 1)).SelectedText = GetINIValue("name", bsho, bini) '（強制的に）武将名選択
            武将名選択(ComboBox(Me, "00" & CStr(i + 1)), Nothing)
        Next
        If syosklflg Then 'スキル詳細設定アリならば
            Form11.スキル詳細設定関連ONOFF(False)
            For i As Integer = 0 To 3
                ReDim cus_addskl(i)(3)
                For j As Integer = 0 To 3
                    Select Case (j Mod 2)
                        Case 0
                            cus_addskl(i)(j) = GetINIValue("add" & CStr(i) & CStr(j \ 2 + 1), "CSKILL", bini)
                        Case 1
                            cus_addskl(i)(j) = スキル関連推定(GetINIValue("add" & CStr(i) & CStr(j \ 2 + 1), "CSKILL", bini), True)
                    End Select
                Next
            Next
        Else 'スキル通常設定ならば
            ComboBox01.Focus()
            ComboBox01.SelectedIndex = ComboBox01.FindString(スキル関連推定(GetINIValue("add1", "武将A", bini), True))
            ComboBox011.SelectedText = GetINIValue("add1", "武将A", bini)
            ComboBox02.Focus()
            ComboBox02.SelectedIndex = ComboBox02.FindString(スキル関連推定(GetINIValue("add2", "武将A", bini), True))
            ComboBox012.SelectedText = GetINIValue("add2", "武将A", bini)
            If Not simu_skeqflg Then 'ランキング武将のスキルがオリジナル
                CheckBox1.Checked = False
                ComboBox41.Focus()
                ComboBox41.SelectedIndex = ComboBox41.FindString(スキル関連推定(GetINIValue("add1", "ランキング武将", bini), True))
                ComboBox041.SelectedText = GetINIValue("add1", "ランキング武将", bini)
                ComboBox42.Focus()
                ComboBox42.SelectedIndex = ComboBox42.FindString(スキル関連推定(GetINIValue("add2", "ランキング武将", bini), True))
                ComboBox042.SelectedText = GetINIValue("add2", "ランキング武将", bini)
            End If
        End If
        If statusflg Then 'ステ振り詳細設定ONならば
            For i As Integer = 0 To 3
                Dim bsho As String = Nothing
                Select Case i
                    Case 0
                        bsho = "武将A"
                    Case 1
                        bsho = "武将B"
                    Case 2
                        bsho = "武将C"
                    Case 3
                        bsho = "ランキング武将"
                End Select
                cus_status(i) = GetINIValue(bsho, "ステ振り設定", bini)
            Next
            cus_status_kb = Val(GetINIValue("攻防", "ステ振り設定", bini))
            cus_status_hei = Val(GetINIValue("兵法", "ステ振り設定", bini))
        End If
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub 設定保存(sender As Object, e As EventArgs) Handles 設定保存ToolStripMenuItem.Click
        Dim fname As String
        While 1
            fname = InputBox("お気に入り設定名", "現在の設定を保存しますか？")
            FILENAME_ranking = My.Application.Info.DirectoryPath & "\settings\" & fname & ".INI" 'INIファイルの保存場所

            If fname = "" Then '無名の場合は抜ける
                Exit Sub
            ElseIf File.Exists(FILENAME_ranking) Then
                Dim yn As Integer = MsgBox("既に同名設定が保存されています。上書きしますか？", vbYesNo)
                If yn = vbYes Then
                    File.Delete(FILENAME_ranking)
                    Exit While
                End If
            Else
                Exit While
            End If
        End While
        Try
            '全体設定
            SetINIValue(ToolStripComboBox1.Text, "kobo", "設定", FILENAME_ranking)
            SetINIValue(ToolStripComboBox2.Text, "heika", "設定", FILENAME_ranking)
            SetINIValue(ToolStripComboBox3.SelectedIndex, "rank", "設定", FILENAME_ranking)
            SetINIValue(ComboBox9.Text, "skilllv", "設定", FILENAME_ranking)
            SetINIValue(ComboBox10.SelectedIndex, "heicount", "設定", FILENAME_ranking)
            SetINIValue(syosklflg, "detail_skillflg", "設定", FILENAME_ranking)
            SetINIValue(statusflg, "detail_statusflg", "設定", FILENAME_ranking)
            SetINIValue(simu_skeqflg, "rankbusho_skillflg", "設定", FILENAME_ranking)
            '対象武将設定
            Dim rarearr() = {"天", "極", "特", "上", "序"}
            For i As Integer = 0 To rarearr.Length - 1
                SetINIValue(taisyo_rare(i), rarearr(i), "対象武将", FILENAME_ranking)
            Next
            '3武将設定と追加スキル
            For i As Integer = 0 To 3
                Dim bsho As String = Nothing
                Select Case i
                    Case 0
                        bsho = "武将A"
                    Case 1
                        bsho = "武将B"
                    Case 2
                        bsho = "武将C"
                    Case 3
                        bsho = "ランキング武将"
                End Select
                With simu_bs(i)
                    '追加スキルについては、シミュ実行しないと内部データに記録されていないので直取り
                    If Not simu_skeqflg Then 'ランキングスキルが他3武将と同様
                        SetINIValue(ComboBox011.Text, "add1", bsho, FILENAME_ranking)
                        SetINIValue(ComboBox012.Text, "add2", bsho, FILENAME_ranking)
                    Else
                        If i < 3 Then
                            SetINIValue(ComboBox011.Text, "add1", bsho, FILENAME_ranking)
                            SetINIValue(ComboBox012.Text, "add2", bsho, FILENAME_ranking)
                        Else
                            SetINIValue(ComboBox041.Text, "add1", bsho, FILENAME_ranking)
                            SetINIValue(ComboBox042.Text, "add2", bsho, FILENAME_ranking)
                        End If
                    End If
                    If i < 3 Then
                        SetINIValue(.name, "name", bsho, FILENAME_ranking)
                        SetINIValue(.rare, "rare", bsho, FILENAME_ranking)
                    End If
                End With
                'ステ振り詳細設定
                If statusflg Then
                    SetINIValue(cus_status(i), bsho, "ステ振り設定", FILENAME_ranking)
                End If
            Next
            'スキル詳細設定
            If syosklflg Then
                For i As Integer = 0 To 3
                    SetINIValue(cus_addskl(i)(0), "add" & CStr(i) & "1", "CSKILL", FILENAME_ranking)
                    SetINIValue(cus_addskl(i)(2), "add" & CStr(i) & "2", "CSKILL", FILENAME_ranking)
                Next
            End If
            'ステ振り詳細設定(お任せの時の設定)
            If statusflg Then
                SetINIValue(cus_status_kb, "攻防", "ステ振り設定", FILENAME_ranking)
                SetINIValue(cus_status_hei, "兵法", "ステ振り設定", FILENAME_ranking)
            End If
            MsgBox("登録完了")
        Catch ex As Exception
            MsgBox("登録内容に漏れがあります。次回復元時に正しく復元されない可能性があります。")
        End Try
    End Sub

    Private Sub 開くボタン(sender As Object, e As EventArgs) Handles ToolStripSplitButton2.ButtonClick
        Call お気に入り設定を開く(sender, e)
    End Sub

    Private Sub ステ振り詳細設定(sender As Object, e As EventArgs) Handles Button2.Click
        Form15.Show()
    End Sub
End Class