Imports System.IO
'Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Imports System.Data.SQLite
Imports Microsoft.VisualBasic.FileIO
Imports System.Text

Public Structure Busho : Implements System.ICloneable

    Public Function Clone() As Object Implements System.ICloneable.Clone
        Return Me.MemberwiseClone()
    End Function

    Public No As Integer '武将No.
    Public rare As String 'レアリティ
    Public name As String '武将名
    Public id As String '武将ID
    Public cost As Decimal 'コスト
    Public job As String '職

    Public tou_d() As Decimal '初期統率
    Public tou_d_a() As String
    Public tou() As Decimal '統率
    Public tou_a() As String

    Public Property Tousotu(Optional ByVal d As Boolean = True, Optional ByVal f As Boolean = False) As String()
        'dは現行の統率ならばTrue, 初期値ならFalse
        'fは初回のみTrue
        Get
            If d = True Then
                Return tou_a
            Else
                Return tou_d_a
            End If
        End Get
        Set(ByVal value As String()) '数値に変換して格納
            If d = True Then
                ReDim tou(3), tou_a(3)
            Else
                If f = True Then
                    ReDim tou(3), tou_a(3)
                    ReDim tou_d(3), tou_d_a(3)
                Else
                    ReDim tou_d(3), tou_d_a(3)
                End If
            End If
            For h As Integer = 0 To 3
                If d = True Then
                    tou_a(h) = value(h)
                Else
                    tou_d_a(h) = value(h)
                End If
            Next
            For h As Integer = 0 To 3
                If d = True Then
                    tou(h) = 統率_数値変換(tou_a(h))
                Else
                    tou_d(h) = 統率_数値変換(tou_d_a(h))
                End If
            Next
            If d = False And f = True Then
                tou_a = tou_d_a.Clone
                tou = tou_d.Clone '初回はtou = tou_d
            End If
        End Set
    End Property

    Public st_d() As Decimal '初期ステータス
    Public st() As Decimal 'ステータス

    Public Property Sta(Optional ByVal d As Boolean = True) As Decimal()
        Get
            If d = True Then
                Return st
            Else
                Return st_d
            End If
        End Get
        Set(ByVal value As Decimal())
            If d = True Then
                ReDim st(2)
            Else
                ReDim st_d(2)
            End If
            For i As Integer = 0 To 2
                If d = True Then
                    st(i) = value(i)
                Else
                    st_d(i) = value(i)
                End If
            Next
            If d = False Then
                st = st_d.Clone '初回時は、st = st_dとなる
            End If
        End Set
    End Property

    Public sta_g() As Decimal 'ステータス成長値

    Public Structure skl : Implements System.ICloneable
        Public name As String 'スキル名
        Public lv As Integer 'スキルLV.
        Public kanren As String '関連
        Public koubou As String '攻撃or防御スキル
        Public kouka_p As Decimal '発動率
        Public kouka_p_b As Decimal '部隊兵法補正後の発動率
        Public kouka_f As Decimal '上昇率
        Public kouka_f_b As Decimal '補正後の上昇率
        Public heika As String '兵科
        Public speed As Decimal '速度上昇率
        Public tokusyu As Integer '特殊スキル判定(0:通常, 1:速度のみ, 2:破壊のみ, 5:データ不足, 9:その他)
        Public t_flg As Boolean '総コストorフラグ存在依存
        Public up_kouka_p As Decimal '部隊内での所定条件を満たしたことで発動率UP
        Public up_kouka_f As Decimal '部隊内での所定条件を満たしたことで上昇率UP
        Public exp_kouka As Decimal '期待値
        Public exp_kouka_b As Decimal '部隊兵法補正後の期待値
        Public Function Clone() As Object Implements System.ICloneable.Clone
            Return Me.MemberwiseClone()
        End Function
        Public Sub スキル計算状態初期化() 'スキルそのもののデータではなく計算に生じて変化する部分を初期化
            kouka_p_b = 0
            kouka_f_b = 0
            exp_kouka = 0
            exp_kouka_b = 0
            up_kouka_p = 0
            up_kouka_f = 0
        End Sub

        Public Sub スキル初期化() 'スキル取得、上書きの際の初期化
            kanren = Nothing
            koubou = Nothing
            kouka_p = 0
            kouka_f = 0
            heika = Nothing
            speed = 0
            tokusyu = Nothing
            t_flg = False
        End Sub
    End Structure

    Public skill_no As Integer 'スキル数
    Public skill() As skl

    Sub 武将設定初期化() '一番最初の最低限の設定
        level = 20
        rank = 0
        rankup_r = 0
        skill_no = 1
        ReDim skill(0)
        skill(0).lv = 1
    End Sub

    Sub スキル取得(ByVal sno As Integer, ByVal sname As String, ByVal slv As Integer, ByVal fss As Integer(), Optional ByRef outlog As String = Nothing)
        'fssはその武将のスキル登録状態
        For i As Integer = 0 To skill_no - 1
            If skill.Length - 1 < i Then 'skillの要素数がiよりも小さければ
                ReDim Preserve skill(i)
            End If
            If fss(i) = sno Then
                Dim tmp() As String
                If Not skill_no = skill.Length And Not skill(i).koubou = "" Then '既にスキルが格納されている場合
                    ReDim Preserve skill(i + 1)
                    If skill(i + 1).koubou = "" Then 'その下が空欄ならば
                        skill(i + 1) = skill(i).Clone
                    End If
                End If
                With skill(i)
                    .スキル初期化()
                    .name = sname
                    .lv = slv
                    tmp = Skill_ref(.name, .lv)
                    If tmp(7) Is Nothing Then 'エントリ自体が存在しない
                        outlog = "データベースに存在しないスキル。要更新。" & "【" & .name & "LV" & .lv & "】"
                        Continue For
                    End If

                    .kanren = tmp(1)
                    .heika = tmp(2)
                    If InStr(.heika, "全") Then
                        .heika = "槍弓馬砲器"
                    End If
                    .koubou = tmp(3)
                    'データ不足の場合を除く
                    If tmp(7) = "U" Then
                        .tokusyu = 5
                        outlog = outlog & "登録情報が無いスキル・LV。結果には反映されません。" & _
                                   "【" & .name & "LV" & .lv & "】"
                        .kouka_p = 0
                        .kouka_f = 0
                        Continue For
                    End If
                    .kouka_p = Decimal.Parse(tmp(0))
                    Select Case (.kanren)
                        '特殊スキル
                        Case "特殊"
                            .tokusyu = 9
                            .kouka_p = 0
                            .kouka_f = 0
                            Continue For
                        Case "条件" ', "童"
                            .tokusyu = 9
                            .kouka_f = Val(tmp(4))
                            .t_flg = フラグ付きスキル参照(skill(i)) '条件付きスキルの場合
                            If .koubou = "速" Then '速度付きならばそこはspeedに格納しておく
                                .speed = Decimal.Parse(tmp(4)) '速度はコスト依存・・・しない・・・（現状
                            ElseIf tmp(5) = "速" Then
                                .speed = Decimal.Parse(tmp(6)) '付与効果に速度がある
                            End If
                        Case Else
                            Select Case (.koubou)
                                Case "速" '速度オンリー
                                    .tokusyu = 1
                                    .speed = Decimal.Parse(tmp(4)) '速度はコスト依存・・・しない・・・（現状
                                Case "破壊" '破壊オンリー
                                    .tokusyu = 2
                                Case Else '通常スキルの場合
                                    .tokusyu = 0
                                    If InStr(tmp(4), "C") Then tmp(4) = 文字列計算(Replace(tmp(4), "C", CStr(cost))) 'コスト依存スキルの扱い。「スキル所持武将の」コストで一括適用
                                    .kouka_f = Decimal.Parse(tmp(4))
                                    If tmp(5) = "速" Then .speed = Decimal.Parse(tmp(6)) '付与効果に速度がある
                            End Select
                    End Select
                End With
            Else
                skill(i) = skill(fss(i)).Clone
            End If
        Next
        ReDim Preserve skill(skill_no - 1)
    End Sub

    '基本効果、付加効果を入力、加速率を出力
    'Function スピードスキル取得(ByVal kihon As String, ByVal huka As String) As Decimal
    '    If InStr(kihon, "速") = 0 And InStr(huka, "速") = 0 Then '一般スキルにはゼロ
    '        Return 0
    '        Exit Function
    '    End If
    '    Dim spd As String = "0"
    '    Dim dsp As String = vbNullString
    '    Dim ss, se, sf As Integer
    '    '基本欄にあるか付加欄にあるか
    '    If InStr(kihon, "速") Then
    '        dsp = kihon
    '    ElseIf InStr(huka, "速") Then
    '        dsp = huka
    '    End If
    '    ss = InStr(dsp, "：") + 1
    '    If InStr(dsp, "上昇") Then
    '        se = InStr(dsp, "上昇") - ss
    '        sf = 1
    '    ElseIf InStr(dsp, "低下") Then
    '        se = InStr(dsp, "低下") - ss
    '        sf = -1
    '    End If
    '    spd = Mid(dsp, ss, se)
    '    If sf = -1 Then '低下スキルならば
    '        spd = "-" & spd
    '    End If
    '    '特殊な付加要素が加わったスピードスキルが増えてきたためエラーチェックFalse
    '    スピードスキル取得 = 文字列計算(spd, False)
    'End Function

    Public Structure heisu : Implements System.ICloneable
        Public name As String '兵種名
        Public bunrui As String '兵種分類
        Public tousotu As String '影響する統率
        Public ts As Decimal '実質統率値
        Public jyk_utuwa As Boolean '「上級」範囲判定
        Public jyk_hou As Boolean '「上級砲」範囲設定
        Public tok_hikyo As Boolean '「秘境兵」範囲設定
        Public pwr As Integer '現在の戦闘力（攻撃力か防御力かのどちらか）
        Public atk As Integer '攻撃力
        Public def As Integer '防御力
        Public spd As Integer '速度
        Public Function Clone() As Object Implements System.ICloneable.Clone
            Return Me.MemberwiseClone()
        End Function
    End Structure

    Public heisyu As heisu

    '以降の_kbflgは、外から攻撃/防御の判定を取ってくるかどうか
    Sub 兵科情報取得(ByVal h As String, Optional ByVal kb_ As String = "")
        If h = Nothing Then '兵科が未確定の時は抜ける必要がある
            Exit Sub
        End If
        Dim kobo_ As String
        If Not kb_ = "" Then
            kobo_ = kb_
        Else
            kobo_ = kb
        End If
        Dim s() As String = _
         DB_DirectOUT("SELECT 統率, 兵種名, 兵科, 攻撃値, 防御値, 移動値 FROM HData WHERE 兵種名=" & ダブルクオート(h) & "", {"兵科", "統率", "攻撃値", "防御値", "移動値"})
        With heisyu
            .name = h
            .bunrui = s(0)
            .tousotu = s(1)
            .atk = Val(s(2))
            .def = Val(s(3))
            .spd = Val(s(4))
            .ts = 0 '初期化
            If InStr(kobo_, "攻撃") Then
                .pwr = .atk
            Else
                .pwr = .def
            End If
            Dim t As Integer = .tousotu.Length '統率に関係する兵科の種類数
            Dim v() As Decimal
            ReDim v(t - 1)
            For i As Integer = 1 To t
                Select Case Mid(.tousotu, i, 1)
                    Case "槍"
                        v(i - 1) = tou(0)
                    Case "弓"
                        v(i - 1) = tou(1)
                    Case "馬"
                        v(i - 1) = tou(2)
                    Case "器"
                        v(i - 1) = tou(3)
                End Select
                .ts = .ts + v(i - 1)
            Next
            .ts = .ts / t '該当兵科の実質統率値

            Dim target_u() As String = {"武士", "弓騎馬", "赤備え", "騎馬鉄砲", "鉄砲足軽", "焙烙火矢", "大筒兵", "破城鎚", "攻城櫓"}
            Dim target_h() As String = {"武士", "弓騎馬", "赤備え", "騎馬鉄砲", "鉄砲足軽", "焙烙火矢", "大筒兵", "雑賀衆"}
            Dim target_hi() As String = {"国人衆", "雑賀衆", "海賊衆", "母衣衆"}
            '上級器判定
            .jyk_utuwa = False
            For i As Integer = 0 To target_u.Length - 1
                If InStr(.name, target_u(i)) Then
                    .jyk_utuwa = True
                    Exit For
                End If
            Next
            '上級砲判定
            .jyk_hou = False
            For i As Integer = 0 To target_h.Length - 1
                If InStr(.name, target_h(i)) Then
                    .jyk_hou = True
                    Exit For
                End If
            Next
            '秘境兵判定
            .tok_hikyo = False
            For i As Integer = 0 To target_hi.Length - 1
                If InStr(.name, target_hi(i)) Then
                    .tok_hikyo = True
                    Exit For
                End If
            Next
        End With
    End Sub

    Public rank As Integer 'ランク
    Public rankup_r As Integer 'ランクアップ残り可能回数
    Public limitbreakflg As Boolean '限界突破

    Sub 残りランクアップ可能回数表示(ByVal g As GroupBox)
        Dim fugou As String
        If rankup_r > 0 Then
            fugou = "+"
            g.ForeColor = Color.Red
        ElseIf rankup_r = 0 Then
            fugou = ""
            g.ForeColor = Color.Black
            g.Text = "統率"
            Exit Sub
        Else
            fugou = ""
            g.ForeColor = Color.Blue
        End If
        g.Text = "統率 " & "( " & fugou & rankup_r & " )"
    End Sub

    Public level As Integer 'レベル
    Public huri As String '極振り設定

    Sub ステ極振り計算()
        Dim stpt As Integer 'ステ振りポイントの合計
        If rank < 2 Then
            stpt = (rank * 20 + level) * 4
        ElseIf rank < 4 Then
            stpt = 160 + ((rank - 2) * 20 + level) * 5
        ElseIf rank < 6 Then
            stpt = 360 + ((rank - 4) * 20 + level) * 6
        ElseIf rank = 6 Then 'つまりは限界突破時
            stpt = 630
        End If

        Select Case huri
            Case "攻撃極振り"
                Sta(True)(0) = Sta(False)(0) + stpt * sta_g(0)
                Sta(True)(1) = Sta(False)(1)
                Sta(True)(2) = Sta(False)(2)
            Case "防御極振り"
                Sta(True)(0) = Sta(False)(0)
                Sta(True)(1) = Sta(False)(1) + stpt * sta_g(1)
                Sta(True)(2) = Sta(False)(2)
            Case "兵法極振り"
                Sta(True)(0) = Sta(False)(0)
                Sta(True)(1) = Sta(False)(1)
                Sta(True)(2) = Sta(False)(2) + stpt * sta_g(2)
        End Select
    End Sub

    Public hei_max As Integer '最大積載兵数
    Public hei_max_d As Integer '最大積載兵数(☆0)
    Public hei_sum As Integer '積載兵数

    Public attack As Decimal '小隊攻
    Public higai As Integer '被害兵数

    Sub 情報入力(ByVal bc As Integer, Optional ByVal fs As Boolean = True) 'fsがFalseなのは初回の情報入力のみ
        Button(Form1, CStr(bc) & "00" & "1").Text = "/ " & hei_max
        Label(Form1, CStr(bc) & "00" & "2").Text = skill(0).name
        TextBox(Form1, CStr(bc) & "01").Text = hei_sum
        TextBox(Form1, CStr(bc) & "02").Text = level
        For k As Integer = 3 To 5
            TextBox(Form1, CStr(bc) & "0" & CStr(k)).Text = Sta(fs)(k - 3)
        Next
        '攻防兵成長
        For k As Integer = 3 To 5
            Label(Form1, CStr(bc) & "0" & CStr(k)).Text = "(+" & sta_g(k - 3).ToString & ")"
        Next
        '職
        GroupBox(Form1, "0" & CStr(bc) & "2").Text = "ステータス[ " & job & " ]"
        'スキルレベルを埋める段階でremovehandlerしていないと二度埋めになる
        Dim cc As ComboBox = ComboBox(Form1, CStr(bc) & "14")
        RemoveHandler cc.SelectedValueChanged, AddressOf Form1.追加スキル追加
        cc.Text = skill(0).lv
        AddHandler cc.SelectedValueChanged, AddressOf Form1.追加スキル追加
        ComboBox(Form1, CStr(bc) & "04").Text = rank
        For k As Integer = 5 To 8
            ComboBox(Form1, CStr(bc) & "0" & CStr(k)).Text = Tousotu(fs)(k - 5)
        Next
    End Sub

    Sub 小隊攻撃力計算(Optional ByVal kb_ As String = "")
        If heisyu.ts = Nothing Or hei_sum = Nothing Then
            MsgBox("必要なデータが不足しています")
            Exit Sub
        End If
        Dim kobo_ As String
        If Not kb_ = "" Then
            kobo_ = kb_
        Else
            kobo_ = kb
        End If
        Dim ad As Decimal
        If InStr(kobo_, "攻撃") Then
            ad = st(0)
        Else
            ad = st(1)
        End If
        attack = (ad + hei_sum * heisyu.pwr) * heisyu.ts
    End Sub

    '外から部隊兵法値を取ってくるかどうかがheihou_, rank_, cost_
    Sub スキル期待値計算(Optional ByVal heihou_ As Decimal = -1, Optional ByVal rank_ As Decimal = -1, Optional ByVal cost_ As Decimal = -1) '部隊兵法値、ランク合計値が既知であることが必要
        Dim heihou_sum_ As Decimal
        Dim rank_sum_ As Decimal
        Dim cost_sum_ As Decimal
        If heihou_ = -1 Then
            heihou_sum_ = heihou_sum
        Else
            heihou_sum_ = heihou_
        End If
        If rank_ = -1 Then
            rank_sum_ = rank_sum
        Else
            rank_sum_ = rank_
        End If
        If cost_ = -1 Then
            cost_sum_ = Costsum
        Else
            cost_sum_ = cost_
        End If
        For j As Integer = 0 To skill_no - 1
            With skill(j)
                If .tokusyu = 0 Then '通常スキルならば
                    .t_flg = False
                    If .kouka_p = 1 Then
                        .kouka_p_b = 1
                    Else
                        .kouka_p_b = .kouka_p + 0.01 * (heihou_sum_ + rank_sum_)
                    End If
                    .kouka_f_b = .kouka_f
                    .exp_kouka = .kouka_p * .kouka_f
                    .exp_kouka_b = .kouka_p_b * .kouka_f_b '期待値
                Else
                    If .tokusyu = 9 Then '特殊スキルの場合は・・・
                        .t_flg = 条件依存スキル・フラグスキル判定(skill(j), cost_sum_, rank_sum_) '怪しいスキルを疑う
                        .t_flg = フラグ付きスキル参照(skill(j))
                        If Not .kouka_f = 0 Then 'ゼロならば単なる特殊スキル
                            If .kouka_p = 1 Then
                                .kouka_p_b = 1
                            Else
                                .kouka_p_b = .kouka_p + 0.01 * (heihou_sum_ + rank_sum_)
                            End If
                            .kouka_f_b = .kouka_f
                            .exp_kouka = .kouka_p * .kouka_f
                            .exp_kouka_b = .kouka_p_b * .kouka_f_b '期待値
                        Else
                            .exp_kouka = 0
                            .exp_kouka_b = 0
                        End If
                    End If
                End If
            End With
        Next
    End Sub
End Structure

Public Structure bskl '部隊スキル関連
    Public flg As Boolean '有効かどうか
    Public bsk() As _bsk
    Public activebsk() As _bsk
    Public Structure _bsk : Implements System.ICloneable
        Public koubou As String '攻防
        Public type As String 'タイプ（攻、防、速）
        'Public name As String '部隊スキル名
        Public kouka_p As Decimal '発動率
        Public kouka_f As Decimal '上昇率
        Public speed As Decimal 'スピード上昇のある場合、加速率
        Public taisyo As String '対象
        'Public qq As Boolean '将攻二乗モードONOFF
        Public Function Clone() As Object Implements System.ICloneable.Clone
            Return Me.MemberwiseClone()
        End Function
    End Structure
    Public ReadOnly Property speed() As Decimal
        Get
            If bsk Is Nothing Then Return 0
            Dim spdtmp As Decimal = 0
            For i As Integer = 0 To bsk.Length - 1
                If spdtmp < bsk(i).speed Then
                    spdtmp = bsk(i).speed
                End If
            Next
            Return spdtmp
        End Get
    End Property 'スピードスキルは複数あった場合でも最速のもの
    '有効な部隊スキルリストを返す
    Public ReadOnly Property activeskl(ByVal kb As String) As _bsk()
        Get
            If bsk Is Nothing Then Return Nothing
            Dim validcount As Decimal = 0
            For i As Integer = 0 To bsk.Length - 1
                If InStr(kb, bsk(i).koubou) And bsk(i).kouka_f > 0 Then
                    If validcount = 0 Then
                        ReDim activebsk(0)
                    Else
                        ReDim Preserve activebsk(validcount)
                    End If
                    activebsk(validcount) = bsk(i).Clone
                    validcount = validcount + 1
                End If
            Next
            Return activebsk
        End Get
    End Property
End Structure

Public Structure flgskl : Implements System.ICloneable 'フラグスキル格納
    Public id As Integer 'id
    Public name As String 'スキル名
    Public lv As Integer 'スキルLV.
    Public koubou As String '攻撃or防御スキル
    Public kouka_p As Decimal '発動率
    Public kouka_p_b As Decimal '部隊兵法補正後の発動率
    Public kouka_f As Decimal '上昇率
    Public heika As String '兵科
    Public speed As Decimal '加速率    
    Public onoff As Boolean 'ONOFF
    Public wflg As Boolean '童フラグ
    Public Function Clone() As Object Implements System.ICloneable.Clone
        Return Me.MemberwiseClone()
    End Function
End Structure

Public Structure confskl : Implements System.ICloneable '置換パラメータ格納
    Public id As Integer 'id
    Public param_name As String 'パラメータ名
    Public param_value As String 'パラメータ設定値
    Public Function Clone() As Object Implements ICloneable.Clone
        Return Me.MemberwiseClone()
    End Function
End Structure

Public Structure _warabe '童情報格納
    Public Structure _atk
        Public yari As Decimal
        Public yumi As Decimal
        Public uma As Decimal
        Public hou As Decimal
        Public utuwa As Decimal
    End Structure
    Public atk As _atk
    Public Structure _def
        Public yari As Decimal
        Public yumi As Decimal
        Public uma As Decimal
        Public hou As Decimal
        Public utuwa As Decimal
    End Structure
    Public def As _def
    Public Structure _speed
        Public yari As Decimal
        Public yumi As Decimal
        Public uma As Decimal
        Public hou As Decimal
        Public utuwa As Decimal
    End Structure
    Public speed As _speed
    Public Sub warabe_clean()
        With atk
            .yari = 0
            .yumi = 0
            .uma = 0
            .hou = 0
            .utuwa = 0
        End With
        With def
            .yari = 0
            .yumi = 0
            .uma = 0
            .hou = 0
            .utuwa = 0
        End With
        With speed
            .yari = 0
            .yumi = 0
            .uma = 0
            .hou = 0
            .utuwa = 0
        End With
    End Sub
    Public Sub warabe_itemset(ByVal kb As String, ByVal t As String, ByVal value As Decimal)
        If InStr(kb, "攻") Then
            Select Case t
                Case "槍"
                    warabe.atk.yari += value
                Case "弓"
                    warabe.atk.yumi += value
                Case "馬"
                    warabe.atk.uma += value
                Case "砲"
                    warabe.atk.hou += value
                Case "器"
                    warabe.atk.utuwa += value
            End Select
        ElseIf InStr(kb, "防") Then
            Select Case t
                Case "槍"
                    warabe.def.yari += value
                Case "弓"
                    warabe.def.yumi += value
                Case "馬"
                    warabe.def.uma += value
                Case "砲"
                    warabe.def.hou += value
                Case "器"
                    warabe.def.utuwa += value
            End Select
        ElseIf InStr(kb, "速") Then
            Select Case t
                Case "槍"
                    warabe.speed.yari += value
                Case "弓"
                    warabe.speed.yumi += value
                Case "馬"
                    warabe.speed.uma += value
                Case "砲"
                    warabe.speed.hou += value
                Case "器"
                    warabe.speed.utuwa += value
            End Select
        End If
    End Sub
    Public Sub warabe_set(ByVal kb As String, ByVal heika As String, ByVal value As Decimal)
        Dim hlen As Integer = heika.Length
        Dim tstr As String = Nothing
        For i As Integer = 0 To hlen - 1
            tstr = Mid(heika, i + 1, 1)
            warabe_itemset(kb, tstr, value)
        Next
    End Sub
    Public Function warabe_get(ByVal kb As String, ByVal t As String) As Decimal
        If InStr(kb, "攻") Then
            Select Case t
                Case "槍"
                    Return warabe.atk.yari
                Case "弓"
                    Return warabe.atk.yumi
                Case "馬"
                    Return warabe.atk.uma
                Case "砲"
                    Return warabe.atk.hou
                Case "器"
                    Return warabe.atk.utuwa
            End Select
        ElseIf InStr(kb, "防") Then
            Select Case t
                Case "槍"
                    Return warabe.def.yari
                Case "弓"
                    Return warabe.def.yumi
                Case "馬"
                    Return warabe.def.uma
                Case "砲"
                    Return warabe.def.hou
                Case "器"
                    Return warabe.def.utuwa
            End Select
        ElseIf InStr(kb, "速") Then
            Select Case t
                Case "槍"
                    Return warabe.speed.yari
                Case "弓"
                    Return warabe.speed.yumi
                Case "馬"
                    Return warabe.speed.uma
                Case "砲"
                    Return warabe.speed.hou
                Case "器"
                    Return warabe.speed.utuwa
            End Select
        End If
        Return 0
    End Function
    Public Function warabe_gets(ByVal kb As String) As Decimal()
        If InStr(kb, "攻") Then
            Return {warabe.atk.yari, warabe.atk.yumi, warabe.atk.uma, warabe.atk.hou, warabe.atk.utuwa}
        ElseIf InStr(kb, "防") Then
            Return {warabe.def.yari, warabe.def.yumi, warabe.def.uma, warabe.def.hou, warabe.def.utuwa}
        ElseIf InStr(kb, "速") Then
            Return {warabe.speed.yari, warabe.speed.yumi, warabe.speed.uma, warabe.speed.hou, warabe.speed.utuwa}
        End If
        Return {0, 0, 0, 0, 0}
    End Function
End Structure

Module Module1
    'INIファイル編集のためにWindows-API使用
    <DllImport("KERNEL32.DLL", CharSet:=CharSet.Auto)> _
    Public Function WritePrivateProfileString( _
        ByVal lpAppName As String, _
        ByVal lpKeyName As String, _
        ByVal lpString As String, _
        ByVal lpFileName As String) As Boolean
    End Function
    <DllImport("KERNEL32.DLL", CharSet:=CharSet.Auto)> _
    Public Function GetPrivateProfileString( _
        ByVal lpAppName As String, _
        ByVal lpKeyName As String, _
        ByVal lpDefault As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Integer, _
        ByVal lpFileName As String) As Integer
    End Function
    <DllImport("KERNEL32.DLL", CharSet:=CharSet.Auto)> _
    Public Function GetPrivateProfileSection( _
        ByVal lpAppName As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Integer, _
        ByVal lpFileName As String) As Integer
    End Function
    '*** 画面再描画を一時停止するために用いる ***
    <DllImport("user32.dll", CharSet:=CharSet.Auto)> _
    Public Function SendMessage(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As IntPtr
    End Function
    Public Const WM_SETREDRAW As Integer = &HB
    Public Const Win32False As Integer = 0
    Public Const Win32True As Integer = 1
    '********************************************

    ' 描画を一時停止
    Public Sub 描画一時停止(ByVal frm As Form)
        SendMessage(frm.Handle, WM_SETREDRAW, Win32False, 0)
    End Sub

    ' 描画を再開
    Public Sub 描画再開(ByVal frm As Form)
        SendMessage(frm.Handle, WM_SETREDRAW, Win32True, 0)
        frm.Invalidate()
    End Sub

    '*** 一般変数 ***
    Public kb As String '攻撃/防衛部隊どちらか
    Public busho_counter As Integer = 0 '武将数
    Public bs() As Busho '武将

    Public Heisum As Long '総兵数
    Public heihou_sum As Decimal '兵法補正
    Public Atksum As Decimal '総戦闘力
    Public Costsum As Decimal '総コスト
    Public Ranksum As Decimal '部隊ランクボーナス適用値
    Public rank_sum As Decimal '部隊ランクボーナス値

    Public zenmetuHp() As String '兵0になるHPの値
    Public atksum_max As Decimal 'MAX値
    Public atksum_maxmax As Decimal '部隊スキルが有効な場合のMAX値

    Public skill_x() As Decimal '確率変数x → xの生起確率
    Public skill_xx() As Decimal '部隊スキルが有効な時の、部隊スキル抜きのx
    Public skill_y(,) As Decimal '確率変数x → 総戦闘力f(x)(k) k:k番目の武将の戦闘力
    Public skill_yk() As Decimal 'skill_yk(i) = SUM{skill_y(i)(k)}
    Public skill_yy(,) As Decimal '部隊スキルが有効な時の、部隊スキル抜きのy
    Public skill_yyk() As Decimal
    Public skill_syo(,) As Decimal '各スキル発動パターンにおける、将攻成分（部隊スキルの時に使う）
    Public skill_ex() As Decimal '確率変数x → 期待値E(x) k:k番目の武将の戦闘力の期待値
    Public skill_exk As Decimal 'skill_exk = SUM{skill_ex(k)}
    Public skill_exx As Decimal = 0
    Public skill_ax As Decimal = 0 '確率変数x → 標準偏差σ(x)
    Public skill_axx As Decimal = 0
    Public can_skill() As Busho.skl '発動して意味のあるスキル一覧 '0:一般, 1:将
    Public warabe As _warabe '童効果格納
    '*** 部隊スキル関係 ***
    Public bskill As bskl '武将スキル
    '*** DB関連の変数 ***
    'Public con, con2, con3 As New OleDbConnection 'DB接続設定に必要な変数 1:武将DB, 2:スキルDB, 3:NPC空き地情報
    Public Connection As New SQLiteConnection
    Public Command As New SQLiteCommand
    'Public Command_sklref As SQLite.SQLiteCommand
    'Public cmd, cmd2, cmd3 As New OleDbCommand
    'Public dbpath As String = Application.StartupPath & "\Busho.mdb" '実行ファイルのある階層に武将DBは置く
    'Public dbpath2 As String = Application.StartupPath & "\Skill.mdb" '同じく、スキルDB
    'Public dbpath3 As String = Application.StartupPath & "\npc.mdb"

    'Public bnpath As String = Application.StartupPath & "\BName2.INI" '同名武将区別ファイル
    'Public espath As String = Application.StartupPath & "\ERRORSKILL.txt" '特殊スキルリストの場所
    Public fdpath As String = Application.StartupPath & "\optionskill.csv" 'フラグ付きスキルリストの場所
    Public fspath As String = Application.StartupPath & "\defoption.csv" 'フラグ付きスキル設定ファイルの場所
    Public fcpath As String = Application.StartupPath & "\optionparam.csv" '置換パラメータ設定ファイルの場所
    Public dlpath As String = Application.StartupPath & "\DBEXEC_LOG.txt" 'データベース上で実行されたSQL履歴
    'Public error_skill() As String '特殊スキルリスト
    Public fskill_data() As flgskl 'フラグ付きスキルデータリスト
    Public cparam_data() As confskl '置換パラメータ設定値リスト
    Public simu_execno As Integer 'どこから計算しているのかを格納 0:シミュレータ本体, 1:ランキングモード
    '*** INIファイルの場所 ***
    Public FILENAME_bs As String
    Public FILENAME_ranking As String
    '*** CSVファイルの場所 ***
    Public FILENAME_csv As String

    'Public Sub DB_Open(ByVal con As OleDbConnection, ByVal cmd As OleDbCommand, ByVal dbpath As String) '開始時にDBを開く設定
    Public Sub DB_Open() '開始時にDBを開く設定(外部キー制約ON)
        Connection.ConnectionString = "Version=3;Data Source=ixadb.db3;New=False;Compress=True;foreign keys=True" 'Read Only=True;"
        Connection.Open()
        Command = Connection.CreateCommand
        'Command_sklref = Connection.CreateCommand
        'Command_sklref.CommandText = "SELECT * FROM SData INNER JOIN SName ON SData.スキル名 = SName.スキル名 WHERE SData.スキル名 = @param1 AND スキルLV = @param2"
        'Command_sklref.CommandType = CommandType.Text
        ''パスワードを変更
        'Connection.ChangePassword("password")
    End Sub

    Public Function ダブルクオート(ByVal Tstr As String) As String
        Return ("""" & Tstr & """")
    End Function


    Public Function TrimJ(ByVal str As String) As String
        Dim ret As String = Trim(str)
        Return (Replace(Replace(ret, " ", ""), "　", ""))
    End Function

    'データベースからSQL文を用いて出力(DataSet出力)
    'Public Function DB_TableOUT(ByVal con As OleDbConnection, ByVal cmd As OleDbCommand, ByVal sql_str As String, ByVal outTable As String) As DataSet
    Public Function DB_TableOUT(ByVal sql_str As String, ByVal outTable As String) As DataSet
        'Dim da As New SQLiteDataAdapter 'OleDbDataAdapter
        Dim ds As DataSet = New DataSet()
        Dim da As SQLiteDataAdapter = New SQLiteDataAdapter(sql_str, Connection)
        da.Fill(ds, outTable)
        Return ds
    End Function

    '(直接出力)
    'Public Function DB_DirectOUT(ByVal con As OleDbConnection, ByVal cmd As OleDbCommand, ByVal sql_str As String, ByVal outlist As String()) As String()
    Public Function DB_DirectOUT(ByVal sql_str As String, ByVal outlist As String()) As String()
        'Dim Command As New SQLiteCommand
        Dim dr As SQLiteDataReader 'OleDbDataReader
        Dim output() As String
        Dim l As Integer = outlist.Length
        ReDim output(l - 1)
        'Command = Connection.CreateCommand
        Command.CommandText = sql_str
        dr = Command.ExecuteReader()
        'Command.Dispose()
        While dr.Read()
            For i As Integer = 0 To l - 1
                If TypeOf dr(outlist(i)) Is DBNull Then 'DBが空欄
                    output(i) = ""
                Else
                    Dim getstr = CStr(dr(outlist(i)))
                    If InStr(getstr, "%%") Then '置換パラメータが埋め込まれている場合は置換
                        output(i) = パラメータ置換(getstr)
                    Else
                        output(i) = getstr
                    End If
                End If
            Next
        End While
        dr.Close()
        Return output
    End Function

    '基本はDB_DirectOUTと同じだが、こちらはリスト形式（縦向き配列）で出力する必要のある場合用いる
    'Public Function DB_DirectOUT2(ByVal con As OleDbConnection, ByVal cmd As OleDbCommand, ByVal sql_str As String, ByVal outlist As String) As String()
    Public Function DB_DirectOUT2(ByVal sql_str As String, ByVal outlist As String) As String()
        Dim dr As SQLiteDataReader 'OleDbDataReader
        Dim output() As String = Nothing
        Command.CommandText = sql_str
        dr = Command.ExecuteReader()
        Dim c As Integer = 0
        While dr.Read()
            ReDim Preserve output(c)
            If TypeOf dr(outlist) Is DBNull Then 'DBが空欄
                output(c) = ""
            Else
                Dim getstr = CStr(dr(outlist))
                If InStr(getstr, "%%") Then '置換パラメータが埋め込まれている場合は置換
                    output(c) = パラメータ置換(getstr)
                Else
                    output(c) = getstr
                End If
            End If
            c = c + 1
        End While
        dr.Close()
        Return output
    End Function

    '基本は上の二つのDB_DirectOUTと同じだが、これは二次元配列で一気に表形式で出力する場合に用いる
    'Public Function DB_DirectOUT3(ByVal con As OleDbConnection, ByVal cmd As OleDbCommand, ByVal sql_str As String, ByVal outlist() As String) As String()()
    Public Function DB_DirectOUT3(ByVal sql_str As String, ByVal outlist() As String) As String()()
        Dim dr As SQLiteDataReader 'OleDbDataReader
        Dim output()() As String = Nothing
        Dim l As Integer = outlist.Length
        ReDim output(l - 1)
        Command.CommandText = sql_str
        dr = Command.ExecuteReader()
        Dim d As Integer = 0
        While dr.Read()
            Dim outputtmp() As String = Nothing
            For i As Integer = 0 To l - 1
                ReDim Preserve outputtmp(i)
                If TypeOf dr(outlist(i)) Is DBNull Then 'DBが空欄
                    outputtmp(i) = ""
                Else
                    Dim getstr = CStr(dr(outlist(i)))
                    If InStr(getstr, "%%") Then '置換パラメータが埋め込まれている場合は置換
                        outputtmp(i) = パラメータ置換(getstr)
                    Else
                        outputtmp(i) = getstr
                    End If
                End If
            Next
            ReDim Preserve output(d)
            output(d) = outputtmp
            d = d + 1
        End While
        dr.Close()
        Return output
    End Function

    'RichtextBoxにおいて、与えた文字列配列の最初に現れた文字を太字にする
    Public Sub RTextBox_BOLD(ByVal richtextbox As RichTextBox, ByVal str() As String)
        Dim f As Integer = -1
        For i As Integer = 0 To str.Length - 1
            With richtextbox
                f = InStr(.Text, str(i))
                If f <> 0 Then
                    If Not Mid(.Text, f, str(i).Length) = str(i) Then
                        Continue For
                    End If
                    .SelectionStart = f - 1
                    .SelectionLength = str(i).Length
                    .SelectionFont = New Font(.SelectionFont, FontStyle.Bold)
                End If
            End With
        Next
    End Sub

    '指定小数点以下を切り捨てる
    Public Function ToRoundDown(ByVal dValue As Decimal, ByVal iDigits As Integer) As Double
        Dim dCoef As Double = System.Math.Pow(10, iDigits)

        If dValue > 0 Then
            Return System.Math.Floor(dValue * dCoef) / dCoef
        Else
            Return System.Math.Ceiling(dValue * dCoef) / dCoef
        End If
    End Function

    'スキルを検索
    Public Function Skill_ref(ByVal skillName As String, ByVal skillLv As Integer) As String()
        Dim s As String = "SELECT * FROM SData INNER JOIN SName ON SData.スキル名 = SName.スキル名 WHERE SData.スキル名 = " & ダブルクオート(skillName) & " AND スキルLV = " & skillLv & ""
        Dim t() As String = _
        DB_DirectOUT(s, {"発動率", "分類", "対象", "攻防", "上昇率", "付与効果", "付与率", "Sunf"})
        'Dim t() As String = {"発動率", "分類", "対象", "攻防", "上昇率", "付与効果", "付与率", "Sunf"}
        'Dim dr As SQLiteDataReader 'OleDbDataReader
        'Dim output() As String
        'ReDim output(t.Length - 1)
        'Dim sp1 As SQLite.SQLiteParameter = New SQLite.SQLiteParameter("@param1", skillName)
        'Dim sp2 As SQLite.SQLiteParameter = New SQLite.SQLiteParameter("@param2", CStr(skillLv))
        'Command_sklref.Parameters.Add(sp1)
        'Command_sklref.Parameters.Add(sp2)
        'dr = Command_sklref.ExecuteReader()
        ''Command.Dispose()
        'While dr.Read()
        '    For i As Integer = 0 To t.Length - 1
        '        If TypeOf dr(t(i)) Is DBNull Then 'DBが空欄"
        '            output(i) = ""
        '        Else
        '            output(i) = CStr(dr(t(i)))
        '        End If
        '    Next
        'End While
        'dr.Close()
        'Return output
        Return t
    End Function

    Public Function Skill_ref_list(ByVal skillName() As String, ByVal skillLv As Integer) As String()()
        Dim skillstr As String = ""
        For i As Integer = 0 To skillName.Length - 1
            skillstr = skillstr & """" & skillName(i) & """" & ","
        Next
        skillstr = skillstr.Remove(skillstr.Length - 1, 1)
        Dim s As String = "SELECT * FROM SData INNER JOIN SName ON SData.スキル名 = SName.スキル名 WHERE スキルLV = " & skillLv & " AND SData.スキル名 IN(" & skillstr & ")" & ""
        s.Remove(s.Length - 1, 1)
        Dim t()() As String = _
        DB_DirectOUT3(s, {"スキル名", "発動率", "分類", "対象", "攻防", "上昇率", "付与効果", "付与率", "Sunf"})
        Return t
    End Function

    '統率変換
    Public Function 統率_数値変換(ByVal alpha As String) As Decimal
        Dim rankstr() As String = {"SSS", "SS", ".S", "A", "B", "C", "D", "E", "F"}
        Dim ret As Decimal
        For i As Integer = 0 To rankstr.Length - 1
            If InStr(alpha, rankstr(i)) Then
                ret = 1.2 - 0.05 * i
                Exit For
            End If
        Next
        Return ret
    End Function
    Public Function 数値_統率変換(ByVal num As Decimal) As String
        Dim rankstr() As String = {"SSS", "SS", ".S", "A", "B", "C", "D", "E", "F"}
        Dim ri As Integer = (1.2 - num) / 0.05
        Return rankstr(ri)
    End Function

    '統率未振り数逆推定
    Public Function 統率未振り数推定(ByVal tou_d() As String, ByVal tou() As String, ByVal rank As Integer) As Integer
        Dim stm As Decimal = 0
        For i As Integer = 0 To 3
            stm = stm + ((統率_数値変換(tou(i)) - 統率_数値変換(tou_d(i))) / 0.05)
        Next
        Return (rank - stm)
    End Function

    'ステ振り逆推定
    'compflgをTrueにすると、ステ振り数と実際に振られているポイント数が一致しない時に「中途半端」と返す
    Public Function ステ振り推定(ByVal st_d() As Decimal, ByVal st() As Decimal, ByVal st_g() As Decimal, ByVal rank As Integer, ByVal lv As Integer, Optional ByVal compflg As Boolean = False) As String
        Dim stpt As Integer 'ステ振りポイントの合計
        Dim ret_str As String '戻り値
        If rank < 2 Then
            stpt = (rank * 20 + lv) * 4
        ElseIf rank < 4 Then
            stpt = 160 + ((rank - 2) * 20 + lv) * 5
        ElseIf rank < 6 Then
            stpt = 360 + ((rank - 4) * 20 + lv) * 6
        ElseIf rank = 6 Then 'つまりは限界突破時
            stpt = 630
        End If

        Dim st_f(2) As Integer '推定されるステ量
        For i As Integer = 0 To 2
            st_f(i) = st_d(i) + stpt * st_g(i)
        Next
        If st(0) = st_f(0) Then
            ret_str = "攻撃極振り"
        ElseIf st(1) = st_f(1) Then
            ret_str = "防御極振り"
        ElseIf st(2) = st_f(2) Then
            ret_str = "兵法極振り"
        Else
            ret_str = "手動"
        End If
        If ret_str = "手動" And compflg = True Then
            Dim stsum As Integer = 0 '実際に振られているステ数
            For i As Integer = 0 To 2
                stsum = stsum + (st(i) - st_d(i)) / st_g(i)
            Next
            If Not stsum = stpt Then '一致すればステポイントを使い切っての手動振り確定
                ret_str = "中途半端"
            End If
        End If
        Return ret_str
    End Function

    'スキル関連推定
    'cflg = Trueならば、combobox関連からの呼び出しとして条件付きスキルを特殊として返す
    Public Function スキル関連推定(ByVal skill_name As String, Optional ByVal cflg As Boolean = False) As String
        Dim stmp() As String = DB_DirectOUT("SELECT * FROM SData INNER JOIN SName ON SData.スキル名 = SName.スキル名 WHERE SData.スキル名 = " & ダブルクオート(skill_name) & " AND スキルLV = 1", {"分類"})
        If cflg And stmp(0) = "条件" Then stmp(0) = "特殊" '条件付きスキルは特殊カテゴリーに含む
        スキル関連推定 = stmp(0)
    End Function

    '部隊条件に依存したスキルの扱い（アドホック対応※）
    'データベースから静的に参照するだけではどうしようもないスキルはココでデータ確定
    Public Function 条件依存スキル・フラグスキル判定(ByRef sk As Busho.skl, ByVal sumcost As Decimal, ByVal sumrank As Decimal) As Boolean
        'Dim sdata() As String
        With sk
            Select Case (.name)
                Case "覇王征軍"
                    '覇王征軍
                    '覇王征軍の増分データ
                    Dim hs_f() As Decimal = {10, 11, 12, 13, 15, 17, 19, 21, 23, 25}
                    Dim hs_d() As Decimal = {1, 1, 1, 2, 3, 4, 5, 7, 9, 11, 13, 16, 20, 25}
                    'sdata = Skill_ref(.name, .lv) '特殊スキルなのでここで改めて取得
                    If sumcost > 9 Then
                        .kouka_f = 0.01 * (hs_f(.lv - 1) + hs_d(((sumcost - 9) / 0.5) - 1))
                    Else
                        .kouka_f = 0.01 * hs_f(.lv - 1)
                    End If
                    .heika = "槍弓馬砲器"
                    Return True
                Case "武神八幡陣"
                    '武神八幡陣
                    '武神八幡陣の増分データ
                    Dim bh_f() As Decimal = {10, 12, 14, 17, 20, 24, 28, 32, 36, 40}
                    Dim bh_d() As Decimal = {1, 1, 1, 2, 3, 3, 4, 5, 6, 7, 9, 10, 12, 14, 16, 19, 22, 25}
                    'sdata = Skill_ref(.name, .lv) '特殊スキルなのでここで改めて取得
                    If sumcost > 7 Then
                        .kouka_f = 0.01 * (bh_f(.lv - 1) + bh_d(((sumcost - 7) / 0.5) - 1))
                    Else
                        .kouka_f = 0.01 * bh_f(.lv - 1)
                    End If
                    .heika = "槍弓馬砲器"
                    Return True
                Case "妙見の矜持"
                    '妙見の矜持
                    '部隊総コストが0.5増える毎に、威力-2% → コメントを頂いて修正。
                    '妙見の矜持の減分データ(0.5C-12.5C)
                    Dim mh_d() As Decimal = {0, 2, 4, 6, 8, 10, 12, 15, 17, 19, 21, 23, 25, 27, 30, 32, 34, 36, 38, 40, _
                                             42, 45, 47, 49, 51}
                    Dim mh_f() As Decimal = {57, 58, 59, 60, 61, 62, 63, 64, 65, 67}
                    Dim minus_f As Decimal = mh_d((sumcost - 0.5) * 2) * 0.01
                    .kouka_f = 0.01 * mh_f(.lv - 1) - minus_f
                    .heika = "槍弓馬砲器"
                    Return True
                Case "遁世影武者"
                    '遁世影武者
                    '部隊長の初期スキルbs(0).skill(0)もしくはForm10.simu_bs(0).skill(0)と同等に
                    '※参照順番依存があるので変更・デバッグは慎重に
                    'sdata = Skill_ref(.name, .lv) '特殊スキルなのでここで改めて取得
                    Dim refskl As Busho.skl = Nothing
                    If simu_execno = 0 Then
                        refskl = bs(0).skill(0).Clone
                    ElseIf simu_execno = 1 Then
                        refskl = Form10.simu_bs(0).skill(0).Clone
                    End If
                    If refskl.name = "遁世影武者" Then '部隊長の初期スキルが自分自身なら無効
                        Return False
                    End If
                    .heika = refskl.heika
                    .koubou = refskl.koubou
                    .kouka_f = refskl.kouka_f
                    Return True
                Case "花魁心操術"
                    '花魁心操術
                    '自分以外の武将に姫がいる場合、そのスキル発動率を操作
                    '※参照順番依存があるので変更・デバッグは慎重に
                    Dim refbs() As Busho = Nothing
                    If simu_execno = 0 Then
                        refbs = bs.Clone
                    ElseIf simu_execno = 1 Then
                        refbs = Form10.simu_bs.Clone
                    End If
                    For i As Integer = 0 To refbs.Length - 1
                        If refbs(i).job = "姫" Then
                            For j As Integer = 0 To refbs(i).skill.Length - 1
                                refbs(i).skill(j).up_kouka_p = refbs(i).skill(j).up_kouka_p + 0.01 * .lv 'LVの分だけ発動率上昇
                            Next
                        End If
                    Next
                    Return False '花魁心操術自体は他に何も影響しない
                Case "仁将"
                    '仁将
                    '「将」武将のスキル発動率を上昇。
                    '※参照順番依存があるので変更・デバッグは慎重に
                    Dim refbs() As Busho = Nothing
                    If simu_execno = 0 Then
                        refbs = bs.Clone
                    ElseIf simu_execno = 1 Then
                        refbs = Form10.simu_bs.Clone
                    End If
                    '仁将の増分データ
                    Dim jh_f() As Decimal = {2.5, 3, 3.5, 4, 4.5, 5, 5.5, 6, 7, 8}
                    For i As Integer = 0 To refbs.Length - 1
                        If refbs(i).job = "将" Then
                            For j As Integer = 0 To refbs(i).skill.Length - 1
                                refbs(i).skill(j).up_kouka_p = refbs(i).skill(j).up_kouka_p + 0.01 * jh_f(.lv - 1) 'LVの分だけ発動率上昇
                            Next
                        End If
                    Next
                    Return False '仁将自体は他に何も影響しない
                Case "穢土の礎"
                    '穢土の礎
                    '姫武将にのみスキル効果を適用。
                    '効果自体は変動しないため、ここではflgをONにするだけ
                    Return True
                Case "堅国の絆"
                    '堅国の絆
                    '自本領に滞在する全部隊に効果を適用。
                    '効果自体は変動しないため、ここではflgをONにするだけ
                    Return True
                Case "天正筆才"
                    '天正筆才
                    '部隊内に「織田信長」という名称の武将が居れば、そのスキル発動率を操作
                    '※参照順番依存があるので変更・デバッグは慎重に
                    Dim nobu_x() As Decimal = {0.5, 0.55, 0.6, 0.65, 0.7, 0.75, 0.8, 0.85, 0.9, 1.0} '倍率
                    Dim refbs() As Busho = Nothing
                    If simu_execno = 0 Then
                        refbs = bs.Clone
                    ElseIf simu_execno = 1 Then
                        refbs = Form10.simu_bs.Clone
                    End If
                    For i As Integer = 0 To refbs.Length - 1
                        If InStr(refbs(i).name, "織田信長") Then
                            For j As Integer = 0 To refbs(i).skill.Length - 1
                                refbs(i).skill(j).up_kouka_p = refbs(i).skill(j).up_kouka_p + refbs(i).skill(j).kouka_p * nobu_x(.lv - 1) 'LVの分だけ発動率上昇
                            Next
                        End If
                    Next
                    Return False '天正筆才自体は他に何も影響しない
                Case "権勢専横"
                    '権勢専横
                    '部隊内の「権勢専横」スキルの個数によって効果が変動
                    'Dim ken_x() As Decimal = {0.055, 0.06, 0.065, 0.07, 0.075, 0.08, 0.085, 0.09, 0.1, 0.11}
                    Dim ken_no As Integer = 0
                    Dim refbs() As Busho = Nothing
                    If simu_execno = 0 Then
                        refbs = bs.Clone
                    ElseIf simu_execno = 1 Then
                        refbs = Form10.simu_bs.Clone
                    End If
                    For i As Integer = 0 To refbs.Length - 1
                        For j As Integer = 0 To refbs(i).skill.Length - 1
                            If refbs(i).skill(j).name = "権勢専横" Then
                                ken_no = ken_no + 1
                            End If
                        Next
                    Next
                    .up_kouka_f = .kouka_f * (ken_no - 1)
                    Return True
                Case "双鞭驃騎兵"
                    '双鞭驃騎兵
                    '全攻+「部隊速度/2」%
                    .kouka_f = (速度計算() / 2) * 0.01
                    .heika = "槍弓馬砲器"
                    Return True
                Case "洞察"
                    '洞察
                    '部隊ランクボーナスによって効果が変動
                    '洞察の増分データ
                    Dim dh_f() As Decimal = {5, 5.5, 6, 6.5, 7, 7.5, 8, 8.5, 9, 10}
                    .kouka_f = 0.01 * sumrank * dh_f(.lv - 1)
                    .heika = "槍弓馬砲器"
                    Return True
            End Select
            Return False
        End With
        Return True
    End Function

    'Form1の速度計測関数を参考に
    'butaiflg = Trueならば部隊スキル込みの速度を返す
    Public Function 速度計算(Optional ByVal butaiflg As Boolean = False) As Decimal
        Dim speed_skl() As Busho.skl = Nothing '速度スキル格納
        Dim ksk() As Decimal = Nothing '加速率
        Dim butai_spd As Decimal '行軍速度（秒）
        Dim busho_spd() As Decimal = Nothing, it() As Decimal = Nothing '武将の移動速度、No.
        Dim heikac As Integer = 0 '兵科セット数
        Dim c As Integer = 0 'カウンター

        '部隊兵法値計算・スキルデータ確定関数にてフラグ付きスキルは読み込み済み。実行順序依存。
        '勝軍地蔵が居るので読み込み
        'Call フラグ付きスキル読み込み() '読込（更新）
        'Call 童ボーナス加算()

        Dim refbs() As Busho = Nothing
        If simu_execno = 0 Then
            refbs = bs.Clone
        ElseIf simu_execno = 1 Then
            refbs = Form10.simu_bs.Clone
        End If

        ReDim ksk(refbs.Length - 1), busho_spd(refbs.Length - 1), it(refbs.Length - 1)

        For i As Integer = 0 To refbs.Length - 1 'スピードスキルのみ抽出
            With refbs(i)
                For j As Integer = 0 To .skill.Length - 1
                    If Not .skill(j).speed = 0 Or .skill(j).tokusyu = 1 Then '加速スキルがあれば
                        ReDim Preserve speed_skl(c)
                        speed_skl(c) = .skill(j).Clone
                        c = c + 1
                    End If
                Next
            End With
        Next
        For i As Integer = 0 To refbs.Length - 1 '各武将の速度を計算
            '童効果適用
            For j As Integer = 0 To refbs.Length - 1
                Select Case (refbs(j).heisyu.bunrui)
                    Case "槍"
                        ksk(i) = ksk(i) + 0.01 * warabe.speed.yari
                    Case "弓"
                        ksk(i) = ksk(i) + 0.01 * warabe.speed.yumi
                    Case "馬"
                        ksk(i) = ksk(i) + 0.01 * warabe.speed.uma
                    Case "砲"
                        ksk(i) = ksk(i) + 0.01 * warabe.speed.hou
                    Case "器"
                        ksk(i) = ksk(i) + 0.01 * warabe.speed.utuwa
                End Select
            Next
            If Not c = 0 Then '加速スキルが見つかった場合
                For j As Integer = 0 To speed_skl.Length - 1
                    If InStr(speed_skl(j).heika, refbs(i).heisyu.bunrui) Or InStr(speed_skl(j).heika, "全") Or InStr(speed_skl(j).heika, "将") Then
                        speed_skl(j).t_flg = フラグ付きスキル参照(speed_skl(j))
                        '特殊条件の速度スキルが増えてきたらアドホック対応では×・・・
                        If refbs.Length = 1 And speed_skl(j).name = "焔槍雷戟" Then '天前田利家のスキルのみは特例
                            ksk(i) = ksk(i) + speed_skl(j).speed * speed_skl(j).lv
                        Else
                            ksk(i) = ksk(i) + speed_skl(j).speed
                        End If
                    End If
                Next
            End If
            busho_spd(i) = refbs(i).heisyu.spd * (1 + ksk(i))
            it(i) = i
        Next
        Array.Sort(busho_spd, it)

        '今は部隊スキルが問題になることは無いので考慮しない
        'If Not bskill.speed = 0 Then
        'butai_spd = Int(3600 / (bs(Int(it(bsc - 1))).heisyu.spd * (1 + bskill.speed) * (1 + ksk(Int(it(bsc - 1))))))
        'butai_spd = kasoku(Int(bsc - 1)) * (1 + bskill.speed)
        'Else
        'butai_spd = kasoku(Int(refbs.Length - 1))
        'End If
        butai_spd = busho_spd(refbs.Length - 1)
        Return butai_spd
    End Function

    '武将データとスキルデータを引数に取り、武将がスキル効果適用条件に合致しなければFalse
    '（攻防一致、兵科一致は既に判定されているとして省略）
    Public Function 条件付スキル適用チェック(ByVal b As Busho, ByVal s As Busho.skl) As Boolean
        Select Case (s.name)
            Case "穢土の礎" '姫武将にのみ適用
                If Not (b.job = "姫") Then Return False
        End Select
        Return True
    End Function

    '全兵数と対象武将の積載兵数、全体の期待値と対象武将の期待値を引数にとり、
    Public Function デッキアウトHP計算(ByVal hei() As Decimal, ByVal ex() As Decimal) As String()
        Dim allhei, allex As Decimal
        Dim retHp() As String '戻り値となる全滅するHP
        ReDim retHp(hei.Length - 1)
        For i As Integer = 0 To hei.Length - 1
            allhei += hei(i)
            allex += ex(i)
            retHp(i) = vbNullString
        Next

        Dim higai_p() As Decimal = 被害分配比計算(ex) '各武将の被害兵数分配比を計算

        '各武将の全滅時HPを計算
        '全パターンを計算する。
        '勝利時と敗北時の、残HPと被害兵数との関係式
        '[勝] 100 - ( 敵総 / 自総 ) × 58
        '[負] ( 自総 / 敵総 ) × 42 - 5
        '勝利時と敗北時それぞれの被害兵数
        '[勝] 0.4 × ( 敵総 / 自総 ) × 自総兵数
        '[負] ( 1 - 0.6 × ( 自総攻 / 敵総) ) × 自総兵数
        '( 自総 / 敵総 ) <= 1/4の時には全滅確定。即ち残HP5.5、小数点以下切捨てHP5以下の時は全滅確定。
        '全滅未満、勝利しない範囲の彼我の戦力倍率を100等分して区間内で近似値を求める形にする
        For i As Decimal = 4 To 0.25 Step -((4 - 0.25) / 500)
            '完全全滅時
            If i = 0.25 Then
                For j As Integer = 0 To hei.Length - 1
                    If retHp(j) Is vbNullString Then
                        retHp(j) = "全滅時以外デッキ落ち無"
                    End If
                Next
                Continue For
            End If

            Dim allhigai As Decimal
            Dim hp As Decimal
            If i > 1 Then '勝利時
                allhigai = 0.4 * (1 / i) * allhei
                hp = Math.Ceiling(100 - (1 / i) * 58) '小数点以下切り上げ
            Else '敗北時
                allhigai = (1 - 0.6 * i) * allhei
                hp = Math.Floor((i * 42) - 5) '小数点以下切り捨て
            End If

            Dim higai() As Decimal
            Dim ahure() As Integer
            Dim ahuresum As Integer
            Dim ahurecnt As Integer = 0
            ReDim higai(hei.Length - 1), ahure(hei.Length - 1)
            'まずはそのまま被害兵数を分配
            For j As Integer = 0 To hei.Length - 1
                higai(j) = allhigai * higai_p(j)
                If ahure(j) = 0 Then
                    If (higai(j) >= hei(j)) Then '溢れ再分配を考えるまでもなく全滅
                        If retHp(j) Is vbNullString Then
                            retHp(j) = "デッキ落: 残HP" & hp.ToString
                        End If
                        ahure(j) = Math.Ceiling(higai(j) - hei(j)) '切り上げ勘定
                        ahuresum += ahure(j)
                        ahurecnt += 1
                    Else
                        ahure(j) = 0
                    End If
                End If
            Next
            '溢れを再分配
            Dim loopcnt As Integer = 0 'whileループカウンター（武将数以上周ることはありえない）
            While (ahuresum > 0 And loopcnt < hei.Length)
                loopcnt += 1
                '被害分配比再計算
                Dim zs_ex() As Decimal
                ReDim zs_ex(hei.Length - ahurecnt - 1)
                Dim cnt As Integer = 0
                For k As Integer = 0 To hei.Length - 1
                    If ahure(k) = 0 Then
                        zs_ex(cnt) = ex(k)
                        cnt += 1
                    End If
                Next
                Dim zs_higai_p() As Decimal = 被害分配比計算(zs_ex)
                Dim tmp_ahuresum As Integer = ahuresum
                ahuresum = 0
                cnt = 0
                For k As Integer = 0 To hei.Length - 1
                    If ahure(k) = 0 Then
                        higai(k) += tmp_ahuresum * zs_higai_p(cnt)
                        If (higai(k) >= hei(k)) Then
                            If retHp(k) Is vbNullString Then
                                retHp(k) = "デッキ落: 残HP" & hp.ToString
                            End If
                            ahure(k) = Math.Ceiling(higai(k) - hei(k)) '切り上げ勘定
                            ahuresum += ahure(k)
                            ahurecnt += 1
                        Else
                            ahure(k) = 0
                        End If
                    End If
                Next
            End While
        Next
        Return retHp
    End Function

    Public Function 被害分配比計算(ByVal Atkex() As Decimal) As Decimal()
        '総期待値
        Dim allAtkex As Decimal = 0
        For i As Integer = 0 To Atkex.Length - 1
            allAtkex += Atkex(i)
        Next
        '各武将の被害兵数分配比を計算
        Dim gyakuhi(), allgyaku As Decimal '逆比、逆比の総和
        Dim higai_p() As Decimal '被害率
        ReDim gyakuhi(Atkex.Length - 1), higai_p(Atkex.Length - 1)
        For i As Integer = 0 To Atkex.Length - 1
            gyakuhi(i) = 1 / (Atkex(i) / allAtkex)
            allgyaku += gyakuhi(i)
        Next
        For i As Integer = 0 To Atkex.Length - 1
            higai_p(i) = gyakuhi(i) / allgyaku '被害率
        Next
        Return higai_p
    End Function

    Public Function CsvLoad(ByVal csvname As String, Optional ByVal output_dgridview As DataGridView = Nothing) As String()()
        Dim parser As TextFieldParser = New TextFieldParser(csvname, Encoding.GetEncoding("Shift_JIS"))
        Dim returnstr()() As String = Nothing
        Dim count As Integer = 0
        parser.TextFieldType = FieldType.Delimited
        parser.SetDelimiters(",") ' 区切り文字はコンマ
        While (Not parser.EndOfData)
            Dim row As String() = parser.ReadFields() ' 1行読み込み
            ' 読み込んだデータ(1行をDataGridViewに表示する)
            If Not (output_dgridview Is Nothing) Then
                output_dgridview.Rows.Add(row)
            End If
            ReDim Preserve returnstr(count)
            returnstr(count) = row.Clone
            count = count + 1
        End While
        Return returnstr
    End Function

    Public Sub SaveToCsv(ByVal tempDgv As DataGridView, ByVal savepath As String)
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
        Dim swText As New System.IO.StreamWriter(savepath, False, System.Text.Encoding.GetEncoding(932))
        swText.Write(strResult.ToString)
        swText.Dispose()
    End Sub

    Public Sub 置換パラメータ設定読み込み()
        Dim sr As New System.IO.StreamReader(fcpath, System.Text.Encoding.GetEncoding(932))
        Dim srbuff As String = sr.ReadToEnd()
        sr.Close()
        Dim tmpbuf() As String = Split(srbuff, vbCrLf)
        ReDim cparam_data(tmpbuf.Length - 2)
        For i As Integer = 0 To tmpbuf.Length - 2
            Dim dmp() As String = Split(tmpbuf(i), ",")
            With cparam_data(i)
                .id = i
                .param_name = dmp(0)
                .param_value = dmp(2)
            End With
        Next
    End Sub

    Public Sub フラグ付きスキル読み込み()
        Dim sr As New System.IO.StreamReader(fdpath, System.Text.Encoding.GetEncoding(932))
        Dim srbuff As String = sr.ReadToEnd()
        sr.Close()
        Dim tmpbuf() As String = Split(srbuff, vbCrLf)
        Dim undefflg As Boolean 'データ抜けチェックフラグ
        Dim onoffstr()() As String = CsvLoad(fspath)
        ReDim fskill_data(tmpbuf.Length - 3)
        For i As Integer = 0 To tmpbuf.Length - 3
            undefflg = False
            Dim dmp() As String = Split(tmpbuf(i + 1), ",") 'csvのインデックスを飛ばす+1
            With fskill_data(i)
                .id = dmp(0)
                .name = dmp(1)
                .koubou = dmp(2)
                .heika = dmp(3)
                .lv = dmp(4)
                'データの抜けは「n」表記→これがあればその行は不完全なので読み込まない
                For j As Integer = 0 To dmp.Length - 1
                    If dmp(j) = "n" Then
                        undefflg = True
                        Exit For
                    End If
                Next
                If undefflg Then Continue For
                For j As Integer = 0 To onoffstr.GetLength(0) - 1
                    If (onoffstr(j)(1) = .name) Then
                        If InStr(onoffstr(j)(2), "[童]") Then
                            .wflg = True
                        Else
                            .wflg = False
                        End If
                        If (onoffstr(j)(3) = "ON") Then
                            .onoff = True
                        Else
                            .onoff = False
                        End If
                        Exit For
                    End If
                Next
                If .onoff Then
                    .kouka_p = dmp(5)
                    .kouka_f = dmp(6)
                    .speed = dmp(7)
                Else
                    .kouka_p = dmp(8)
                    .kouka_f = dmp(9)
                    .speed = dmp(10)
                End If
            End With
        Next
    End Sub

    'フラグ付きスキルのフラグNoとスキル性能の結合
    Public Function フラグ付きスキル参照(ByRef skl As Busho.skl) As Boolean
        Dim count As Integer = 0
        While (count < fskill_data.Length)
            '今のところコスト依存が無いから楽・・・
            With skl
                If (fskill_data(count).name = .name) And (fskill_data(count).lv = .lv) Then
                    .koubou = fskill_data(count).koubou
                    .heika = fskill_data(count).heika
                    If InStr(.heika, "全") Then
                        .heika = "槍弓馬砲器"
                    End If
                    .kouka_p = fskill_data(count).kouka_p
                    .kouka_f = fskill_data(count).kouka_f
                    .speed = fskill_data(count).speed
                    Return True
                End If
                count = count + 1
            End With
        End While
        If skl.t_flg Then
            Return True
        Else
            Return False
        End If
    End Function

    'csvからの文字列で童情報を抽出
    Public Sub 童ボーナス加算()
        Call warabe.warabe_clean()
        '今は単調ボーナスだけなのでここに記述したほうが軽いかな・・・
        For i As Integer = 0 To fskill_data.Length - 1
            If Not (fskill_data(i).wflg) Or Not (fskill_data(i).onoff) Then
                Continue For
            End If
            With fskill_data(i)
                Dim h As String = .heika
                Dim f As Decimal = Nothing
                If h = "全" Then h = "槍弓馬砲器"
                If InStr(kb, "速") Then f = .speed * 100 Else f = .kouka_f * 100
                warabe.warabe_set(.koubou, h, f)
                'Select Case (.name)
                '    Case "傾奇爛漫"
                '        warabe.def.yari += 3
                '    Case "日輪の子"
                '        warabe.def.yari += 5
                '    Case "華の童子"
                '        warabe.def.yumi += 5
                '    Case "勝軍地蔵"
                '        warabe.speed.yari += 3
                '        warabe.speed.yumi += 3
                '        warabe.speed.uma += 3
                '        warabe.speed.hou += 3
                '        warabe.speed.utuwa += 3
                '    Case Else
                '        Continue For
                'End Select
            End With
        Next
    End Sub

    Public Function 童効果文字列出力(ByVal kb As String)
        Dim harr() As String = {"槍", "弓", "馬", "砲", "器"}
        Dim warr() As Decimal = warabe.warabe_gets(kb)
        Dim strret As String = ""
        Dim kbstr As String = Nothing
        If InStr(kb, "攻") Then kbstr = "攻+" Else kbstr = "防+"
        For i As Integer = 0 To harr.Length - 1
            If warr(i) > 0 Then
                strret = strret & ", " & harr(i) & kbstr & CInt(warr(i)) & "%" '童効果は現在整数値のみ

            End If
        Next
        If Not strret = "" Then
            strret = strret.Remove(0, 2)
        Else
            strret = "適用無"
        End If
        Return strret
    End Function

    '部隊ランクボーナス
    Public Function 部隊ランクボーナス計算(ByVal ranksum As Decimal) As Decimal
        Dim rbutaibonus As Decimal = 0
        If ranksum < 4 Then
            rbutaibonus = ranksum * 0.1
        ElseIf ranksum < 8 Then
            rbutaibonus = 0.3 + (ranksum - 3) * 0.15
        ElseIf ranksum < 12 Then
            rbutaibonus = 0.9 + (ranksum - 7) * 0.2
        ElseIf ranksum < 16 Then
            rbutaibonus = 1.7 + (ranksum - 11) * 0.25
        ElseIf ranksum < 20 Then
            rbutaibonus = 2.7 + (ranksum - 15) * 0.3
        ElseIf ranksum < 24 Then
            rbutaibonus = 3.9 + (ranksum - 19) * 0.4
        Else 'ranksum = 24
            rbutaibonus = 6.0
        End If
        Return rbutaibonus
    End Function

    '兵科1に対する兵科2の相性
    Public Function 相性計算(ByVal heika1 As String, ByVal heika2 As String) As Decimal
        Select Case heika1
            Case "槍"
                If InStr(heika2, "弓") Then
                    Return 0.5
                ElseIf InStr(heika2, "槍") Then
                    Return 1
                ElseIf InStr(heika2, "馬") Then
                    Return 2
                ElseIf InStr(heika2, "砲") Then
                    Return 0.8
                Else
                    Return 1
                End If
            Case "弓"
                If InStr(heika2, "馬") Then
                    Return 0.5
                ElseIf InStr(heika2, "弓") Then
                    Return 1
                ElseIf InStr(heika2, "槍") Then
                    Return 2
                ElseIf InStr(heika2, "砲") Then
                    Return 1.3
                Else
                    Return 1
                End If
            Case "馬"
                If InStr(heika2, "槍") Then
                    Return 0.5
                ElseIf InStr(heika2, "馬") Then
                    Return 1
                ElseIf InStr(heika2, "弓") Then
                    Return 2
                ElseIf InStr(heika2, "砲") Then
                    Return 0.8
                Else
                    Return 1
                End If
        End Select
        Return 1
    End Function

    'DBから読み込んだ値に置換パラメータを適用
    Public Function パラメータ置換(ByVal target As String) As String
        For i As Integer = 0 To cparam_data.Length - 1
            Dim confstr As String = "%%" & StringRep(cparam_data(i).param_name, {""""}) & "%%"
            If InStr(target, confstr) Then '一致すれば置換
                Return 文字列計算(Replace(target, confstr, CStr(StringRep(cparam_data(i).param_value, {""""}))))
            End If
        Next
        Return 0
    End Function

    '全てクリア(2つまで例外設定可)
    Public Sub ClearTextBox(ByVal hParent As Control, Optional ByVal ex As Control = Nothing, Optional ByVal ex2 As Control = Nothing)
        ' hParent 内のすべてのコントロールを列挙する
        For Each cControl As Control In hParent.Controls
            ' 列挙したコントロールにコントロールが含まれている場合は再帰呼び出しする
            If cControl.HasChildren Then
                ClearTextBox(cControl)
            End If
            ' コントロールの型が TextBoxBase||Combobox からの派生型の場合は Text をクリアする
            If TypeOf cControl Is TextBoxBase Or TypeOf cControl Is ComboBox Then
                If Not (cControl Is ex Or cControl Is ex2) Then
                    If TypeOf cControl Is ComboBox Then
                        Dim ctmp As ComboBox = CType(cControl, ComboBox)
                        If ctmp.DropDownStyle = ComboBoxStyle.DropDownList Then 'DropDownList型のcomboboxがある場合、Textをクリアではダメ
                            ctmp.SelectedIndex = -1
                            Continue For
                        End If
                        ctmp.Text = String.Empty
                    Else
                        cControl.Text = String.Empty
                    End If
                End If
            End If
        Next cControl
    End Sub

    'テキスト加工(引数に与えた文字列を削る)
    Public Function StringRep(ByVal s As String, ByVal d As String()) As String
        Dim s_d As String = s
        For i As Integer = d.Length - 1 To 0 Step -1
            s_d = Replace(s_d, d(i), "")
        Next
        Return s_d
    End Function

    '現在の入力されている武将数カウント(0～3)
    Public Function Count_Busho() As Integer
        Dim bno As Integer = 4
        Dim c As Integer = 0
        For i As Integer = 0 To 3
            If ComboBox(Form1, CStr(i) & "02").Text = "" Then
                c = c + 1
            End If
        Next
        Return (bno - c - 1)
    End Function

    'テキストから数字文字列のみを抜き出す
    Public Function String_onlyNumber(ByVal s As String) As String
        Dim valstr As String = ""
        For i As Integer = 1 To s.Length
            Dim tmpd As String = Mid(s, i, 1)
            If StringRep(tmpd, {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}) = "" Then
                valstr = valstr & tmpd
            End If
        Next
        Return valstr
    End Function

    '配列の中から最大の数字の格納されている添え字番号を返す
    Public Function Array_MAX(ByVal ary() As Decimal, Optional ByVal when_eq() As Decimal = Nothing) As Integer
        Dim amaxi As Integer = 0
        Dim amax As Decimal = ary(0)
        For i As Integer = 1 To ary.Length - 1
            If amax < ary(i) Then
                amax = ary(i)
                amaxi = i
            ElseIf amax = ary(i) Then '同じ場合どっちをMAXにするか
                If Not when_eq Is Nothing Then
                    If when_eq(amaxi) > when_eq(i) Then
                        amax = ary(i)
                        amaxi = i
                    End If
                End If
            End If
        Next
        Return amaxi
    End Function

    'グループボックスを非表示にし、内部のテキストクリア
    Public Sub OFF_GroupBox(ByVal hParent As GroupBox)
        hParent.Visible = False
        ClearTextBox(hParent)
    End Sub

    Public Function 文字列計算(ByVal exp As String, Optional ByVal errmsg As Boolean = True) As Decimal
        Dim t As New DataTable()
        exp = Replace(exp, "×", "*")
        exp = Replace(exp, "÷", "/")
        exp = Replace(exp, "＋", "+")
        exp = Replace(exp, "－", "-")
        exp = Replace(exp, " ", "")
        exp = Replace(exp, "%", "*0.01")
        Try
            t.Columns.Add("数式", GetType(String), exp)
        Catch es As Exception
            If errmsg = True Then
                MsgBox("スキル計算式エラー" & vbCrLf & CStr(exp))
            End If
            Return 0
        End Try
        Return t.Rows.Add()("数式")
    End Function

    '引数indexに番号を受け取って、その番号が付いたLabelコントロールを返す
    Public Function Label(ByVal f As Form, ByVal index As String) As Label
        Dim ctrl As Control() = f.Controls.Find("Label" & index, True)
        Return DirectCast(ctrl(0), Label)
    End Function

    '引数indexに番号を受け取って、その番号が付いたTextBoxコントロールを返す
    Public Function TextBox(ByVal f As Form, ByVal index As String) As TextBox
        Dim ctrl As Control() = f.Controls.Find("TextBox" & index, True)
        Return DirectCast(ctrl(0), TextBox)
    End Function

    '引数indexに番号を受け取って、その番号が付いたComboBoxコントロールを返す
    Public Function ComboBox(ByVal f As Form, ByVal index As String) As ComboBox
        Dim ctrl As Control() = f.Controls.Find("ComboBox" & index, True)
        Return DirectCast(ctrl(0), ComboBox)
    End Function

    '引数indexに番号を受け取って、その番号が付いたGroupBoxコントロールを返す
    Public Function GroupBox(ByVal f As Form, ByVal index As String) As GroupBox
        Dim ctrl As Control() = f.Controls.Find("GroupBox" & index, True)
        Return DirectCast(ctrl(0), GroupBox)
    End Function

    '引数indexに番号を受け取って、その番号が付いたButtonコントロールを返す
    Public Function Button(ByVal f As Form, ByVal index As String) As Button
        Dim ctrl As Control() = f.Controls.Find("Button" & index, True)
        Return DirectCast(ctrl(0), Button)
    End Function

    '引数indexに番号を受け取って、その番号が付いたRichTextBoxコントロールを返す
    Public Function RichTextBox(ByVal f As Form, ByVal index As String) As RichTextBox
        Dim ctrl As Control() = f.Controls.Find("RichTextBox" & index, True)
        Return DirectCast(ctrl(0), RichTextBox)
    End Function

    '引数indexに番号を受け取って、その番号が付いたCheckBoxコントロールを返す
    Public Function CheckBox(ByVal f As Form, ByVal index As String) As CheckBox
        Dim ctrl As Control() = f.Controls.Find("CheckBox" & index, True)
        Return DirectCast(ctrl(0), CheckBox)
    End Function

    '■Convert10to2
    '■機能：10進数を2進数に変換する。
    Public Function Convert10to2(ByVal Value As Long, ByVal keta As Integer) As String

        Dim lngBit As Long
        Dim strData As String = Nothing

        Do Until (Value < 2 ^ lngBit)
            If (Value And 2 ^ lngBit) <> 0 Then
                strData = "1" & strData
            Else
                strData = "0" & strData
            End If
            lngBit = lngBit + 1
        Loop
        If strData = "" Then 'value = 0の時にnothingで出てこないようにする
            strData = 0
        End If
        Convert10to2 = strData.PadLeft(keta, "0")
    End Function

    '正規分布計算
    Public Function pnorm(ByVal qn As Decimal) As Decimal
        Dim b() As Decimal = {1.570796288, 0.03706987906, -0.0008364353589,
              -0.0002250947176, 0.000006841218299, 0.000005824238515,
              -0.00000104527497, 0.00000008360937017, -0.000000003231081277,
               0.00000000003657763036, 0.0000000000006936233982}

        Dim w1, w3 As Decimal
        If qn < 0 Or 1 < qn Then
            MsgBox("Error : qn <= 0 or qn >= 1  in pnorm()!\n")
            Return 0
        End If
        If qn = 0.5 Then
            Return 0
        End If

        w1 = qn

        If qn > 0.5 Then
            w1 = 1 - w1
        End If
        w3 = (-1) * Math.Log(4 * w1 * (1 - w1))
        w1 = b(0)
        For i As Integer = 1 To 10
            w1 = w1 + (b(i) * w3 ^ (i))
        Next

        If qn > 0.5 Then
            Return Math.Sqrt(w1 * w3)
        Else
            Return (-1) * Math.Sqrt(w1 * w3)
        End If
    End Function
    Public Function qnorm(ByVal u As Decimal) As Decimal
        Dim a() As Decimal = {0.000124818987, -0.001075204047, 0.005198775019,
           -0.019198292004, 0.059054035642, -0.151968751364,
            0.319152932694, -0.5319230073, 0.797884560593}
        Dim b() As Decimal = {-0.000045255659, 0.00015252929, -0.000019538132,
              -0.000676904986, 0.001390604284, -0.00079462082,
              -0.002034254874, 0.006549791214, -0.010557625006,
               0.011630447319, -0.009279453341, 0.005353579108,
              -0.002141268741, 0.000535310549, 0.999936657524}
        Dim w, y, z As Decimal

        If u = 0 Then
            Return 0.5
        End If
        y = u / 2
        If y < -6 Then
            Return 0
        End If
        If y > 6 Then
            Return 1
        End If
        If y < 0 Then
            y = -y
        End If
        If y < 1 Then
            w = y * y
            z = a(0)
            For i As Integer = 1 To 8
                z = z * w + a(i)
            Next
            z = z * (y * 2)
        Else
            y = y - 2
            z = b(0)
            For i As Integer = 1 To 14
                z = z * y + b(i)
            Next
        End If
        If u < 0 Then
            Return (1 - z) / 2
        Else
            Return (1 + z) / 2
        End If
    End Function
    Public Function Nx(ByVal x As Decimal, ByVal exp As Decimal, ByVal a As Decimal) As Decimal
        Nx = 1 / (Math.Sqrt(2 * Math.PI) * a) * Math.Exp((-1) * (x - exp) ^ 2 / (2 * a ^ 2))
    End Function

    ' クイックソート昇順(各武将の戦闘力→その和の昇順ソート，生起確率)
    Public Sub sQuickSort2(ByRef myArray(,) As Decimal, ByRef dArray() As Decimal)
        Dim indexarray(), myArrayk(), tmpArray(,), tmpdArray() As Decimal
        ReDim indexarray(UBound(dArray)), skill_yk(UBound(dArray))
        'インデックス配列を用意
        For t As Integer = 0 To indexarray.Length - 1
            indexarray(t) = t
        Next
        'ソートのためにmyArrayk更新
        myArrayk = Array_to_Arrayk(myArray)
        Array.Sort(myArrayk, indexarray)
        'myArrayをソート
        ReDim tmpArray(UBound(myArray), UBound(myArray, 2))
        ReDim tmpdArray(UBound(dArray))
        For i As Integer = 0 To UBound(myArray)
            For j As Integer = 0 To UBound(myArray, 2)
                tmpArray(i, j) = myArray(indexarray(i), j)
            Next
            tmpdArray(i) = dArray(indexarray(i))
        Next
        myArray = tmpArray.Clone
        dArray = tmpdArray.Clone
    End Sub

    '各武将の戦闘力が入っている2次元配列から、総戦闘力の1次元配列を返す
    Public Function Array_to_Arrayk(ByVal myArray(,) As Decimal) As Decimal()
        Dim tmparray() As Decimal = Nothing
        ReDim tmparray(UBound(myArray))
        For i As Integer = 0 To UBound(myArray)
            For j As Integer = 0 To UBound(myArray, 2)
                tmparray(i) = tmparray(i) + myArray(i, j)
            Next
        Next
        Return tmparray
    End Function

    '' クイックソート昇順(1次元:1次元)
    'Public Sub sQuickSort(ByRef myArray() As Decimal, ByRef dArray() As Decimal)
    '    Dim lngLow As Decimal, lngHigh As Decimal
    '    lngLow = LBound(myArray)
    '    lngHigh = UBound(myArray)
    '    Call sAQuick(myArray, dArray, lngLow, lngHigh)
    'End Sub
    'Private Sub sAQuick(ByRef myArray() As Decimal, ByRef dArray() As Decimal, _
    ' ByVal lngLeft As Decimal, _
    ' ByVal lngRight As Decimal)
    '    Dim tmpArray As Decimal
    '    Dim tmpa, tmpb As Decimal
    '    Dim i As Long, j As Long

    '    If lngLeft < lngRight Then
    '        tmpArray = myArray((lngLeft + lngRight) \ 2)
    '        i = lngLeft
    '        j = lngRight
    '        Do While (True)
    '            Do While (myArray(i) < tmpArray)
    '                i = i + 1
    '            Loop
    '            Do While (myArray(j) > tmpArray)
    '                j = j - 1
    '            Loop
    '            If i >= j Then Exit Do
    '            'myArray(i) = myArray(i) Xor myArray(j)
    '            'myArray(j) = myArray(i) Xor myArray(j)
    '            'myArray(i) = myArray(i) Xor myArray(j)
    '            tmpa = myArray(i)
    '            myArray(i) = myArray(j)
    '            myArray(j) = tmpa
    '            tmpb = dArray(i)
    '            dArray(i) = dArray(j)
    '            dArray(j) = tmpb
    '            i = i + 1
    '            j = j - 1
    '        Loop
    '        Call sAQuick(myArray, dArray, lngLeft, i - 1)
    '        Call sAQuick(myArray, dArray, j + 1, lngRight)
    '    End If
    'End Sub

    'INIファイル読み書き
    'GETの場合はINIFILEにはフルパス指定
    Public Function GetINIValue(ByVal KEY As String, ByVal Section As String, ByVal INIFILE As String) As String
        Dim strResult As String = Space(255)
        Call GetPrivateProfileString(Section, KEY, "－", _
                                 strResult, Len(strResult), INIFILE)
        GetINIValue = Left(strResult, InStr(strResult, Chr(0)) - 1)
    End Function
    Public Function GetINISection(ByVal Section As String, ByVal INIFILE As String) As Hashtable
        Dim strResult As String = Space(1023)
        Dim setcount As Integer = GetPrivateProfileSection(Section, strResult, Len(strResult), INIFILE)
        If setcount <= 0 Then Return Nothing
        Dim strVal As String = Left(strResult, InStr(1, strResult, vbNullChar & vbNullChar) - 1)
        Dim splVal As String() = Split(strVal, vbNullChar)
        Dim retval As Hashtable = New Hashtable
        For i As Integer = 0 To splVal.Length - 1
            Dim stmp As String() = Split(splVal(i), "=")
            retval.Add(stmp(0), stmp(1))
        Next
        Return retval
    End Function
    'SETの場合はINIFILEにはフルパス指定
    Public Function SetINIValue(ByVal Value As String, ByVal KEY As String, ByVal Section As String, ByVal INIFILE As String) As Boolean
        Dim Ret As Long
        Ret = WritePrivateProfileString(Section, KEY, Value, INIFILE)
        SetINIValue = CBool(Ret)
    End Function
    Public Function DeleteINISection(ByVal Section As String, ByVal INIFILE As String) As Boolean
        Dim Ret As Long
        Ret = WritePrivateProfileString(Section, Nothing, vbNullString, _
                                  My.Application.Info.DirectoryPath & "\" & INIFILE)
        DeleteINISection = CBool(Ret)
    End Function
    Public Function CopyINIValue(ByVal KEY As String, ByVal SectionA As String, ByVal SectionB As String, ByVal INIFILE As String) As Boolean
        'A:コピー元, B:コピー先
        Dim Ret As Long
        Dim SKey As String = GetINIValue(KEY, SectionA, INIFILE)
        Ret = WritePrivateProfileString(SectionB, KEY, SKey, _
                                  My.Application.Info.DirectoryPath & "\" & INIFILE)
        CopyINIValue = CBool(Ret)
    End Function
End Module