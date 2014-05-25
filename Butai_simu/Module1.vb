Imports System.IO
Imports System.Data.OleDb
Imports System.Runtime.InteropServices

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
        Public heika As String '兵科
        Public speed As Decimal '速度上昇率
        Public tokusyu As Integer '特殊スキル判定(0:通常, 1:速度のみ, 2:破壊のみ, 5:データ不足, 9:その他)
        Public sc_dflg As Boolean '総コスト依存フラグ
        Public exp_kouka As Decimal '期待値
        Public exp_kouka_b As Decimal '部隊兵法補正後の期待値
        Public Function Clone() As Object Implements System.ICloneable.Clone
            Return Me.MemberwiseClone()
        End Function
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

    'outflgがOFFならば、スキルが取得できなかった時のエラーMsgboxを表示せずにkanrenにスキル効果説明を放り込む
    Sub スキル取得(ByVal sno As Integer, ByVal sname As String, ByVal slv As Integer, ByVal fss As Integer(), Optional ByVal outflg As Boolean = True)
        'fssはその武将のスキル登録状態
        For i As Integer = 0 To skill_no - 1
            If skill.Length - 1 < i Then 'skillの要素数がiよりも小さければ
                ReDim Preserve skill(i)
            End If
            If fss(i) = sno Then
                Dim tmp(), u, s As String
                If Not skill_no = skill.Length And Not skill(i).koubou = "" Then '既にスキルが格納されている場合
                    ReDim Preserve skill(i + 1)
                    If skill(i + 1).koubou = "" Then 'その下が空欄ならば
                        skill(i + 1) = skill(i).Clone
                    End If
                End If
                With skill(i)
                    .name = sname
                    .lv = slv
                    tmp = Skill_ref(.name, .lv)
                    If tmp(0) = "" Then '最近wikiの更新で、データが無い部分が"+"も無い場合がある
                        tmp(0) = 0
                    End If
                    .kouka_p = Decimal.Parse(tmp(0))
                    .heika = tmp(1)

                    s = Mid(tmp(2), 1, InStr(tmp(2), "："))
                    .koubou = Replace(Replace(s, "：", ""), .heika, "")
                    .speed = スピードスキル取得(tmp(2), tmp(3)) 'スピードスキル確認
                    '特殊スキルリスト確認
                    .tokusyu = 0 'リセット
                    For j As Integer = 0 To error_skill.Length - 1
                        If sname = error_skill(j) Then
                            .tokusyu = 9
                            Exit For
                        End If
                    Next
                    If (InStr(.koubou, "攻") = 0 And InStr(.koubou, "防") = 0) Or .tokusyu = 9 Then '特殊スキルの扱い
                        Select Case .koubou
                            Case "速"
                                .tokusyu = 1
                            Case "破壊"
                                .tokusyu = 2
                            Case Else
                                .tokusyu = 9
                        End Select
                    Else
                        .tokusyu = 0
                        u = Replace(tmp(2), s, "")
                        u = Replace(u, "上昇", "")
                        '***** 情報がない、もしくはコスト限定情報 *****
                        If u = "%" Then
                            .tokusyu = 5 'データが無い
                            If outflg Then
                                MsgBox("このスキル・LVはWikiに登録情報がありません。" & vbCrLf & "シミュレーション結果には反映されません" & vbCrLf & _
                                       "【" & .name & "LV" & .lv & "】")
                            End If
                        End If
                        For k As Decimal = 1 To 4 Step 0.5
                            If InStr(u, "(コスト" & k & ")") Then
                                .tokusyu = 5
                                If outflg Then
                                    MsgBox("このスキル・LVは限定されたコスト下での情報しかありません" & vbCrLf & "シミュレーションエラーを起こす場合があります" & vbCrLf & _
                                           "【" & .name & "LV" & .lv & "】")
                                End If
                                u = Replace(u, "(コスト " & k & " )", "")
                            End If
                        Next
                        '**********************************************
                        If outflg Then
                            .kanren = u 'kanrenにスキル効果文字列を
                        End If
                        If InStr(u, "コスト") Then 'コスト依存スキルの扱い
                            u = Replace(u, "コスト", CStr(cost)) '変更。「スキル所持武将の」コストで一括適用
                        End If
                        If Not .tokusyu = 5 Then 'データが無いスキルの場合は計算しない
                            u = 文字列計算(u)
                            .kouka_f = Decimal.Parse(u)
                        Else
                            .kouka_f = 0
                        End If

                        If InStr(.heika, "全") Then
                            .heika = "槍弓馬砲器"
                        End If
                    End If
                End With
            Else
                skill(i) = skill(fss(i)).Clone
            End If
        Next
        ReDim Preserve skill(skill_no - 1)
    End Sub

    '基本効果、付加効果を入力、加速率を出力
    Function スピードスキル取得(ByVal kihon As String, ByVal huka As String) As Decimal
        If InStr(kihon, "速") = 0 And InStr(huka, "速") = 0 Then '一般スキルにはゼロ
            Return 0
            Exit Function
        End If
        Dim spd As String = "0"
        Dim dsp As String = vbNullString
        Dim ss, se, sf As Integer
        '基本欄にあるか付加欄にあるか
        If InStr(kihon, "速") Then
            dsp = kihon
        ElseIf InStr(huka, "速") Then
            dsp = huka
        End If
        ss = InStr(dsp, "：") + 1
        If InStr(dsp, "上昇") Then
            se = InStr(dsp, "上昇") - ss
            sf = 1
        ElseIf InStr(dsp, "低下") Then
            se = InStr(dsp, "低下") - ss
            sf = -1
        End If
        spd = Mid(dsp, ss, se)
        If sf = -1 Then '低下スキルならば
            spd = "-" & spd
        End If
        '特殊な付加要素が加わったスピードスキルが増えてきたためエラーチェックFalse
        スピードスキル取得 = 文字列計算(spd, False)
    End Function

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
         DB_DirectOUT(con, cmd, "SELECT 統率,兵種名,兵科,攻撃,防御,移動 FROM Heika WHERE 兵種名=""" & h & """", {"兵科", "統率", "攻撃", "防御", "移動"})
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
        For k As Integer = 3 To 5
            Label(Form1, CStr(bc) & "0" & CStr(k)).Text = "(+" & sta_g(k - 3).ToString & ")"
        Next
        ComboBox(Form1, CStr(bc) & "14").Text = skill(0).lv
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

    '外から部隊兵法値を取ってくるかどうかがheihou_
    Sub スキル期待値計算(Optional ByVal heihou_ As Decimal = -1) '部隊兵法値が既知であることが必要
        Dim heihou_sum_ As Decimal
        If heihou_ = -1 Then
            heihou_sum_ = heihou_sum
        Else
            heihou_sum_ = heihou_
        End If
        For j As Integer = 0 To skill_no - 1
            With skill(j)
                If .tokusyu = 0 Then '通常スキルならば
                    .sc_dflg = False
                    If .kouka_p = 1 Then
                        .kouka_p_b = 1
                    Else
                        .kouka_p_b = .kouka_p + 0.01 * heihou_sum_
                    End If
                    .exp_kouka = .kouka_p * .kouka_f
                    .exp_kouka_b = .kouka_p_b * .kouka_f '期待値
                Else
                    If .tokusyu = 9 Then '特殊スキルの場合は・・・
                        .sc_dflg = 部隊コスト依存スキル判定(skill(j), Costsum) '怪しいスキルを疑う
                        If Not .kouka_f = 0 Then 'ゼロならば単なる特殊スキル
                            If .kouka_p = 1 Then
                                .kouka_p_b = 1
                            Else
                                .kouka_p_b = .kouka_p + 0.01 * heihou_sum_
                            End If
                            .exp_kouka = .kouka_p * .kouka_f
                            .exp_kouka_b = .kouka_p_b * .kouka_f '期待値
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
    Public Property ONOFF() As Boolean
        Get
            If flg = True Then
                Return True
            Else
                Return False
            End If
        End Get
        Set(ByVal value As Boolean)
            With Form1.ToolStripButton4
                If value = True Then
                    flg = True '部隊スキルボタンの画像を変更
                    .Image = Bitmap.FromFile(My.Application.Info.DirectoryPath & "\settings\ico\prettyxstickxstripe_p24_rd_nl_l.png")
                    Dim bkb As String
                    If koubou = "攻" Then
                        bkb = "攻撃"
                    Else
                        bkb = "防衛"
                    End If
                    .ToolTipText = "部隊スキル : 有効" & " [" & bkb & "時発動]" & vbCrLf & _
                        "発動率: " & kouka_p & "/ 上昇率: " & kouka_f
                    If Not bskill.speed = 0 Then '加速有効
                        .ToolTipText = .ToolTipText & vbCrLf & "加速: " & speed
                    End If
                Else
                    flg = False
                    .Image = Bitmap.FromFile(My.Application.Info.DirectoryPath & "\settings\ico\prettyxstickxstripe_p24_bk_nl_l.png")
                    .ToolTipText = "部隊スキル : 無効"
                End If
            End With
        End Set
    End Property
    Public koubou As String '攻防
    Public name As String '部隊スキル名
    Public lv As Integer 'LV
    Public kouka_p As Decimal '発動率
    Public kouka_f As Decimal '上昇率
    Public speed As Decimal 'スピード上昇のある場合、加速率
    Public taisyo As String '対象
    Public qq As Boolean '将攻二乗モードONOFF
End Structure

Module Module1
    'INIファイル編集のためにWindows-API使用
    <DllImport("KERNEL32.DLL")> _
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
    '*** 部隊スキル関係 ***
    Public bskill As bskl '武将スキル
    '*** DB関連の変数 ***
    Public con, con2, con3 As New OleDbConnection 'DB接続設定に必要な変数 1:武将DB, 2:スキルDB, 3:NPC空き地情報
    Public cmd, cmd2, cmd3 As New OleDbCommand
    Public dbpath As String = Application.StartupPath & "\Busho.mdb" '実行ファイルのある階層に武将DBは置く
    Public dbpath2 As String = Application.StartupPath & "\Skill.mdb" '同じく、スキルDB
    public dbpath3 as string = Application.StartupPath & "\npc.mdb"
    Public bnpath As String = Application.StartupPath & "\BName2.INI" '同名武将区別ファイル
    Public espath As String = Application.StartupPath & "\ERRORSKILL.txt" '特殊スキルリストの場所
    Public error_skill() As String '特殊スキルリスト
    '*** INIファイルの場所 ***
    Public FILENAME_bs As String
    '*** CSVファイルの場所 ***
    Public FILENAME_csv As String

    Public Sub DB_Open(ByVal con As OleDbConnection, ByVal cmd As OleDbCommand, ByVal dbpath As String) '開始時にDBを開く設定
        If Not con.State = ConnectionState.Open Then
            'If dbcomp = False Then
            'Dim jroJet As New JRO.JetEngine '最適化
            'jroJet.CompactDatabase( _
            '    "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
            '   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath & "\Busho_new.mdb")
            'My.Computer.FileSystem.RenameFile(Application.StartupPath & "\Busho_new.mdb", "Busho.mdb")
            'dbcomp = True
            'End If
            con.ConnectionString = _
              "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & dbpath
            cmd.Connection = con
        End If
    End Sub

    Public Sub DB_Connection(ByVal con As OleDbConnection, ByVal cmd As OleDbCommand) 'オープンorクローズ
        Try
            con.Open()
        Catch ex As Exception
            con.Close()
            con.Open()
        End Try
    End Sub

    'データベースからSQL文を用いて出力(DataSet出力)
    Public Function DB_TableOUT(ByVal con As OleDbConnection, ByVal cmd As OleDbCommand, ByVal sql_str As String, ByVal outTable As String) As DataSet
        Dim da As New OleDbDataAdapter
        Dim ds As DataSet = New DataSet
        Call DB_Connection(con, cmd)
        cmd.CommandText = sql_str
        da.SelectCommand = cmd
        ds.Clear()
        da.Fill(ds, outTable)
        Return ds
    End Function

    '(直接出力)
    Public Function DB_DirectOUT(ByVal con As OleDbConnection, ByVal cmd As OleDbCommand, ByVal sql_str As String, ByVal outlist As String()) As String()
        Dim dr As OleDbDataReader
        Dim output() As String
        Dim l As Integer = outlist.Length
        ReDim output(l - 1)
        Call DB_Connection(con, cmd)
        cmd.CommandText = sql_str
        dr = cmd.ExecuteReader()
        While dr.Read()
            For i As Integer = 0 To l - 1
                If TypeOf dr(outlist(i)) Is DBNull Then 'DBが空欄
                    output(i) = ""
                Else
                    output(i) = CStr(dr(outlist(i)))
                End If
            Next
        End While
        dr.Close()
        Return output
    End Function

    '基本はDB_DirectOUTと同じだが、こちらはリスト形式（縦向き配列）で出力する必要のある場合用いる
    Public Function DB_DirectOUT2(ByVal con As OleDbConnection, ByVal cmd As OleDbCommand, ByVal sql_str As String, ByVal outlist As String) As String()
        Dim dr As OleDbDataReader
        Dim output() As String = Nothing
        Call DB_Connection(con, cmd)
        cmd.CommandText = sql_str
        dr = cmd.ExecuteReader()
        Dim c As Integer = 0
        While dr.Read()
            ReDim Preserve output(c)
            If TypeOf dr(outlist) Is DBNull Then 'DBが空欄
                output(c) = ""
            Else
                output(c) = CStr(dr(outlist))
            End If
            c = c + 1
        End While
        dr.Close()
        Return output
    End Function

    '基本は上の二つのDB_DirectOUTと同じだが、これは二次元配列で一気に表形式で出力する場合に用いる
    Public Function DB_DirectOUT3(ByVal con As OleDbConnection, ByVal cmd As OleDbCommand, ByVal sql_str As String, ByVal outlist() As String) As String()()
        Dim dr As OleDbDataReader
        Dim output()() As String = Nothing
        Dim l As Integer = outlist.Length
        ReDim output(l - 1)
        Call DB_Connection(con, cmd)
        cmd.CommandText = sql_str
        dr = cmd.ExecuteReader()
        Dim d As Integer = 0
        While dr.Read()
            Dim outputtmp() As String = Nothing
            For i As Integer = 0 To l - 1
                ReDim Preserve outputtmp(i)
                If TypeOf dr(outlist(i)) Is DBNull Then 'DBが空欄
                    outputtmp(i) = ""
                Else
                    outputtmp(i) = CStr(dr(outlist(i)))
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
        Dim s As String = "SELECT * FROM Skill WHERE 名前=""" & skillName & """ AND LV=" & skillLv & ""
        Dim t() As String = _
        DB_DirectOUT(con, cmd, s, {"確率", "対象", "基本効果", "付加効果"})
        Return t
    End Function

    Public Function Skill_ref_list(ByVal skillName() As String, ByVal skillLv As Integer) As String()()
        Dim skillstr As String = ""
        For i As Integer = 0 To skillName.Length - 1
            skillstr = skillstr & """" & skillName(i) & """" & ","
        Next
        skillstr = skillstr.Remove(skillstr.Length - 1, 1)
        Dim s As String = "SELECT * FROM Skill WHERE LV=" & skillLv & " AND 名前 IN(" & skillstr & ")" & ""
        s.Remove(s.Length - 1, 1)
        Dim t()() As String = _
        DB_DirectOUT3(con, cmd, s, {"名前", "確率", "対象", "基本効果", "付加効果"})
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
    Public Function スキル関連推定(ByVal skill_name As String) As String
        Dim stmp() As String = DB_DirectOUT(con, cmd, "SELECT * FROM Skill WHERE 名前 = """ & skill_name & _
                         """ AND LV=1", {"分類"})
        スキル関連推定 = stmp(0)
    End Function

    '部隊コストに依存したスキルの扱い（覇王征軍 and 武神八幡陣）
    Public Function 部隊コスト依存スキル判定(ByRef sk As Busho.skl, ByVal sumcost As Decimal) As Boolean
        '覇王征軍の増分データ
        Dim hs() As Decimal = {1, 1, 1, 2, 3, 4, 5, 7, 9, 11, 13, 16, 20, 25}
        '武神八幡陣の増分データ
        Dim bh() As Decimal = {1, 1, 1, 2, 3, 3, 4, 5, 6, 7, 9, 10, 12, 14, 16, 19, 22, 25}
        Dim sdata() As String

        With sk
            Select Case .name
                Case "覇王征軍"
                    '覇王征軍
                    sdata = Skill_ref(.name, .lv) '特殊スキルなのでここで改めて取得
                    If sumcost > 9 Then
                        .kouka_f = 0.01 * (Val(String_onlyNumber(sdata(2))) + hs(((sumcost - 9) / 0.5) - 1))
                    Else
                        .kouka_f = 0.01 * Val(String_onlyNumber(sdata(2)))
                    End If
                    .heika = "槍弓馬砲器"
                Case "武神八幡陣"
                    '武神八幡陣
                    sdata = Skill_ref(.name, .lv) '特殊スキルなのでここで改めて取得
                    If sumcost > 7 Then
                        .kouka_f = 0.01 * (Val(String_onlyNumber(sdata(2))) + bh(((sumcost - 7) / 0.5) - 1))
                    Else
                        .kouka_f = 0.01 * Val(String_onlyNumber(sdata(2)))
                    End If
                    .heika = "槍弓馬砲器"
                Case Else
                    Return False
            End Select
        End With
        Return True
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
                    cControl.Text = String.Empty
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
    'SETの場合はINIFILEにはファイル名のみ指定
    Public Function SetINIValue(ByVal Value As String, ByVal KEY As String, ByVal Section As String, ByVal INIFILE As String) As Boolean
        Dim Ret As Long
        Ret = WritePrivateProfileString(Section, KEY, Value, My.Application.Info.DirectoryPath & "\butai\" & INIFILE)
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