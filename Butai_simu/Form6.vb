Public Class Form6  
    Public Structure Kouho
        Public rare As String 'レア度
        Public name As String 'スキル名
        Public pp As Decimal '成功確率（銅銭での場合）
        Public Function pp_kin() As Decimal
            pp_kin = pp + 5
        End Function
    End Structure

    Public gyakuflg As Boolean '逆引きモードフラグ
    Public errorflg As Integer '各種エラーフラグ
    Public s1, s2 As String 'スロ1、スロ2武将No.
    Public kouho_sum As Integer '合成候補数
    Public skilllv_sum As Integer 'スキルレベル合計
    Public st() As Kouho
    Public s1flg As Boolean = False 's1出現フラグ
    Public s2flg As Boolean = False '同一合成フラグ
    Public suro2sno As Integer 'スロ2武将スキル数→これによって基本候補が変わる
    Public suro1sno As Integer 'スロ1武将スキル数
    Public suro1() As String
    Public suro2stable()() 'スロ2のスキルの基本テーブル
    Public okikaeindex As Integer '置換対象のスキル

    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles Me.Load
        GroupBox1.AllowDrop = True
        GroupBox2.AllowDrop = True
        Me.Opacity = 0.9 '初期の透過度は80%
        TrackBar1.Value = 9 '初期位置8
        Me.TopMost = True
    End Sub

    Private Sub 透過度変更(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.Scroll
        Me.Opacity = 0.1 * Val(sender.value)
    End Sub

    Private Sub 常に手前に表示(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            Me.TopMost = True
        Else
            Me.TopMost = False
        End If
    End Sub

    Private Function 武将取得(ByVal sender As System.Object) As Integer
        Dim bc As String = Mid(CStr(sender.Name), 9, 1)
        Select Case bc
            Case 0
                Return 0
            Case 1
                Return 1
        End Select
        Return -1
    End Function

    Private Sub R選択(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles ComboBox001.SelectedIndexChanged, ComboBox101.SelectedIndexChanged
        Dim cc As ComboBox = ComboBox(Me, CStr(武将取得(sender)) & "02")
        RemoveHandler cc.SelectedValueChanged, AddressOf Me.武将名選択
        Dim p As DataSet = _
        DB_TableOUT("SELECT id, 武将R, 武将名 FROM BData WHERE 武将R = " & ダブルクオート(sender.SelectedItem) & " ORDER BY Bid ASC", "BData")
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
        Handles ComboBox002.SelectedValueChanged, ComboBox102.SelectedValueChanged
        Dim bc As String = 武将取得(sender)
        Dim s() As String = _
        DB_DirectOUT("SELECT * FROM BData WHERE 武将R = " & ダブルクオート(ComboBox(Me, CStr(bc) & "01").Text) & _
                     " AND 武将名 = " & ダブルクオート(sender.Text) & " AND Bunf = 'F'", {"id", "初期スキル名"})
        Label(Me, CStr(bc) & "02").Text = s(1)
        If bc = 0 Then
            s1 = s(0)
        Else
            s2 = s(0)
        End If
        ComboBox(Me, CStr(bc) & "12").Text = 1
    End Sub

    Private Sub 追加スキル表示(ByVal sender As Object, ByVal e As System.EventArgs) _
        Handles ComboBox020.SelectedIndexChanged, ComboBox030.SelectedIndexChanged, _
                ComboBox120.SelectedIndexChanged, ComboBox130.SelectedIndexChanged, _
                ComboBox222.SelectedIndexChanged
        Dim p As DataSet
        Dim s As String = sender.Text 'スキル分類
        Dim cc As ComboBox
        If Equals(sender, ComboBox020) Then '武将1-スロ2
            cc = ComboBox021
        ElseIf Equals(sender, ComboBox030) Then '武将1-スロ3
            cc = ComboBox031
        ElseIf Equals(sender, ComboBox120) Then '武将2-スロ2
            cc = ComboBox121
        ElseIf Equals(sender, ComboBox130) Then '武将2-スロ3
            cc = ComboBox131
        Else '逆引き
            cc = ComboBox111
        End If
        Dim sqlwhere As String = ダブルクオート(s)
        If sqlwhere = ダブルクオート("特殊") Then '特殊項目には、条件付きスキルも含む
            sqlwhere = sqlwhere & " OR 分類 = " & ダブルクオート("条件")
        End If
        p = DB_TableOUT("SELECT id, 分類, スキル名 FROM SName WHERE 分類 = " & sqlwhere & " ORDER BY id", "SName")
        RemoveHandler cc.SelectedIndexChanged, AddressOf 追加スキル入力
        With cc
            .BeginUpdate()
            .ValueMember = "id"
            .DisplayMember = "スキル名"
            .DataSource = p.Tables("SName")
            .SelectedIndex = -1
            .EndUpdate()
        End With
        AddHandler cc.SelectedIndexChanged, AddressOf 追加スキル入力
    End Sub

    Private Sub 追加スキル入力(ByVal sender As Object, ByVal e As System.EventArgs) _
        Handles ComboBox021.SelectedIndexChanged, ComboBox031.SelectedIndexChanged, ComboBox121.SelectedIndexChanged, ComboBox131.SelectedIndexChanged
        Dim cbi As Integer = String_onlyNumber(sender.Name.ToString)
        If Not sender.text = "" Then
            If Equals(sender, ComboBox(Me, Format(cbi, "000"))) And (Not ComboBox(Me, Format((cbi - 1), "000")).Text = "") Then
                With ComboBox(Me, Format((cbi + 1), "000"))
                    .Enabled = True
                    .Text = 1
                End With
            End If
        Else
            If Equals(sender, ComboBox(Me, Format(cbi, "000"))) Then
                With ComboBox(Me, Format((cbi + 1), "000"))
                    .Enabled = False
                    .Text = ""
                End With
            End If
        End If
        'If Not sender.text = "" Then
        '    If Equals(sender, ComboBox021) And (Not ComboBox020.Text = "") Then
        '        ComboBox022.Enabled = True
        '        ComboBox022.Text = 1
        '    ElseIf Equals(sender, ComboBox031) And (Not ComboBox030.Text = "") Then
        '        ComboBox032.Enabled = True
        '        ComboBox032.Text = 1

        '    Else
        '        If Equals(sender, ComboBox121) And (Not ComboBox120.Text = "") Then
        '            ComboBox122.Enabled = True
        '            ComboBox122.Text = 1
        '        ElseIf Equals(sender, ComboBox131) And (Not ComboBox130.Text = "") Then
        '            ComboBox132.Enabled = True
        '            ComboBox132.Text = 1
        '        End If
        '    End If
        'Else
        '    If Equals(sender, ComboBox021) Then
        '        ComboBox022.Enabled = False
        '        ComboBox022.Text = ""
        '    Else
        '        If Equals(sender, ComboBox121) Then
        '            ComboBox122.Enabled = False
        '            ComboBox122.Text = ""
        '        ElseIf Equals(sender, ComboBox131) Then
        '            ComboBox132.Enabled = False
        '            ComboBox132.Text = ""
        '        End If
        '    End If
        'End If
    End Sub

    'Private Function skill_to_SKILL_T(ByVal skill As String)
    '    Dim sl As Integer = InStr(skill, "：")
    '    Return Mid(skill, sl + 1, skill.Length - sl)
    'End Function

    Private Sub 合成関連初期化()
        st = Nothing
        s1flg = Nothing
        s2flg = Nothing
        suro1sno = 0
        suro2sno = 0
        suro1 = Nothing
        kouho_sum = 0
        errorflg = 0
        okikaeindex = 0
    End Sub

    'Private Function 置換スキル決定() As Integer
    '    If CheckBox02.Checked Then 'スロ2にチェック
    '        If CheckBox03.Checked Then Return -1 'スロ2も3にもチェックはエラー
    '        Return 1
    '    ElseIf CheckBox03.Checked Then
    '        Return 2
    '    End If
    '    Return 0
    'End Function

    Private Sub 合成実行(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Call 合成関連初期化()
        'Call DB_Open() 'スキルDBを開く

        If gyakuflg = True Then '逆引きモードの時はそちらへ
            Call スキル逆引きモード()
            Exit Sub
        End If

        If Label002.Text = "" Or Label102.Text = "" Then '武将がどちらか空
            MsgBox("設定が未完成です")
            Exit Sub
        End If
        Dim tmpcombo() As ComboBox = {ComboBox012, ComboBox022, ComboBox112, ComboBox122, ComboBox132}
        For i As Integer = 0 To tmpcombo.Length - 1
            If tmpcombo(i).Enabled = True And tmpcombo(i).Text = "" Then
                MsgBox("スキルLVが未設定の箇所があります")
                Exit Sub
            End If
        Next
        'okikaeindex = 置換スキル決定()
        'If okikaeindex = -1 Then
        '    MsgBox("付け替えスキルを複数選ぶことはできません")
        '    Exit Sub
        'End If

        '*** スキル数を数える ***
        suro2sno = 1 '最初は初期スキルのみ
        For i As Integer = 2 To 3 'スロ2武将のスキル数
            If Not ComboBox(Me, "1" & CStr(i) & "1").Text = "" Then
                suro2sno += 1
            End If
        Next
        suro1sno = 1 '最初は初期スキルのみ
        ReDim suro1(0)
        suro1(0) = Label002.Text
        For i As Integer = 2 To 3 'スロ2武将のスキル数
            If Not ComboBox(Me, "0" & CStr(i) & "1").Text = "" Then
                ReDim Preserve suro1(i - 1)
                suro1(i - 1) = ComboBox(Me, "0" & CStr(i) & "1").Text
                suro1sno += 1
            End If
        Next
        'If Not ComboBox021.Text = "" Then 'スロ1武将のスキル数
        '    suro1sno = 2
        '    ReDim suro1(1)
        '    suro1(0) = Label002.Text
        '    suro1(1) = ComboBox021.Text
        'Else
        '    suro1sno = 1
        '    ReDim suro1(0)
        '    suro1(0) = Label002.Text
        'End If
        '************************
        If (Not ComboBox131.Text = "") And suro2sno = 2 Then 'スロ2が空いていてスロ3が埋まっている状態
            MsgBox("生贄側付加スキル内容の設定異常")
            Exit Sub
        End If
        If (Not ComboBox031.Text = "") And suro1sno = 2 Then 'スロ2が空いていてスロ3が埋まっている状態
            MsgBox("合成側付加スキル内容の設定異常")
            Exit Sub
        End If

        Call 基本テーブル読み込み()
        If errorflg = 1 Then
            RichTextBox1.Text = "※Wikiのデータが埋まり切っていないスキルを含んでいます＾＾；"
            errorflg = 0
            Exit Sub
        End If
        Dim tflgcheck As Integer = 合成テーブル作成()
        If Not tflgcheck = 0 Then Exit Sub 'テーブル作成過程でデータが無い場合は弾く
        Call 合成確率計算()
        Dim dflg As Boolean
        If errorflg = 2 Then
            dflg = False
        Else
            dflg = True
        End If
        RichTextBox1.Text = 追加合成シミュ出力(dflg)

        Dim boldstr() As String = Nothing
        For i As Integer = 0 To st.Length - 1
            ReDim Preserve boldstr(i)
            boldstr(i) = st(i).name
        Next
        Call RTextBox_BOLD(RichTextBox1, boldstr) 'スキル名太文字処理
    End Sub

    'wiki情報読み込み、完全にデータが揃っていないスキルを含んでいればFalseを返す
    Private Sub 基本テーブル読み込み()
        ReDim st(2) 'スキル候補数は3がデフォ
        For i As Integer = 0 To suro2sno - 1 'スロ2武将の各スキルの合成テーブルを読み込み
            ReDim Preserve suro2stable(i)
            Dim r() As String = {"スキルR", "候補A", "候補B", "候補C", "候補S1", "候補S2"}
            If i = 0 Then '初期スキル
                suro2stable(i) = _
                    DB_DirectOUT("SELECT * FROM SName WHERE スキル名 = " & ダブルクオート(Label102.Text) & "", r)
            Else
                suro2stable(i) = _
                    DB_DirectOUT("SELECT * FROM SName WHERE スキル名 = " & ダブルクオート(ComboBox(Me, "1" & CStr(i + 1) & "1").Text) & "", r)
            End If
            '致命的な空白がある（＝wikiが埋まり切っていない、未解明の部分があるスキルが含まれている）場合はエラー
            'ちょっと緩和（A,B,Cがそもそも埋まっていない:1, S1が埋まっていない:2, S2が埋まっていない:3, S1,S2とも埋まっていない:6）
            If i = 0 Then '初期スキルの場合は全枠埋まっていないと弾く
                For j As Integer = 1 To suro2stable(i).Length - 1
                    If suro2stable(i)(j) = "" Then
                        If j = 4 Then 'S1が埋まっていない
                            errorflg = 2
                            'Exit For
                        ElseIf j = 5 Then 'S2が埋まっていない
                            If errorflg = 2 Then 'S1も埋まっていない
                                errorflg = 6
                                Exit For
                            End If
                            errorflg = 3
                            Exit For
                        Else 'A, B, Cが埋まっていない
                            errorflg = 1
                            Exit For
                        End If
                    End If
                Next
            Else '初期スキルでない場合は、S1が埋まっていないと弾く
                If suro2stable(i)(4) = "" Then
                    errorflg = 1
                End If
            End If
        Next
    End Sub

    Private Sub スキル逆引きモード()
        If ComboBox111.Text = "" Then
            MsgBox("設定が未完成です")
            Exit Sub
        End If
        Dim output As String = "******** 逆引き探索結果 ********" '出力文字列
        Dim r() As String = {"候補A", "候補B", "候補C", "候補S1", "候補S2"}
        Dim tmpsl As String()
        For i As Integer = 0 To 4
            tmpsl = 逆引き検索(r(i), ComboBox111.Text)
            If tmpsl Is Nothing Then '候補が見当たらない
                Continue For
            End If
            For j As Integer = 0 To tmpsl.Length - 1
                If j = 0 Then
                    output = output & vbCrLf & "～～ 合成候補" & r(i) & " ～～"
                    output = output & vbCrLf & tmpsl(j)
                Else
                    output = output & ", " & tmpsl(j)
                End If
            Next
        Next
        If output = "******** 逆引き探索結果 ********" Then
            output = output & vbCrLf & "固有スキル。追加合成不可"
        End If
        RichTextBox1.Clear()
        RichTextBox1.Text = output
    End Sub

    Private Function 逆引き検索(ByVal col_name As String, ByVal skillname As String) As String() 'col_name: A,B,C,S1,S2
        逆引き検索 = _
        DB_DirectOUT2("SELECT * FROM SName WHERE " & col_name & " = " & ダブルクオート(skillname) & "", "スキル名")
    End Function

    '基本テーブルを代入、合成テーブル作成
    Private Function 合成テーブル作成() As Integer
        '基本となるテーブル
        Select Case suro2sno
            Case 1 '初期スキルのみ
                For i As Integer = 0 To 2
                    st(i).name = suro2stable(0)(i + 1)
                Next
            Case 2
                Dim stmp() As String = {suro2stable(0)(2), suro2stable(0)(3), suro2stable(1)(4)}
                For i As Integer = 0 To 2
                    st(i).name = stmp(i)
                Next
            Case 3
                Dim stmp() As String = {suro2stable(0)(3), suro2stable(1)(4), suro2stable(2)(4)}
                For i As Integer = 0 To 2
                    st(i).name = stmp(i)
                Next
        End Select

        If s1 = s2 Then '同一合成の場合
            s2flg = True
            If errorflg = 3 Or errorflg = 6 Then 'S2がデータベースで埋まっていない
                RichTextBox1.Text = "※S2データがデータベースに登録されていません＾＾；"
                errorflg = 0
                Return -3
            End If
            ReDim Preserve st(4) '4候補目にS2が出てくる（S1が出現する場合を考慮して都合5候補目に登録）
            st(3).name = "-"
            st(4).name = suro2stable(0)(5)
        End If

        For i As Integer = 0 To st.Length - 1
            Dim tskl As String = st(i).name
            Dim tn As Integer = i
            For j As Integer = 0 To suro1sno - 1
                If st(i).name = suro1(j) Then '候補のいずれかとスロ1武将のスキルが被る
                    st(i).name = "-"
                    s1flg = True
                End If
            Next
            For k As Integer = tn + 1 To st.Length - 1
                If tskl = st(k).name And (Not st(k).name = "-") Then '候補内スキル名がダブっていたら
                    st(k).name = "-"
                    s1flg = True
                End If
            Next
        Next

        If s1flg = True Then 'S1出現の場合
            If errorflg = 2 Or errorflg = 6 Then 'S1がデータベースで埋まっていない
                RichTextBox1.Text = "※S1データがデータベースに登録されていません＾＾；"
                errorflg = 0
                Return -2
            End If
            Dim s1name As String = suro2stable(0)(4)
            For i As Integer = 0 To suro1sno - 1
                If s1name = suro1(i) Then
                    s1flg = False
                End If
            Next
            For i As Integer = 0 To st.Length - 1
                If st(i).name = s1name Then
                    s1flg = False
                End If
            Next

            If s1flg = True Then
                If Not s2flg = True Then 's2が出てくる場合、5枠都合あるが、無い場合は4枠目を作る必要がある
                    ReDim Preserve st(3)
                End If
                st(3).name = s1name
            End If
        End If

        Dim c As Integer = 0
        For i As Integer = 0 To st.Length - 1
            If Not st(i).name = "-" Then '"-"ではない場合＝ちゃんとスキルが埋まっている場合
                st(c).name = st(i).name
                st(c).rare = スキルレア度取得(st(c).name)
                c += 1
            End If
        Next
        ReDim Preserve st(c - 1)
        Return 0
    End Function

    Private Function 追加合成シミュ出力(ByVal dataflg As Boolean) As String 'dataflgがFalseの場合、合成成功確率は***で表示
        RichTextBox1.Clear()
        Dim ssppzeroflg As Boolean = False 'スキルレア度の問題で成功確率が不明なスキルのある場合
        Dim output As String = Nothing
        If dataflg = False Then
            output = "※一部の合成の成功確率がデータ不足により表示できません＾＾；" & vbCrLf
        End If
        output = output & "******** 出現する合成候補 ********"
        For i As Integer = 0 To st.Length - 1
            Dim p, pkin As String
            If st(i).pp = 0 Or dataflg = False Then 'データが無い場合
                p = "***"
                pkin = "***"
                ssppzeroflg = True
            Else
                p = st(i).pp
                pkin = st(i).pp_kin
            End If
            output = output & vbCrLf & "【第" & (i + 1) & "候補】" & " " & "成功率：" & p & "% (金使用で" & pkin & "%)" & vbTab & _
                "スキル名：" & st(i).name & "LV1"
        Next
        If ssppzeroflg = True Then
            output = output & vbCrLf & "※一部のスキルで成功確率不明です。。"
        End If
        output = output & vbCrLf & "******** 材料スキルの合成テーブル（順にA,B,C,S1,S2） ********"
        For i As Integer = 0 To suro2sno - 1
            Dim suro2name As String
            If i = 0 Then
                suro2name = Label102.Text
            Else
                suro2name = ComboBox(Me, "1" & CStr(i + 1) & "1").Text
            End If
            output = output & vbCrLf & "『" & suro2name & "』 → "
            For j As Integer = 0 To 4
                output = output & " | " & suro2stable(i)(j + 1)
            Next
        Next
        Return output
    End Function

    '合成テーブル確定後、確率を計算。レア度が空欄のスキルがあればFalseを返す
    Private Sub 合成確率計算()
        skilllv_sum = _
            Val(ComboBox112.Text) + Val(ComboBox122.Text) + Val(ComboBox132.Text)
        For i As Integer = 0 To st.Length - 1
            Dim tmp() As String = _
                DB_DirectOUT("SELECT * FROM UTable WHERE スキルR = " & ダブルクオート(st(i).rare) & "", {"第" & (i + 1) & "候補"})
            If tmp Is Nothing Then
                errorflg = 2
                Exit Sub
            End If
            If tmp(0) = Nothing Then 'レア度の情報が無い場合
                st(i).pp = 0
            Else
                st(i).pp = 文字列計算(Replace(Replace(tmp(0), "L", skilllv_sum), "%", ""))
            End If
        Next
    End Sub

    Private Function スキルレア度取得(ByVal skillname As String) As String
        Dim si() As String = _
        DB_DirectOUT("SELECT * FROM SName WHERE スキル名 = " & ダブルクオート(skillname) & "", {"スキルR"})
        'Dim rare As New Hashtable
        'Dim alpha() As String = {"F", "E", "D", "C", "B", "A", "S"}
        If si(0) Is Nothing Or si(0) = "-" Then Return Nothing
        Return si(0)
        'For i As Integer = 0 To alpha.Length - 1
        '    rare.Add(alpha(i), i + 1)
        'Next
        'If si(0) = Nothing Or si(0) = "-" Then
        '    Return -1
        'Else
        '    Return Val(rare(si(0)))
        'End If
    End Function

    Private Sub 逆引きモード切替(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then 'チェックONならば逆引きモードへ
            gyakuflg = True
            GroupBox1.Enabled = False
            GroupBox2.Enabled = False
            ComboBox111.Enabled = True
            ComboBox222.Enabled = True
        Else
            gyakuflg = False
            GroupBox1.Enabled = True
            GroupBox2.Enabled = True
            ComboBox111.Enabled = False
            ComboBox222.Enabled = False
        End If
    End Sub

    Private Sub 出力クリア(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        RichTextBox1.Clear()
    End Sub

    Private Sub スキル所持武将検索(ByVal sender As Object, ByVal e As MouseEventArgs) Handles RichTextBox1.MouseUp
        If (e.Button = MouseButtons.Left) Or RichTextBox1.SelectedText = "" Then '左クリックもしくは選択した文字列が空
            Exit Sub
        End If
        Dim sstr As String = RichTextBox1.SelectedText
        sstr = Trim(Replace(sstr, ",", ""))
        sstr = Trim(Replace(sstr, "|", ""))
        sstr = Trim(Replace(sstr, ":", ""))
        sstr = Trim(Replace(sstr, "：", ""))
        sstr = Trim(Replace(sstr, vbLf, "")) '選択を楽にする（「,」, LF,「|」,「:」を含んでも大丈夫に）
        Dim sbsho As String() = DB_DirectOUT2("SELECT * FROM BData WHERE 初期スキル名 = " & ダブルクオート(sstr) & "", "武将名")
        Dim sbshor As String() = DB_DirectOUT2("SELECT * FROM BData WHERE 初期スキル名 = " & ダブルクオート(sstr) & "", "武将R")
        ContextMenuStrip1.Items.Clear()
        If sbsho Is Nothing Then '所持武将がいないスキル、つまり合成のみでしか出て来ないスキル
            ContextMenuStrip1.Items.Add("所持武将不明")
        Else
            ContextMenuStrip1.Items.Add("所持武将:")
            For i As Integer = 0 To sbsho.Length - 1
                ContextMenuStrip1.Items.Add(sbshor(i) & ":" & sbsho(i))
            Next
        End If
        RichTextBox1.ContextMenuStrip.Show()
    End Sub

    Private Sub データDD許可(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) _
        Handles GroupBox1.DragEnter, GroupBox2.DragEnter
        'データ形式の確認
        If e.Data.GetDataPresent(DataFormats.StringFormat) = False Then
            Return
        End If
        'ドロップ可能な場合は、エフェクトを変える
        e.Effect = DragDropEffects.Copy
    End Sub

    Private Sub 武将DD(ByVal sender As Object, ByVal e As DragEventArgs) _
        Handles GroupBox1.DragDrop, GroupBox2.DragDrop
        Dim rare As String = Nothing, bname As String = Nothing
        Dim skillname() As String = Nothing
        Dim skilllv() As Integer = Nothing
        If sender Is GroupBox1 Then
            Call 武将クリア(Button3, Nothing)
        Else
            Call 武将クリア(Button4, Nothing)
        End If
        Try
            Dim stmp() = Split(e.Data.GetData(GetType(String)), vbCrLf)
            If stmp.Length <= 4 Then '取引
                Dim ttmp() As String = Split(stmp(0), " ")
                Select Case Val(Mid(ttmp(0), 1, 1)) 'レアリティ
                    Case 1
                        rare = "天"
                    Case 2
                        rare = "極"
                    Case 3
                        rare = "特"
                    Case 4
                        rare = "上"
                    Case 5
                        rare = "序"
                End Select
                bname = Replace(Replace(ttmp(1), "★", ""), "☆", "")
                bname = Mid(bname, 1, bname.Length - 1) '最後の変な半角空白を消す
                For i As Integer = 1 To stmp.Length - 1
                    If InStr(stmp(i), "LV") = 0 Then Exit For
                    ReDim Preserve skillname(i - 1), skilllv(i - 1)
                    Dim tmp As String = Replace(Replace(stmp(i), "攻:", ""), "防:", "") '攻防を消す
                    If i = 1 Then '初期スキル
                        Dim tttmp() As String = Split(tmp, "	")
                        tmp = Trim(tttmp(1))
                    End If
                    skillname(i - 1) = Mid(tmp, 1, InStr(tmp, "LV") - 1)
                    skilllv(i - 1) = Val(Replace(tmp, skillname(i - 1) & "LV", ""))
                Next
            Else
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
                'レアリティと武将名
                Dim rares() As String = {"天", "極", "特", "上", "序"}
                bname = Replace(tmp(0), "名", "")
                Dim rb, ra As Integer
                For i = 0 To rares.Length - 1
                    rb = bname.Length
                    bname = Replace(bname, "レア" & rares(i), "")
                    ra = bname.Length
                    If Not rb = ra Then rare = rares(i)
                Next
                'スキル名
                For j As Integer = 0 To tmp.Length - 11
                    ReDim Preserve skillname(j), skilllv(j)
                    Dim ttmp As String = Replace(tmp(10 + j), "技" & (j + 1) & vbTab, "") '"技1"みたいなのが付いてる場合、外す
                    skillname(j) = Mid(ttmp, 1, InStr(ttmp, "LV") - 1)
                    skilllv(j) = Val(Mid(ttmp, InStr(ttmp, "LV") + 2))
                Next
            End If
            If sender Is GroupBox1 Then '合成される側
                'If skillname.Length = 3 Then 'スロットに空きが無ければ
                '    Me.Focus()
                '    MsgBox("スロットに空きがありません")
                '    Exit Sub
                'End If
                Call 武将データ代入(0, rare, bname, skillname, skilllv)
            Else '生贄側
                Call 武将データ代入(1, rare, bname, skillname, skilllv)
            End If
        Catch ex As Exception
            Me.Focus()
            MsgBox("データの読み込みが正常に行われなかった可能性があります")
        End Try
    End Sub
    Private Sub 武将データ代入(ByVal bf As Integer, ByVal rare As String, ByVal name As String, ByVal sname() As String, ByVal slv() As Integer)
        ComboBox(Me, CStr(bf) & "01").SelectedIndex = ComboBox(Me, CStr(bf) & "01").FindString(rare) '（強制的に）R選択
        'R選択(ComboBox(Me, CStr(bf) & "01"), Nothing)
        'Dim ntmp As String = GetINIValue(sname(0), name & "・" & rare, bnpath) '同名武将区別
        'If Not ntmp = "－" Then
        '    name = ntmp
        'End If
        Dim repstr As String = "replace(replace(初期スキル名, " & """ "" , """"), " & """　"" , """") = "
        'Dim repstr As String = "replace(初期スキル名, " & """ "" , """") = "
        name = DB_DirectOUT("SELECT 武将名, 初期スキル名 FROM BData WHERE 武将R = " _
            & ダブルクオート(rare) & " AND 武将名 LIKE """ & name & "%""" & " AND " & repstr & ダブルクオート(TrimJ(sname(0))), _
            {"武将名", "初期スキル名"})(0)
        ComboBox(Me, CStr(bf) & "02").SelectedText = name '（強制的に）武将名選択
        武将名選択(ComboBox(Me, CStr(bf) & "02"), Nothing)
        'If Not Label(Me, CStr(bf) & "02").Text = sname(0) Then
        '    Label(Me, CStr(bf) & "02").Text = sname(0)
        'End If
        ComboBox(Me, CStr(bf) & "12").Text = slv(0) '初期スキルのLV
        For i As Integer = 1 To sname.Length - 1
            ComboBox(Me, CStr(bf) & CStr(i + 1) & "0").Focus()
            ComboBox(Me, CStr(bf) & CStr(i + 1) & "0").SelectedIndex = ComboBox(Me, CStr(bf) & CStr(i + 1) & "0").FindString(スキル関連推定(sname(i), True))
            ComboBox(Me, CStr(bf) & CStr(i + 1) & "1").Text = sname(i)
            Call 追加スキル入力(ComboBox(Me, CStr(bf) & CStr(i + 1) & "1"), Nothing)
            ComboBox(Me, CStr(bf) & CStr(i + 1) & "2").Text = slv(i)
        Next
    End Sub
    'Private Sub 武将データ代入2(ByVal bf As Integer, ByVal sname() As String, ByVal slv() As Integer)
    '    Label(Me, CStr(bf) & "02").Text = sname(0)
    '    ComboBox(Me, CStr(bf) & "12").Text = slv(0) '初期スキルのLV
    '    For i As Integer = 1 To sname.Length - 1
    '        ComboBox(Me, CStr(bf) & CStr(i + 1) & "0").Focus()
    '        ComboBox(Me, CStr(bf) & CStr(i + 1) & "0").Text = スキル関連推定(sname(i))
    '        ComboBox(Me, CStr(bf) & CStr(i + 1) & "1").Text = sname(i)
    '        Call 追加スキル入力(ComboBox(Me, CStr(bf) & CStr(i + 1) & "1"), Nothing)
    '        ComboBox(Me, CStr(bf) & CStr(i + 1) & "2").Text = slv(i)
    '    Next
    'End Sub
    Private Sub 武将クリア(sender As Object, e As EventArgs) Handles Button3.Click, Button4.Click
        If sender Is Button3 Then '合成される側
            Call ClearTextBox(GroupBox1)
            Label002.Text = " -------------"
            ComboBox022.Enabled = False
            ComboBox032.Enabled = False
        Else '生贄側
            Call ClearTextBox(GroupBox2)
            Label102.Text = " -------------"
            ComboBox122.Enabled = False
            ComboBox132.Enabled = False
        End If
    End Sub

End Class