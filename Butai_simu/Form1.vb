Imports System.Data.OleDb
Imports System.IO

Public Class Form1
    Public bc As Integer '現在の武将
    Public ss()() As Integer 'Skill Status（各武将のスキル状態）
    Public can_skillp() As String 'スキルの有効状態行列 '有効:1, 無効:0
    Public boldtext()() As String '表示時に太字にする文字列格納
    Public contextmenuflg As Integer 'コンテキストメニューのフラグ
    Public heikaht As Hashtable = New Hashtable '兵科の紐づけ

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ' 画面サイズを取得して、それを基準にフォームサイズを決める。
        Dim screenw As Integer = Screen.GetBounds(Me).Width
        If screenw - 24 < Me.Width Then '画面サイズがForm1の幅よりも短い場合
            StartPosition = FormStartPosition.Manual
            Me.AutoScroll = True
            Me.Width = screenw - 24
            ' フォームの配置位置を指定
            Dim LeftPosition As Integer = Me.Left + 12 ' フォームの左端位置指定
            Dim TopPosition As Integer = Me.Top + 12 ' フォームの上端位置指定
            Me.Location = New Point(LeftPosition, TopPosition)
        End If
        Call DB_Open()
        ReDim ss(busho_counter - 1)
        '一括入力設定時の兵科紐づけ
        For i As Integer = 0 To ToolStripComboBox3.Items.Count - 1
            heikaht(ToolStripComboBox3.Items(i)) = ComboBox003.Items(i)
        Next

        '特殊スキルリストを読み込み
        'Dim sr As New System.IO.StreamReader(espath)
        'Dim srbuff As String = sr.ReadToEnd()
        'sr.Close()
        'error_skill = Split(srbuff, vbCrLf)
        'フラグ付きスキル情報を読み込み
        Call フラグ付きスキル読み込み()

        'ApplicationExitイベントハンドラを追加
        AddHandler Application.ApplicationExit, AddressOf Application_ApplicationExit
    End Sub

    'ApplicationExitイベントハンドラ
    Private Sub Application_ApplicationExit(ByVal sender As Object, ByVal e As EventArgs)
        Command.Dispose()
        Connection.Close()
        Connection.Dispose()
        'ApplicationExitイベントハンドラを削除
        RemoveHandler Application.ApplicationExit, AddressOf Application_ApplicationExit
    End Sub

    Private Sub 攻撃防衛部隊スイッチ(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        If InStr(sender.text, "攻撃") Then
            kb = "攻撃"
        Else
            kb = "防御"
        End If
    End Sub

    Public Sub 武将数スイッチ(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox2.SelectedIndexChanged
        busho_counter = Val(sender.text)
        Select Case Val(sender.text)
            Case 1
                OFF_GroupBox(GroupBox5)
                OFF_GroupBox(GroupBox9)
                OFF_GroupBox(GroupBox13)
            Case 2
                OFF_GroupBox(GroupBox9)
                OFF_GroupBox(GroupBox13)
                GroupBox5.Visible = True
            Case 3
                OFF_GroupBox(GroupBox13)
                GroupBox5.Visible = True
                GroupBox9.Visible = True
            Case 4
                GroupBox5.Visible = True
                GroupBox9.Visible = True
                GroupBox13.Visible = True
        End Select
        ReDim Preserve ss(busho_counter - 1)
        ReDim Preserve bs(busho_counter - 1) 'ここで武将をRedimしている！！
    End Sub

    Public Sub 武将データ消去(ByVal bc As Integer, Optional ByVal nextbs As Boolean = True) 'bcの武将データを消去
        'nextbsがTrueならば、次武将が入ってくるのが前提の消去（→R,武将名残す）
        bs(bc) = Nothing
        If nextbs = True Then
            Dim cc As ComboBox
            Select Case bc
                Case 0
                    cc = ComboBox001
                    Call ClearTextBox(GroupBox1, cc, ComboBox002)
                Case 1
                    cc = ComboBox101
                    Call ClearTextBox(GroupBox5, cc, ComboBox102)
                Case 2
                    cc = ComboBox201
                    Call ClearTextBox(GroupBox9, cc, ComboBox202)
                Case Else
                    cc = ComboBox301
                    Call ClearTextBox(GroupBox13, cc, ComboBox302)
            End Select
        Else
            Select Case bc
                Case 0
                    Call ClearTextBox(GroupBox1)
                Case 1
                    Call ClearTextBox(GroupBox5)
                Case 2
                    Call ClearTextBox(GroupBox9)
                Case Else
                    Call ClearTextBox(GroupBox13)
            End Select
        End If
        Button(Me, CStr(bc) & "001").Text = "/ ----" '指揮兵数クリア
        Label(Me, CStr(bc) & "002").Text = "------------" '初期スキル名クリア
        Button(Me, CStr(bc + 3)).ForeColor = Color.Black '要ステ計算解除
        ComboBox(Me, CStr(bc) & "15").Enabled = False 'スキルLV欄のリセット
        ComboBox(Me, CStr(bc) & "16").Enabled = False
        For i As Integer = 3 To 5 'ステ手動入力不許可
            TextBox(Me, CStr(bc) & "0" & CStr(i)).Enabled = False
        Next
        For i As Integer = 3 To 5 'ステUP率表示の初期化
            Label(Me, CStr(bc) & "0" & CStr(i)).Text = "(---)"
        Next
        GroupBox(Me, "0" & CStr(bc) & "2").Text = "ステータス" '職表示の初期化
        With GroupBox(Me, 4 * (bc + 1)) '統率表示の初期化
            .Text = "統率"
            .ForeColor = Color.Black
        End With
        CheckBox(Me, CStr(bc) & "1").Checked = False '限界突破解除
        ss(bc) = {}
    End Sub

    '子コントロールから、その武将No.を取得
    Private Function 武将取得(ByVal cParent As Control) As Integer
        Dim hparent As Control = cParent.Parent
        If hparent Is GroupBox1 Then
            Return 0
        ElseIf hparent Is GroupBox5 Then
            Return 1
        ElseIf hparent Is GroupBox9 Then
            Return 2
        Else
            If Not hparent Is Me Then 'さかのぼってForm1になるまで探す
                Return 武将取得(hparent)
            Else
                Return 3
            End If
        End If
    End Function

    Public Sub R選択(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles ComboBox001.SelectedIndexChanged, _
                ComboBox101.SelectedIndexChanged, _
                ComboBox201.SelectedIndexChanged, _
                ComboBox301.SelectedIndexChanged
        Dim cc As ComboBox
        bc = 武将取得(sender)
        Select Case bc
            Case 0
                cc = ComboBox002
            Case 1
                cc = ComboBox102
            Case 2
                cc = ComboBox202
            Case Else
                cc = ComboBox302
        End Select

        RemoveHandler cc.SelectedValueChanged, AddressOf Me.武将名選択 'これが無いと武将名を選べなくなる
        Dim p As DataSet
        p = DB_TableOUT("SELECT id, 武将R, 武将名 FROM BData WHERE 武将R = " & ダブルクオート(sender.SelectedItem) & " AND Bunf = 'F' ORDER BY Bid ASC", "BData")
        cc.DisplayMember = "武将名"
        cc.ValueMember = "id"
        cc.DataSource = p.Tables("BData")
        cc.SelectedIndex = -1
        AddHandler cc.SelectedValueChanged, AddressOf Me.武将名選択
    End Sub

    Public Sub 武将名選択(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles ComboBox002.SelectedValueChanged, _
                ComboBox102.SelectedValueChanged, _
                ComboBox202.SelectedValueChanged, _
                ComboBox302.SelectedValueChanged
        bc = 武将取得(sender)

        武将データ消去(bc)
        bs(bc).武将設定初期化()
        Dim r() As String = _
            {"Bid", "武将R", "Cost", "指揮兵数", "槍統率", "弓統率", "馬統率", "器統率", "初期攻撃", "初期防御", "初期兵法", "攻成長", "防成長", "兵成長", "初期スキル名", "職"}
        Dim s() As String = _
        DB_DirectOUT("SELECT * FROM BData WHERE 武将R = " & ダブルクオート(ComboBox(Me, CStr(bc) & "01").SelectedItem) & _
                     " AND 武将名 = " & ダブルクオート(sender.Text) & " AND Bunf = 'F'", r)
        'ここから武将初期化
        With bs(bc)
            .No = bc
            .name = sender.Text
            .id = s(0)
            .rare = s(1)
            .cost = s(2)
            .hei_max_d = s(3)
            .hei_sum = .hei_max_d
            .hei_max = .hei_max_d '初期値設定
            .Tousotu(False, True) = {s(4), s(5), s(6), s(7)}
            .Sta(False) = {s(8), s(9), s(10)}
            ReDim .sta_g(2)
            .sta_g = {s(11), s(12), s(13)}
            .skill(0).name = s(14)
            .job = s(15)
            Dim sklerror As String = Nothing
            .スキル取得(0, .skill(0).name, .skill(0).lv, {0}, sklerror)
            If Not sklerror Is Nothing Then ToolStripLabel3.Text = "[警告ログ]" & sklerror
            .情報入力(bc, False)
        End With
        ss(bc) = {0}
    End Sub

    Public Sub データクリア(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        ReDim bs(busho_counter - 1) '武将情報初期化
        For i As Integer = 0 To busho_counter - 1
            武将データ消去(i, False)
        Next
        ToolTip1.RemoveAll()
        ToolStripLabel6.Text = "---------"
        ToolStripLabel3.Text = "----------"
    End Sub

    Public Sub ステ振り設定(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles ComboBox013.SelectedValueChanged, _
                ComboBox113.SelectedValueChanged, _
                ComboBox213.SelectedValueChanged, _
                ComboBox313.SelectedValueChanged
        If sender.text = "" Then '武将切り替え等で空白の場合もあり得る
            Exit Sub
        End If
        bc = 武将取得(sender)
        If ComboBox(Me, CStr(bc) & "02").Text = "" Then '武将指定していない場合はそちらを先に
            Exit Sub
        End If
        bs(bc).huri = sender.Text
        If sender.text = "手動" Then
            TextBox(Me, CStr(bc) & "03").Enabled = True
            TextBox(Me, CStr(bc) & "04").Enabled = True
            TextBox(Me, CStr(bc) & "05").Enabled = True
        Else
            TextBox(Me, CStr(bc) & "03").Enabled = False
            TextBox(Me, CStr(bc) & "04").Enabled = False
            TextBox(Me, CStr(bc) & "05").Enabled = False
        End If
        Call ステ計算(Button(Me, CStr(bc + 3)), Nothing)
        Button(Me, CStr(bc + 3)).ForeColor = Color.Black '要ステ計算解除
    End Sub

    Public Sub ステ振り手動(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        TextBox003.TextChanged, TextBox004.TextChanged, TextBox005.TextChanged, _
        TextBox103.TextChanged, TextBox104.TextChanged, TextBox105.TextChanged, _
        TextBox203.TextChanged, TextBox204.TextChanged, TextBox205.TextChanged, _
        TextBox303.TextChanged, TextBox304.TextChanged, TextBox305.TextChanged
        bc = 武将取得(sender)
        Dim kbh As Integer = Val(Mid(sender.name, 8, 3)) Mod 100
        If bs(bc).huri = "手動" Then '手動での確定ならば
            Select Case kbh
                Case 3 '将攻変更
                    bs(bc).st(0) = Val(sender.text)
                Case 4
                    bs(bc).st(1) = Val(sender.text)
                Case 5
                    bs(bc).st(2) = Val(sender.text)
            End Select
        End If
    End Sub

    Private Sub 追加スキル表示(ByVal sender As Object, ByVal e As System.EventArgs) _
        Handles ComboBox009.SelectedIndexChanged, ComboBox010.SelectedIndexChanged, _
                ComboBox109.SelectedIndexChanged, ComboBox110.SelectedIndexChanged, _
                ComboBox209.SelectedIndexChanged, ComboBox210.SelectedIndexChanged, _
                ComboBox309.SelectedIndexChanged, ComboBox310.SelectedIndexChanged
        Dim p As DataSet
        Dim cc As ComboBox
        Dim s As String = sender.Name
        bc = 武将取得(sender)
        Dim v As String = s.Where(Function(C) C Like "[.0-9]").ToArray()
        If InStr(v, CStr(bc) & "09") Then 'スロ2
            s = ComboBox(Me, CStr(bc) & "09").Text
            cc = ComboBox(Me, CStr(bc) & "11")
        Else 'スロ3
            s = ComboBox(Me, CStr(bc) & "10").Text
            cc = ComboBox(Me, CStr(bc) & "12")
        End If
        Dim sqlwhere As String = ダブルクオート(s)
        If sqlwhere = ダブルクオート("特殊") Then '特殊項目には、条件付きスキルも含む
            sqlwhere = sqlwhere & " OR 分類 = " & ダブルクオート("条件")
        End If
        p = DB_TableOUT("SELECT id, 分類, スキル名 FROM SName WHERE 分類 = " & sqlwhere & " ORDER BY id", "SName")
        RemoveHandler cc.SelectedValueChanged, AddressOf Me.スキル名入力 'これが無いとスキル名を選べなくなる
        With cc
            .ValueMember = "id"
            .DisplayMember = "スキル名"
            .DataSource = p.Tables("SName")
            .SelectedIndex = -1
        End With
        AddHandler cc.SelectedValueChanged, AddressOf Me.スキル名入力
    End Sub

    Public Sub 追加スキル追加(ByVal sender As Object, ByVal e As System.EventArgs) _
        Handles ComboBox014.SelectedValueChanged, ComboBox015.SelectedValueChanged, ComboBox016.SelectedValueChanged, _
                ComboBox114.SelectedValueChanged, ComboBox115.SelectedValueChanged, ComboBox116.SelectedValueChanged, _
                ComboBox214.SelectedValueChanged, ComboBox215.SelectedValueChanged, ComboBox216.SelectedValueChanged, _
                ComboBox314.SelectedValueChanged, ComboBox315.SelectedValueChanged, ComboBox316.SelectedValueChanged, ComboBox316.MouseUp, ComboBox315.MouseUp, ComboBox314.MouseUp, ComboBox216.MouseUp, ComboBox215.MouseUp, ComboBox214.MouseUp, ComboBox116.MouseUp, ComboBox115.MouseUp, ComboBox114.MouseUp, ComboBox016.MouseUp, ComboBox015.MouseUp, ComboBox014.MouseUp
        bc = 武将取得(sender)
        If ComboBox(Me, CStr(bc) & "02").Text = "" Then '武将指定していない場合はそちらを先に
            Exit Sub
        End If

        'スキル数を数える
        If ComboBox(Me, CStr(bc) & "15").Enabled = False Or ComboBox(Me, CStr(bc) & "11").Text = Nothing Then
            If ComboBox(Me, CStr(bc) & "16").Enabled = False Or ComboBox(Me, CStr(bc) & "12").Text = Nothing Then
                bs(bc).skill_no = 1
                ss(bc) = {0}
            Else
                bs(bc).skill_no = 2
                ss(bc) = {0, 2}
            End If
        Else
            If ComboBox(Me, CStr(bc) & "16").Enabled = False Or ComboBox(Me, CStr(bc) & "12").Text = Nothing Then
                bs(bc).skill_no = 2
                ss(bc) = {0, 1}
            Else
                bs(bc).skill_no = 3
                ss(bc) = {0, 1, 2}
            End If
        End If

        If sender.text = "" Then 'スキル削除、武将切り替え等で空白の場合もあり得る
            Exit Sub
        End If

        Dim sklerror As String = Nothing
        Select Case CInt(Mid(CStr(sender.Name), 10, 2))
            Case 14 '初期スキルに関する変更
                bs(bc).スキル取得(0, Label(Me, CStr(bc) & "002").Text, CInt(sender.text), ss(bc), sklerror)
            Case 15 'スロ2に関する変更
                bs(bc).スキル取得(1, ComboBox(Me, CStr(bc) & "11").Text, CInt(sender.text), ss(bc), sklerror)
                If bs(bc).skill.Length - 1 >= 1 Then
                    bs(bc).skill(1).kanren = ComboBox(Me, CStr(bc) & "09").Text
                End If
            Case 16 'スロ3に関する変更
                bs(bc).スキル取得(2, ComboBox(Me, CStr(bc) & "12").Text, CInt(sender.text), ss(bc), sklerror)
                If bs(bc).skill.Length - 1 >= 2 Then
                    bs(bc).skill(2).kanren = ComboBox(Me, CStr(bc) & "10").Text
                End If
        End Select
        If Not sklerror Is Nothing Then ToolStripLabel3.Text = "[警告ログ]" & sklerror
    End Sub

    Public Sub スキルレベル一括変更(ByVal sender As Object, ByVal e As MouseEventArgs) _
        Handles ComboBox014.MouseUp, ComboBox015.MouseUp, ComboBox016.MouseUp, _
                ComboBox114.MouseUp, ComboBox115.MouseUp, ComboBox116.MouseUp, _
                ComboBox214.MouseUp, ComboBox215.MouseUp, ComboBox216.MouseUp, _
                ComboBox314.MouseUp, ComboBox315.MouseUp, ComboBox316.MouseUp
        If e.Button = MouseButtons.Right Then '右クリックなら
            Dim mp As Point = Control.MousePosition
            contextmenuflg = 1 'レベル一括変更は1
            ContextMenuStrip1.Tag = Val(sender.text) 'TagにLVを渡す
            ContextMenuStrip1.Show(mp)
        End If
        bc = 武将取得(sender)
    End Sub

    Private Sub スキル除去(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles ComboBox015.EnabledChanged, ComboBox016.EnabledChanged, _
                ComboBox115.EnabledChanged, ComboBox116.EnabledChanged, _
                ComboBox215.EnabledChanged, ComboBox216.EnabledChanged, _
                ComboBox315.EnabledChanged, ComboBox316.EnabledChanged
        If sender.Name = "" Then 'これが無いとFormをロードする前に乙る
            Exit Sub
        End If
        If sender.Text = "" Then
            Exit Sub
        End If
        'Dim cc As ComboBox = sender
        'RemoveHandler cc.SelectedValueChanged, AddressOf Me.追加スキル追加 'これが無いと追加時と競合
        Call 追加スキル追加(sender, e)
        'AddHandler cc.SelectedValueChanged, AddressOf Me.追加スキル追加
        Select Case CInt(Mid(CStr(sender.Name), 10, 2))
            Case 15 'スロ2消去
                ToolTip1.SetToolTip(Label(Me, CStr(bc) & "004"), vbNullString)
            Case 16 'スロ3消去
                ToolTip1.SetToolTip(Label(Me, CStr(bc) & "005"), vbNullString)
        End Select
    End Sub

    Private Sub ステ計算(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles Button3.Click, Button4.Click, Button5.Click, Button6.Click
        bc = 武将取得(sender)
        With bs(bc)
            If .huri = "" Then
                MsgBox("ステ計算を行うには極振り設定が必要です。（手動以外）")
                Exit Sub
            ElseIf .huri = "手動" Then
                Exit Sub
            End If
            Call .ステ極振り計算()
            Call .情報入力(bc)
        End With
        Button(Me, CStr(bc + 3)).ForeColor = Color.Black '要ステ計算解除
    End Sub

    Public Sub ランク変更(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles ComboBox004.SelectedValueChanged, _
                ComboBox104.SelectedValueChanged, _
                ComboBox204.SelectedValueChanged, _
                ComboBox304.SelectedValueChanged
        If sender.Text = "" Then '武将切り替え等で空白の場合もあり得る
            Exit Sub
        End If
        bc = 武将取得(sender)
        If ComboBox(Me, CStr(bc) & "02").Text = "" Then '武将指定していない場合はそちらを先に
            Exit Sub
        End If
        Dim be, af As Integer
        With bs(bc)
            be = .rank 'before
            If .limitbreakflg Then '限界突破時
                .rank = 6
            Else
                If .rank = 6 Then '限界突破していなくてrank=6の時
                    .rank = 0 '初期化
                    ComboBox(Me, CStr(bc) & "04").Text = 0
                End If
                .rank = CInt(sender.text)
            End If
            af = .rank 'after
            If Not .job = "剣" Then
                If .job = "覇" Then
                    .hei_max = .hei_max_d + .rank * 200 'ランクアップで兵数一律+200
                Else
                    .hei_max = .hei_max_d + .rank * 100 'ランクアップで兵数一律+100
                End If
            End If
            .rankup_r = .rankup_r + (af - be)
            .残りランクアップ可能回数表示(GroupBox(Me, 4 * (bc + 1)))
        End With
        If Not be = af Then '初期の代入など、前後で値が変わっていない場合を除く
            Button(Me, CStr(bc + 3)).ForeColor = Color.Red '要ステ計算
            Button(Me, CStr(bc) & "001").Text = "/ " & bs(bc).hei_max
        End If
    End Sub

    Public Sub 限界突破(sender As Object, e As EventArgs) _
        Handles CheckBox01.CheckedChanged, CheckBox11.CheckedChanged, CheckBox21.CheckedChanged, CheckBox31.CheckedChanged
        bc = 武将取得(sender)
        If CheckBox(Me, CStr(bc) & "1").Checked Then
            CheckBox(Me, CStr(bc) & "1").ForeColor = Color.FromArgb(192, 64, 0)
            ComboBox(Me, CStr(bc) & "04").Enabled = False
            TextBox(Me, CStr(bc) & "02").Enabled = False
            bs(bc).limitbreakflg = True
        Else
            CheckBox(Me, CStr(bc) & "1").ForeColor = SystemColors.ControlText
            ComboBox(Me, CStr(bc) & "04").Enabled = True
            TextBox(Me, CStr(bc) & "02").Enabled = True
            bs(bc).limitbreakflg = False
        End If
        'ランク処理
        Call ランク変更(ComboBox(Me, CStr(bc) & "04"), Nothing)
    End Sub

    Public Sub LV変更(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles TextBox002.TextChanged, _
                TextBox102.TextChanged, _
                TextBox202.TextChanged, _
                TextBox302.TextChanged
        If sender.Text = "" Then '武将切り替え等で空白の場合もあり得る
            Exit Sub
        End If
        bc = 武将取得(sender)
        Dim be As Integer = bs(bc).level '要ステ計算判定で用いる
        If ComboBox(Me, CStr(bc) & "02").Text = "" Then '武将指定していない場合はそちらを先に
            Exit Sub
        End If
        bs(bc).level = CInt(sender.text)
        If Not be = bs(bc).level Then '実際にLV変更が起こった場合のみ
            Button(Me, CStr(bc + 3)).ForeColor = Color.Red '要ステ計算
        End If
    End Sub

    Public Sub 積載兵数変更(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles TextBox001.TextChanged, _
                TextBox101.TextChanged, _
                TextBox201.TextChanged, _
                TextBox301.TextChanged
        If sender.Text = "" Then '武将切り替え等で空白の場合もあり得る
            Exit Sub
        End If
        bc = 武将取得(sender)
        If ComboBox(Me, CStr(bc) & "02").Text = "" Then '武将指定していない場合はそちらを先に
            Exit Sub
        End If
        bs(bc).hei_sum = CInt(sender.text)
    End Sub

    'Contextstripmenuから積載兵数変更
    Public Sub 積載兵数変更2(ByVal sender As Object, ByVal e As MouseEventArgs) _
        Handles Button0001.Click, Button1001.Click, Button2001.Click, Button3001.Click
        Dim mousePosition As Point = Control.MousePosition
        ContextMenuStrip2.Tag = 武将取得(sender)
        ContextMenuStrip2.Show(mousePosition)
    End Sub
    Private Sub コンテキストメニュー操作2(ByVal sender As System.Object, ByVal e As ToolStripItemClickedEventArgs) Handles ContextMenuStrip2.ItemClicked
        Cursor.Current = Cursors.WaitCursor
        Dim cit As String = e.ClickedItem.Text
        Dim selbs As Integer = Int(ContextMenuStrip2.Tag)
        If ComboBox(Me, CStr(selbs) & "02").Text = "" Then '武将未選択
            MsgBox("武将未選択です")
            Exit Sub
        End If
        Select Case cit
            Case "MAX"
                TextBox(Me, CStr(selbs) & "01").Text = bs(selbs).hei_max
            Case "1"
                TextBox(Me, CStr(selbs) & "01").Text = 1
            Case "10"
                TextBox(Me, CStr(selbs) & "01").Text = 10
            Case Else
                Dim ds, dk, tset As Integer
                If InStr(kb, "攻") Then
                    ds = bs(selbs).st(0)
                    dk = bs(selbs).heisyu.atk
                Else
                    ds = bs(selbs).st(1)
                    dk = bs(selbs).heisyu.def
                End If
                If dk = 0 Then '兵科がセットされていないとここがゼロになる
                    MsgBox("兵科がセットされていないのでキャップ計算ができません")
                    Exit Select
                End If
                tset = Math.Ceiling((0.02 * ds * bs(selbs).heisyu.ts) / (1 - 0.02 * dk * bs(selbs).heisyu.ts))
                If tset <= bs(selbs).hei_max Then 'キャップ兵数が指揮兵数以下
                    TextBox(Me, CStr(selbs) & "01").Text = tset
                Else
                    TextBox(Me, CStr(selbs) & "01").Text = bs(selbs).hei_max
                End If
        End Select
        Cursor.Current = Cursors.Default
    End Sub

    Public Sub 積載兵科変更(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles ComboBox003.SelectedIndexChanged, _
                ComboBox103.SelectedIndexChanged, _
                ComboBox203.SelectedIndexChanged, _
                ComboBox303.SelectedIndexChanged
        If sender.Text = "" Or sender.Text = "-----" Then '武将切り替え等で空白の場合もあり得る
            Exit Sub
        End If
        bc = 武将取得(sender)
        If ComboBox(Me, CStr(bc) & "02").Text = "" Then '武将指定していない場合はそちらを先に
            Exit Sub
        End If
        bs(bc).兵科情報取得(sender.text)
    End Sub

    Private Sub 統率変更(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles ComboBox005.SelectedIndexChanged, ComboBox006.SelectedIndexChanged, ComboBox007.SelectedIndexChanged, ComboBox008.SelectedIndexChanged,
                ComboBox105.SelectedIndexChanged, ComboBox106.SelectedIndexChanged, ComboBox107.SelectedIndexChanged, ComboBox108.SelectedIndexChanged, _
                ComboBox205.SelectedIndexChanged, ComboBox206.SelectedIndexChanged, ComboBox207.SelectedIndexChanged, ComboBox208.SelectedIndexChanged, _
                ComboBox305.SelectedIndexChanged, ComboBox306.SelectedIndexChanged, ComboBox307.SelectedIndexChanged, ComboBox308.SelectedIndexChanged
        If sender.text = "" Then '武将切り替え等で空白の場合もあり得る
            Exit Sub
        End If
        bc = 武将取得(sender)
        If ComboBox(Me, CStr(bc) & "02").Text = "" Then '武将指定していない場合はそちらを先に
            Exit Sub
        End If

        For i = 5 To 8
            If ComboBox(Me, CStr(bc) & "0" & CStr(i)).Text = "" Then
                Exit Sub
            End If
        Next
        Dim Tousotu_af(3) As String
        Dim be, af As Decimal
        Dim d As Integer '変更幅
        Dim s() As String = {"槍統率", "弓統率", "馬統率", "器統率"}

        For i As Integer = 0 To 3
            Tousotu_af(i) = ComboBox(Me, CStr(bc) & "0" & CStr(i + 5)).Text
        Next
        With bs(bc)
            For k = 0 To 3 '変更前の統率値の合計
                be = be + .tou(k)
            Next
            For k = 0 To 3 '変更後、あり得ない変更がなされていれば強制終了
                If Tousotu_af(k) = "" Then '初回時は必ずこれだと引っかかるのでここで抜け出す
                    Exit For
                End If
                If .tou_d(k) > 統率_数値変換(Tousotu_af(k)) Then
                    MsgBox("この武将の★0での" & s(k) & "は「 " & .tou_d_a(k) & " 」です。" & vbCrLf & _
                            "これ以上低い統率には設定できません。")
                    ComboBox(Me, CStr(bc) & "0" & CStr(k + 5)).Text = .tou_a(k)
                    Exit Sub
                End If
            Next
            .Tousotu = Tousotu_af
            For k = 0 To 3 '変更後の統率値の合計
                af = af + .tou(k)
            Next
            d = (af - be) / 0.05
            .rankup_r = .rankup_r - d
            .残りランクアップ可能回数表示(GroupBox(Me, 4 * (bc + 1)))
            .兵科情報取得(.heisyu.name) 'ランクが変われば当然兵科に対する統率値も変わる
        End With
    End Sub

    Private Sub 部隊兵法値計算・スキルデータ確定() 'いわば前段階
        '部隊兵法値計算
        Dim maxheihou As Decimal
        Dim heihou_kei As Decimal
        maxheihou = 0
        For i As Integer = 0 To busho_counter - 1
            If (maxheihou < bs(i).st(2)) Then
                maxheihou = bs(i).st(2)
            End If
            heihou_kei = heihou_kei + bs(i).st(2)
        Next
        heihou_sum = (maxheihou + (heihou_kei - maxheihou) / 6) / 100

        Call フラグ付きスキル読み込み() '読込（更新）
        Call 童ボーナス加算()

        '部隊総戦闘力計算
        For i As Integer = 0 To busho_counter - 1
            bs(i).小隊攻撃力計算()
            Heisum = Heisum + bs(i).hei_sum
            Atksum = Atksum + bs(i).attack
            bs(i).スキル期待値計算()
        Next
        '海野六郎のようなスキル（他武将のスキル条件に影響を与えるスキル）がある。その部分を補正
        '※今は発動率（kouka_p)のみ
        For i As Integer = 0 To busho_counter - 1
            For j As Integer = 0 To bs(i).skill_no - 1
                With bs(i).skill(j)
                    If (.kouka_p_b + .up_kouka_p) > 1 Then '合計100%を超えるならば
                        .kouka_p_b = 1
                    Else
                        .kouka_p_b = .kouka_p_b + .up_kouka_p
                    End If
                    .exp_kouka_b = .kouka_p_b * .kouka_f
                End With
            Next
        Next
        Call 発動スキル候補決定()
    End Sub

    Private Sub 部隊初期化()
        Heisum = 0
        heihou_sum = 0
        Atksum = 0
        Call スキル状態初期化()
        'スキル計算に使う変数をクリア
        For i As Integer = 0 To busho_counter - 1
            For j As Integer = 0 To bs(i).skill_no - 1
                bs(i).skill(j).スキル計算状態初期化()
            Next
        Next
    End Sub

    'Private Sub デフォルト統率ポップアップ(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    '    Handles Label009.MouseHover, Label010.MouseHover, Label011.MouseHover, Label012.MouseHover _
    '           , Label109.MouseHover, Label110.MouseHover, Label111.MouseHover, Label112.MouseHover _
    '           , Label209.MouseHover, Label210.MouseHover, Label211.MouseHover, Label212.MouseHover _
    '           , Label309.MouseHover, Label310.MouseHover, Label311.MouseHover, Label312.MouseHover
    '    bc = 武将取得(sender)
    '    If ComboBox(Me, CStr(bc) & "02").Text = "" Then '武将指定していない場合はそちらを先に
    '        Exit Sub
    '    End If
    '    Dim p As Integer = CInt(Mid(CStr(sender.Name), 8, 1)) - 9 '参照する統率の種類
    '    Dim s As String = bs(bc).tou_d_a(p) 'ポップアップ文字列
    '    ToolTip1.SetToolTip(Label(Me, CStr(bc) & "0" & CStr(p + 9)), s)
    'End Sub

    Private Sub スキルポップアップ(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles Label0003.MouseHover, Label0004.MouseHover, Label0005.MouseHover, _
                Label1003.MouseHover, Label1004.MouseHover, Label1005.MouseHover, _
                Label2003.MouseHover, Label2004.MouseHover, Label2005.MouseHover, _
                Label3003.MouseHover, Label3004.MouseHover, Label3005.MouseHover
        bc = 武将取得(sender)
        If ComboBox(Me, CStr(bc) & "02").Text = "" Then '武将指定していない場合はそちらを先に
            Exit Sub
        End If
        Dim p As Integer = CInt(Mid(CStr(sender.Name), 9, 1)) - 3 '参照するスキルスロットの位置.
        Dim m As Integer = -1 '参照するスキルNo.
        For i As Integer = 0 To ss(bc).Length - 1
            If ss(bc)(i) = p Then
                m = i
            End If
        Next
        If m = -1 Then '該当スロットは空
            Exit Sub
        End If

        With bs(bc)
            If (.skill_no - 1) < m Then
                Exit Sub
            End If
            With .skill(m)
                Dim t, s As String
                Try
                    If InStr(.heika, "槍弓馬砲器") Then
                        t = "全"
                    Else
                        t = .heika
                    End If
                    If .tokusyu = 9 Then '特殊スキル
                        s = "特殊スキルです。"
                    ElseIf .tokusyu = 1 Then '速度スキル
                        s = "速度のみ適用スキルです。" & vbCrLf & _
                            "加速: " & t & "速+" & .speed.ToString("p")
                    ElseIf .tokusyu = 2 Then '破壊スキル
                        s = "破壊のみ適用スキルです。"
                    ElseIf .tokusyu = 5 Then 'データ不足スキル
                        s = "Wikiにデータが無いスキルです。" & vbCrLf & _
                            "（シミュレート結果に影響しません）"
                    Else
                        s = "発動率: +" & .kouka_p.ToString("p") & vbCrLf & _
                            "上昇率: " & t & .koubou & "+" & .kouka_f.ToString("p")
                        If Not .speed = 0 Then '加速要素もある
                            If .speed < 0 Then '減速スキルならば
                                s = s & vbCrLf & _
                                "減速: " & t & "速" & .speed.ToString("p")
                            Else
                                s = s & vbCrLf & _
                                    "加速: " & t & "速+" & .speed.ToString("p")
                            End If
                        End If
                    End If
                    ToolTip1.SetToolTip(Label(Me, CStr(bc) & "00" & CStr(p + 3)), s)
                Catch ex As Exception
                End Try
            End With
        End With
    End Sub

    Public Sub スキル名削除(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles ComboBox011.TextChanged, ComboBox012.TextChanged, _
                ComboBox111.TextChanged, ComboBox112.TextChanged, _
                ComboBox211.TextChanged, ComboBox212.TextChanged, _
                ComboBox311.TextChanged, ComboBox312.TextChanged
        bc = 武将取得(sender)
        If ComboBox(Me, CStr(bc) & "02").Text = "" Then '武将指定していない場合はそちらを先に
            Exit Sub
        End If
        If sender.Text = "" Then
            If CInt(Mid(CStr(sender.Name), 10, 2)) = 11 Then
                ComboBox(Me, CStr(bc) & "15").Enabled = False
                ComboBox(Me, CStr(bc) & "15").Text = ""
            Else
                ComboBox(Me, CStr(bc) & "16").Enabled = False
                ComboBox(Me, CStr(bc) & "16").Text = ""
            End If
        End If
    End Sub

    Public Sub スキル名入力(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles ComboBox011.SelectedIndexChanged, ComboBox012.SelectedIndexChanged, _
                ComboBox111.SelectedIndexChanged, ComboBox112.SelectedIndexChanged, _
                ComboBox211.SelectedIndexChanged, ComboBox212.SelectedIndexChanged, _
                ComboBox311.SelectedIndexChanged, ComboBox312.SelectedIndexChanged
        bc = 武将取得(sender)
        If ComboBox(Me, CStr(bc) & "02").Text = "" Then '武将指定していない場合はそちらを先に
            Exit Sub
        End If

        If CInt(Mid(CStr(sender.Name), 10, 2)) = 11 Then
            ComboBox(Me, CStr(bc) & "15").Enabled = True
            '他スキルの登録が済んでいない場合そちらを先に
            If ComboBox(Me, CStr(bc) & "16").Enabled = True And ComboBox(Me, CStr(bc) & "16").Text = "" Then
                ComboBox(Me, CStr(bc) & "16").SelectedIndex = 1
                '追加スキル追加(ComboBox(Me, CStr(bc) & "16"), Nothing)
            End If
            追加スキル追加(ComboBox(Me, CStr(bc) & "15"), Nothing)
        Else
            ComboBox(Me, CStr(bc) & "16").Enabled = True
            If ComboBox(Me, CStr(bc) & "15").Enabled = True And ComboBox(Me, CStr(bc) & "15").Text = "" Then
                ComboBox(Me, CStr(bc) & "15").SelectedIndex = 1
                '追加スキル追加(ComboBox(Me, CStr(bc) & "15"), Nothing)
            End If
            追加スキル追加(ComboBox(Me, CStr(bc) & "16"), Nothing)
        End If
    End Sub

    Private Sub 武将入れ替え(ByVal fromb As Integer, ByVal tob As Integer, Optional ByVal fmv As Boolean = False) 'frombの武将データを、tobのデータ位置へ。（tobへ上書き）
        'fmv=Trueの時は、コピー先が空白。
        If fmv = False Then
            Dim tmpb As Busho
            tmpb = bs(tob).Clone '武将情報入れ替え
            bs(tob) = bs(fromb).Clone
            bs(fromb) = tmpb.Clone
            武将データ消去(fromb)
            武将コピー(fromb, bs(fromb))
            武将データ消去(tob)
            武将コピー(tob, bs(tob))
        Else
            bs(tob) = bs(fromb).Clone
            武将データ消去(fromb)
            武将コピー(tob, bs(tob))
        End If
    End Sub

    'removeすべきハンドラをまとめた
    Public Sub 武将データ手入力用(ByVal bno As Integer, ByVal op As Boolean, Optional ByVal bsf As Boolean = False) 'op = TrueならばAdd, FalseならばRemove
        'bsf = Trueならば武将名選択, 追加スキル追加(動的に計算して埋める範囲)はRemove, addしない
        'R選択、追加スキル表示は動かないと他コントローラが使えないので切らない
        If op = False Then
            If bsf = False Then
                RemoveHandler ComboBox(Me, CStr(bno) & "02").SelectedValueChanged, AddressOf 武将名選択
                For i As Integer = 14 To 16
                    RemoveHandler ComboBox(Me, CStr(bno) & CStr(i)).SelectedValueChanged, AddressOf 追加スキル追加
                Next
            End If
            RemoveHandler ComboBox(Me, CStr(bno) & "13").SelectedValueChanged, AddressOf ステ振り設定
            For i As Integer = 15 To 16
                RemoveHandler ComboBox(Me, CStr(bno) & CStr(i)).EnabledChanged, AddressOf スキル除去
            Next
            RemoveHandler ComboBox(Me, CStr(bno) & "04").SelectedValueChanged, AddressOf ランク変更
            RemoveHandler TextBox(Me, CStr(bno) & "02").TextChanged, AddressOf LV変更
            RemoveHandler TextBox(Me, CStr(bno) & "01").TextChanged, AddressOf 積載兵数変更
            RemoveHandler ComboBox(Me, CStr(bno) & "03").SelectedIndexChanged, AddressOf 積載兵科変更
            For i As Integer = 5 To 8
                RemoveHandler ComboBox(Me, CStr(bno) & "0" & CStr(i)).SelectedIndexChanged, AddressOf 統率変更
            Next
        Else
            If bsf = False Then
                AddHandler ComboBox(Me, CStr(bno) & "02").SelectedValueChanged, AddressOf 武将名選択
                For i As Integer = 14 To 16
                    AddHandler ComboBox(Me, CStr(bno) & CStr(i)).SelectedValueChanged, AddressOf 追加スキル追加
                Next
            End If
            AddHandler ComboBox(Me, CStr(bno) & "13").SelectedValueChanged, AddressOf ステ振り設定
            For i As Integer = 15 To 16
                AddHandler ComboBox(Me, CStr(bno) & CStr(i)).EnabledChanged, AddressOf スキル除去
            Next
            AddHandler ComboBox(Me, CStr(bno) & "04").SelectedValueChanged, AddressOf ランク変更
            AddHandler TextBox(Me, CStr(bno) & "02").TextChanged, AddressOf LV変更
            AddHandler TextBox(Me, CStr(bno) & "01").TextChanged, AddressOf 積載兵数変更
            AddHandler ComboBox(Me, CStr(bno) & "03").SelectedIndexChanged, AddressOf 積載兵科変更
            For i As Integer = 5 To 8
                AddHandler ComboBox(Me, CStr(bno) & "0" & CStr(i)).SelectedIndexChanged, AddressOf 統率変更
            Next
        End If
    End Sub

    Private Sub 武将コピー(ByVal bno As Integer, ByVal bsd As Busho) 'コピー先位置と、武将データ
        Call 武将データ手入力用(bno, False, True)
        With bsd
            'この2つは面倒なので入力→更新する
            '----------
            ComboBox(Me, CStr(bno) & "01").SelectedText = ComboBox(Me, CStr(bno) & "01").FindString(.rare) '（強制的に）R選択
            'R選択(ComboBox(Me, CStr(bno) & "01"), Nothing)
            ComboBox(Me, CStr(bno) & "02").Text = .name '（強制的に）武将名選択
            武将名選択(ComboBox(Me, CStr(bno) & "02"), Nothing)
            For i As Integer = 1 To .skill_no - 1
                If Not .skill(i).kanren = "" Then '関連スキルが空ではない＝追加スキルがある
                    Select Case i
                        Case 1 'スロ2
                            ComboBox(Me, CStr(bno) & "09").Focus()
                            ComboBox(Me, CStr(bno) & "09").SelectedIndex = ComboBox(Me, CStr(bno) & "09").FindString(.skill(i).kanren)
                            ComboBox(Me, CStr(bno) & "11").Text = .skill(i).name
                        Case 2 'スロ3
                            ComboBox(Me, CStr(bno) & "10").Focus()
                            ComboBox(Me, CStr(bno) & "10").SelectedIndex = ComboBox(Me, CStr(bno) & "10").FindString(.skill(i).kanren)
                            ComboBox(Me, CStr(bno) & "12").Text = .skill(i).name
                    End Select
                End If
            Next
            '----------
            '他は“表面上フォームを埋めていくだけ”
            If .limitbreakflg Then '限界突破
                CheckBox(Me, CStr(bc) & "1").Checked = True
            Else
                ComboBox(Me, CStr(bno) & "04").Text = .rank
            End If
            ComboBox(Me, CStr(bno) & "03").Text = .heisyu.name
            For i As Integer = 5 To 8
                ComboBox(Me, CStr(bno) & "0" & CStr(i)).Text = .tou_a(i - 5)
            Next
            ComboBox(Me, CStr(bno) & "13").Text = .huri
            For i As Integer = 14 To 14 + .skill_no - 1
                ComboBox(Me, CStr(bno) & CStr(i)).Text = .skill(i - 14).lv
            Next

            TextBox(Me, CStr(bno) & "01").Text = .hei_sum
            TextBox(Me, CStr(bno) & "02").Text = .level
            For i As Integer = 3 To 5 'ステ選択
                TextBox(Me, CStr(bno) & "0" & CStr(i)).Text = .st(i - 3)
            Next
            .rankup_r = 統率未振り数推定(.tou_d_a, .tou_a, .rank)
            If ステ振り推定(.st_d, .st, .sta_g, .rank, .level, True) = "中途半端" Then
                Button(Me, CStr(bc + 3)).ForeColor = Color.Red '要ステ計算
            Else
                Button(Me, CStr(bc + 3)).ForeColor = Color.Black
            End If
            .残りランクアップ可能回数表示(GroupBox(Me, 4 * (bno + 1)))
            '---------
            武将データ手入力用(bno, True)
            bs(bno) = .Clone
            bs(bno).No = bno '武将Noのみ付け替え必要
        End With
    End Sub

    'can_skill, can_skillp生成
    Private Sub 発動スキル候補決定()
        can_skill = Nothing '初期化
        can_skillp = Nothing
        Dim c As Integer = 0 'カウンター
        '全有効スキル数カウント
        For i As Integer = 0 To busho_counter - 1
            For j As Integer = 0 To bs(i).skill_no - 1 '攻防一致、特殊スキル排除
                If InStr(kb, bs(i).skill(j).koubou) Then
                    If bs(i).skill(j).tokusyu = 0 Or bs(i).skill(j).t_flg Then '通常スキルもしくは総コスト・フラグ依存スキル
                        ReDim Preserve can_skill(c)
                        can_skill(c) = bs(i).skill(j).Clone
                        c = c + 1
                    End If
                ElseIf InStr(kb, "攻") And InStr(bs(i).skill(j).koubou, "上級器") Then '攻撃かつ上級器→上級器攻スキルの条件を満たす
                    ReDim Preserve can_skill(c)
                    can_skill(c) = bs(i).skill(j).Clone
                    c = c + 1
                ElseIf InStr(kb, "防") And InStr(bs(i).skill(j).koubou, "上級砲") Then '防御かつ上級砲→上級砲防スキルの条件を満たす
                    ReDim Preserve can_skill(c)
                    can_skill(c) = bs(i).skill(j).Clone
                    c = c + 1
                ElseIf InStr(kb, "攻") And InStr(bs(i).skill(j).koubou, "秘境兵") Then '攻撃かつ秘境兵→秘境兵攻スキルの条件を満たす
                    ReDim Preserve can_skill(c)
                    can_skill(c) = bs(i).skill(j).Clone
                    c = c + 1
                End If
            Next
        Next
        c = 2 ^ (c) '全スキル有効状態数
        ReDim can_skillp(c - 1)
        For i As Integer = 0 To c - 1
            can_skillp(i) = Convert10to2(i, Math.Ceiling(Math.Log10(c) / Math.Log10(2)))
        Next
    End Sub

    Public Function スキル実質上昇率(ByVal sk As Busho.skl) As Decimal(,) 'そのスキルが発動することによりプラスされる上昇率
        '二次元配列が戻り値になる。(x,0)->将スキル, (x,1)->一般スキル
        Dim tmpatk(,) As Decimal
        Dim refbs() As Busho = Nothing
        Dim refbsc As Integer = 0

        If simu_execno = 0 Then
            refbs = bs.Clone
            refbsc = busho_counter
        ElseIf simu_execno = 1 Then
            refbs = Form10.simu_bs.Clone
            refbsc = 4
        End If
        ReDim tmpatk(refbsc - 1, 1)

        With sk
            For i As Integer = 0 To refbsc - 1
                If (InStr(.heika, refbs(i).heisyu.bunrui) Or InStr(.heika, "将")) And 条件付スキル適用チェック(refbs(i), sk) Then '兵科が合致するか＋条件付チェック
                    '（攻防一致、特殊スキル排除、通常スキル確認は発動スキル候補決定の段階で除外
                    If InStr(.heika, "将") Then '将スキルかどうか
                        tmpatk(i, 0) = .kouka_f
                    Else
                        tmpatk(i, 1) = .kouka_f
                    End If
                ElseIf InStr(kb, "攻") And InStr(.koubou, "上級器") Then '上級器スキルならば
                    If refbs(i).heisyu.jyk_utuwa Then '上級器に含まれる兵科を積んでいれば（今は『上級器攻』のみ）
                        tmpatk(i, 1) = .kouka_f
                    End If
                ElseIf InStr(kb, "防") And InStr(.koubou, "上級砲") Then '上級砲スキルならば
                    If refbs(i).heisyu.jyk_hou Then '上級砲に含まれる兵科を積んでいれば（今は『上級砲防』のみ）
                        tmpatk(i, 1) = .kouka_f
                    End If
                ElseIf InStr(kb, "攻") And InStr(.koubou, "秘境兵") Then '秘境兵スキルならば
                    If refbs(i).heisyu.tok_hikyo Then '秘境兵に含まれる兵科を積んでいれば（今は『秘境兵攻』のみ）
                        tmpatk(i, 1) = .kouka_f
                    End If
                End If
            Next
        End With
        スキル実質上昇率 = tmpatk
    End Function

    '各部隊スキルの発動率を計算（重複発動はしない、内部でダブったらランダム）
    Private Function 部隊スキル発動確率計算(ByVal activebsk() As bskl._bsk) As Decimal()
        If activebsk Is Nothing Then Return {"0"}
        Dim retp() As Decimal '戻す各スキルの発動率が入った配列
        Dim decp() As String '発動状態パターン文字列
        Dim sumpdecp() As String 'decpのパターンの生起確率に、指定の重みづけを加えたもの
        Dim p() As Decimal, rp() As Decimal
        ReDim retp(activebsk.Length), p(activebsk.Length - 1), rp(activebsk.Length - 1)
        'pnと^pnを計算
        For i As Integer = 0 To activebsk.Length - 1
            With activebsk(i)
                p(i) = .kouka_p
                rp(i) = 1 - .kouka_p
            End With
        Next
        '各結合確率を計算
        'ex: (length=3) → c=8, b1～b3のうち、b1が発動する確率は
        'b1 = p1^p2^p3 + (1/2)p1p2^p3 + (1/2)p1^p2p3 + (1/3)p1p2p3
        Dim c As Long = 2 ^ (activebsk.Length) '全部隊スキル有効状態数 ex: "000", "001", "010",etc
        ReDim decp(c - 1), sumpdecp(c - 1)
        For i As Integer = 0 To c - 1
            decp(i) = CStr(Convert10to2(i, Math.Ceiling(Math.Log10(c) / Math.Log10(2))))
            Dim tmpp As Decimal = 1
            Dim countz As Integer = 0
            Dim ktmpp As Decimal = 1
            For j As Integer = 1 To activebsk.Length
                If Mid(decp(i), j, 1) = 1 Then '発動ならば
                    tmpp = tmpp * p(j - 1)
                Else
                    tmpp = tmpp * rp(j - 1)
                    countz = countz + 1
                End If
            Next
            If Not i = 0 Then ktmpp = 1 / (activebsk.Length - countz)
            sumpdecp(i) = ktmpp * tmpp
        Next
        '各部隊スキルの発動確率を計算
        For i As Integer = 1 To activebsk.Length
            retp(i) = 0
            For j As Integer = 0 To c - 1
                If Mid(decp(j), i, 1) = 1 Then '自分のpが1（発動ビット）ならば
                    retp(i) = retp(i) + sumpdecp(j)
                End If
            Next
        Next
        '全ての部隊スキルが未発動
        retp(0) = sumpdecp(0)
        'For i As Integer = 0 To activebsk.Length - 1
        '    retp(0) = retp(0) - retp(i)
        'Next
        Return retp
    End Function

    Private Sub スキル状態初期化()
        skill_x = Nothing '初期化
        skill_y = Nothing
        skill_yk = Nothing
        skill_xx = Nothing
        skill_yy = Nothing
        skill_yyk = Nothing
        'skill_syo = Nothing
        skill_ex = Nothing
        skill_exk = 0
        skill_exx = 0
        skill_ax = 0
        skill_axx = 0
    End Sub

    '部隊スキル複数がアリになったので、状態数も可変
    Private Sub スキル状態計算()

        Call スキル状態初期化()
        Dim activebsk() As bskl._bsk = bskill.activeskl(kb)
        Dim bskillcount As Integer
        If activebsk Is Nothing Then
            bskillcount = 0
        Else
            bskillcount = activebsk.Length '有効な部隊スキル数
        End If
        Dim activebskp() As Decimal = 部隊スキル発動確率計算(activebsk)

        If bskill.flg And bskillcount > 0 Then '部隊スキルが有効ならば
            ReDim skill_yk((bskillcount + 1) * can_skillp.Length - 1), _
            skill_x((bskillcount + 1) * can_skillp.Length - 1), _
            skill_y((bskillcount + 1) * can_skillp.Length - 1, busho_counter - 1)
        Else '部隊スキル無
            ReDim skill_yk(can_skillp.Length - 1)
            ReDim skill_x(can_skillp.Length - 1), skill_y(can_skillp.Length - 1, busho_counter - 1)
        End If
        'ReDim skill_syo(can_skillp.Length - 1, busho_counter - 1) '将攻成分(スキルが乗る時はそのUPも含む)

        '童関係
        Dim harr() As String = {"槍", "弓", "馬", "砲", "器"}
        'Dim warr() As Decimal = {warabe.def.yari, warabe.def.yumi, warabe.def.uma, warabe.def.hou, warabe.def.utuwa}
        Dim warr() As Decimal = warabe.warabe_gets(kb)
        Dim wflg As Boolean = False
        For i As Integer = 0 To warr.Length - 1
            If Not (warr(i) = 0) Then
                wflg = True
                Exit For
            End If
        Next

        'スキル状態計算
        For i As Integer = 0 To can_skillp.Length - 1
            skill_x(i) = 1
            For k As Integer = 0 To busho_counter - 1
                skill_y(i, k) = bs(k).attack
            Next

            Dim syoplus() As Decimal '将UP率
            Dim heiplus() As Decimal '一般UP率
            ReDim syoplus(busho_counter - 1)
            ReDim heiplus(busho_counter - 1)

            '***** DUMMY *****
            If can_skill Is Nothing Then '特殊スキルしかない等でそもそも有効状態が存在しない場合
                If bskillcount = 0 Or bskill.flg = False Then 'かつ部隊スキルも有効でない
                    If Not (wflg) Then Exit For 'さらに童も有効でない
                End If
                '部隊スキルのみ、童のみ有効な場合、無意味なダミースキルを一つ作ってエラー回避
                ReDim can_skill(0)
                With can_skill(0)
                    If InStr(kb, "攻") Then
                        .koubou = "防"
                    Else
                        .koubou = "攻"
                    End If
                    .kouka_p_b = 0
                End With
            End If
            '***** DUMMY *****

            For j As Integer = 1 To can_skill.Length
                If Mid(can_skillp(i), j, 1) = 1 Then '発動ならば
                    Dim ttmp(,) As Decimal = スキル実質上昇率(can_skill(j - 1))
                    skill_x(i) = skill_x(i) * can_skill(j - 1).kouka_p_b
                    For k As Integer = 0 To busho_counter - 1
                        syoplus(k) = syoplus(k) + ttmp(k, 0)
                        heiplus(k) = heiplus(k) + ttmp(k, 1)
                    Next
                Else
                    skill_x(i) = skill_x(i) * (1 - can_skill(j - 1).kouka_p_b)
                End If
            Next

            '実際のスキル発動時の戦闘力を計算
            'syo＝各スキルの上乗せ分＋素攻
            For k As Integer = 0 To busho_counter - 1
                Dim ds, dk As Integer
                With bs(k)
                    If InStr(kb, "攻") Then
                        ds = .st(0)
                        dk = .heisyu.atk
                    Else
                        ds = .st(1)
                        dk = .heisyu.def
                    End If
                    skill_y(i, k) = (ds * .heisyu.ts * (1 + syoplus(k) + heiplus(k))) + (.hei_sum * dk * .heisyu.ts * (1 + heiplus(k)))
                    'skill_syo(i, k) = ds * .heisyu.ts * (1 + syoplus(k)) * (1 + heiplus(k))
                    '童適用(今のところ単科防スキルのみ対応)
                    For l As Integer = 0 To harr.Length - 1
                        If InStr(.heisyu.bunrui, harr(l)) Then
                            skill_y(i, k) = skill_y(i, k) * (1 + 0.01 * warr(l))
                        End If
                    Next
                End With
            Next
        Next

        '部隊スキルは、「全」しかない事を前提にしている。後々各兵科別に適用・未適用を考えないといけない場合は改良必要
        '↑「全」以外のものにも対応 InStr(.heika, bs(i).heisyu.bunrui) Or InStr(.heika, "将") Then '兵科が合致するか
        If bskill.flg And bskillcount > 0 Then '部隊スキルが有効ならば
            ReDim skill_xx(can_skillp.Length - 1), skill_yy(can_skillp.Length - 1, busho_counter - 1)
            skill_xx = skill_x.Clone
            skill_yy = skill_y.Clone
            '部隊スキル未発動時
            For i As Integer = 0 To can_skillp.Length - 1
                skill_x(i) = skill_x(i) * activebskp(0) 'xのみ変化、yそのまま
            Next
            '部隊スキル発動時
            For i As Integer = 1 To bskillcount
                For j As Integer = i * can_skillp.Length To (i + 1) * can_skillp.Length - 1
                    skill_x(j) = skill_xx(j - i * can_skillp.Length) * activebskp(i) '後半は発動時。yも変化。
                    For k As Integer = 0 To busho_counter - 1
                        If InStr(activebsk(i - 1).taisyo, bs(k).heisyu.bunrui) Or (activebsk(i - 1).taisyo = "全") Then
                            skill_y(j, k) = skill_yy(j - i * can_skillp.Length, k) * (1 + activebsk(i - 1).kouka_f) '乗算適用。
                            'If bskill.qq = True Then '将攻成分が2乗
                            '    skill_y(i, k) = skill_y(i, k) + _
                            '        skill_syo(i - can_skillp.Length, k) * (1 + bskill.kouka_f) * bskill.kouka_f
                            'End If
                        Else
                            skill_y(j, k) = skill_yy(j - i * can_skillp.Length, k) '未適用
                        End If
                    Next
                Next
            Next
        End If

        skill_yk = Array_to_Arrayk(skill_y)

        'skill_yk(＝x軸)が重複する場合があるかをチェック、あれば統合
        Dim tmp As Decimal, tmpx() As Decimal = Nothing, tmpy(,) As Decimal = Nothing, tend() As Decimal = Nothing
        Dim c As Integer = 0
        Dim flg As Boolean '重複の時はFalse
        For i As Integer = 0 To skill_x.Length - 1
            flg = True
            If c > 0 Then
                For k As Integer = 0 To c - 1
                    If skill_yk(i) = tend(k) Then '調査済みの値かどうか
                        flg = False
                    End If
                Next
            End If
            If flg = False Then
                Continue For
            End If
            tmp = skill_yk(i)
            For j As Integer = i + 1 To skill_x.Length - 1
                If skill_yk(j) = tmp Then '重複を見つけたら
                    skill_x(i) = skill_x(i) + skill_x(j) 'その部分の確率密度を加算
                    skill_x(j) = 0 '足し終わればゼロに
                End If
            Next
            ReDim Preserve tend(c)
            tend(c) = tmp '重複済みの値を格納
            c = c + 1
        Next
        c = 0 'カウンター再利用
        For i As Integer = 0 To skill_x.Length - 1
            If Not skill_x(i) = 0 Then '意味のあるxの個数を数え上げる（tmpyが2次元になったせいでRedimを一発でしないといけないため）
                c = c + 1
            End If
        Next
        ReDim tmpx(c - 1), tmpy(c - 1, busho_counter - 1)
        c = 0 'カウンター再々利用
        For i As Integer = 0 To skill_x.Length - 1
            If Not skill_x(i) = 0 Then
                tmpx(c) = skill_x(i)
                For k As Integer = 0 To busho_counter - 1
                    tmpy(c, k) = skill_y(i, k)
                Next
                c = c + 1
            End If
        Next
        ReDim skill_x(c - 1), skill_y(c - 1, busho_counter - 1)
        skill_x = tmpx.Clone
        skill_y = tmpy.Clone
        sQuickSort2(skill_y, skill_x)
        'ソートのためにskill_yk更新
        ReDim skill_yk(UBound(skill_y))
        skill_yk = Array_to_Arrayk(skill_y)
        atksum_max = skill_yk(skill_yk.Length - 1) 'MAX値

        '母集団の期待値、分散
        ReDim skill_ex(busho_counter - 1)
        For i As Integer = 0 To skill_x.Length - 1
            For k As Integer = 0 To busho_counter - 1
                skill_ex(k) = skill_ex(k) + skill_x(i) * skill_y(i, k) '各武将のスキル発動時の個別の戦闘力が分からないと、NPC側の見かけの防御力期待値が分からない
            Next
        Next
        For k As Integer = 0 To busho_counter - 1
            skill_exk = skill_exk + skill_ex(k)
        Next
        For i As Integer = 0 To skill_x.Length - 1
            skill_ax = skill_ax + skill_x(i) * ((skill_yk(i) - skill_exk) ^ 2) '分散
        Next
        '部隊スキルが有効な場合、スキルを無視した時の期待値、分散、MAX値
        If bskill.flg = True And bskillcount > 0 Then
            'sQuickSort2(skill_yy, skill_xx)
            'skill_yykも昇順ソートで再生成
            ReDim skill_yyk(UBound(skill_yy))
            skill_yyk = Array_to_Arrayk(skill_yy)
            Array.Sort(skill_yyk, skill_xx)
            For i As Integer = 0 To skill_xx.Length - 1
                skill_exx = skill_exx + skill_xx(i) * skill_yyk(i)
            Next
            For i As Integer = 0 To skill_xx.Length - 1
                skill_axx = skill_axx + skill_xx(i) * ((skill_yyk(i) - skill_exx) ^ 2) '分散
            Next
            atksum_maxmax = skill_yyk(skill_yyk.Length - 1)
        End If
        skill_ax = Math.Sqrt(skill_ax)

        Call Form2.グラフ描画()
    End Sub

    Private Sub 分析実行(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        If bs.Length = 0 Or bs(0).name = "" Then '武将ゼロで実行しようとしている
            MsgBox("武将数がゼロです")
            Exit Sub
        End If
        Cursor.Current = Cursors.WaitCursor

        simu_execno = 0

        '武将数があっていない場合自動で訂正
        Dim ts() As Integer = Nothing 'データの抜けている部分
        Dim c As Integer = 0, d As Integer = 0
        Dim tbc As Integer 'データの入っていない武将数
        For i = 0 To busho_counter - 1
            If ComboBox(Me, CStr(i) & "02").Text = "" Then
                ReDim Preserve ts(c)
                ts(c) = i '空き番号
                c = c + 1
                tbc = tbc + 1
            Else
                If c > 0 Then '空でなければ→その前の番号が空いている
                    武将入れ替え(i, ts(d), True)
                    d = d + 1
                    ReDim Preserve ts(c)
                    ts(c) = i '移動した元の場所がまた開く
                    c = c + 1
                End If
            End If
        Next
        ToolStripComboBox2.Text = busho_counter - tbc
        Costsum = 0
        Ranksum = 0

        '攻撃/防御変更によって数値が変わる可能性のあるものを更新する必要がある → 兵科能力値のみ？
        For i As Integer = 0 To busho_counter - 1
            Call bs(i).兵科情報取得(bs(i).heisyu.name)
            Costsum = Costsum + bs(i).cost 'ついでに総コスト計算
            Ranksum = Ranksum + bs(i).rank '部隊ランクも計算
        Next

        Try
            rank_sum = 部隊ランクボーナス計算(Ranksum)
            Call 部隊初期化()
            Call 部隊兵法値計算・スキルデータ確定()
            Call スキル状態計算()

            Form2.RichTextBox1.Text = データテキスト出力(0) '全体情報
            'Call RTextBox_BOLD(Form2.RichTextBox1, boldtext(0))
            For i As Integer = 0 To busho_counter - 1
                RichTextBox(Form2, i + 2).Text = データテキスト出力(i + 1) '各武将情報
                For j As Integer = 1 To boldtext.Length - 1
                    Call RTextBox_BOLD(RichTextBox(Form2, i + 2), boldtext(i + 1)) '太字処理
                Next
            Next
        Catch ex As Exception
            MsgBox("必要なデータが不足しているか原因不明のエラーです＞＜")
            Exit Sub
        End Try
        Cursor.Current = Cursors.Default
        Form2.Show()
    End Sub

    'bno = 0なら全体、1～4で武将1～4
    Public Function データテキスト出力(ByVal bn As Integer) As String
        データテキスト出力 = ""
        Dim tmp() As String = Nothing
        Dim fg As Boolean = False '部隊スキル
        ReDim Preserve boldtext(bn)
        Dim bstr() As String = Nothing '太字にする行
        If bn = 0 Then '全体情報
            If (Not bskill.activebsk Is Nothing) And bskill.flg = True Then
                'bstr = Split("4,5,6,9", ",")
                fg = True
                ReDim tmp(13)
                tmp(10) = "+++ 部隊スキルON (" & Int(bskill.activebsk.Length) & "個) +++"
                For i As Integer = 0 To bskill.activebsk.Length - 1
                    If i = 0 Then
                        tmp(11) = "[" & Int(i + 1) & "] " & _
                        "発動率: " & bskill.activebsk(i).kouka_p & "/ " & "上昇率: " & bskill.activebsk(i).kouka_f & "/ 対象: " & bskill.activebsk(i).taisyo
                    Else
                        tmp(11) = tmp(11) & vbCrLf & "[" & Int(i + 1) & "] " & _
                        "発動率: " & bskill.activebsk(i).kouka_p & "/ " & "上昇率: " & bskill.activebsk(i).kouka_f & "/ 対象: " & bskill.activebsk(i).taisyo
                    End If
                Next
                tmp(12) = "※部隊スキルを全て無視した場合" & vbCrLf & _
                        "   |- ※期待値: " & Int(skill_exx) & vbCrLf & _
                        "   |- ※MAX値: " & Int(atksum_maxmax) & vbCrLf & "++++++++++++"
            Else
                fg = False
                'bstr = Split("4,5,6", ",")
                ReDim tmp(9)
            End If
            tmp(0) = "武将数: " & busho_counter
            tmp(1) = "総コスト: " & Costsum
            tmp(2) = "総兵数: " & Heisum
            tmp(3) = "総" & kb & "力: "
            tmp(4) = "  |- 素" & Mid(kb, 1, 1) & ": " & Int(Atksum)
            tmp(5) = "  |- 期待値: " & Int(skill_exk) & " (+" & Format(((skill_exk / Atksum) - 1) * 100, "#0.00") & "%上昇)"
            tmp(6) = "  |- MAX値: " & Int(atksum_max) & " (+" & Format(((atksum_max / Atksum) - 1) * 100, "#0.00") & "%上昇)"
            tmp(7) = "部隊兵法補正値: +" & Math.Ceiling(heihou_sum * 100) / 100 & "%"
            tmp(8) = "部隊ランクボーナス: +" & Format(rank_sum, "#0.00") & "% (★" & Format(Ranksum, "#0") & ")"
            If InStr(kb, "防") Then
                tmp(4) = tmp(4) & vbCrLf & "      |- コス1あたり -> " & Int(Atksum / Costsum)
                tmp(5) = tmp(5) & vbCrLf & "      |- コス1あたり -> " & Int(skill_exk / Costsum)
                tmp(6) = tmp(6) & vbCrLf & "      |- コス1あたり -> " & Int(atksum_max / Costsum)
            End If
            tmp(9) = "童効果: " + 童効果文字列出力(kb)
        Else
            bstr = Split("0", ",")
            With bs(bn - 1)
                ReDim tmp(10)
                tmp(0) = "武将名: " & .name
                tmp(1) = "レアリティ: " & .rare & " | " & "コスト: " & .cost
                tmp(2) = "兵数: " & .hei_sum
                If .limitbreakflg Then '限界突破時
                    tmp(3) = "☆限界突破☆"
                Else
                    tmp(3) = "★" & .rank & " LV" & .level
                End If
                tmp(4) = "攻: " & .st(0) & "/ " & "防: " & .st(1) & "/ " & "兵法: " & .st(2)
                tmp(5) = "統率: " & "槍" & .tou_a(0) & "/ " & "弓" & .tou_a(1) & "/ " & "馬" & .tou_a(2) & "/ " & "砲" & .tou_a(3)
                tmp(6) = "積載兵科: " & .heisyu.name
                tmp(7) = "素" & kb & "力: " & Int(.attack) & " (全体の" & Format((.attack / Atksum) * 100, "#0.00") & "%)"
                tmp(8) = "----------"
                Dim c As Integer = 0 'カウンター
                For i As Integer = 0 To .skill_no - 1
                    With .skill(i)
                        ReDim Preserve tmp(10 + 2 * i)
                        If .tokusyu = 0 Or .t_flg Then '特殊スキル（総コスト依存スキルを除く）じゃなければ
                            Dim tmps As String = .heika
                            If InStr(.heika, "槍弓馬砲器") Then
                                tmps = "全"
                            End If
                            tmp(9 + 2 * i) = "【スキル" & i + 1 & "】" & vbCrLf & _
                                .name & "LV" & .lv & " /" & .koubou & " /" & tmps
                            tmp(10 + 2 * i) = "発動率: " & .kouka_p & "/ " & "上昇率: " & .kouka_f & vbCrLf & _
                                "→ ◆期待値 " & Math.Ceiling(.exp_kouka_b * 10000) / 10000
                            If .up_kouka_p > 0 Then
                                tmp(10 + 2 * i) = tmp(10 + 2 * i) & vbCrLf & " [発動率上昇中 +" & Format(.up_kouka_p * 100, "#0.00") & "%]"
                            End If
                            If .t_flg Then
                                tmp(10 + 2 * i) = tmp(10 + 2 * i) & vbCrLf & "★☆特殊条件スキル☆★"
                            End If
                        Else
                            tmp(9 + 2 * i) = "【スキル" & i + 1 & "】" & vbCrLf & .name & "LV" & .lv
                            tmp(10 + 2 * i) = "***********"
                        End If
                    End With
                Next
            End With
        End If
        For i As Integer = 0 To tmp.Length - 1
            データテキスト出力 = データテキスト出力 & tmp(i) & vbCrLf
        Next
        If Not bstr Is Nothing Then '太文字にする部分の抽出
            For i As Integer = 0 To bstr.Length - 1
                ReDim Preserve boldtext(bn)(i)
                Dim btemp As String = tmp(CInt(bstr(i)))
                boldtext(bn)(i) = Mid(btemp, InStr(btemp, ": ") + 2, btemp.Length - InStr(btemp, ": ") - 1)
            Next
        End If
    End Function

    Private Sub 一括入力フォーム起動(ByVal sender As System.Object, ByVal e As MouseEventArgs) Handles ToolStripButton3.MouseUp
        If (e.Button = MouseButtons.Left) Then
            Form3.Show()
        Else
            Form5.Show()
        End If
    End Sub

    Private Sub 部隊スキル入力(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Form4.Show()
    End Sub

    Public Sub INIファイルから読み込み(ByVal bini As String)
        Cursor.Current = Cursors.WaitCursor
        'まず、武将数のみ取得してデータクリア
        Dim b_c As Integer = Val(GetINIValue("busho_counter", "部隊長", bini))
        ToolStripComboBox2.Text = b_c 'Form1の武将数を更新
        Call 武将数スイッチ(ToolStripComboBox2, Nothing)
        Call データクリア(vbNull, Nothing)

        For i As Integer = 0 To b_c - 1
            Dim bsho As String = Nothing
            Call 武将データ手入力用(i, False, True)
            Select Case i
                Case 0
                    bsho = "部隊長"
                Case 1
                    bsho = "小隊長A"
                Case 2
                    bsho = "小隊長B"
                Case 3
                    bsho = "小隊長C"
            End Select
            ReDim Preserve bs(i)
            With bs(i)
                '--- 武将データをINIファイルから読み込み ---
                .rare = GetINIValue("rare", bsho, bini)
                .name = GetINIValue("name", bsho, bini)

                ComboBox(Me, CStr(i) & "01").SelectedIndex = ComboBox(Me, CStr(i) & "01").FindString(.rare) '（強制的に）R選択
                'R選択(ComboBox(Me, CStr(i) & "01"), Nothing)
                ComboBox(Me, CStr(i) & "02").SelectedText = .name '（強制的に）武将名選択
                武将名選択(ComboBox(Me, CStr(i) & "02"), Nothing)

                .heisyu.name = GetINIValue("heisyu.name", bsho, bini)
                .hei_sum = Val(GetINIValue("hei_sum", bsho, bini))
                .rank = Val(GetINIValue("rank", bsho, bini))
                If .rank = 6 Then '限界突破時
                    CheckBox(Me, CStr(i) & "1").Checked = True
                End If
                If Not .job = "剣" Then
                    If .job = "覇" Then
                        .hei_max = .hei_max_d + .rank * 200 'ランクアップで兵数一律+200
                    Else
                        .hei_max = .hei_max_d + .rank * 100 'ランクアップで兵数一律+100
                    End If
                End If
                .level = Val(GetINIValue("level", bsho, bini))
                .skill_no = Val(GetINIValue("skill_no", bsho, bini))

                For j As Integer = 0 To 2
                    .st(j) = Val(GetINIValue("st(" & j & ")", bsho, bini))
                Next
                For j As Integer = 0 To 3
                    .tou_a(j) = GetINIValue("tou_a(" & j & ")", bsho, bini)
                Next

                'この2つは面倒なので入力→更新する
                Dim tmpskill_no As Integer = .skill_no '変わっていく値なので一旦移し替え
                '----------
                For j As Integer = 0 To tmpskill_no - 1
                    ReDim Preserve .skill(tmpskill_no - 1) '途中、追加スキル追加の部分で空白部分を削られてしまうので逐一Redimする必要がある
                    .skill(j).name = GetINIValue("skill(" & j & ").name", bsho, bini)
                    .skill(j).lv = Val(GetINIValue("skill(" & j & ").lv", bsho, bini))
                    If Not j = 0 Then '初期スキルはスルー
                        .skill(j).kanren = スキル関連推定(.skill(j).name, True)
                    End If

                    If Not .skill(j).kanren = "" And Not j = 0 Then '関連スキルが空ではない＝追加スキルがある
                        Select Case j
                            Case 1 'スロ2
                                ComboBox(Me, CStr(i) & "09").Focus()
                                ComboBox(Me, CStr(i) & "09").SelectedIndex = ComboBox(Me, CStr(i) & "09").FindString(スキル関連推定(.skill(j).name, True))
                                ComboBox(Me, CStr(i) & "11").SelectedText = .skill(j).name
                                スキル名入力(ComboBox(Me, CStr(i) & "11"), Nothing)
                                ComboBox(Me, CStr(i) & "15").Text = .skill(j).lv
                                追加スキル追加(ComboBox(Me, CStr(i) & "15"), Nothing)
                            Case 2 'スロ3
                                ComboBox(Me, CStr(i) & "10").Focus()
                                ComboBox(Me, CStr(i) & "10").SelectedIndex = ComboBox(Me, CStr(i) & "10").FindString(スキル関連推定(.skill(j).name, True))
                                ComboBox(Me, CStr(i) & "12").SelectedText = .skill(j).name
                                スキル名入力(ComboBox(Me, CStr(i) & "12"), Nothing)
                                ComboBox(Me, CStr(i) & "16").Text = .skill(j).lv
                                追加スキル追加(ComboBox(Me, CStr(i) & "16"), Nothing)
                        End Select
                    ElseIf j = 0 Then '初期スキルの場合は
                        ComboBox(Me, CStr(i) & "14").Text = .skill(0).lv
                        追加スキル追加(ComboBox(Me, CStr(i) & "14"), Nothing)
                    End If
                Next
                '----------
                '他は“表面上フォームを埋めていくだけ”
                Button(Me, CStr(i) & "001").Text = "/ " & .hei_max
                ComboBox(Me, CStr(i) & "03").Text = .heisyu.name
                ComboBox(Me, CStr(i) & "04").Text = .rank
                For j As Integer = 5 To 8
                    ComboBox(Me, CStr(i) & "0" & CStr(j)).Text = Nothing
                    ComboBox(Me, CStr(i) & "0" & CStr(j)).Text = .tou_a(j - 5)
                Next
                .Tousotu(True) = .tou_a '統率値を変換
                .兵科情報取得(bs(i).heisyu.name)
                TextBox(Me, CStr(i) & "01").Text = .hei_sum
                TextBox(Me, CStr(i) & "02").Text = .level
                For j As Integer = 3 To 5 'ステ選択
                    TextBox(Me, CStr(i) & "0" & CStr(j)).Text = .st(j - 3)
                Next
                .rankup_r = 統率未振り数推定(.tou_d_a, .tou_a, .rank)
                .huri = ステ振り推定(.st_d, .st, .sta_g, .rank, .level)
                ComboBox(Me, CStr(i) & "13").Text = .huri
                .残りランクアップ可能回数表示(GroupBox(Me, 4 * (i + 1)))
                '---------
                武将データ手入力用(i, True, True)
                bs(i).No = i '武将No付け替え必要
            End With
        Next
        '部隊スキル読み込み
        Dim bskillno As Integer
        bskillno = Val(GetINIValue("bskill_no", "部隊スキル", bini))
        If bskillno > 0 Then
            For i As Integer = 0 To bskillno - 1
                ' bskillを読み込んでいく
                If i = 0 Then
                    ReDim bskill.bsk(0)
                Else
                    ReDim Preserve bskill.bsk(i)
                End If
                With bskill.bsk(i)
                    .koubou = GetINIValue("koubou(" & i & ")", "部隊スキル", bini)
                    .type = GetINIValue("type(" & i & ")", "部隊スキル", bini)
                    .kouka_p = Val(GetINIValue("kouka_p(" & i & ")", "部隊スキル", bini))
                    .kouka_f = Val(GetINIValue("kouka_f(" & i & ")", "部隊スキル", bini))
                    .speed = Val(GetINIValue("speed(" & i & ")", "部隊スキル", bini))
                    .taisyo = GetINIValue("taisyo(" & i & ")", "部隊スキル", bini)
                End With
            Next
            Form4.部隊スキルONOFF(True)
        End If
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub 名前を付けて現在の部隊を保存(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 名前を付けて現在の部隊を保存ToolStripMenuItem.Click
        Dim fname As String
        While 1
            fname = InputBox("部隊名", "現在表示されている部隊を保存しますか？")
            FILENAME_bs = My.Application.Info.DirectoryPath & "\butai\" & fname & ".INI" 'INIファイルの保存場所

            If fname = "" Then '無名の場合は抜ける
                Exit Sub
            ElseIf File.Exists(FILENAME_bs) Then
                Dim yn As Integer = MsgBox("既に同名の部隊が保存されています。上書きしますか？", vbYesNo)
                If yn = vbYes Then
                    File.Delete(FILENAME_bs)
                    Exit While
                End If
            Else
                Exit While
            End If
        End While

        Dim c As Integer '武将数を見直し
        For i = 0 To busho_counter - 1
            If ComboBox(Me, CStr(i) & "02").Text = "" Then
                c = c + 1
            End If
        Next
        busho_counter = busho_counter - c

        Try
            For i As Integer = 0 To busho_counter - 1
                Dim bsho As String = Nothing
                Select Case i
                    Case 0
                        bsho = "部隊長"
                    Case 1
                        bsho = "小隊長A"
                    Case 2
                        bsho = "小隊長B"
                    Case 3
                        bsho = "小隊長C"
                End Select
                With bs(i)
                    If i = 0 Then '部隊長記憶時
                        SetINIValue(busho_counter, "busho_counter", bsho, FILENAME_bs) '武将数も記憶
                    End If
                    '無いと復元できないもの（主キー）のみをINIへ保存
                    SetINIValue(.rare, "rare", bsho, FILENAME_bs)
                    SetINIValue(.name, "name", bsho, FILENAME_bs)
                    SetINIValue(.heisyu.name, "heisyu.name", bsho, FILENAME_bs)
                    SetINIValue(.hei_sum, "hei_sum", bsho, FILENAME_bs)
                    If .limitbreakflg Then '限界突破時
                        SetINIValue(6, "rank", bsho, FILENAME_bs)
                    Else
                        SetINIValue(.rank, "rank", bsho, FILENAME_bs)
                    End If
                    SetINIValue(.level, "level", bsho, FILENAME_bs)
                    'SetINIValue(.rankup_r, "rankup_r", bsho, FILENAME_bs)
                    'SetINIValue(.huri, "huri", bsho, FILENAME_bs)
                    For j As Integer = 0 To 2
                        SetINIValue(.st(j), "st(" & j & ")", bsho, FILENAME_bs)
                    Next
                    For j As Integer = 0 To 3
                        SetINIValue(.tou_a(j), "tou_a(" & j & ")", bsho, FILENAME_bs)
                    Next
                    SetINIValue(.skill_no, "skill_no", bsho, FILENAME_bs)
                    For j As Integer = 0 To .skill_no - 1
                        SetINIValue(.skill(j).name, "skill(" & j & ").name", bsho, FILENAME_bs)
                        SetINIValue(.skill(j).lv, "skill(" & j & ").lv", bsho, FILENAME_bs)
                        'SetINIValue(.skill(j).kanren, "skill(" & j & ").kanren", bsho, FILENAME_bs)
                    Next
                End With
            Next
            '部隊スキル
            If bskill.flg Then
                With bskill
                    SetINIValue(.bsk.Length, "bskill_no", "部隊スキル", FILENAME_bs)
                    For i As Integer = 0 To .bsk.Length - 1
                        SetINIValue(.bsk(i).koubou, "koubou(" & i & ")", "部隊スキル", FILENAME_bs)
                        SetINIValue(.bsk(i).type, "type(" & i & ")", "部隊スキル", FILENAME_bs)
                        SetINIValue(.bsk(i).kouka_p, "kouka_p(" & i & ")", "部隊スキル", FILENAME_bs)
                        SetINIValue(.bsk(i).kouka_f, "kouka_f(" & i & ")", "部隊スキル", FILENAME_bs)
                        SetINIValue(.bsk(i).speed, "speed(" & i & ")", "部隊スキル", FILENAME_bs)
                        SetINIValue(.bsk(i).taisyo, "taisyo(" & i & ")", "部隊スキル", FILENAME_bs)
                    Next
                End With
            Else
                SetINIValue(0, "bskill_no", "部隊スキル", FILENAME_bs)
            End If
            MsgBox("登録完了")
        Catch ex As Exception
            MsgBox("登録内容に漏れがあります。次回復元時に正しく復元されない可能性があります。")
        End Try
    End Sub

    Private Sub 保存部隊を開く(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 保存部隊を開くToolStripMenuItem.Click
        Dim bini As String '適用するINIファイル
        'OpenFileDialogクラスのインスタンスを作成
        Dim ofd As New OpenFileDialog()

        'はじめに表示されるフォルダを指定する
        '指定しない（空の文字列）の時は、現在のディレクトリが表示される
        ofd.InitialDirectory = My.Application.Info.DirectoryPath & "\butai"
        'タイトルを設定する
        ofd.Title = "部隊を選択してください"
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

    Private Sub 速度計測(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        Dim speed_skl() As Busho.skl = Nothing '速度スキル格納
        Dim ksk() As Decimal = Nothing '加速率
        Dim butai_spd As Integer '行軍速度（秒）
        Dim kasoku() As Decimal = Nothing, it() As Decimal = Nothing '武将の移動速度、No.
        Dim bsc As Integer '現時点での武将数
        Dim heikac As Integer = 0 '兵科セット数
        Dim c As Integer = 0

        For i As Integer = 0 To busho_counter - 1
            If ComboBox(Me, CStr(i) & "02").Text = "" Then
                c = c + 1
            End If
            If Not ComboBox(Me, CStr(i) & "03").Text = "" Then
                heikac = heikac + 1
            End If
        Next
        bsc = busho_counter - c
        If bsc = 0 Then '武将数ゼロ
            MsgBox("武将数がゼロです")
            Exit Sub
        End If
        If Not heikac = bsc Then '兵科セット数と武将数が一致しない、つまり兵科が決まっていない武将がいる
            MsgBox("兵科設定に異常があります")
            Exit Sub
        End If

        '勝軍地蔵が居るので読み込み
        Call フラグ付きスキル読み込み() '読込（更新）
        Call 童ボーナス加算()

        ReDim ksk(bsc - 1), kasoku(bsc - 1), it(bsc - 1)
        c = 0

        For i As Integer = 0 To bsc - 1 'スピードスキルのみ抽出
            With bs(i)
                .兵科情報取得(bs(i).heisyu.name)
                For j As Integer = 0 To .skill.Length - 1
                    If Not .skill(j).speed = 0 Or .skill(j).tokusyu = 1 Then '加速スキルがあれば
                        ReDim Preserve speed_skl(c)
                        speed_skl(c) = .skill(j).Clone
                        c = c + 1
                    End If
                Next
            End With
        Next
        For i As Integer = 0 To bsc - 1 '各武将の速度を計算
            '童効果適用
            For j As Integer = 0 To bsc - 1
                Select Case (bs(j).heisyu.bunrui)
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
                    If InStr(speed_skl(j).heika, bs(i).heisyu.bunrui) Or InStr(speed_skl(j).heika, "全") Or InStr(speed_skl(j).heika, "将") Then
                        speed_skl(j).t_flg = フラグ付きスキル参照(speed_skl(j))
                        ksk(i) = ksk(i) + speed_skl(j).speed
                    End If
                Next
            End If
            kasoku(i) = Int(3600 / (bs(i).heisyu.spd * (1 + ksk(i))))
            it(i) = i
        Next

        Array.Sort(kasoku, it)
        If Not bskill.speed = 0 Then
            butai_spd = Int(3600 / (bs(Int(it(bsc - 1))).heisyu.spd * (1 + bskill.speed) * (1 + ksk(Int(it(bsc - 1))))))
            'butai_spd = kasoku(Int(bsc - 1)) * (1 + bskill.speed)
        Else
            butai_spd = kasoku(Int(bsc - 1))
        End If

        With ToolStripLabel6
            .Text = (butai_spd \ 60) & "分" & (butai_spd Mod 60).ToString("00") & "秒"
            '.Visible = True
            .ForeColor = Color.Brown
            .ToolTipText = "加速率: " & ksk(Int(it(bsc - 1))).ToString("p")
            If Not bskill.speed = 0 Then
                .ToolTipText = .ToolTipText & vbCrLf & _
                    "部隊スキル適用 +" & bskill.speed.ToString("p") & vbCrLf & _
                    "(部隊スキル抜き速度: " & (Int(3600 / (bs(Int(it(bsc - 1))).heisyu.spd * (1 + ksk(Int(it(bsc - 1)))))) \ 60) & "分" & _
                    (Int(3600 / (bs(Int(it(bsc - 1))).heisyu.spd * (1 + ksk(Int(it(bsc - 1)))))) Mod 60) & "秒)"
            End If
            '勝軍地蔵等が入る場合、その表示
            If Not warabe.speed.uma = 0 Then
                .ToolTipText = .ToolTipText & vbCrLf & "☆童アリ"
            End If
        End With
    End Sub

    'クリックで速度更新
    Private Sub 速度更新(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel6.Click
        Call 速度計測(sender, Nothing)
    End Sub

    Private Sub コンテキストメニュー(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening
        ContextMenuStrip1.Items.Clear()
        Select Case contextmenuflg
            Case 1
                ContextMenuStrip1.Items.Add("このスキルのLVを全てのスキルの適用")
            Case 2
                ContextMenuStrip1.Items.Add("積載兵科一括変更")
            Case 3
                ContextMenuStrip1.Items.Add("兵満載セット")
                ContextMenuStrip1.Items.Add("兵1ALLセット")
                ContextMenuStrip1.Items.Add("部隊長の指揮兵数に合わせる")
            Case Else
                Exit Sub
        End Select
    End Sub

    Private Sub 兵科一括変更(ByVal sender As Object, ByVal e As MouseEventArgs) Handles Label2.MouseUp
        If (e.Button = MouseButtons.Right) Then
            Dim mousePosition As Point = Control.MousePosition
            contextmenuflg = 2 'Flgは2
            ContextMenuStrip1.Tag = ComboBox003.Text
            ContextMenuStrip1.Show(mousePosition)
        End If
    End Sub

    Private Sub 兵数一括変更(ByVal sender As Object, ByVal e As MouseEventArgs) Handles Label5.MouseUp
        If (e.Button = MouseButtons.Right) Then
            Dim mousePosition As Point = Control.MousePosition
            contextmenuflg = 3 'Flgは3
            ContextMenuStrip1.Tag = Val(TextBox001.Text)
            ContextMenuStrip1.Show(mousePosition)
        End If
    End Sub

    Private Sub コンテキストメニュー操作(ByVal sender As System.Object, ByVal e As ToolStripItemClickedEventArgs) Handles ContextMenuStrip1.ItemClicked
        Cursor.Current = Cursors.WaitCursor
        Dim cit As String = e.ClickedItem.Text
        '適用武将数をカウント
        Dim c, bsc As Integer
        For i As Integer = 0 To busho_counter - 1
            If ComboBox(Me, CStr(i) & "02").Text = "" Then
                c = c + 1
            End If
        Next
        bsc = busho_counter - c
        If bsc = 0 Then '武将数ゼロ
            MsgBox("武将数がゼロです")
            Exit Sub
        End If

        Select Case contextmenuflg
            Case 1 '"このスキルのLVを全てのスキルの適用"がクリックされた
                If Int(ContextMenuStrip1.Tag) = 0 Then
                    MsgBox("スキルLVの値が空です")
                    Exit Sub
                End If
                For i As Integer = 0 To bsc - 1
                    For j As Integer = 14 To 14 + bs(i).skill.Length - 1
                        If Not ComboBox(Me, CStr(i) & CStr(j)).Text = "" Then '未設定
                            ComboBox(Me, CStr(i) & CStr(j)).Text = Int(ContextMenuStrip1.Tag)
                        End If
                    Next
                Next
            Case 2
                If cit = "積載兵科一括変更" Then
                    If ContextMenuStrip1.Tag = "" Then
                        MsgBox("部隊長の兵科が未選択です")
                        Exit Sub
                    End If
                    For i As Integer = 0 To bsc - 1
                        ComboBox(Me, CStr(i) & "03").Text = ContextMenuStrip1.Tag
                        Me.積載兵科変更(ComboBox(Me, CStr(i) & "03"), Nothing)
                    Next
                End If
            Case 3
                For i As Integer = 0 To bsc - 1
                    If cit = "兵満載セット" Then
                        TextBox(Me, CStr(i) & "01").Text = bs(i).hei_max
                    ElseIf cit = "兵1ALLセット" Then
                        TextBox(Me, CStr(i) & "01").Text = 1
                    Else
                        If Val(ContextMenuStrip1.Tag) = 0 Then '部隊長の兵数0
                            MsgBox("部隊長の兵数がゼロもしくは未設定です")
                            Exit Sub
                        End If
                        TextBox(Me, CStr(i) & "01").Text = Val(ContextMenuStrip1.Tag)
                        Me.積載兵数変更(TextBox(Me, CStr(i) & "01"), Nothing)
                    End If
                Next
        End Select
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub 追加合成シミュ起動(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 追加合成シミュToolStripMenuItem.Click
        Form6.Show()
    End Sub

    Private Sub スキル期待値シミュ起動(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles スキル期待値シミュToolStripMenuItem.Click
        Form7.Show()
    End Sub

    Private Sub 武将シミュ起動(sender As Object, e As EventArgs) Handles 武将シミュToolStripMenuItem.Click
        Form8.Show()
    End Sub

    '********** 以下、一括入力に関する関数 **********
    Private Sub 一括入力_追加スキル表示(sender As Object, e As EventArgs) Handles ToolStripComboBox7.GotFocus
        Dim sqlwhere As String = ダブルクオート(ToolStripComboBox6.Text)
        If sqlwhere = ダブルクオート("特殊") Then '特殊項目には、条件付きスキルも含む
            sqlwhere = sqlwhere & " OR 分類 = " & ダブルクオート("条件")
        End If
        Dim p As DataSet = _
            DB_TableOUT("SELECT id, 分類, スキル名 FROM SName WHERE 分類 = " & sqlwhere & " ORDER BY id", "SName")
        With ToolStripComboBox7.ComboBox
            .BindingContext = Me.BindingContext
            .DisplayMember = "スキル名"
            .ValueMember = "id"
            .DataSource = p.Tables("SName")
            .SelectedIndex = -1
        End With
    End Sub

    '一括入力追加スキル設定
    Private Sub スロ2設定(sender As Object, e As EventArgs) Handles スロ2設定ToolStripMenuItem.Click
        ToolStripLabel9.Text = "【" & ToolStripComboBox7.Text & Val(ToolStripComboBox8.Text) & "】"
        ToolStripLabel9.Tag = ToolStripComboBox6.Text
    End Sub
    Private Sub スロ3設定(sender As Object, e As EventArgs) Handles スロ3設定ToolStripMenuItem.Click
        ToolStripLabel10.Text = "【" & ToolStripComboBox7.Text & Val(ToolStripComboBox8.Text) & "】"
        ToolStripLabel10.Tag = ToolStripComboBox6.Text
    End Sub
    Private Delegate Sub スロ設定(sender As Object, e As EventArgs)

    '一括入力追加スキル削除
    Private Sub スロ2解除(ByVal sender As Object, ByVal e As MouseEventArgs) Handles ToolStripLabel9.MouseUp
        If (e.Button = MouseButtons.Right) Then
            ToolStripLabel9.Text = "【スロ2 : 空】"
            ToolStripLabel9.Tag = ""
        End If
    End Sub
    Private Sub スロ3解除(ByVal sender As Object, ByVal e As MouseEventArgs) Handles ToolStripLabel10.MouseUp
        If (e.Button = MouseButtons.Right) Then
            ToolStripLabel10.Text = "【スロ3 : 空】"
            ToolStripLabel10.Tag = ""
        End If
    End Sub

    Private Sub 一括入力(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        '武将数
        Dim tbno As Integer = Count_Busho()
        'Me.ToolStripComboBox2.Text = tbno + 1 'Form1の武将数を更新
        'Call Me.武将数スイッチ(Me.ToolStripComboBox2, Nothing)
        'Call Me.データクリア(sender, Nothing)

        For i As Integer = 0 To tbno
            '兵科変更
            Dim heika As String = CType(heikaht(ToolStripComboBox3.Text), String)
            ComboBox(Me, CStr(i) & "03").Text = heika '兵科入力
            Me.積載兵科変更(ComboBox(Me, CStr(i) & "03"), Nothing)

            '★/LV入力
            Dim rank As Integer
            If ToolStripComboBox4.Text = "限突" Then
                CheckBox(Me, CStr(i) & "1").Checked = True '限界突破
                rank = 6
            Else
                If bs(i).limitbreakflg Then '限界突破ONならば
                    CheckBox(Me, CStr(i) & "1").Checked = False
                End If
                rank = Val(Mid(ToolStripComboBox4.Text, 2, 1))
                ComboBox(Me, CStr(i) & "04").Text = rank 'ランク入力
                TextBox(Me, CStr(i) & "02").Text = "20"
                Me.ランク変更(ComboBox(Me, CStr(i) & "04"), Nothing)
                Me.LV変更(TextBox(Me, CStr(i) & "02"), Nothing)
            End If

            '統率変更
            If autotou = True Then '自動統率変更ONならば
                bs(i).tou_a = 自動統率(heika, bs(i).tou_d, bs(i).rank).Clone
                '自動統率設定は一括で設定してしまうため、removehandler
                For j As Integer = 5 To 8
                    RemoveHandler ComboBox(Me, CStr(i) & "0" & CStr(j)).SelectedIndexChanged, AddressOf 統率変更
                    ComboBox(Me, CStr(i) & "0" & CStr(j)).Text = bs(i).tou_a(j - 5)
                    bs(i).tou(j - 5) = 統率_数値変換(bs(i).tou_a(j - 5))
                    AddHandler ComboBox(Me, CStr(i) & "0" & CStr(j)).SelectedIndexChanged, AddressOf 統率変更
                Next
                With bs(i)
                    .rankup_r = 0
                    .残りランクアップ可能回数表示(GroupBox(Me, 4 * (bc + 1)))
                    .兵科情報取得(.heisyu.name) 'ランクが変われば当然兵科に対する統率値も変わる
                End With
            End If

            'ステ振り変更
            Dim sutetmp As String = ""
            Select Case ToolStripComboBox5.Text
                Case "攻"
                    sutetmp = "攻撃極振り"
                Case "防"
                    sutetmp = "防御極振り"
                Case "兵"
                    sutetmp = "兵法極振り"
            End Select
            ComboBox(Me, CStr(i) & "13").Text = sutetmp 'ステ振り入力
            Me.ステ振り設定(ComboBox(Me, CStr(i) & "13"), Nothing)

            '兵数変更
            Dim heisutmp As Integer = bs(i).hei_max_d
            Dim setheisu As Integer 'SETする兵数(デフォはMAX)
            If Not bs(i).job = "剣" Then
                If bs(i).job = "覇" Then
                    setheisu = heisutmp + 200 * rank
                Else
                    setheisu = heisutmp + 100 * rank
                End If
            Else
                setheisu = heisutmp
            End If
            If 積載兵数ToolStripMenuItem.Tag = 0 Then
                setheisu = sethei
            Else
                If sethei = -2 Then '2%キャップ
                    Dim ds, dk, tset As Integer
                    If InStr(kb, "攻") Then
                        ds = bs(i).st(0)
                        dk = bs(i).heisyu.atk
                    Else
                        ds = bs(i).st(1)
                        dk = bs(i).heisyu.def
                    End If
                    tset = Math.Ceiling((0.02 * ds * bs(i).heisyu.ts) / (1 - 0.02 * dk * bs(i).heisyu.ts))
                    If tset <= setheisu Then 'キャップ兵数が指揮兵数以下
                        setheisu = tset
                    End If
                End If
            End If
            TextBox(Me, CStr(i) & "01").Text = setheisu '積載兵数入力
            Me.積載兵数変更(TextBox(Me, CStr(i) & "01"), Nothing)

            '初期スキル
            ComboBox(Me, CStr(i) & "14").Text = syokiLV

            '追加スキル
            Dim stu As Integer = 0 '追加スキル数
            If ToolStripLabel9.Tag = "" Then
                If Not ToolStripLabel10.Tag = "" Then 'スロ2が無くてスロ3がある場合
                    Dim stmp, ttmp As String
                    stmp = ToolStripLabel10.Text
                    ttmp = ToolStripLabel10.Tag
                    ToolStripLabel10.Text = "【スロ3 : 空】"
                    ToolStripLabel10.Tag = ""
                    ToolStripLabel9.Text = stmp
                    ToolStripLabel9.Tag = ttmp
                End If
            End If
            If Not ToolStripLabel9.Tag = "" Then
                stu = 1
                If Not ToolStripLabel10.Tag = "" Then
                    stu = 2
                End If
            End If

            If stu >= 1 Then
                ComboBox(Me, CStr(i) & "09").Focus()
                ComboBox(Me, CStr(i) & "09").Text = ToolStripLabel9.Tag
                ComboBox(Me, CStr(i) & "11").Text = _
                    StringRep(ToolStripLabel9.Text, {"【", "】", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"})
                ComboBox(Me, CStr(i) & "15").Enabled = True
                ComboBox(Me, CStr(i) & "15").Text = Val(String_onlyNumber(ToolStripLabel9.Text))
                Me.追加スキル追加(ComboBox(Me, CStr(i) & "15"), Nothing) 'スキルスロ2入力
            End If
            If stu >= 2 Then
                ComboBox(Me, CStr(i) & "10").Focus()
                ComboBox(Me, CStr(i) & "10").Text = ToolStripLabel10.Tag
                ComboBox(Me, CStr(i) & "12").Text = _
                    StringRep(ToolStripLabel10.Text, {"【", "】", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"})
                ComboBox(Me, CStr(i) & "16").Enabled = True
                ComboBox(Me, CStr(i) & "16").Text = Val(String_onlyNumber(ToolStripLabel10.Text))
                Me.追加スキル追加(ComboBox(Me, CStr(i) & "16"), Nothing) 'スキルスロ3入力
            End If
        Next
    End Sub

    '自動設定
    Private Sub スキル入力設定(sender As Object, e As EventArgs) Handles ToolStripSplitButton2.ButtonClick
        Dim suro As スロ設定 = New スロ設定(AddressOf スロ2設定)
        If Not ToolStripLabel9.Tag = "" Then 'スロ2が埋まっている
            suro = New スロ設定(AddressOf スロ3設定)
            If Not ToolStripLabel10.Tag = "" Then 'スロ3が埋まっている
                suro = New スロ設定(AddressOf スロ2設定)
            End If
        End If
        Call suro(sender, Nothing)
    End Sub

    Public syokiLV As Integer = 10
    Public autotou As Boolean = True
    Public sethei As Integer = -1

    '設定
    Private Sub 初期スキルLV設定(sender As Object, e As EventArgs) _
        Handles ToolStripMenuItem2.Click, ToolStripMenuItem3.Click, ToolStripMenuItem5.Click, ToolStripMenuItem6.Click
        syokiLV = Val(sender.text)
        初期スキルLVToolStripMenuItem.Text = "初期スキルLV(" & syokiLV & ")"
    End Sub
    Private Sub 自動統率設定(sender As Object, e As EventArgs) _
        Handles ONToolStripMenuItem.Click, OFFToolStripMenuItem.Click
        Dim o As String = Mid(CStr(sender.Name), 1, 2)
        If o = "ON" Then 'ONならば
            autotou = True
            自動統率設定ToolStripMenuItem.Text = "自動統率設定(ON)"
            自動統率設定ToolStripMenuItem.Tag = 1
            自動統率設定ToolStripMenuItem.ForeColor = Color.Black
        Else 'OFFならば
            autotou = False
            自動統率設定ToolStripMenuItem.Text = "自動統率設定(OFF)"
            自動統率設定ToolStripMenuItem.Tag = 0
            自動統率設定ToolStripMenuItem.ForeColor = Color.Gray
        End If
    End Sub
    Private Sub 積載兵数設定(sender As Object, e As EventArgs) _
        Handles hei1.Click, hei10.Click, cap2.Click, heiMAX.Click
        Select Case (sender.Name)
            Case "hei1"
                sethei = 1
                積載兵数ToolStripMenuItem.Text = "積載兵数(1)"
                積載兵数ToolStripMenuItem.Tag = 0
            Case "hei10"
                sethei = 10
                積載兵数ToolStripMenuItem.Text = "積載兵数(10)"
                積載兵数ToolStripMenuItem.Tag = 0
            Case "cap2"
                sethei = -2
                積載兵数ToolStripMenuItem.Text = "積載兵数(2%cap)"
                積載兵数ToolStripMenuItem.Tag = 1
            Case "heiMAX"
                sethei = -1
                積載兵数ToolStripMenuItem.Text = "積載兵数(MAX)"
                積載兵数ToolStripMenuItem.Tag = 1
        End Select
    End Sub

    'heika:兵科, st:ランクアップ前統率, rc:ランクアップ残回数
    Private Function 自動統率(ByVal heika As String, ByVal st() As Decimal, ByVal rc As Integer) As String()
        Dim stt() As Decimal = st.Clone
        Dim s() As String = _
         DB_DirectOUT("SELECT 統率, 兵種名 FROM HData WHERE 兵種名= " & ダブルクオート(heika) & "", {"統率"})
        Dim tt As New Hashtable
        Dim touk As String() = {"槍", "弓", "馬", "器"}
        Dim output As String() = {数値_統率変換(st(0)), 数値_統率変換(st(1)), 数値_統率変換(st(2)), 数値_統率変換(st(3))}
        Dim toukt As String() = touk.Clone
        Dim rup_zan As Integer() = {(1.2 - st(0)) / 0.05, (1.2 - st(1)) / 0.05, (1.2 - st(2)) / 0.05, (1.2 - st(3)) / 0.05}
        For i As Integer = 0 To 3
            tt(touk(i)) = i
        Next
        Dim un() As Integer = {0, 0, 0, 0} 'ランクアップ優先順位
        '現段階でst降順にtoukをソート
        Array.Sort(stt, toukt)
        Array.Reverse(toukt)

        Dim sn() As Integer = Nothing '対象兵科の位置
        Dim k As Integer = 0
        For i As Integer = 1 To s(0).Length
            Dim td As Integer = tt(Mid(s(0), i, 1))
            ReDim Preserve sn(k)
            sn(k) = td
            k = k + 1
        Next
        If s(0).Length = 1 Then '複数統率が関わらない兵科
            un(sn(0)) = 1
            Dim crank As Integer = 2
            For i As Integer = 0 To 3
                If Not s(0) = toukt(i) Then
                    un(tt(toukt(i))) = crank
                    crank = crank + 1
                End If
            Next
        Else
            If (st(sn(0)) < st(sn(1))) Then
                un(sn(1)) = 1
                un(sn(0)) = 2
            Else
                un(sn(0)) = 1
                un(sn(1)) = 2
            End If
            Dim crank As Integer = 3
            For i As Integer = 0 To 3
                If InStr(s(0), toukt(i)) = 0 Then
                    un(tt(toukt(i))) = crank
                    crank = crank + 1
                End If
            Next
        End If
        Array.Sort(un, touk)
        '実際のランクアップ
        Dim r As Integer = rc
        For i As Integer = 0 To 3
            If rup_zan(tt(touk(i))) < r Then
                output(tt(touk(i))) = "SSS" 'とりあえずSSS
                r = r - rup_zan(tt(touk(i)))
            Else
                output(tt(touk(i))) = 数値_統率変換(st(tt(touk(i))) + r * 0.05)
                Exit For
            End If
        Next
        Return output
    End Function

    Private Sub 武将ランキング起動(sender As Object, e As EventArgs) Handles 武将ランキングToolStripMenuItem.Click
        Form10.Show()
    End Sub

    'Private Sub 条件設定起動(sender As Object, e As EventArgs) Handles 条件設定ToolStripMenuItem.Click

    'End Sub

    Private Sub 条件付スキルオプション表示(sender As Object, e As EventArgs) Handles ToolStripButton7.Click
        Form12.Show()
    End Sub

    Private Sub 開くボタン(sender As Object, e As EventArgs) Handles ToolStripSplitButton1.ButtonClick
        Call 保存部隊を開く(sender, e)
    End Sub

    Private Sub DB編集ツールを開く(sender As Object, e As EventArgs) Handles ToolStripButton8.Click
        Form13.Show()
    End Sub
End Class