Public Class Form5
    Public Structure bData : Implements System.ICloneable

        Public Function Clone() As Object Implements System.ICloneable.Clone
            Return Me.MemberwiseClone()
        End Function
        Public rare As String
        Public name As String
        Public heisyu_name As String
        Public hei_sum As Integer
        Public rank As Integer
        Public level As Integer
        Public st() As Decimal
        Public tou_a() As String
        Public skill_no As Integer
        Public skill_name() As String
        Public skill_lv() As Integer
    End Structure
    Private bd() As bData
    Private ss() As String 'Richtextboxに表示する文字列
    'Public busho_counter As Integer = 4

    Private Sub Form5_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Opacity = 0.8 '初期の透過度は80%
        TrackBar1.Value = 8 '初期位置8
        Form1.Hide()
        Me.TopMost = True
        ReDim bd(3)
        ReDim ss(3)
        '武将数は4に初期化
        busho_counter = 4
    End Sub

    Private Sub Form5_Closing(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.FormClosing
        Form1.Show()
    End Sub

    Private Function 部隊枠番号(ByVal controlname As String) As Integer
        部隊枠番号 = Val(Mid(controlname, controlname.Length, 1)) - 1
    End Function

    Private Sub データDD(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) _
            Handles RichTextBox1.DragDrop, RichTextBox2.DragDrop, RichTextBox3.DragDrop, RichTextBox4.DragDrop
        Dim cc As Integer = 部隊枠番号(CStr(sender.name))
        Dim stmp As String = e.Data.GetData(GetType(String))
        Dim growstate As String = Nothing
        'ハンゲーム版であるかどうか
        Dim hangameflg As Boolean = False
        'ハンゲ版は、ステータス欄が「攻撃」「防御」「兵法」ではなく「攻」「防」「兵」となっている。
        If 正規表現マッチ("\b攻撃\s*[0-9]+.[0-9]+", stmp) Is Nothing Then
            hangameflg = True
        End If

        With bd(cc)
            Try
                If InStr(stmp, "限界突破") Then '限界突破時
                    .rank = 6
                    .level = 20
                    growstate = "☆限界突破☆"
                Else
                    .rank = Val(正規表現マッチ("[0-9]", 正規表現マッチ("★[0-9]", stmp)(0))(0))
                    .level = Val(正規表現マッチ("[0-9]+", 正規表現マッチ("｜\b\S+\b", stmp)(0))(0))
                    growstate = "レベル ★" & .rank & "｜" & .level
                End If
                ReDim .st(2), .tou_a(3)
                .st(0) = Val(正規表現マッチ("[0-9]+", 正規表現マッチ("\b(攻撃|攻)\s*[0-9]+.[0-9]+", stmp)(0))(0))
                .st(1) = Val(正規表現マッチ("[0-9]+", 正規表現マッチ("\b(防御|防)\s*[0-9]+.[0-9]+", stmp)(0))(0))
                .st(2) = Val(正規表現マッチ("[0-9]+", 正規表現マッチ("\b(兵法|兵)\s*[0-9]+.[0-9]+", stmp)(0))(0))
                .tou_a(0) = 正規表現マッチ("[A-Z]+", 正規表現マッチ("\b槍\s*[A-Z]+", stmp)(0))(0)
                .tou_a(1) = 正規表現マッチ("[A-Z]+", 正規表現マッチ("\b弓\s*[A-Z]+", stmp)(0))(0)
                .tou_a(2) = 正規表現マッチ("[A-Z]+", 正規表現マッチ("\b馬\s*[A-Z]+", stmp)(0))(0)
                .tou_a(3) = 正規表現マッチ("[A-Z]+", 正規表現マッチ("\b器\s*[A-Z]+", stmp)(0))(0)
                For j As Integer = 0 To 3
                    If .tou_a(j) = "S" Then
                        .tou_a(j) = ".S"
                    End If
                Next
                Dim slv() As String = 正規表現マッチ("LV[0-9]+", stmp)
                .skill_no = slv.Length
                ReDim Preserve .skill_name(.skill_no - 1), .skill_lv(.skill_no - 1)
                .name = 正規表現マッチ("^\s*\S+", stmp)(0).Trim()
                If hangameflg Then
                    'ハンゲームの処理
                    If InStr(.name, "レア") Then
                        '武将名は、ハンゲームの場合そのまま取れないことがある。
                        .name = .name.Replace(正規表現マッチ("レア.*", .name)(0), "")
                    End If
                    For i As Integer = 0 To slv.Length - 1
                        Dim wazastr As String = "技" & i + 1
                        .skill_name(i) = 正規表現マッチ(wazastr & ".*" & slv(i), stmp)(0) _
                            .Replace(正規表現マッチ(wazastr & "\s+", stmp)(0), "").Replace(slv(i), "")
                        .skill_lv(i) = Val(正規表現マッチ("[0-9]+", slv(i))(0))
                        ss(cc) = ss(cc) & vbCrLf & (.skill_name(i) & slv(i))
                        stmp = stmp.Replace(.skill_name(i) & slv(i), "")
                    Next
                Else
                    'Yahooでの処理
                    For i As Integer = 0 To slv.Length - 1
                        Dim wazastr As String = "技" & i + 1
                        .skill_name(i) = 正規表現マッチ(wazastr & ".*" & slv(i), stmp)(0).Replace(wazastr, "").Replace(slv(i), "")
                        .skill_lv(i) = Val(正規表現マッチ("[0-9]+", slv(i))(0))
                        ss(cc) = ss(cc) & vbCrLf & (.skill_name(i) & slv(i))
                        stmp = stmp.Replace(.skill_name(i) & slv(i), "")
                    Next
                End If
                .heisyu_name = 正規表現マッチ("\S+", 正規表現マッチ("兵種\s+\S+", stmp)(0).Replace("兵種", ""))(0)
                '兵種が空の場合、ハンゲームならば「攻」、Yahooなら「指揮」が返っている。
                If .heisyu_name = "攻" Or .heisyu_name = "指揮" Then
                    .heisyu_name = ""
                End If
                .hei_sum = Val(正規表現マッチ("[0-9]+", 正規表現マッチ("(指揮兵|指揮)\s+\S+", stmp)(0))(0))
                Dim repstr As String = "replace(replace(初期スキル名, " & """ "" , """"), " & """　"" , """") = "
                'Dim repstr As String = "replace(初期スキル名, " & """ "" , """") = "
                Dim s() As String = DB_DirectOUT("SELECT 武将名, 武将R, 初期スキル名 FROM BData WHERE " _
                                & " 武将名 LIKE """ & .name & "%""" & " AND " & repstr & ダブルクオート(TrimJ(.skill_name(0))), _
                                {"武将名", "武将R"})
                .name = s(0)
                .rare = s(1)
                '武将名とレアリティがDBから正常に取得できていない場合、エラーにする
                If (.name = "" Or .rare = "") Then
                    Throw New Exception
                End If
                '出力部
                ss(cc) = "<" & .rare & ">" & vbCrLf & growstate & ss(cc)
                ss(cc) = .name & vbCrLf & ss(cc)
            Catch ex As Exception
                Me.Focus()
                MsgBox("データの読み込みが正常に行われなかった可能性があります")
                RichTextBox(Me, Val(cc) + 1).Tag = "E" 'エラータグ
                Call テキスト確定(RichTextBox(Me, Val(cc) + 1), Nothing)
                RichTextBox(Me, Val(cc) + 1).Tag = Nothing
            End Try
        End With
    End Sub

    Private Sub テキスト確定(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles RichTextBox1.TextChanged, RichTextBox2.TextChanged, RichTextBox3.TextChanged, RichTextBox4.TextChanged
        Dim cc As Integer = 部隊枠番号(CStr(sender.Name))
        If sender.Tag = "D" Then '削除マークがあれば
            sender.clear()
            ss(cc) = Nothing
        ElseIf sender.Tag = "E" Then 'エラーマークがあれば
            sender.clear()
            ss(cc) = "*** 読込異常 ***"
            sender.text = ss(cc)
        Else
            sender.clear()
            sender.text = ss(cc)
        End If
    End Sub

    Private Sub 透過度変更(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.Scroll
        Me.Opacity = 0.1 * Val(sender.value)
    End Sub

    Private Sub 武将クリア(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles Button1.Click, Button2.Click, Button3.Click, Button4.Click
        Dim tRichbox As RichTextBox = RichTextBox(Me, Val(部隊枠番号(CStr(sender.Name))) + 1)
        tRichbox.Tag = "D" '削除マーク
        Call テキスト確定(tRichbox, Nothing)
        bd(Val(部隊枠番号(CStr(sender.Name)))) = Nothing
        tRichbox.Tag = Nothing
    End Sub

    Private Sub 常に手前に表示(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Me.TopMost = True
        Else
            Me.TopMost = False
        End If
    End Sub

    Private Sub INIファイル書き込み(ByVal rogini As String)
        If IO.File.Exists(rogini) Then '既に存在すれば削除して作り直し
            IO.File.Delete(rogini)
        End If

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
                With bd(i)
                    If i = 0 Then '部隊長記憶時
                        SetINIValue(busho_counter, "busho_counter", bsho, rogini) '武将数も記憶
                    End If
                    '無いと復元できないもの（主キー）のみをINIへ保存
                    SetINIValue(.rare, "rare", bsho, rogini)
                    SetINIValue(.name, "name", bsho, rogini)
                    SetINIValue(.heisyu_name, "heisyu.name", bsho, rogini)
                    SetINIValue(.hei_sum, "hei_sum", bsho, rogini)
                    SetINIValue(.rank, "rank", bsho, rogini)
                    SetINIValue(.level, "level", bsho, rogini)
                    For j As Integer = 0 To 2
                        SetINIValue(.st(j), "st(" & j & ")", bsho, rogini)
                    Next
                    For j As Integer = 0 To 3
                        SetINIValue(.tou_a(j), "tou_a(" & j & ")", bsho, rogini)
                    Next
                    SetINIValue(.skill_no, "skill_no", bsho, rogini)
                    For j As Integer = 0 To bd(i).skill_no - 1
                        SetINIValue(.skill_name(j), "skill(" & j & ").name", bsho, rogini)
                        SetINIValue(.skill_lv(j), "skill(" & j & ").lv", bsho, rogini)
                    Next
                End With
            Next
        Catch ex As Exception
            MsgBox("DRAGDROP_BUTAI記録中にエラーがあります")
        End Try
    End Sub

    Private Sub 読み込みデータ確定(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Cursor.Current = Cursors.WaitCursor
        '武将数があっていない場合自動で訂正
        Dim ts() As Integer = Nothing 'データの抜けている部分
        Dim c As Integer = 0, d As Integer = 0
        Dim tbc As Integer 'データの入っていない武将数
        For i = 0 To 3
            If bd(i).name = Nothing Then
                ReDim Preserve ts(c)
                ts(c) = i '空き番号
                c = c + 1
                tbc = tbc + 1
            Else
                If c > 0 Then '空でなければ→その前の番号が空いている
                    bd(ts(d)) = bd(i).Clone
                    bd(i) = Nothing
                    d = d + 1
                    ReDim Preserve ts(c)
                    ts(c) = i '移動した元の場所がまた開く
                    c = c + 1
                End If
            End If
        Next
        busho_counter = 4 - tbc
        ReDim Preserve bd(busho_counter - 1)
        If bd.Length = 0 Then '武将ゼロで実行しようとしている
            MsgBox("読み込みデータがありません")
            Exit Sub
        End If
        '取り込んだ部隊情報はDRAGDROP_BUTAI.INIに保存
        Dim rogini As String = "DRAGDROP_BUTAI" & ".INI"
        Dim rogini_path As String = My.Application.Info.DirectoryPath & "\butai\" & rogini
        Call INIファイル書き込み(rogini_path)
        Call Form1.INIファイルから読み込み(rogini_path)
        Me.Close()
        Cursor.Current = Cursors.Default
    End Sub
End Class