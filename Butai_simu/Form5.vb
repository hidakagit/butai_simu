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
        With bd(cc)
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
                ss(cc) = "<" & .rare & ">" & vbCrLf & tmp(2)
                ReDim .skill_name(.skill_no - 1), .skill_lv(.skill_no - 1)
                Dim syokisk As String = Nothing '初期スキル名
                For j As Integer = 0 To bd(cc).skill_no - 1
                    Dim ttmp As String = Replace(tmp(10 + j), "技" & (j + 1) & vbTab, "") '"技1"みたいなのが付いてる場合、外す
                    If j = 0 Then
                        syokisk = Mid(ttmp, 1, InStr(ttmp, "LV") - 1)
                    End If
                    .skill_name(j) = Mid(ttmp, 1, InStr(ttmp, "LV") - 1)
                    .skill_lv(j) = Val(Mid(ttmp, InStr(ttmp, "LV") + 2))
                    ss(cc) = ss(cc) & vbCrLf & ttmp
                Next
                Dim ntmp As String = GetINIValue(syokisk, Replace(Replace(tmp(0), "名", ""), "レア", "・"), bnpath)
                If Not ntmp = "－" Then
                    .name = ntmp
                Else
                    .name = Mid(tmp(0), 2, InStr(tmp(0), "レア") - 2)
                End If
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

    Private Sub INIファイル書き込み(ByVal path As String, ByVal rogini As String)
        If IO.File.Exists(path) Then '既に存在すれば削除して作り直し
            IO.File.Delete(path)
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
        Call INIファイル書き込み(rogini_path, rogini)
        Call Form1.INIファイルから読み込み(rogini_path)
        Me.Close()
        Cursor.Current = Cursors.Default
    End Sub
End Class