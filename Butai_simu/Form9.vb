Public Class Form9
    Public syou As Integer = 10 '章設定
    Public ki As Integer = 10 '期設定
    Public d As Decimal '距離補正率
    Public dist As Decimal = Nothing '距離
    Public d_kakin As Boolean = False '距離補正課金の有無
    Public g_kakin As Decimal = 0 '経験値課金の有無
    Public domei_b As Decimal = 0 '同盟ボーナス(%)
    Public touzai As Boolean = True '東西無双ボーナス
    Public simai As Boolean = True '浅井三姉妹ボーナス
    Public Structure akiti_
        Public rank As Integer
        Public toti As String
        Public Structure npc_hei_
            'Public name As String '取得文字列（分類:兵科）
            Public heika As String '兵科
            Public sum As Decimal '兵数
            Public bunrui As String '分類
            Public heika_def As Decimal '兵科防御力
            Public keiken As Integer '兵科経験値
            Public def As Decimal '防御力
            Public higai As Integer '被害分担兵数
        End Structure
        Public npc_hei() As npc_hei_
        Public win_p As Decimal '勝率
        Public ex_changed_def As Decimal '見かけの防御力（期待値）
        Public higai_sum As Integer 'NPC被害兵数の総数
    End Structure
    Public akiti As akiti_

    'Private Sub Form9_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    '    Call DB_Open() '空き地DBを開く
    'End Sub

    Private Sub NPC情報初期化()
        akiti.npc_hei = Nothing
    End Sub

    Private Sub 空き地討伐シミュ実行(sender As Object, e As EventArgs) Handles Button1.Click
        If akiti.npc_hei Is Nothing Then
            MsgBox("空き地情報が不足しています")
            Exit Sub
        End If

        'Call NPC兵素防御力()
        Call 距離減衰計算()
        Call 勝率計算()

        Dim tmp() As String = Nothing '出力文字列
        RichTextBox2.Clear() 'クリア
        ReDim tmp(3)
        tmp(0) = "*** NPC側 *** " & vbCrLf & "見かけの防御力: " & ToRoundDown(Val(見かけの防御力計算(skill_ex)), 0)
        tmp(1) = "*** 攻撃側 ***" & vbCrLf & "距離補正率: " & Val(d) * 100 & "(%) / 距離: " & dist
        tmp(2) = "総攻撃力期待値: " & Int(skill_exk * d)
        tmp(3) = "勝率: " & Math.Round((akiti.win_p * 100), 2) & "%"

        For i As Integer = 0 To tmp.Length - 1
            RichTextBox2.Text = RichTextBox2.Text & tmp(i) & vbCrLf
        Next
        RTextBox_BOLD(RichTextBox2, {tmp(3)}) '勝率を太字
    End Sub

    'Atklistは，各将の素攻を格納した配列（{bs(0),bs(1),bs(2),bs(3)}の順固定）
    Public Function 見かけの防御力計算(ByVal Atklist() As Decimal) As Decimal
        Dim Atk_sum As Decimal = 0 '総攻撃力
        Dim bs_r() As Decimal '攻撃側の各武将の戦闘力比率
        ReDim bs_r(busho_counter - 1)
        Dim changed_Defsum As Decimal = 0 '見かけの防御力

        For i As Integer = 0 To busho_counter - 1
            Atk_sum = Atk_sum + Atklist(i)
        Next
        For i As Integer = 0 To busho_counter - 1
            bs_r(i) = Atklist(i) / Atk_sum
        Next
        For i As Integer = 0 To akiti.npc_hei.Length - 1
            With akiti.npc_hei(i)
                Dim tmp As Decimal = 0
                For j As Integer = 0 To busho_counter - 1
                    tmp = tmp + bs_r(j) * (.def * 相性計算(.bunrui, bs(j).heisyu.bunrui))
                Next
                changed_Defsum = changed_Defsum + tmp
            End With
        Next
        Return changed_Defsum
    End Function
    'Atklistを受け取って，勝利判定ならばその生起確率を返す（敗北ならばゼロ）
    Public Function 勝敗判定(ByVal x As Decimal, ByVal Atklist() As Decimal) As Decimal
        Dim Atk_sum As Decimal = 0
        For i As Integer = 0 To Atklist.Length - 1
            Atk_sum = Atk_sum + Atklist(i)
        Next
        If Atk_sum > 見かけの防御力計算(Atklist) Then
            Return x
        Else
            Return 0
        End If
    End Function

    Public Sub 被害兵数計算()
        akiti.ex_changed_def = 見かけの防御力計算(skill_ex) '攻撃側が期待値通りの攻撃力の時の、みかけの防衛力

        Dim ex_Atk_sum As Decimal = 0 '総攻撃力
        Dim ex_bs_r() As Decimal '攻撃側の各武将の戦闘力比率
        ReDim ex_bs_r(busho_counter - 1)
        For i As Integer = 0 To busho_counter - 1
            ex_Atk_sum = ex_Atk_sum + skill_ex(i)
        Next
        For i As Integer = 0 To busho_counter - 1
            ex_bs_r(i) = skill_ex(i) / ex_Atk_sum
        Next

        If akiti.ex_changed_def >= ex_Atk_sum Then '防衛側勝利（＝攻撃側敗北）

        End If
    End Sub

    Private Sub 勝率計算()
        akiti.win_p = 0
        Dim tmp_skill_y() As Decimal
        ReDim tmp_skill_y(UBound(skill_y, 2))
        For i As Integer = 0 To UBound(skill_y)
            For j As Integer = 0 To UBound(skill_y, 2)
                tmp_skill_y(j) = skill_y(i, j) * d
            Next
            akiti.win_p = akiti.win_p + 勝敗判定(skill_x(i), tmp_skill_y)
        Next
    End Sub

    Private Sub 空き地ランク変更(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If sender.Text = "" Then
            Exit Sub
        End If
        akiti.rank = Val(sender.Text)
        RemoveHandler ComboBox3.SelectedIndexChanged, AddressOf Me.空き地変更 'これが無いと空き地を選べなくなる
        Dim p As DataSet
        p = DB_TableOUT("SELECT 土地ランク, パネル配置 FROM LName WHERE 土地ランク = " & ダブルクオート("☆" & akiti.rank) & "", "LName")
        With ComboBox3
            .BeginUpdate()
            .DataSource = p.Tables("LName")
            .DisplayMember = "パネル配置"
            '.ValueMember = "Index"
            .SelectedIndex = -1
            .EndUpdate()
        End With
        AddHandler ComboBox3.SelectedIndexChanged, AddressOf Me.空き地変更
    End Sub

    '現在は10章以降のみ対応
    Private Sub 章期変更(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If sender.Text = "" Then
            Exit Sub
        End If
        syou = 10 '10章固定
        ki = Val(sender.text)
        '章期設定が変われば一旦クリア
        ComboBox2.Text = Nothing
        ComboBox3.Text = Nothing
        RichTextBox1.Text = Nothing
    End Sub

    Private Sub 空き地変更(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If sender.text = "" Then
            Exit Sub
        End If
        Call NPC情報初期化()
        akiti.toti = sender.text
        'NPC兵を取得・表示   兵種名, 土地ランク, パネル配置, NPC兵数
        RichTextBox1.Clear()
        Dim sqlstr As String, heika()() As String
        sqlstr = "SELECT * FROM HData INNER JOIN (LData INNER JOIN LName ON LData.id = LName.id) ON HData.兵種名 = LData.兵種名 " & _
            "WHERE 期数 = " & ki & " AND 土地ランク = " & ダブルクオート("☆" & akiti.rank) & " AND パネル配置 = " & ダブルクオート(akiti.toti) & ""
        heika = DB_DirectOUT3(sqlstr, {"兵科", "兵種名", "防御値", "経験値", "NPC兵数"})
        For i As Integer = 0 To heika.Length - 1
            ReDim Preserve akiti.npc_hei(i)
            With akiti.npc_hei(i)
                '.name = heika(i)
                .bunrui = heika(i)(0)
                .heika = heika(i)(1)
                .heika_def = heika(i)(2)
                .keiken = heika(i)(3)
                .sum = heika(i)(4)
                .def = .heika_def * .sum
                RichTextBox1.Text = RichTextBox1.Text & .bunrui & ":" & .heika & " " & .sum & vbCrLf
            End With
        Next
    End Sub

    'Private Sub NPC兵素防御力()
    '    For i As Integer = 0 To akiti.npc_hei.Length - 1
    '        With akiti.npc_hei(i)
    '            Dim tmp() As String = Split(Mid(.name, 3), " ")
    '            .heika = tmp(0)
    '            .bunrui = Mid(.name, 1, 1)
    '            .sum = tmp(1)
    '            Dim s() As String = _
    '                DB_DirectOUT("SELECT 兵種名, 防御値, 経験値 FROM HData WHERE 兵種名 = " & ダブルクオート(.heika) & "", {"防御値", "経験値"})
    '            .heika_def = s(0)
    '            .keiken = s(1)
    '            .def = .heika_def * .sum
    '        End With
    '    Next
    'End Sub

    Private Sub 距離課金の有無(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then '有ならば
            d_kakin = True
            Label7.ForeColor = Color.Red
        Else
            d_kakin = False
            Label7.ForeColor = Color.Black
        End If
    End Sub

    Private Sub 経験値課金の有無(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then '有ならば
            g_kakin = 30

        Else
            g_kakin = 0
        End If
    End Sub

    Private Sub 距離減衰計算()
        dist = Val(TextBox2.Text)
        If dist = 0 Then Exit Sub '距離が数字でなければ抜ける
        If dist <= 10 Then
            d = 1
        ElseIf dist > 10 And dist <= 20 Then
            If d_kakin Then
                d = 1
            Else
                d = ToRoundDown(19 / (dist + 9), 2)
            End If
        Else
            d = ToRoundDown(19 / (dist + 9), 2)
        End If
    End Sub

    Private Sub 出力クリア(sender As Object, e As EventArgs) Handles Button2.Click
        RichTextBox2.Clear()
        RichTextBox3.Clear()
    End Sub

    Private Sub 同盟ボーナス変更(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Dim tmp As String = sender.text
        If InStr(tmp, "%") Then
            tmp = Replace(tmp, "%", "")
        End If
        domei_b = 0.01 * Val(tmp)
    End Sub

    Private Sub 東西無双(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then '有ならば
            touzai = True
        Else
            touzai = False
        End If
    End Sub

    Private Sub 三姉妹(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then '有ならば
            simai = True
        Else
            simai = False
        End If
    End Sub
End Class