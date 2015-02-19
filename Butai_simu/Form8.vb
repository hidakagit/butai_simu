Public Class Form8
    Public kobo As String '攻防
    Public busho_heika As String '兵科
    Public busho_rank As Integer 'コスト
    Public kosuhi As Integer 'コス比
    Public tous As Decimal '統率
    Public heiho As Decimal '兵法成長
    Public Busho() As _Busho '武将情報を格納
    Public bc As Integer '武将数(-1)
    Public Structure _Busho
        Public bno As Integer
        Public r As String
        Public name As String
        Public cost As Decimal
        Public hei As Integer
        Public job As String
        Public Structure _Tousotu
            Public yari As String
            Public yumi As String
            Public uma As String
            Public utuwa As String
        End Structure
        Public Structure _Ini
            Public kou As Decimal
            Public bou As Decimal
            Public hei As Decimal
        End Structure
        Public Structure _Grow
            Public kou As Decimal
            Public bou As Decimal
            Public hei As Decimal
        End Structure
        Public Structure _Skill
            Public name As String
            Public taisyo As String
            Public p As Decimal
            Public k As Decimal
        End Structure
        Public tousotu As _Tousotu
        Public ini As _Ini
        Public grow As _Grow
        Public skill As _Skill
    End Structure

    Private Function 条件入力() As Boolean
        Dim flg As Boolean = True
        '攻防
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
        '兵科
        If ToolStripComboBox2.Text = "" Or ToolStripComboBox2.Text = "-----" Then
            MsgBox("兵科が未設定です")
            flg = False
        Else
            busho_heika = ToolStripComboBox2.Text
        End If
        '武将ランク
        Select Case ToolStripComboBox3.Text
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
            Case Else
                MsgBox("武将ランクが未設定です")
                flg = False
        End Select
        '統率
        Select Case ToolStripComboBox4.Text
            Case "1.2 (SSS)"
                tous = 1.2
            Case "1.15 (SS)"
                tous = 1.15
            Case "1.1 (S)"
                tous = 1.1
            Case "1.05 (A)"
                tous = 1.05
            Case "1.0 (B)"
                tous = 1.0
            Case Else
                MsgBox("色づけ統率が未設定です")
                flg = False
        End Select
        '兵法成長
        If ToolStripComboBox5.Text = "" Then
            MsgBox("色づけ兵法成長値が未設定です")
            flg = False
        Else
            heiho = Val(ToolStripComboBox5.Text)
        End If
        '実質コス比
        If ToolStripComboBox6.Text = "" Then
            MsgBox("色づけ実質コス比値が未設定です")
            flg = False
        Else
            kosuhi = Val(ToolStripComboBox6.Text)
        End If
        Return flg
    End Function

    Private Sub プログレスバー初期化()
        ToolStripProgressBar1.Value = 0
    End Sub

    Private Sub プログレスバー変化(ByVal val As Decimal)
        ToolStripProgressBar1.Value = val
    End Sub

    Private Function 武将データ取得() As Integer
        Dim sl()() As String = Nothing
        Dim slbl() As String = _
        {"Bid", "武将R", "武将名", "Cost", "指揮兵数", "槍統率", "弓統率", "馬統率", "器統率", "初期攻撃", "初期防御", "初期兵法", "攻成長", "防成長", "兵成長", "初期スキル名", "職"}
        Dim sqlstr As String = "SELECT * FROM BData WHERE ( NOT 武将R IN('祝', '雅', '化')) AND Bunf = 'F' ORDER BY Bid ASC"
        sl = DB_DirectOUT3(sqlstr, slbl)
        Dim sc As Integer = sl.GetLength(0) - 1 '合致したデータ個数
        For i As Integer = 0 To sc
            ReDim Preserve Busho(i)
            With Busho(i)
                .bno = Val(sl(i)(0))
                .r = sl(i)(1)
                .name = sl(i)(2)
                .cost = Val(sl(i)(3))
                With .tousotu
                    .yari = sl(i)(5)
                    .yumi = sl(i)(6)
                    .uma = sl(i)(7)
                    .utuwa = sl(i)(8)
                End With
                With .ini
                    .kou = Val(sl(i)(9))
                    .bou = Val(sl(i)(10))
                    .hei = Val(sl(i)(11))
                End With
                With .grow
                    .kou = Val(sl(i)(12))
                    .bou = Val(sl(i)(13))
                    .hei = Val(sl(i)(14))
                End With
                With .skill
                    .name = sl(i)(15)
                    'Dim sk() As String = _
                    '    Skill_ref(Mid(.name, InStr(.name, "：") + 1, .name.Length - InStr(.name, "：")), 1)
                    '.taisyo = sk(1)
                End With
                .job = sl(i)(16)
                If Not .job = "剣" Then
                    If .job = "覇" Then
                        .hei = Val(sl(i)(4)) + busho_rank * 200
                    Else
                        .hei = Val(sl(i)(4)) + busho_rank * 100
                    End If
                Else
                    .hei = Val(sl(i)(4))
                End If
            End With
        Next
        Return sc
    End Function

    Private Sub 表更新(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim jf As Boolean = 条件入力()
        If jf = False Then '条件に不備があれば
            MsgBox("必要な条件が設定されていない個所があります")
            Exit Sub
        End If
        Call プログレスバー初期化()
        DataGridView1.Rows.Clear() '表をクリア
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        Call プログレスバー変化(250)
        Dim sklt() As String
        Dim td(), heih(), syou(), tanh(), jih(), tank(), jik() As Decimal
        Dim target_u() As String = {"武士", "弓騎馬", "赤備え", "騎馬鉄砲", "鉄砲足軽", "焙烙火矢", "大筒兵", "破城鎚", "攻城櫓"}
        Dim target_h() As String = {"武士", "弓騎馬", "赤備え", "騎馬鉄砲", "鉄砲足軽", "焙烙火矢", "大筒兵", "雑賀衆"}
        Dim target_hi() As String = {"国人衆", "雑賀衆", "海賊衆", "母衣衆"}
        Dim jkf_u As Boolean = False '上級器フラグ
        Dim jkf_h As Boolean = False '上級砲フラグ
        Dim hik_f As Boolean = False '秘境兵フラグ
        For i As Integer = 0 To target_u.Length - 1 '上級器対応かどうか
            If InStr(busho_heika, target_u(i)) Then
                jkf_u = True
            End If
        Next
        For i As Integer = 0 To target_h.Length - 1 '上級器対応かどうか
            If InStr(busho_heika, target_h(i)) Then
                jkf_h = True
            End If
        Next
        For i As Integer = 0 To target_hi.Length - 1 '上級器対応かどうか
            If InStr(busho_heika, target_hi(i)) Then
                hik_f = True
            End If
        Next
        Dim s() As String = _
         DB_DirectOUT("SELECT 統率, 兵種名, 兵科, 攻撃値, 防御値 FROM HData WHERE 兵種名 = " & ダブルクオート(busho_heika) & "", {"兵科", "統率", "攻撃値", "防御値"})
        Dim t As Integer = s(1).Length '統率に関係する兵科の種類数
        Dim heika As String = s(0) '兵種兵科
        Dim tousotu As String = s(1) '兵種の対応兵科
        Dim heika_kou As Integer = s(2) '兵科攻撃力
        Dim heika_bou As Integer = s(3) '兵科防御力

        '表示データを全取得
        Call プログレスバー変化(500)
        'If Busho Is Nothing Then 'まだ武将データが無い場合(初回起動時)
        bc = 武将データ取得() 'ランクが上がって指揮兵数が変わるようになったので毎回更新が要る＾＾；
        'End If
        ReDim sklt(bc), td(bc), heih(bc), syou(bc), tanh(bc), jih(bc), tank(bc), jik(bc)
        '表示データを順に加工
        For i As Integer = 0 To bc
            '実質統率
            Dim v() As Decimal
            ReDim v(t - 1)
            For j As Integer = 1 To t
                Select Case Mid(tousotu, j, 1)
                    Case "槍"
                        v(j - 1) = 統率_数値変換(Busho(i).tousotu.yari)
                    Case "弓"
                        v(j - 1) = 統率_数値変換(Busho(i).tousotu.yumi)
                    Case "馬"
                        v(j - 1) = 統率_数値変換(Busho(i).tousotu.uma)
                    Case "器"
                        v(j - 1) = 統率_数値変換(Busho(i).tousotu.utuwa)
                End Select
                td(i) = td(i) + v(j - 1)
            Next
            td(i) = td(i) / t '該当兵科の実質統率値
            If td(i) + busho_rank * (0.05 / t) < 1.2 Then 'ランクアップ時の統率変化
                td(i) = td(i) + busho_rank * (0.05 / t)
            Else
                td(i) = 1.2
            End If
            'スキル適用、不適用
            'If InStr(sl(13)(i), kobo) Then '攻防一致
            '    Dim sk() As String = _
            '    Skill_ref(Mid(sl(13)(i), InStr(sl(13)(i), "：") + 1, sl(13)(i).Length - InStr(sl(13)(i), "：")), 1)
            '    If InStr(sk(1), heika) Then
            '        sklt(i) = "◎"
            '    ElseIf jkf = True And InStr(sk(1), "上級器") Then
            '        sklt(i) = "◎"
            '    End If
            'End If
            'LV20になった時の将攻/防の値、実質兵数
            Dim stpt As Integer 'ステ振りポイントの合計
            If busho_rank < 2 Then
                stpt = (busho_rank * 20 + 20) * 4
            ElseIf busho_rank < 4 Then
                stpt = 160 + ((busho_rank - 2) * 20 + 20) * 5
            ElseIf busho_rank < 6 Then
                stpt = 360 + ((busho_rank - 4) * 20 + 20) * 6
            ElseIf busho_rank = 6 Then 'つまりは限界突破時
                stpt = 630
            End If
            With Busho(i)
                heih(i) = .ini.hei + stpt * .grow.hei
                If InStr(kobo, "攻") Then
                    syou(i) = .ini.kou + stpt * .grow.kou
                    jih(i) = Math.Floor(((.hei * heika_kou + syou(i)) * td(i)) / heika_kou)
                    tanh(i) = Math.Floor(((.hei * heika_kou + .ini.kou) * td(i)) / heika_kou)
                Else
                    syou(i) = .ini.bou + stpt * .grow.bou
                    jih(i) = Math.Floor(((.hei * heika_bou + syou(i)) * td(i)) / heika_bou)
                    tanh(i) = Math.Floor(((.hei * heika_bou + .ini.bou) * td(i)) / heika_bou)
                End If
                If .cost = 0 Then
                    jik(i) = 99999
                    tank(i) = 99999
                Else
                    jik(i) = Math.Floor(jih(i) / .cost)
                    tank(i) = Math.Floor(tanh(i) / .cost)
                End If
            End With
        Next

        'Datagridに追加
        'DataGridView1.VirtualMode = True
        Dim rows As DataGridViewRow()
        ReDim rows(bc)

        For i As Integer = 0 To bc
            Call プログレスバー変化(500 + ((i + 1) / (bc + 1)) * 500)
            Dim row As DataGridViewRow = New DataGridViewRow()
            Dim cell As DataGridViewCell
            row.CreateCells(DataGridView1)
            With Busho(i)
                If .cost = 193 Then Continue For 'デッキセットできないカードは飛ばす
                'Dim jikX As String
                'If jik(i) = -1 Then jikX = "∞" Else jikX = jik(i)
                'Dim tankX As String
                'If tank(i) = -1 Then tankX = "∞" Else tankX = tank(i)
                row.SetValues(New Object() {.bno, .r, .name, .cost, .hei, td(i), .skill.name, heih(i), .grow.hei, syou(i), tanh(i), jih(i), tank(i), jik(i)})
                'DataGridView1.Rows.AddRange(row)
                'cell = DataGridView1.Rows(DataGridView1.Rows.Count - 2).Cells(8)
                cell = row.Cells(8)
                If cell.Value >= heiho Then '兵法成長色づけ
                    cell.Style.ForeColor = Color.DarkRed
                    cell.Style.Font = New Font("Consolas", 10, FontStyle.Bold)
                End If
                'cell = DataGridView1.Rows(DataGridView1.Rows.Count - 2).Cells(13)
                'If cell.Value = "∞" Then 'コス比色づけ
                'cell.Style.ForeColor = Color.DarkGreen
                'cell.Style.Font = New Font("Consolas", 10, FontStyle.Bold)
                cell = row.Cells(13)
                If cell.Value >= kosuhi Then
                    cell.Style.ForeColor = Color.DarkGreen
                    cell.Style.Font = New Font("Consolas", 10, FontStyle.Bold)
                End If
                'cell = DataGridView1.Rows(DataGridView1.Rows.Count - 2).Cells(5)
                cell = row.Cells(5)
                If cell.Value >= tous Then '統率色づけ
                    cell.Style.ForeColor = Color.MidnightBlue
                    cell.Style.Font = New Font("Consolas", 10, FontStyle.Bold)
                End If
            End With
            'cell = DataGridView1.Rows(DataGridView1.Rows.Count - 2).Cells(1) '武将レアリティによる色分け
            cell = row.Cells(1)
            Dim cell2 As DataGridViewCell = row.Cells(2) 'DataGridView1.Rows(DataGridView1.Rows.Count - 2).Cells(2)
            Select Case Mid(CStr(row.Cells(0).Value), 1, 2) 'Mid(CStr(DataGridView1.Rows(DataGridView1.Rows.Count - 2).Cells(0).Value), 1, 2)
                Case "10" '天
                    cell.Style.ForeColor = Color.Goldenrod
                    cell2.Style.ForeColor = Color.Goldenrod
                Case "19" '祝
                    cell.Style.ForeColor = Color.Magenta
                    cell2.Style.ForeColor = Color.Magenta
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
            rows(i) = row
        Next
        DataGridView1.Rows.AddRange(rows)
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Call プログレスバー変化(0)
    End Sub
    'Private Sub DataGridView1_CellValueNeeded(ByVal sender As Object, _
    '            ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs) _
    '                                        Handles DataGridView1.CellValueNeeded
    '    e.Value = e.RowIndex.ToString & "," & e.ColumnIndex.ToString
    'End Sub

End Class