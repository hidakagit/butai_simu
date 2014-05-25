Public Class Form2

    Public Xitv As Long = 10000 'X軸の間隔はデフォ1万
    Public hstitv As Long = 1000 'ヒストグラムの間隔
    Public xhst() As Decimal 'ヒストグラムの度数
    Public yhst() As Long 'ヒストグラムの横軸
    Dim minhst, maxhst As Long 'ヒストグラムの最大値・最小値

    Public max_x As Decimal 'X軸の最大値
    Public min_x As Decimal 'X軸の最小値
    Public max_y As Decimal 'Y軸の最大値
    Public min_y As Decimal 'Y軸の最小値
    
    Private Sub ChartClar(ByVal cht As System.Windows.Forms.DataVisualization.Charting.Chart)
        'Chart の設定を初期値に戻す(通常は必要ありません)
        With cht
            .Titles.Clear()                  'タイトルの初期化
            '.BackGradientStyle = GradientStyle.None
            .BackColor = Color.White         '背景色を白色に
            '外形をデフォルトに
            '.BorderSkin.SkinStyle = BorderSkinStyle.None
            .Legends.Clear()                 '凡例の初期化
            .Legends.Add("Legend1")
            .Series.Clear()                  '系列(データ関係)の初期化
            .ChartAreas.Clear()              '軸メモリ・3D 表示関係の初期化
            .ChartAreas.Add("ChartArea1")
            .Annotations.Clear()
        End With
    End Sub

    Public Sub グラフ描画()
        max_x = skill_x(0)
        min_x = skill_x(0)
        max_y = skill_yk(0)
        min_y = skill_yk(0)
        For i As Integer = 0 To skill_x.Length - 1
            If max_x < skill_x(i) Then
                max_x = skill_x(i)
            End If
            If min_x > skill_x(i) Then
                min_x = skill_x(i)
            End If
            If max_y < skill_yk(i) Then
                max_y = skill_yk(i)
            End If
            If min_y > skill_yk(i) Then
                min_y = skill_yk(i)
            End If
        Next

        With Chart1
            Call ChartClar(Chart1)
            .Titles.Add("総戦闘力分布グラフ")
            .Series.Add("棒グラフ")
            .Legends(0).Enabled = False '凡例を非表示
            With .ChartAreas(0)
                With .AxisX
                    .Title = "総" & kb & "力"
                    .Minimum = Math.Floor(min_y / 10000) * 10000
                    ToolStripTextBox1.Text = .Minimum
                    .Maximum = Math.Ceiling(max_y / 10000) * 10000
                    ToolStripTextBox2.Text = .Maximum
                    .Interval = Xitv
                    .MajorGrid.Interval = 10000000 'あり得ないインターバルにすることで期待値部分1本のみにする
                    .MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.DashDot
                    .MajorGrid.IntervalOffset = Val(skill_exk) - .Minimum
                    .MajorGrid.LineColor = Color.Red
                End With
                With .AxisY
                    .Title = "生起確率"
                    .Minimum = 0
                    If max_x >= 0.1 Then '0.1よりも小さければ0.01単位のグラフにする
                        .Maximum = (Math.Ceiling(max_x * 10)) / 10
                        .Interval = 0.1
                    Else
                        .Maximum = (Math.Ceiling(max_x * 100)) / 100
                        .Interval = 0.01
                    End If
                End With
            End With
            For i As Integer = 0 To skill_x.Length - 1
                .Series("棒グラフ").Points.AddXY(skill_yk(i), skill_x(i))
            Next
        End With

        '期待値、分散記入
        ToolStripLabel8.Text = "期待値: " & Val(Int(skill_exk))
        ToolStripLabel9.Text = "標準偏差: " & Val(Int(skill_ax))
    End Sub

    Private Sub ツールチップ表示(ByVal sender As Object, _
         ByVal e As System.Windows.Forms.DataVisualization.Charting.ToolTipEventArgs) _
                                                   Handles Chart1.GetToolTipText
        'データの上にマウスを持ってきたときに表示するツールチップ
        If e.HitTestResult.ChartElementType = DataVisualization.Charting.ChartElementType.DataPoint Then
            Dim i As Integer = e.HitTestResult.PointIndex
            Dim dp As DataVisualization.Charting.DataPoint = e.HitTestResult.Series.Points(i)
            '科目名、生徒名、点数 をツールチップテキストに表示
            e.Text = String.Format("生起確率 = {0}, 総戦闘力 = {1}", Math.Ceiling(dp.YValues(0) * 10000) / 10000, dp.XValue)
        End If
        '目盛（攻撃力数字）の上にマウスを持ってきたときに表示するツールチップ
        If e.HitTestResult.ChartElementType = DataVisualization.Charting.ChartElementType.AxisLabels Then
            If e.HitTestResult.Axis.AxisName = 0 Then 'X軸ならば
                Dim j As Double = Val(e.HitTestResult.Object.text)
                e.Text = String.Format("{0}↑{1}{2}%で達成", _
                                       j, vbCrLf, Math.Ceiling(要求攻撃力達成率(j) * 10000) / 100)
            End If
        End If
    End Sub

    Private Function 要求攻撃力達成率(ByVal p As Double) As Double
        Dim k As Integer = -1
        Dim sum As Double = 0 '達成率
        For i As Integer = 0 To skill_yk.Length - 1
            If skill_yk(i) >= p Then '要求攻撃力を超える
                k = i
                Exit For
            End If
        Next
        If k = -1 Then Return -1
        For i As Integer = k To skill_x.Length - 1
            sum = sum + skill_x(i)
        Next
        If sum > 1 Then sum = 1 '100.01%みたいにならないように
        Return sum
    End Function

    Private Function 実現範囲戦闘力計算(ByVal p As Double) As Decimal()
        Dim lpoint As Decimal = min_y
        Dim hpoint As Integer = max_y '下限値と上限値
        Dim sum As Double = 0 '累積確率

        '下限計算
        For i As Integer = 0 To skill_x.Length - 1
            sum = sum + skill_x(i)
            If 0.5 - (p / 2) <= sum Then
                lpoint = skill_yk(i)
                Exit For
            End If
        Next
        '上限計算
        sum = 0
        For i As Integer = skill_x.Length - 1 To 0 Step -1
            sum = sum + skill_x(i)
            If 0.5 - (p / 2) <= sum Then '上限に到達
                hpoint = skill_yk(i)
                Exit For
            End If
        Next
        Return {lpoint, hpoint}
    End Function

    Private Sub X軸間隔変更(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        If sender.text = "" Then
            Exit Sub
        End If
        Xitv = Val(Replace(sender.text, "万", "0000"))
        Chart1.ChartAreas(0).AxisX.Interval = Xitv
    End Sub

    Private Sub 表示領域切り出し(ByVal sender As System.Object, ByVal e As MouseEventArgs) Handles ToolStripButton1.MouseUp
        If e.Button = MouseButtons.Right Then '右クリックなら
            With Chart1.ChartAreas(0).AxisX '切り出しをデフォに戻す
                .Minimum = Math.Floor(min_y / 10000) * 10000
                .Maximum = Math.Ceiling(max_y / 10000) * 10000
                ToolStripTextBox1.Text = .Minimum
                ToolStripTextBox2.Text = .Maximum
            End With
        Else
            Dim AxisX_max, AxisX_min As Decimal
            If ToolStripTextBox1.Text = "" Or ToolStripTextBox2.Text = "" Then
                MsgBox("最小値・最大値を両方設定してください")
                Exit Sub
            End If
            If Val(ToolStripTextBox1.Text) > Val(ToolStripTextBox2.Text) Then
                Dim tmp As Decimal
                tmp = Val(ToolStripTextBox2.Text)
                ToolStripTextBox2.Text = Val(ToolStripTextBox1.Text)
                ToolStripTextBox1.Text = tmp
            ElseIf Val(ToolStripTextBox1.Text) = Val(ToolStripTextBox2.Text) Then
                MsgBox("最小値と最大値が等しくなっています")
                Exit Sub
            End If
            AxisX_max = Val(ToolStripTextBox2.Text)
            AxisX_min = Val(ToolStripTextBox1.Text)
            With Chart1.ChartAreas(0).AxisX
                .Maximum = AxisX_max
                .Minimum = AxisX_min
            End With
        End If
        Chart1.ChartAreas(0).AxisX.MajorGrid.IntervalOffset = _
            Val(skill_exk) - Chart1.ChartAreas(0).AxisX.Minimum 'minimumが変わる都度更新
    End Sub

    Private Sub ヒストグラム間隔変更(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox2.SelectedIndexChanged
        If Chart1.Series.Count = 0 Then 'これが無いとFormをロードする前に乙る
            Exit Sub
        End If
        Call ヒストグラム描画(sender, Nothing)
    End Sub

    Private Sub ヒストグラム描画(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ヒストグラムToolStripMenuItem.Click

        ToolStripComboBox2.Visible = True
        ToolStripLabel4.Visible = True
        '*** 目標火力測定部分 ***
        ToolStripLabel5.Visible = False
        ToolStripTextBox3.Visible = False
        ToolStripLabel7.Visible = False
        ToolStripButton3.Visible = False
        ToolStripSeparator4.Visible = False
        ToolStripLabel6.Visible = False



        Select Case ToolStripComboBox2.Text
            Case "default"
                hstitv = (3.49 * skill_ax) / ((skill_y.Length) ^ (1 / 3)) 'Scottの選択により区間を決める
            Case "1万"
                hstitv = 10000
            Case "5万"
                hstitv = 50000
            Case "10万"
                hstitv = 100000
            Case "20万"
                hstitv = 200000
            Case Else
                hstitv = Val(ToolStripComboBox2.Text)
        End Select

        If hstitv = 0 Then '区間ゼロになった場合
            hstitv = 1 + (Math.Log10(skill_y.Length) / Math.Log10(2)) '1 + log2n
        End If
        Dim c As Long = 0
        xhst = Nothing
        yhst = Nothing

        minhst = Math.Floor(min_y / hstitv) * hstitv
        maxhst = Math.Ceiling(max_y / hstitv) * hstitv
        Dim mhc As Decimal = 0 '最大度数

        Dim sd As Integer = 0 '振り分け済の数値数
        For i As Integer = 0 To (maxhst - minhst) / hstitv - 1
            ReDim Preserve xhst(c), yhst(c)
            yhst(c) = minhst + i * hstitv
            For d As Long = sd To UBound(skill_y)
                If (minhst + i * hstitv) <= skill_yk(d) And skill_yk(d) < (minhst + (i + 1) * hstitv) Then 'min以上、(min+1)未満
                    xhst(c) = xhst(c) + skill_x(d)
                    sd = sd + 1
                Else
                    c = c + 1
                    Exit For
                End If
            Next
        Next

        If xhst Is Nothing Then 'minhst = maxhstになったりしたらこうなる
            mhc = 1
            ReDim xhst(0)
            xhst(0) = 1
        End If

        For i As Integer = 0 To xhst.Length - 1
            If xhst(i) > mhc Then
                mhc = xhst(i)
            End If
        Next
        With Chart1
            .Series("棒グラフ").Enabled = False
            If .Series.Count = 2 Then '既にヒストグラム作成済ならば
                .Series.RemoveAt(1) 'つぶして作り直す
            End If
            With .Series.Add("ヒストグラム")
                .ChartType = DataVisualization.Charting.SeriesChartType.Column
            End With
            .Series("ヒストグラム")("PointWidth") = "1.0"
            With .ChartAreas(0)
                With .AxisX
                    .Title = "ヒストグラム"
                    .Minimum = Math.Floor(minhst / 10000) * 10000
                    ToolStripTextBox1.Text = .Minimum
                    .Maximum = Math.Ceiling(maxhst / 10000) * 10000
                    ToolStripTextBox2.Text = .Maximum
                    .Interval = Xitv
                    .MajorGrid.IntervalOffset = Val(skill_exk) - .Minimum 'minimumが変わる都度更新
                End With
                With .AxisY
                    .Title = "相対度数"
                    .Minimum = 0
                    If mhc < 0.1 Then '最大度数が0.1より小さい
                        .Maximum = Math.Ceiling(mhc * 100) / 100 '最大値はとりうる最大度数切り上げ
                        .Interval = 0.01
                    Else
                        .Maximum = Math.Ceiling(mhc * 10) / 10 '最大値はとりうる最大度数切り上げ
                        .Interval = 0.1
                    End If
                End With
            End With
            For i As Integer = 0 To xhst.Length - 1
                Dim med As Long = minhst + (i + 0.5) * hstitv
                .Series("ヒストグラム").Points.AddXY(med, xhst(i))
            Next
        End With
    End Sub

    Private Sub 棒グラフ表示(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 棒グラフToolStripMenuItem.Click

        ToolStripComboBox2.Visible = False
        ToolStripLabel4.Visible = False
        ToolStripLabel6.Text = "該当確率: -----"
        '*** 目標火力測定部分 ***
        ToolStripLabel5.Visible = True
        ToolStripTextBox3.Visible = True
        ToolStripLabel7.Visible = True
        ToolStripButton3.Visible = True
        ToolStripSeparator4.Visible = True
        ToolStripLabel6.Visible = True

        Call グラフ描画()
    End Sub

    Private Sub 威力推定(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        Dim hope_power As Decimal = Val(ToolStripTextBox3.Text) * 10000 '希望火力
        If hope_power <= 0 Then '数値でなければ
            ToolStripTextBox3.Clear()
            ToolStripLabel6.Text = "該当確率: -----"
            Exit Sub
        End If
        Dim hp As Decimal = 要求攻撃力達成率(hope_power) '希望火力の達成率
        Dim hps As String '表示文字列
        If hope_power > max_y Then
            hps = " ⇒ " & "0%で達成 (達成不可能)"
        ElseIf hope_power <= min_y Then
            If (hope_power / Atksum) - 1 < 0 Then '上昇率0%で達成できる
                hps = " ⇒ " & "100%で達成 (絶対火力)"
            Else
                hps = " ⇒ " & "100%で達成 (+" & Format(((hope_power / Atksum) - 1) * 100, "#.00") & "%上昇が必要)"
            End If
        Else
            hps = " ⇒ " & Format(hp * 100, "#.00") & "%で達成 (+" & _
            Format(((hope_power / Atksum) - 1) * 100, "#.00") & "%上昇が必要)"
        End If
        ToolStripLabel6.Text = hps
    End Sub

    'Private Sub Z推定(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
    '    If ToolStripComboBox3.Text = "" Then '信頼区間が空白ならば
    '        Exit Sub
    '    End If
    '    Dim confd As Decimal = 0.01 * Val(ToolStripComboBox3.Text) '信頼区間
    '    Dim z As Decimal = -3
    '    Dim sumz, lc, rc As Decimal
    '    While z <= 0
    '        sumz = qnorm(Math.Abs(z)) - qnorm(z)
    '        If sumz <= confd Then
    '            Exit While
    '        End If
    '        z = z + 0.001
    '    End While

    '    If z = 0 Then 'z値が見つからなかった
    '        MsgBox("エラー。z値がみつかりません。")
    '        Exit Sub
    '    End If
    '    lc = skill_exk - z * skill_ax
    '    rc = skill_exk + z * skill_ax
    '    If Atksum > rc Then '素戦闘力よりも下側が小さい場合
    '        rc = Atksum
    '        ToolStripLabel6.Text = " ※⇒ " & Int(rc) & " ～ " & Int(lc)
    '    Else
    '        ToolStripLabel6.Text = " ⇒ " & Int(rc) & " ～ " & Int(lc)
    '    End If

    '    'With Chart1 後々、期待値もグラフ内に表示するかも？
    '    '    If .Series.Count = 3 Then '既に期待値・分散が存在
    '    '        .Series.RemoveAt(2) '削除して作り直す
    '    '    End If
    '    '    With .Series.Add("期待値・分散")
    '    '        .Points.AddXY(skill_ex, 1)
    '    '    End With
    '    'End With
    'End Sub

    Private Sub グラフを保存(ByVal sender As System.Object, ByVal e As MouseEventArgs) Handles ToolStripSplitButton1.MouseUp
        If e.Button = MouseButtons.Left Then '左クリックなら
            Exit Sub
        End If
        Dim yn As String
        yn = InputBox("保存ファイル名(.bmp)", "現在表示されているグラフを保存しますか？")
        If yn = "" Then
            Exit Sub
        End If
        Using st As New System.IO.FileStream(My.Application.Info.DirectoryPath & "\" & yn & ".bmp", System.IO.FileMode.Create)
            Chart1.SaveImage(st, System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Bmp)
        End Using
    End Sub

    Private Sub 空地討伐シミュ起動(sender As Object, e As EventArgs) Handles Button1.Click
        Form9.Show()
    End Sub
End Class