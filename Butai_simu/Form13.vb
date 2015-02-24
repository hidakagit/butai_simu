Imports System.Data.SQLite
Imports System.Linq
Imports System.Reflection

Public Class Form13
    'database 関係
    'Public Connection As New SQLiteConnection
    'Public Command As SQLiteCommand
    'Public DataReader As SQLiteDataReader
    'Public SqlString As String

    Public BData_items As String() = {"武将R", "武将名", "Bid", "Cost", "指揮兵数", "槍統率", "弓統率", "馬統率", "器統率", _
                                   "初期攻撃", "初期防御", "初期兵法", "攻成長", "防成長", "兵成長", "初期スキル名", "職", "Bunf"}
    Public BData_dists As String() = {"武将のレアリティ。設定可能な値は『天, 童, 祝, 極, 特, 上, 序, 雅, 化』", _
                                     "武将名。", _
                                     "武将id。", _
                                     "武将のコスト。", _
                                     "武将の指揮兵数", _
                                     "武将の槍統率。設定可能な値は『SSS, SS, .S, A ,B, C, D, E, F』", _
                                     "武将の弓統率。設定可能な値は『SSS, SS, .S, A ,B, C, D, E, F』", _
                                     "武将の馬統率。設定可能な値は『SSS, SS, .S, A ,B, C, D, E, F』", _
                                     "武将の器統率。設定可能な値は『SSS, SS, .S, A, B, C, D, E, F』", _
                                     "武将の初期将攻値。", _
                                     "武将の初期将防値。", _
                                     "武将の初期兵法値。", _
                                     "武将の将攻成長値。", _
                                     "武将の将防成長値。", _
                                     "武将の兵法成長値。", _
                                     "武将の初期スキル名。(SNameのスキル名に従属)", _
                                     "武将の職。設定可能な値は『将, 剣, 姫, 忍, 童, 覇, 文』", _
                                     "未完成の項目がある場合は『U』、全項目が埋まっている場合は『F』"}
    Public BData_reds As Integer() = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 15, 16, 17}

    Public SName_items As String() = {"スキル名", "スキルR", "分類", "攻防", "対象", "付与効果", "候補A", "候補B", "候補C", "候補S1", "候補S2", "Uunf"}
    Public SName_dists As String() = {"スキル名。", _
                                     "スキルのレアリティ。選択可能な値は『F, E, D, C, B, A, S』", _
                                     "スキルの分類。選択可能な値は『槍, 弓, 馬, 砲, 器, 複数, 全, 将, 速, 童, 条件, 特殊, 不可』", _
                                     "スキルの攻防分類。選択可能な値は『攻, 防, 速, 破壊, 特殊』", _
                                     "スキルの適用対象となる兵科。", _
                                     "スキルの付与効果。破壊や速度が付与される場合はそれぞれ『破壊, 速』を指定。ただし、攻防分類が「破壊, 速」の場合は空欄。", _
                                     "合成テーブルにおけるA候補。", _
                                     "合成テーブルにおけるB候補。", _
                                     "合成テーブルにおけるC候補。", _
                                     "合成テーブルにおけるS1候補。", _
                                     "合成テーブルにおけるS2候補。", _
                                     "未完成の項目がある場合は『U』、全項目が埋まっている場合は『F』"}
    Public SName_reds As Integer() = {0, 2, 3, 4, 11}

    Public SData_items As String() = {"スキル名", "スキルLV", "発動率", "上昇率", "付与率", "Sunf"}
    Public SData_dists As String() = {"スキル名。 (SNameのスキル名に従属)", _
                                     "スキルレベル。", _
                                     "スキルの発動率。例：5%発動[0.05]", _
                                     "スキルの上昇率。コスト依存の場合は'C'を含む計算式。コスト依存入力例：[+(18+C*8)%]", _
                                     "付加効果を設定した場合の、付加効果の上昇率。", _
                                     "未完成の項目がある場合は『U』、全項目が埋まっている場合は『F』"}
    Public SData_reds As Integer() = {0, 1, 5}

    Public L1items As String()
    Public L1dists As String()
    Public L1reds As Integer()
    Public L1 As String()
    Public SelectedTable As String
    Public SelectedMode As String
    Public ViewTable As String

    Public viewunflg As Boolean = True

    Private Sub Form13_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ''接続文字列を設定
        'Connection.ConnectionString = "Version=3;Data Source=ixadb.db3;New=False;Compress=True;" 'Read Only=True;"
        ''オープン
        'Connection.Open()
        ' ''パスワードを変更
        ''Connection.ChangePassword("password")
        ''コマンド作成
        'Command = Connection.CreateCommand

        'Combobox初期値を設定
        ToolStripComboBox1.SelectedIndex = 1
        ToolStripComboBox2.SelectedIndex = 1
        ToolStripComboBox3.SelectedIndex = 1
        ToolStripComboBox4.SelectedIndex = 0
        Form1.Hide()
    End Sub

    ''ApplicationExitイベントハンドラ
    'Private Sub Application_ApplicationExit(ByVal sender As Object, ByVal e As EventArgs)
    '    'DataReader.Close()
    '    Command.Dispose()
    '    Connection.Close()
    '    Connection.Dispose()
    '    'ApplicationExitイベントハンドラを削除
    '    RemoveHandler Application.ApplicationExit, AddressOf Application_ApplicationExit
    'End Sub

    Private Function List_setItem(ByRef LBox As ListBox, ByVal Setitems As String()) As Integer
        LBox.Items.AddRange(Setitems)
        Return LBox.Items.Count
    End Function

    Private Function List_removeItem(ByRef LBox As ListBox, ByVal Deleteitem As String) As Integer
        LBox.Items.Remove(Deleteitem)
        Return LBox.Items.Count
    End Function

    Private Sub ListBox1_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ListBox1.DrawItem
        If e.Index = -1 Then
            Exit Sub
        End If
        e.DrawBackground()
        Dim myBrush As Brush
        If Array.IndexOf(L1reds, e.Index) > -1 Then
            myBrush = Brushes.Red
        Else
            myBrush = New SolidBrush(ListBox1.ForeColor)
        End If
        e.Graphics.DrawString(ListBox1.Items(e.Index), e.Font, myBrush, New RectangleF(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height))
        e.DrawFocusRectangle()
    End Sub

    Private Sub TableChange(sender As Object, e As EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        Select Case (ToolStripComboBox1.Text.ToString)
            Case "BData"
                SelectedTable = "BData"
                L1items = BData_items.Clone
                L1dists = BData_dists.Clone
                L1reds = BData_reds.Clone
            Case "SName"
                SelectedTable = "SName"
                L1items = SName_items.Clone
                L1dists = SName_dists.Clone
                L1reds = SName_reds.Clone
            Case "SData"
                SelectedTable = "SData"
                L1items = SData_items.Clone
                L1dists = SData_dists.Clone
                L1reds = SData_reds.Clone
        End Select
        ReDim L1(L1items.Length - 1)
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        List_setItem(ListBox1, L1items)
    End Sub

    Private Sub 値入力項目選択(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        If sender.SelectedIndex = -1 Then Exit Sub
        'Label update
        Label3.Text = "値 (項目: " & sender.SelectedItem & ")"
        If Array.IndexOf(L1reds, sender.SelectedIndex) > -1 Then
            Label4.Text = "[必須値]"
            Label4.ForeColor = Color.Red
        Else
            Label4.Text = "[任意値]"
            Label4.ForeColor = Color.Black
        End If
        RichTextBox1.Clear()
        '項目の説明を更新
        RichTextBox4.Clear()
        RichTextBox4.Text = L1dists(sender.SelectedIndex)
    End Sub

    Private Sub 値更新(sender As Object, e As EventArgs) Handles Button2.Click
        If ListBox1.SelectedIndex = -1 Then Exit Sub
        Dim index As Integer = ListBox1.SelectedIndex
        If Not L1(index) Is Nothing Then List_removeItem(ListBox2, L1items(index) & ": " & L1(index))
        L1(index) = RichTextBox1.Text 'update
        RichTextBox1.Clear()
        List_setItem(ListBox2, {L1items(index) & ": " & L1(index)})
    End Sub

    Private Sub 値削除(sender As Object, e As MouseEventArgs) Handles ListBox2.MouseDown
        If e.Button = System.Windows.Forms.MouseButtons.Right Then
            'Dim index As Integer = ListBox2.SelectedIndex
            Dim L2dindex As Integer = ListBox2.IndexFromPoint(e.X, e.Y)
            If L2dindex = -1 Then Exit Sub
            Dim Ldindex As Integer = -1
            Dim rmitem As String = Split(ListBox2.Items(L2dindex).ToString, ":")(0)
            For i As Integer = 0 To ListBox1.Items.Count - 1
                If ListBox1.Items(i).ToString = rmitem Then
                    Ldindex = i
                    Exit For
                End If
            Next
            ListBox2.Items.RemoveAt(L2dindex)
            L1(Ldindex) = Nothing 'delete
        End If
    End Sub

    Private Sub 条件式表示切替(ByVal onoff As Boolean)
        RichTextBox2.Visible = onoff
        Button3.Visible = onoff
        Button4.Visible = onoff
        Button5.Visible = onoff
    End Sub

    Private Sub 項目表示切替(ByVal onoff As Boolean)
        ListBox2.Visible = onoff
        Label2.Visible = onoff
        Button1.Visible = onoff
        Button2.Visible = onoff
    End Sub

    Private Sub 処理選択(sender As Object, e As EventArgs) Handles ToolStripComboBox2.SelectedIndexChanged
        Select Case (ToolStripComboBox2.Text.ToString)
            Case "Insert"
                SelectedMode = "Insert"
                条件式表示切替(False)
                項目表示切替(True)
                Label4.Visible = True
                Label5.Text = ""
            Case "Delete"
                SelectedMode = "Delete"
                条件式表示切替(True)
                項目表示切替(False)
                Label4.Visible = False
                Label5.Text = "削除条件式 (WHERE)"
            Case "Update"
                SelectedMode = "Update"
                条件式表示切替(True)
                項目表示切替(True)
                Label4.Visible = False
                Label5.Text = "更新条件式 (WHERE)"
        End Select
        Lクリア()
    End Sub

    Private Sub View表示処理(ByRef Dgv As DataGridView, ByVal sqlstr As String)
        Dgv.DataSource = ""
        'Dgv.Rows.Clear()
        Dim ds As DataSet = New DataSet()
        Dim da As SQLiteDataAdapter = New SQLiteDataAdapter(sqlstr, Connection)
        da.Fill(ds)
        Dgv.DataSource = ds.Tables(0).DefaultView
    End Sub

    Private Sub View表示(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        View表示処理(DataGridView1, viewSQL生成())
    End Sub

    Private Sub unf表示切替(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        If viewunflg Then 'unfのみ表示
            ToolStripButton3.Text = "全てのデータ"
            viewunflg = False
        Else 'すべて表示
            ToolStripButton3.Text = "未更新データ"
            viewunflg = True
        End If
    End Sub

    Private Sub ViewTable切替(sender As Object, e As EventArgs) Handles ToolStripComboBox3.SelectedIndexChanged
        ViewTable = ToolStripComboBox3.Text
        Dim setitems As String() = {}
        Select Case (ViewTable)
            Case "BData"
                setitems = BData_items.Clone
            Case "SName"
                setitems = SName_items.Clone
            Case "SData"
                setitems = SData_items.Clone
        End Select
        ToolStripComboBox4.Items.Clear()
        ToolStripComboBox4.Items.AddRange(setitems)
    End Sub

    Private Function viewSQL生成() As String
        Dim unfstr As String = ""
        If viewunflg Then
            Select Case (ViewTable)
                Case "BData"
                    unfstr = "WHERE Bunf = 'U'"
                Case "SName"
                    unfstr = "WHERE Uunf = 'U'"
                Case "SData"
                    unfstr = "WHERE Sunf = 'U'"
            End Select
        End If
        Return ("SELECT * FROM " & ViewTable & " " & unfstr)
    End Function

    Private Sub Listbox2_clear(sender As Object, e As EventArgs) Handles Button1.Click
        Lクリア()
    End Sub

    Private Sub Lクリア()
        ReDim L1(L1items.Length - 1)
        ListBox2.Items.Clear()
    End Sub

    Private Sub データベース更新実行(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Dim SqlString As String = ""
        Dim retcount As Integer = -1
        Select Case (SelectedMode)
            Case "Insert"
                SqlString = Insert文生成()
            Case "Delete"
                SqlString = Delete文生成()
            Case "Update"
                SqlString = Update文生成()
        End Select
        RichTextBox3.Clear()
        RichTextBox3.Text = SqlString
        Command.CommandText = SqlString
        retcount = Command.ExecuteNonQuery()
        RichTextBox3.Text = RichTextBox3.Text & vbCrLf & "( " & retcount & " data affected )"
    End Sub

    Private Function Insert文生成() As String
        Dim sstr As String = ""
        Dim vstr As String = ""
        For i As Integer = 0 To L1.Length - 1
            If L1(i) Is Nothing Then Continue For
            sstr = sstr & L1items(i) & ", "
            vstr = vstr & ダブルクオート(L1(i)) & ", "
        Next
        sstr = sstr.Substring(0, sstr.Length - 2)
        vstr = vstr.Substring(0, vstr.Length - 2)
        Return ("INSERT INTO " & SelectedTable & " (" & sstr & ") values (" & vstr & ")")
    End Function

    Private Function Delete文生成() As String
        Dim wstr As String = RichTextBox2.Text
        Return ("DELETE FROM " & SelectedTable & " WHERE " & wstr)
    End Function

    Private Function Update文生成() As String
        Dim sstr As String = ""
        Dim wstr As String = RichTextBox2.Text
        For i As Integer = 0 To L1.Length - 1
            If L1(i) Is Nothing Then Continue For
            sstr = sstr & L1items(i) & " = " & ダブルクオート(L1(i)) & ", "
        Next
        sstr = sstr.Substring(0, sstr.Length - 2)
        Return ("UPDATE " & SelectedTable & " SET " & sstr & " WHERE " & wstr)
    End Function

    Private Sub WHERE句更新(sender As Object, e As EventArgs) Handles Button3.Click
        If ListBox1.SelectedIndex = -1 Then Exit Sub
        Dim index As Integer = ListBox1.SelectedIndex
        RichTextBox2.Text = RichTextBox2.Text & L1items(index) & " = " & ダブルクオート(RichTextBox1.Text)
    End Sub

    Private Sub WHERE_AND(sender As Object, e As EventArgs) Handles Button4.Click
        RichTextBox2.Text = RichTextBox2.Text & " AND "
    End Sub

    Private Sub WHERE_OR(sender As Object, e As EventArgs)
        RichTextBox2.Text = RichTextBox2.Text & " OR "
    End Sub

    Private Sub WHEREクリア(sender As Object, e As EventArgs) Handles Button5.Click
        RichTextBox2.Clear()
    End Sub

    Private Sub 値取得(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        RichTextBox1.Text = DataGridView1.SelectedCells.Item(0).Value
        TabControl1.SelectedTab = TabPage1
    End Sub

    Private Sub view検索(sender As Object, e As MouseEventArgs) Handles ToolStripButton4.MouseDown
        Dim eqstr As String = ""
        If e.Button = System.Windows.Forms.MouseButtons.Right Then '右クリックなら曖昧検索
            eqstr = " Like "
        Else
            eqstr = " = "
        End If

        Dim keyValue As String = ToolStripTextBox1.Text
        Dim keyindex As String = ToolStripComboBox4.Text
        If ToolStripComboBox4.SelectedIndex = -1 Then Exit Sub

        Dim SqlString As String = _
            "SELECT * FROM " & ViewTable & " WHERE " & ダブルクオート(keyindex) & eqstr & ダブルクオート(keyValue)
        View表示処理(DataGridView1, SqlString)

        'Dim rList As IEnumerable(Of DataGridViewRow) 'DataGridViewRowCollection
        'Try
        '    rList = DataGridView1.Rows.Cast(Of DataGridViewRow)().Where(Function(row) row.Cells(keyindex).Value.Equals(keyValue))
        '    MsgBox(rList(1).Cells(1).Value)
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message)
        'End Try
    End Sub

    Private Sub 最適化(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        Command.CommandText = "vacuum"
        Command.ExecuteNonQuery()
        MsgBox("最適化完了")
    End Sub

    Private Sub Form13_Closing(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.FormClosing
        Form1.Show()
    End Sub
End Class