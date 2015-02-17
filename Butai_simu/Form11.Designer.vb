<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form11
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.ComboBox012 = New System.Windows.Forms.ComboBox()
        Me.ComboBox011 = New System.Windows.Forms.ComboBox()
        Me.ComboBox010 = New System.Windows.Forms.ComboBox()
        Me.ComboBox009 = New System.Windows.Forms.ComboBox()
        Me.Label001 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.ComboBox112 = New System.Windows.Forms.ComboBox()
        Me.ComboBox111 = New System.Windows.Forms.ComboBox()
        Me.ComboBox110 = New System.Windows.Forms.ComboBox()
        Me.ComboBox109 = New System.Windows.Forms.ComboBox()
        Me.Label002 = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.ComboBox212 = New System.Windows.Forms.ComboBox()
        Me.ComboBox211 = New System.Windows.Forms.ComboBox()
        Me.ComboBox210 = New System.Windows.Forms.ComboBox()
        Me.ComboBox209 = New System.Windows.Forms.ComboBox()
        Me.Label003 = New System.Windows.Forms.Label()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.ComboBox312 = New System.Windows.Forms.ComboBox()
        Me.ComboBox311 = New System.Windows.Forms.ComboBox()
        Me.ComboBox310 = New System.Windows.Forms.ComboBox()
        Me.ComboBox309 = New System.Windows.Forms.ComboBox()
        Me.Label004 = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(195, 12)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "各武将に対する追加スキルを細かく設定"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 93)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(29, 12)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "スロ3"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 71)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(29, 12)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "スロ2"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(12, 49)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(29, 12)
        Me.Label6.TabIndex = 17
        Me.Label6.Text = "初期"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ComboBox012)
        Me.GroupBox1.Controls.Add(Me.ComboBox011)
        Me.GroupBox1.Controls.Add(Me.ComboBox010)
        Me.GroupBox1.Controls.Add(Me.ComboBox009)
        Me.GroupBox1.Controls.Add(Me.Label001)
        Me.GroupBox1.Location = New System.Drawing.Point(47, 29)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(140, 89)
        Me.GroupBox1.TabIndex = 28
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "部隊長"
        '
        'ComboBox012
        '
        Me.ComboBox012.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.ComboBox012.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.ComboBox012.FormattingEnabled = True
        Me.ComboBox012.Location = New System.Drawing.Point(49, 61)
        Me.ComboBox012.Name = "ComboBox012"
        Me.ComboBox012.Size = New System.Drawing.Size(85, 20)
        Me.ComboBox012.TabIndex = 28
        '
        'ComboBox011
        '
        Me.ComboBox011.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.ComboBox011.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.ComboBox011.FormattingEnabled = True
        Me.ComboBox011.Location = New System.Drawing.Point(49, 39)
        Me.ComboBox011.Name = "ComboBox011"
        Me.ComboBox011.Size = New System.Drawing.Size(85, 20)
        Me.ComboBox011.TabIndex = 29
        '
        'ComboBox010
        '
        Me.ComboBox010.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox010.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.ComboBox010.FormattingEnabled = True
        Me.ComboBox010.Items.AddRange(New Object() {"槍", "弓", "馬", "砲", "器", "複数", "全", "将", "速", "特殊"})
        Me.ComboBox010.Location = New System.Drawing.Point(7, 61)
        Me.ComboBox010.Name = "ComboBox010"
        Me.ComboBox010.Size = New System.Drawing.Size(36, 20)
        Me.ComboBox010.TabIndex = 30
        '
        'ComboBox009
        '
        Me.ComboBox009.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox009.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.ComboBox009.FormattingEnabled = True
        Me.ComboBox009.Items.AddRange(New Object() {"槍", "弓", "馬", "砲", "器", "複数", "全", "将", "速", "特殊"})
        Me.ComboBox009.Location = New System.Drawing.Point(7, 39)
        Me.ComboBox009.Name = "ComboBox009"
        Me.ComboBox009.Size = New System.Drawing.Size(36, 20)
        Me.ComboBox009.TabIndex = 31
        '
        'Label001
        '
        Me.Label001.AutoSize = True
        Me.Label001.Location = New System.Drawing.Point(53, 20)
        Me.Label001.Name = "Label001"
        Me.Label001.Size = New System.Drawing.Size(77, 12)
        Me.Label001.TabIndex = 32
        Me.Label001.Text = "------------"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.ComboBox112)
        Me.GroupBox2.Controls.Add(Me.ComboBox111)
        Me.GroupBox2.Controls.Add(Me.ComboBox110)
        Me.GroupBox2.Controls.Add(Me.ComboBox109)
        Me.GroupBox2.Controls.Add(Me.Label002)
        Me.GroupBox2.Location = New System.Drawing.Point(193, 29)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(140, 89)
        Me.GroupBox2.TabIndex = 29
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "小隊長A"
        '
        'ComboBox112
        '
        Me.ComboBox112.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.ComboBox112.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.ComboBox112.FormattingEnabled = True
        Me.ComboBox112.Location = New System.Drawing.Point(49, 61)
        Me.ComboBox112.Name = "ComboBox112"
        Me.ComboBox112.Size = New System.Drawing.Size(85, 20)
        Me.ComboBox112.TabIndex = 28
        '
        'ComboBox111
        '
        Me.ComboBox111.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.ComboBox111.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.ComboBox111.FormattingEnabled = True
        Me.ComboBox111.Location = New System.Drawing.Point(49, 39)
        Me.ComboBox111.Name = "ComboBox111"
        Me.ComboBox111.Size = New System.Drawing.Size(85, 20)
        Me.ComboBox111.TabIndex = 29
        '
        'ComboBox110
        '
        Me.ComboBox110.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox110.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.ComboBox110.FormattingEnabled = True
        Me.ComboBox110.Items.AddRange(New Object() {"槍", "弓", "馬", "砲", "器", "複数", "全", "将", "速", "特殊"})
        Me.ComboBox110.Location = New System.Drawing.Point(7, 61)
        Me.ComboBox110.Name = "ComboBox110"
        Me.ComboBox110.Size = New System.Drawing.Size(36, 20)
        Me.ComboBox110.TabIndex = 30
        '
        'ComboBox109
        '
        Me.ComboBox109.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox109.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.ComboBox109.FormattingEnabled = True
        Me.ComboBox109.Items.AddRange(New Object() {"槍", "弓", "馬", "砲", "器", "複数", "全", "将", "速", "特殊"})
        Me.ComboBox109.Location = New System.Drawing.Point(7, 39)
        Me.ComboBox109.Name = "ComboBox109"
        Me.ComboBox109.Size = New System.Drawing.Size(36, 20)
        Me.ComboBox109.TabIndex = 31
        '
        'Label002
        '
        Me.Label002.AutoSize = True
        Me.Label002.Location = New System.Drawing.Point(53, 20)
        Me.Label002.Name = "Label002"
        Me.Label002.Size = New System.Drawing.Size(77, 12)
        Me.Label002.TabIndex = 32
        Me.Label002.Text = "------------"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.ComboBox212)
        Me.GroupBox3.Controls.Add(Me.ComboBox211)
        Me.GroupBox3.Controls.Add(Me.ComboBox210)
        Me.GroupBox3.Controls.Add(Me.ComboBox209)
        Me.GroupBox3.Controls.Add(Me.Label003)
        Me.GroupBox3.Location = New System.Drawing.Point(339, 29)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(140, 89)
        Me.GroupBox3.TabIndex = 30
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "小隊長B"
        '
        'ComboBox212
        '
        Me.ComboBox212.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.ComboBox212.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.ComboBox212.FormattingEnabled = True
        Me.ComboBox212.Location = New System.Drawing.Point(49, 61)
        Me.ComboBox212.Name = "ComboBox212"
        Me.ComboBox212.Size = New System.Drawing.Size(85, 20)
        Me.ComboBox212.TabIndex = 28
        '
        'ComboBox211
        '
        Me.ComboBox211.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.ComboBox211.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.ComboBox211.FormattingEnabled = True
        Me.ComboBox211.Location = New System.Drawing.Point(49, 39)
        Me.ComboBox211.Name = "ComboBox211"
        Me.ComboBox211.Size = New System.Drawing.Size(85, 20)
        Me.ComboBox211.TabIndex = 29
        '
        'ComboBox210
        '
        Me.ComboBox210.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox210.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.ComboBox210.FormattingEnabled = True
        Me.ComboBox210.Items.AddRange(New Object() {"槍", "弓", "馬", "砲", "器", "複数", "全", "将", "速", "特殊"})
        Me.ComboBox210.Location = New System.Drawing.Point(7, 61)
        Me.ComboBox210.Name = "ComboBox210"
        Me.ComboBox210.Size = New System.Drawing.Size(36, 20)
        Me.ComboBox210.TabIndex = 30
        '
        'ComboBox209
        '
        Me.ComboBox209.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox209.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.ComboBox209.FormattingEnabled = True
        Me.ComboBox209.Items.AddRange(New Object() {"槍", "弓", "馬", "砲", "器", "複数", "全", "将", "速", "特殊"})
        Me.ComboBox209.Location = New System.Drawing.Point(7, 39)
        Me.ComboBox209.Name = "ComboBox209"
        Me.ComboBox209.Size = New System.Drawing.Size(36, 20)
        Me.ComboBox209.TabIndex = 31
        '
        'Label003
        '
        Me.Label003.AutoSize = True
        Me.Label003.Location = New System.Drawing.Point(53, 20)
        Me.Label003.Name = "Label003"
        Me.Label003.Size = New System.Drawing.Size(77, 12)
        Me.Label003.TabIndex = 32
        Me.Label003.Text = "------------"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.CheckBox1.Location = New System.Drawing.Point(12, 131)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(141, 16)
        Me.CheckBox1.TabIndex = 31
        Me.CheckBox1.Text = "追加スキル詳細設定ON"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Location = New System.Drawing.Point(550, 124)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 32
        Me.Button1.Text = "確定"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.ComboBox312)
        Me.GroupBox4.Controls.Add(Me.ComboBox311)
        Me.GroupBox4.Controls.Add(Me.ComboBox310)
        Me.GroupBox4.Controls.Add(Me.ComboBox309)
        Me.GroupBox4.Controls.Add(Me.Label004)
        Me.GroupBox4.Location = New System.Drawing.Point(485, 29)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(140, 89)
        Me.GroupBox4.TabIndex = 33
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "ランキング武将"
        '
        'ComboBox312
        '
        Me.ComboBox312.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.ComboBox312.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.ComboBox312.FormattingEnabled = True
        Me.ComboBox312.Location = New System.Drawing.Point(49, 61)
        Me.ComboBox312.Name = "ComboBox312"
        Me.ComboBox312.Size = New System.Drawing.Size(85, 20)
        Me.ComboBox312.TabIndex = 28
        '
        'ComboBox311
        '
        Me.ComboBox311.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.ComboBox311.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.ComboBox311.FormattingEnabled = True
        Me.ComboBox311.Location = New System.Drawing.Point(49, 39)
        Me.ComboBox311.Name = "ComboBox311"
        Me.ComboBox311.Size = New System.Drawing.Size(85, 20)
        Me.ComboBox311.TabIndex = 29
        '
        'ComboBox310
        '
        Me.ComboBox310.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox310.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.ComboBox310.FormattingEnabled = True
        Me.ComboBox310.Items.AddRange(New Object() {"槍", "弓", "馬", "砲", "器", "複数", "全", "将", "速", "特殊"})
        Me.ComboBox310.Location = New System.Drawing.Point(7, 61)
        Me.ComboBox310.Name = "ComboBox310"
        Me.ComboBox310.Size = New System.Drawing.Size(36, 20)
        Me.ComboBox310.TabIndex = 30
        '
        'ComboBox309
        '
        Me.ComboBox309.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox309.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.ComboBox309.FormattingEnabled = True
        Me.ComboBox309.Items.AddRange(New Object() {"槍", "弓", "馬", "砲", "器", "複数", "全", "将", "速", "特殊"})
        Me.ComboBox309.Location = New System.Drawing.Point(7, 39)
        Me.ComboBox309.Name = "ComboBox309"
        Me.ComboBox309.Size = New System.Drawing.Size(36, 20)
        Me.ComboBox309.TabIndex = 31
        '
        'Label004
        '
        Me.Label004.AutoSize = True
        Me.Label004.Location = New System.Drawing.Point(53, 20)
        Me.Label004.Name = "Label004"
        Me.Label004.Size = New System.Drawing.Size(77, 12)
        Me.Label004.TabIndex = 32
        Me.Label004.Text = "------------"
        '
        'Button2
        '
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Location = New System.Drawing.Point(469, 124)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 34
        Me.Button2.Text = "消去"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Form11
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(634, 152)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Form11"
        Me.Text = "追加スキル詳細設定"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ComboBox012 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox011 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox010 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox009 As System.Windows.Forms.ComboBox
    Friend WithEvents Label001 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents ComboBox112 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox111 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox110 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox109 As System.Windows.Forms.ComboBox
    Friend WithEvents Label002 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents ComboBox212 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox211 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox210 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox209 As System.Windows.Forms.ComboBox
    Friend WithEvents Label003 As System.Windows.Forms.Label
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents ComboBox312 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox311 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox310 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox309 As System.Windows.Forms.ComboBox
    Friend WithEvents Label004 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
End Class
