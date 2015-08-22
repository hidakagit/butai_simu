<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form15
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
        Me.ComboBox01 = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.ComboBox31 = New System.Windows.Forms.ComboBox()
        Me.ComboBox21 = New System.Windows.Forms.ComboBox()
        Me.ComboBox11 = New System.Windows.Forms.ComboBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(180, 12)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "各武将に対するステ振りを細かく設定"
        '
        'ComboBox01
        '
        Me.ComboBox01.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.ComboBox01.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.ComboBox01.FormattingEnabled = True
        Me.ComboBox01.Items.AddRange(New Object() {"攻撃極振り", "防御極振り", "兵法極振り"})
        Me.ComboBox01.Location = New System.Drawing.Point(108, 29)
        Me.ComboBox01.Name = "ComboBox01"
        Me.ComboBox01.Size = New System.Drawing.Size(85, 20)
        Me.ComboBox01.TabIndex = 29
        Me.ComboBox01.Text = "兵法極振り"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(59, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(41, 12)
        Me.Label2.TabIndex = 35
        Me.Label2.Text = "部隊長"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(51, 58)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(49, 12)
        Me.Label3.TabIndex = 36
        Me.Label3.Text = "小隊長A"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(51, 84)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(49, 12)
        Me.Label4.TabIndex = 37
        Me.Label4.Text = "小隊長B"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(26, 110)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(74, 12)
        Me.Label5.TabIndex = 38
        Me.Label5.Text = "ランキング武将"
        '
        'ComboBox31
        '
        Me.ComboBox31.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.ComboBox31.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.ComboBox31.FormattingEnabled = True
        Me.ComboBox31.Items.AddRange(New Object() {"攻撃極振り", "防御極振り", "兵法極振り", "適正お任せ"})
        Me.ComboBox31.Location = New System.Drawing.Point(108, 107)
        Me.ComboBox31.Name = "ComboBox31"
        Me.ComboBox31.Size = New System.Drawing.Size(85, 20)
        Me.ComboBox31.TabIndex = 39
        Me.ComboBox31.Text = "適正お任せ"
        '
        'ComboBox21
        '
        Me.ComboBox21.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.ComboBox21.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.ComboBox21.FormattingEnabled = True
        Me.ComboBox21.Items.AddRange(New Object() {"攻撃極振り", "防御極振り", "兵法極振り"})
        Me.ComboBox21.Location = New System.Drawing.Point(108, 81)
        Me.ComboBox21.Name = "ComboBox21"
        Me.ComboBox21.Size = New System.Drawing.Size(85, 20)
        Me.ComboBox21.TabIndex = 40
        Me.ComboBox21.Text = "兵法極振り"
        '
        'ComboBox11
        '
        Me.ComboBox11.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.ComboBox11.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.ComboBox11.FormattingEnabled = True
        Me.ComboBox11.Items.AddRange(New Object() {"攻撃極振り", "防御極振り", "兵法極振り"})
        Me.ComboBox11.Location = New System.Drawing.Point(108, 55)
        Me.ComboBox11.Name = "ComboBox11"
        Me.ComboBox11.Size = New System.Drawing.Size(85, 20)
        Me.ComboBox11.TabIndex = 41
        Me.ComboBox11.Text = "兵法極振り"
        '
        'Button1
        '
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Location = New System.Drawing.Point(147, 232)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 42
        Me.Button1.Text = "確定"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.ForeColor = System.Drawing.Color.Green
        Me.Label6.Location = New System.Drawing.Point(14, 139)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(198, 12)
        Me.Label6.TabIndex = 43
        Me.Label6.Text = "▼ランキング武将で「適正お任せ」の場合"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(14, 162)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(65, 12)
        Me.Label7.TabIndex = 44
        Me.Label7.Text = "1. 攻防成長"
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Consolas", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(85, 157)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(35, 22)
        Me.TextBox1.TabIndex = 45
        Me.TextBox1.Text = "40"
        Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(126, 162)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(95, 12)
        Me.Label8.TabIndex = 46
        Me.Label8.Text = "以上なら攻防極振"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(126, 187)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(95, 12)
        Me.Label9.TabIndex = 49
        Me.Label9.Text = "以上なら兵法極振"
        '
        'TextBox2
        '
        Me.TextBox2.Font = New System.Drawing.Font("Consolas", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox2.Location = New System.Drawing.Point(85, 182)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(35, 22)
        Me.TextBox2.TabIndex = 48
        Me.TextBox2.Text = "2.5"
        Me.TextBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(14, 187)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(65, 12)
        Me.Label10.TabIndex = 47
        Me.Label10.Text = "2. 兵法成長"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(14, 209)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(151, 12)
        Me.Label11.TabIndex = 50
        Me.Label11.Text = "3. 上の2条件以外は攻防極振"
        '
        'Form15
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(234, 267)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.ComboBox11)
        Me.Controls.Add(Me.ComboBox21)
        Me.Controls.Add(Me.ComboBox31)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.ComboBox01)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Form15"
        Me.Text = "ステ振り詳細設定"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ComboBox01 As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ComboBox31 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox21 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox11 As System.Windows.Forms.ComboBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
End Class
