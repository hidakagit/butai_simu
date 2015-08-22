'Imports System.Security.Permissions

Public Class Form15

    '閉じるボタンをオーバーライドで消す
    'Protected Overrides ReadOnly Property CreateParams() As  _
    '    System.Windows.Forms.CreateParams
    '    <SecurityPermission(SecurityAction.Demand, _
    '        Flags:=SecurityPermissionFlag.UnmanagedCode)> _
    '    Get
    '        Const CS_NOCLOSE As Integer = &H200
    '        Dim cp As CreateParams = MyBase.CreateParams
    '        cp.ClassStyle = cp.ClassStyle Or CS_NOCLOSE

    '        Return cp
    '    End Get
    'End Property

    Private Sub Form15_load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            With Form10
                If Not .cus_status Is Nothing Then '空じゃなければ
                    For i As Integer = 0 To 3
                        Select Case .cus_status(i)
                            Case 0
                                ComboBox(Me, Val(i) & "1").Text = "攻撃極振り"
                            Case 1
                                ComboBox(Me, Val(i) & "1").Text = "防御極振り"
                            Case 2
                                ComboBox(Me, Val(i) & "1").Text = "兵法極振り"
                            Case 3
                                ComboBox(Me, Val(i) & "1").Text = "適正お任せ"
                        End Select
                    Next
                End If
                If Val(.cus_status_kb) > 0 Then TextBox1.Text = .cus_status_kb
                If Val(.cus_status_hei) > 0 Then TextBox2.Text = .cus_status_hei
            End With
        Catch ex As Exception
        End Try
    End Sub

    Public Sub ステ振り詳細設定関連ONOFF(ByVal TorF As Boolean)
        With Form10
            If TorF Then
                With .Label15
                    .ForeColor = Color.Blue
                    .Text = "OFF"
                End With
                .Button2.Enabled = False
            Else
                With .Label15
                    .ForeColor = Color.Red
                    .Text = "ON"
                End With
                .Button2.Enabled = True
            End If
        End With
    End Sub

    Private Sub 確定(sender As Object, e As EventArgs) Handles Button1.Click
        '入力チェック
        If ComboBox01.Text = "" Or ComboBox11.Text = "" Or ComboBox21.Text = "" Or ComboBox31.Text = "" Then
            MsgBox("ステ振り設定が完全に完了していません")
            Exit Sub
        Else
            For i As Integer = 0 To 3
                With Form10
                    Select Case ComboBox(Me, CStr(i) & "1").Text
                        Case "攻撃極振り"
                            .cus_status(i) = 0
                        Case "防御極振り"
                            .cus_status(i) = 1
                        Case "兵法極振り"
                            .cus_status(i) = 2
                        Case "適正お任せ"
                            .cus_status(i) = 3
                    End Select
                End With
            Next
        End If
        If Val(TextBox1.Text) <= 0 Then
            MsgBox("攻防成長値は0以上の数字を入力して下さい")
            Exit Sub
        Else
            Form10.cus_status_kb = Val(TextBox1.Text)
        End If
        If Val(TextBox2.Text) <= 0 Then
            MsgBox("兵法成長値は0以上の数字を入力して下さい")
            Exit Sub
        Else
            Form10.cus_status_hei = Val(TextBox2.Text)
        End If
        Me.Close()
    End Sub
End Class