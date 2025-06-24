' 模块级变量（放在最上方）
Dim PasswordVerified As Boolean

Private Sub Worksheet_Activate()
    PasswordVerified = False ' 每次激活都重置状态

    Dim pwd As Variant
    pwd = Application.InputBox("请输入密码:", "密码保护", Type:=2)

    If pwd = "6688" Then
        PasswordVerified = True
        Cells.EntireRow.Hidden = False
    Else
        MsgBox "密码错误，您无权查看"
        Application.EnableEvents = False
        ThisWorkbook.Sheets(1).Select
        Application.EnableEvents = True
    End If
End Sub

Private Sub Worksheet_Deactivate()
    If PasswordVerified Then
        Cells.EntireRow.Hidden = True
        PasswordVerified = False
    End If
End Sub
