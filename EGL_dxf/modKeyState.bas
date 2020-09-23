Attribute VB_Name = "modKeyboard"

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Sub UpdatePartPos()

    With ObjPart
        If frmCanvas.chkAnimation.Value = 1 Then
            Call UpdateAnim
        Else
            Call UpdateKey
        End If
        .Direction.X = .Direction.X Mod 360
        .Direction.Y = .Direction.Y Mod 360
        .Direction.Z = .Direction.Z Mod 360
        If .Scale < 0.001 Then .Scale = 0.001
        If .Scale > 1000 Then .Scale = 1000
    End With

End Sub

'---------------------------------------
'Sub: UpdateKey
'Basýlan tuþa göre parçanýn konumu belirleniyor.
'---------------------------------------

Private Sub UpdateKey()
 
    Dim R As Single, S As Single, T As Single
    
    With ObjPart
        R = 5
        S = .Scale * 0.02
        T = 0.1 '5 / .Scale
        If State(vbKeyControl) Then
            If State(vbKeyRight) Then .Position.X = .Position.X + T
            If State(vbKeyLeft) Then .Position.X = .Position.X - T
            If State(vbKeyUp) Then .Position.Y = .Position.Y + T
            If State(vbKeyDown) Then .Position.Y = .Position.Y - T
            If State(vbKeyPageDown) Then .Position.Z = .Position.Z + T
            If State(vbKeyPageUp) Then .Position.Z = .Position.Z - T
        ElseIf State(vbKeyShift) Then
            If State(vbKeyPageDown) Then .Scale = .Scale - S
            If State(vbKeyPageUp) Then .Scale = .Scale + S
        Else
            If State(vbKey1) Then RenderOption (0)
            If State(vbKey2) Then RenderOption (1)
            If State(vbKey3) Then RenderOption (2)
            If State(vbKey4) Then RenderOption (3)
            If State(vbKey5) Then RenderOption (4)
            If State(vbKey6) Then RenderOption (5)
            If State(vbKey7) Then RenderOption (6)
            If State(vbKeyDown) Then .Direction.X = .Direction.X + R
            If State(vbKeyUp) Then .Direction.X = .Direction.X - R
            If State(vbKeyRight) Then .Direction.Y = .Direction.Y + R
            If State(vbKeyLeft) Then .Direction.Y = .Direction.Y - R
            If State(vbKeyPageUp) Then .Direction.Z = .Direction.Z + R
            If State(vbKeyPageDown) Then .Direction.Z = .Direction.Z - R
            If State(vbKeyC) Then .Scale = .Scale - S
            If State(vbKeyZ) Then .Scale = .Scale + S
            If State(vbKeyX) Then Call ResetPos
            If State(vbKeyEscape) Then Unload frmCanvas:    End
        End If
    End With

End Sub

Private Sub UpdateAnim()

    Dim i As Integer
    Dim R(2) As Single, T(2) As Single
    
    For i = 0 To 2
        R(i) = VerifyText(frmCanvas.txtRot(i))
        T(i) = VerifyText(frmCanvas.txtTrans(i))
    Next
    
    With ObjPart
        If T(0) <> 0 Then .Position.X = .Position.X + T(0)
        If T(1) <> 0 Then .Position.Y = .Position.Y + T(1)
        If T(2) <> 0 Then .Position.Z = .Position.Z + T(2)
        If R(0) <> 0 Then .Direction.X = .Direction.X + R(0)
        If R(1) <> 0 Then .Direction.Y = .Direction.Y + R(1)
        If R(2) <> 0 Then .Direction.Z = .Direction.Z + R(2)
    End With
   
End Sub

Private Function State(key As Long) As Boolean
 
    Dim lngKeyState As Long
    
    lngKeyState = GetKeyState(key)
    State = IIf((lngKeyState And &H8000), True, False)

End Function

Private Function VerifyText(txt As TextBox) As Single
  
    If IsNumeric(txt.Text) And txt.Text <> 0 Then
        VerifyText = CSng(txt.Text)
    Else
        VerifyText = 0
    End If

End Function

Private Sub ResetPos()

    With ObjPart
        .Direction.X = 0: .Direction.Y = 0: .Direction.Z = 0
        .Position.X = 0: .Position.Y = 0: .Position.Z = 0
        .Scale = 1
    End With
    
End Sub
