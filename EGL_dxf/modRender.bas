Attribute VB_Name = "modRender"
Option Explicit

Public Enum RenderType      'Render tipi
    Dot                     'Nokta
    Dothidden               'Görünen yüzeylerin noktalarý
    Wireframe               'Tel kafes
    Hidden                  'Görünen yüzeyler kullanýlýyor
    SolidFrame              'Katý model ile wireframe birlikte
    Solid                   'Katý model
    Smooth                  'Daha gerçekçi
End Enum

Public Enum BackGroundType
    Black
    SolidColor
    Gradient
End Enum


'CreatePen(nPenStyle)
Private Const PS_SOLID = 0                   '  _______
'Private Const PS_DASH = 1                    '  -------
'Private Const PS_DOT = 2                     '  .......
'Private Const PS_DASHDOT = 3                 '  _._._._
'Private Const PS_DASHDOTDOT = 4              '  _.._.._
'Private Const PS_NULL = 5
'Private Const PS_INSIDEFRAME = 6

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type RenderOption
    RType As RenderType
    tColor As ColorRGB
    Luminance As Integer
    Hidden As Boolean
    Shade As Boolean
    LightOrbit As Boolean
    Show As Boolean
    ShowIndex As Integer
End Type

Type TRIVERTEX
    X As Long
    Y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type

Private Type GRADIENT_TRIANGLE
    Vertex1 As Long
    Vertex2 As Long
    Vertex3 As Long
End Type

Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Const GRADIENT_FILL_TRIANGLE As Long = &H2
Const GRADIENT_FILL_RECT_H As Long = &H0
Const GRADIENT_FILL_RECT_V  As Long = &H1

Declare Function GradientFillTriangle Lib "msimg32" Alias "GradientFill" _
    (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, _
    pMesh As GRADIENT_TRIANGLE, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long

Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" _
    (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, _
    pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public BackBuffer As Long
Public BackBitmap As Long

Public BackgroundColor(2) As Long
Public oldBackBitmap As Long

Public RO As RenderOption
Public HalfWidth As Long
Public HalfHeight As Long
Public triVert() As TRIVERTEX
Public Rec As RECT
Public BackType As BackGroundType

Public Sub InitializeDC(Canvas As PictureBox)

    With Canvas
        .ScaleMode = 3
        .AutoRedraw = False
        .Visible = True
        .FillStyle = vbSolid
        .DrawMode = vbCopyPen
        .DrawStyle = vbSolid
        BackBuffer = CreateCompatibleDC(.hDC)
        BackBitmap = CreateCompatibleBitmap(.hDC, .ScaleWidth, .ScaleHeight)
        oldBackBitmap = SelectObject(BackBuffer, BackBitmap)
    End With

End Sub

Public Sub TerminateDC()
    
    DeleteDC BackBuffer
    DeleteObject BackBitmap

End Sub

Public Sub RenderOption(Index As Integer)

    Dim i As Integer
    
    RO.RType = Index
    RO.Hidden = IIf(RO.RType = Dot Or RO.RType = Wireframe, False, True)
    For i = 0 To frmCanvas.mnuRender.Count - 1
        frmCanvas.mnuRender(i).Checked = IIf(i = Index, True, False)
    Next
    
End Sub

Public Sub Render(Canvas As PictureBox)
 
    Dim i As Integer, iV As Integer
    Dim fromcol As ColorRGB
    Dim tocol As ColorRGB
    
    DoEvents
    Select Case BackType
        Case Black
            BitBlt BackBuffer, 0, 0, Canvas.ScaleWidth, Canvas.ScaleHeight, BackBuffer, 0, 0, vbBlackness
        Case SolidColor
            fromcol = ColorLongToRGB(BackgroundColor(0))
            tocol = ColorLongToRGB(BackgroundColor(0))
            Call GradientRectangle(fromcol, tocol)
        Case Gradient
            fromcol = ColorLongToRGB(BackgroundColor(1))
            tocol = ColorLongToRGB(BackgroundColor(2))
            Call GradientRectangle(fromcol, tocol)
    End Select
    
    With ObjPart
        If RO.Hidden Then
            If ShortVisibleFaces > -1 Then
                If RO.RType = Smooth Then
                    ReDim triVert(UBound(.FaceV) * 3)
                Else
                    Erase triVert
                End If
                For i = 0 To UBound(.FaceV)
                    iV = .FaceV(i).iVisible
                    Call Draw(iV)
                Next
            End If
        Else
            For i = 0 To .NumFaces
                Call Draw(i)
            Next i
        End If
        If RO.Show Then
            Call DrawPoint(BackBuffer, .ScreenCoord(.Faces(RO.ShowIndex).A).X, .ScreenCoord(.Faces(RO.ShowIndex).A).Y, 3, vbRed)
            Call DrawPoint(BackBuffer, .ScreenCoord(.Faces(RO.ShowIndex).B).X, .ScreenCoord(.Faces(RO.ShowIndex).B).Y, 3, vbGreen)
            Call DrawPoint(BackBuffer, .ScreenCoord(.Faces(RO.ShowIndex).C).X, .ScreenCoord(.Faces(RO.ShowIndex).C).Y, 3, vbBlue)
        End If
    End With
    
    BitBlt Canvas.hDC, 0, 0, Canvas.ScaleWidth, Canvas.ScaleHeight, BackBuffer, 0, 0, vbSrcCopy
    
End Sub

Private Sub Draw(i As Integer)
    
    Dim idx As Integer
    Dim tmp(2) As POINTAPI
    Dim L As Single
    Dim PartColor As ColorRGB
    Dim lngColor As Long
    
    
    Dim BrushSelect As Long
    Dim PenSelect As Long
    
    With ObjPart
        tmp(0) = .ScreenCoord(.Faces(i).A)
        tmp(1) = .ScreenCoord(.Faces(i).B)
        tmp(2) = .ScreenCoord(.Faces(i).C)
    End With
    
    If RO.Shade Then
        L = VectorDot(ObjPart.NormalT(i), LightT)
        If L < 0 Then L = 0
        PartColor = ColorScale(RO.tColor, L)
    Else
        PartColor = RO.tColor
    End If
    
    If RO.Show And i = RO.ShowIndex Then PartColor = ColorInvert(PartColor)
    lngColor = ColorRGBToLong(ColorPlus(PartColor, RO.Luminance))
        
    Select Case RO.RType
        Case Wireframe
            PenSelect = SelectObject(BackBuffer, CreatePen(PS_SOLID, 1, lngColor))
            DrawTriangle BackBuffer, tmp
        Case Hidden
            PenSelect = SelectObject(BackBuffer, CreatePen(PS_SOLID, 1, lngColor))
            BrushSelect = SelectObject(BackBuffer, CreateSolidBrush(BackgroundColor(0)))
            Polygon BackBuffer, tmp(0), 3
        Case SolidFrame
            PenSelect = SelectObject(BackBuffer, CreatePen(PS_SOLID, 1, 0)) 'BackgroundColor(0)))
            BrushSelect = SelectObject(BackBuffer, CreateSolidBrush(lngColor))
            Polygon BackBuffer, tmp(0), 3
        Case Solid
            PenSelect = SelectObject(BackBuffer, CreatePen(PS_SOLID, 1, lngColor))
            BrushSelect = SelectObject(BackBuffer, CreateSolidBrush(lngColor))
            Polygon BackBuffer, tmp(0), 3
        Case Smooth
            Dim vert(2) As TRIVERTEX
            Dim gTRi As GRADIENT_TRIANGLE
            
            With ObjPart
                If triVert(.Faces(i).A).Alpha = 0 Then
                    triVert(.Faces(i).A).X = .ScreenCoord(.Faces(i).A).X
                    triVert(.Faces(i).A).Y = .ScreenCoord(.Faces(i).A).Y
                    triVert(.Faces(i).A).Red = PartColor.R
                    triVert(.Faces(i).A).Green = PartColor.G
                    triVert(.Faces(i).A).Blue = PartColor.B
                    triVert(.Faces(i).A).Alpha = 1
                End If
                If triVert(.Faces(i).B).Alpha = 0 Then
                    triVert(.Faces(i).B).X = .ScreenCoord(.Faces(i).B).X
                    triVert(.Faces(i).B).Y = .ScreenCoord(.Faces(i).B).Y
                    triVert(.Faces(i).B).Red = PartColor.R
                    triVert(.Faces(i).B).Green = PartColor.G
                    triVert(.Faces(i).B).Blue = PartColor.B
                    triVert(.Faces(i).B).Alpha = 1
                End If
                If triVert(.Faces(i).C).Alpha = 0 Then
                    triVert(.Faces(i).C).X = .ScreenCoord(.Faces(i).C).X
                    triVert(.Faces(i).C).Y = .ScreenCoord(.Faces(i).C).Y
                    triVert(.Faces(i).C).Red = PartColor.R
                    triVert(.Faces(i).C).Green = PartColor.G
                    triVert(.Faces(i).C).Blue = PartColor.B
                    triVert(.Faces(i).C).Alpha = 1
                End If
                vert(0).X = triVert(.Faces(i).A).X
                vert(0).Y = triVert(.Faces(i).A).Y
                vert(0).Red = Val("&h" & Hex(triVert(.Faces(i).A).Red) & "00")
                vert(0).Green = Val("&h" & Hex(triVert(.Faces(i).A).Green) & "00")
                vert(0).Blue = Val("&h" & Hex(triVert(.Faces(i).A).Blue) & "00")
                
                vert(1).X = triVert(.Faces(i).B).X
                vert(1).Y = triVert(.Faces(i).B).Y
                vert(1).Red = Val("&h" & Hex(triVert(.Faces(i).B).Red) & "00")
                vert(1).Green = Val("&h" & Hex(triVert(.Faces(i).B).Green) & "00")
                vert(1).Blue = Val("&h" & Hex(triVert(.Faces(i).B).Blue) & "00")
        
                vert(2).X = triVert(.Faces(i).C).X
                vert(2).Y = triVert(.Faces(i).C).Y
                vert(2).Red = Val("&h" & Hex(triVert(.Faces(i).C).Red) & "00")
                vert(2).Green = Val("&h" & Hex(triVert(.Faces(i).C).Green) & "00")
                vert(2).Blue = Val("&h" & Hex(triVert(.Faces(i).C).Blue) & "00")
            End With
            gTRi.Vertex1 = 0
            gTRi.Vertex2 = 1
            gTRi.Vertex3 = 2
            Call GradientFillTriangle(BackBuffer, vert(0), 3, gTRi, 1, GRADIENT_FILL_TRIANGLE)
        Case Dot, Dothidden
            For idx = 0 To 2
                Call DrawPoint(BackBuffer, tmp(idx).X, tmp(idx).Y, frmCanvas.scrDot.Value, lngColor)
            Next
    End Select

    DeleteObject PenSelect
    DeleteObject BrushSelect
    
End Sub

Private Sub DrawTriangle(hDC As Long, tmp() As POINTAPI)

    Dim p As POINTAPI
    
    MoveToEx hDC, tmp(0).X, tmp(0).Y, p
    LineTo hDC, tmp(1).X, tmp(1).Y
    LineTo hDC, tmp(2).X, tmp(2).Y
    LineTo hDC, tmp(0).X, tmp(0).Y

End Sub

Private Sub DrawPoint(hDC As Long, ByVal X As Single, ByVal Y As Single, R As Integer, color As Long)
    
    Dim BrushSelect As Long
    Dim PenSelect As Long
    
    PenSelect = SelectObject(BackBuffer, CreatePen(PS_SOLID, 1, color))
    BrushSelect = SelectObject(BackBuffer, CreateSolidBrush(color))
    Ellipse hDC, X - R, Y - R, X + R, Y + R
    DeleteObject PenSelect
    DeleteObject BrushSelect

End Sub

'Public Sub ColorRectangle(pic As PictureBox)
'
'    Dim BrushSelect As Long
'
'    BrushSelect = SelectObject(pic.hDC, CreateSolidBrush(BackgroundColor(0)))
'    FillRect BackBuffer, Rec, BrushSelect
'    DeleteObject BrushSelect
'
'End Sub

Public Sub GradientRectangle(FromColor As ColorRGB, ToColor As ColorRGB)

    Dim vert(1) As TRIVERTEX
    Dim gRect As GRADIENT_RECT
    
    With vert(0)
        .X = Rec.Left
        .Y = Rec.Top
        .Red = Val("&h" & Hex(FromColor.R) & "00")
        .Green = Val("&h" & Hex(FromColor.G) & "00")
        .Blue = Val("&h" & Hex(FromColor.B) & "00")
        .Alpha = 0&
    End With

    With vert(1)
        .X = Rec.Left + Rec.Right
        .Y = Rec.Top + Rec.Bottom
        .Red = Val("&h" & Hex(ToColor.R) & "00")
        .Green = Val("&h" & Hex(ToColor.G) & "00")
        .Blue = Val("&h" & Hex(ToColor.B) & "00")
        .Alpha = 0&
    End With

    gRect.UpperLeft = 0
    gRect.LowerRight = 1

    Call GradientFillRect(BackBuffer, vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_V)

End Sub

Private Function ShortVisibleFaces() As Integer
    
    Dim IsVisible As Boolean
    Dim i As Integer
    Dim iV As Integer

    With ObjPart
        iV = -1
        Erase .FaceV
        For i = 0 To .NumFaces
            IsVisible = IIf(VectorDot(ObjPart.NormalT(i), Camera) > 0, True, False)
            If IsVisible Then
                iV = iV + 1
                ReDim Preserve .FaceV(iV)
                .FaceV(iV).ZValue = ( _
                    .VerticesT(.Faces(i).A).Z + _
                    .VerticesT(.Faces(i).B).Z + _
                    .VerticesT(.Faces(i).C).Z)
                .FaceV(iV).iVisible = i
            End If
        Next
        If iV > -1 Then SortFaces 0, iV
        ShortVisibleFaces = iV
    End With

End Function

Private Sub SortFaces(ByVal First As Long, ByVal Last As Long)

    Dim FirstIdx  As Long
    Dim MidIdx As Long
    Dim LastIdx  As Long
    Dim MidVal As Single
    Dim TempOrder  As Order
    
    If (First < Last) Then
        With ObjPart
            MidIdx = (First + Last) \ 2
            MidVal = .FaceV(MidIdx).ZValue
            FirstIdx = First
            LastIdx = Last
            Do
                Do While .FaceV(FirstIdx).ZValue < MidVal
                    FirstIdx = FirstIdx + 1
                Loop
                Do While .FaceV(LastIdx).ZValue > MidVal
                    LastIdx = LastIdx - 1
                Loop
                If (FirstIdx <= LastIdx) Then
                    TempOrder = .FaceV(LastIdx)
                    .FaceV(LastIdx) = .FaceV(FirstIdx)
                    .FaceV(FirstIdx) = TempOrder
                    FirstIdx = FirstIdx + 1
                    LastIdx = LastIdx - 1
                End If
            Loop Until FirstIdx > LastIdx

            If (LastIdx <= MidIdx) Then
                SortFaces First, LastIdx
                SortFaces FirstIdx, Last
            Else
                SortFaces FirstIdx, Last
                SortFaces First, LastIdx
            End If
        End With
    End If

End Sub
