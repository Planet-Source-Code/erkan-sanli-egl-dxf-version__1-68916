Attribute VB_Name = "modObject"
Option Explicit

Public Type Vertex
    A       As Integer
    B       As Integer
    C       As Integer
End Type

Public Type Order           'Y�zeylerin Z y�ksekliklerinin s�ralanmas� i�in kullan�l�yor.
    ZValue   As Single      'Z de�eri
    iVisible As Integer     'FaceV indexi
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type Part            'Par�an�n tan�mland��� de�i�ken
    Caption As String       'Par�an�n ad�
    Position As Vector      'Orijine g�re par�an�n konumu
    Direction As Vector     'Koordinat sistemine g�re par�an�n d�n�kl���,yat�kl���
    Vertices() As Vector    'Orjinal nokta tan�mlamalar�
    VerticesT() As Vector   'Transform olmu� noktalar(rotate ve/veya transform olmu�)
    VerticesV() As Vector   'G�r�nen nokta tan�mlamalar�
    ScreenCoord() As POINTAPI
    Normal()  As Vector     'Orjinal y�zey normalleri.Dosya y�klendi�inde bir kere hesaplan�yor.
    NormalT()  As Vector    'Orjinal y�zey normalleri.Dosya y�klendi�inde bir kere hesaplan�yor.
    Faces() As Vertex
    NumVertices As Integer
    NumFaces As Integer
    FaceV() As Order        '"Visible Face" G�r�nen y�zeylerin Face indexlerinin sakland��� de�i�ken
    Scale As Single         'Par�an�n b�y�kl���.Ayn� dosyadan farkl� b�y�kl�kte par�alar elde etmek i�in.
    color As ColorRGB
End Type

Public Camera As Vector
Public Light As Vector
Public LightT As Vector
Public ObjPart As Part
Dim eofflag As Boolean

Public Sub LoadObject(ByVal strFileName As String)

    Dim i      As Integer
    Dim rgbcol As ColorRGB

    With ObjPart
        Open strFileName For Input As 1
            Input #1, .Caption                                  'Obje ad�
            Input #1, .Scale                                    '�l�ek
            Input #1, .color.R, .color.G, .color.B              'RGB, renk bilgileri
            Input #1, .Direction.X, .Direction.Y, .Direction.Z  'Rotasyon
            Input #1, .Position.X, .Position.Y, .Position.Z     'Pozisyon
            Input #1, .NumVertices, .NumFaces                   'Nokta ve y�zey adetleri
            ReDim .Vertices(.NumVertices)
            ReDim .ScreenCoord(.NumVertices)
            ReDim .Faces(.NumFaces)
            ReDim .Normal(.NumFaces)
            For i = 0 To (.NumVertices)                         'Noktalar(Vertices)
                Input #1, .Vertices(i).X, _
                          .Vertices(i).Y, _
                          .Vertices(i).Z
                          .Vertices(i).W = 1
            Next i
            For i = 0 To (.NumFaces)                            'Y�zeyler(Faces)
                Input #1, .Faces(i).A, _
                          .Faces(i).B, _
                          .Faces(i).C
            Next i
            .VerticesT = .Vertices
            Call CalculateNormal                                'Y�klenen par�an�n normali hesaplan�yor.
        Close #1
    End With

End Sub

Public Sub LoadDXF(ByVal strFileName As String)
    
    Dim X As String
   
    
    With ObjPart
        .Caption = "DXF File"
        .Scale = 700
        .color.R = 210: .color.G = 100: .color.B = 0
        .Direction.X = -90: .Direction.Y = 0: .Direction.Z = 0
        .Position.X = 0: .Position.Y = 0: .Position.Z = -0.3
        eofflag = False
        Open strFileName For Input As 1
            Do Until eofflag
                Call FindCommand("3DFACE")
                If Not eofflag Then Call ParseSection
            Loop
            .NumVertices = .NumVertices - 1
            .NumFaces = .NumFaces - 1
            ReDim .ScreenCoord(.NumVertices)
            ReDim .Normal(.NumFaces)
            .VerticesT = .Vertices
            Call CalculateNormal                                'Y�klenen par�an�n normali hesaplan�yor.
        Close #1
    End With
End Sub

Sub FindCommand(Command As String)
    
    Dim X As String
    
    Do While UCase(Trim(X)) <> UCase(Command)
        Line Input #1, X
        If UCase(Trim(X)) = "EOF" Then
            eofflag = True
            Exit Sub
        End If
    Loop

End Sub

Sub ParseSection()
    
    Dim X As String
    Dim sngX As Single
    
    With ObjPart
        
        ReDim Preserve .Faces(.NumFaces)
        
        Line Input #1, X        '8
        Line Input #1, X        '0
        
        .Faces(.NumFaces).A = .NumVertices
        ReDim Preserve .Vertices(.NumVertices)
        Line Input #1, X        '10
        Line Input #1, X        '0.039410
        .Vertices(.NumVertices).X = Val(X)
        Line Input #1, X        '20
        Line Input #1, X        '0.005292
        .Vertices(.NumVertices).Y = Val(X)
        Line Input #1, X        '30
        Line Input #1, X        '0.582973
        .Vertices(.NumVertices).Z = Val(X)
        .Vertices(.NumVertices).W = 1
        .NumVertices = .NumVertices + 1
        
        .Faces(.NumFaces).B = .NumVertices
        ReDim Preserve .Vertices(.NumVertices)
        Line Input #1, X        '11
        Line Input #1, X        '0.035595
        .Vertices(.NumVertices).X = Val(X)
        Line Input #1, X        '21
        Line Input #1, X        '0.013980
        .Vertices(.NumVertices).Y = Val(X)
        Line Input #1, X        '31
        Line Input #1, X        '0.586025
        .Vertices(.NumVertices).Z = Val(X)
        .Vertices(.NumVertices).W = 1
        .NumVertices = .NumVertices + 1
        
        .Faces(.NumFaces).C = .NumVertices
        ReDim Preserve .Vertices(.NumVertices)
        Line Input #1, X        '12
        Line Input #1, X        '0.029246
        .Vertices(.NumVertices).X = Val(X)
        Line Input #1, X        '22
        Line Input #1, X        '-0.000407
        .Vertices(.NumVertices).Y = Val(X)
        Line Input #1, X        '32
        Line Input #1, X        '0.583105
        .Vertices(.NumVertices).Z = Val(X)
        .Vertices(.NumVertices).W = 1
        .NumVertices = .NumVertices + 1
        .NumFaces = .NumFaces + 1
    End With
        
'        Line Input #FileNum, X        '13
'        Line Input #FileNum, X        '0.039410
'        Line Input #FileNum, X        '23
'        Line Input #FileNum, X        '0.005292
'        Line Input #FileNum, X        '33
'        Line Input #FileNum, X        '0.582973
    
        
        

End Sub
