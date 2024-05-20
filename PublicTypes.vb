Imports System.Runtime.InteropServices
Imports System.Drawing

Public Enum GridKind
    FullLines = 0
    Points = 1
    Crosses = 2
End Enum

Public Enum enClickAction
    None = 0
    Zoom = 2
    MeasureDistance = 3
End Enum

Public Enum ResizeMode
    Normal
    Stretch
End Enum

<CLSCompliantAttribute(True), Serializable(), StructLayout(LayoutKind.Sequential, Pack:=4), _
 DebuggerDisplay("Left={left} Top={top} Right={right} Bottom={bottom} [Width={Width},Height={Height}]")> _
Public Structure RECT
    Public Shared Function InvalidPoint() As System.Drawing.Point
        Dim retVal As New System.Drawing.Point
        retVal.X = Integer.MaxValue
        retVal.Y = Integer.MaxValue
        Return retVal
    End Function

#Region "Public members"
    Public left As Integer
    Public top As Integer
    Public right As Integer
    Public bottom As Integer
#End Region

#Region "Operators"
    Public Shared Operator =(ByVal R1 As RECT, ByVal R2 As RECT) As Boolean
        Return R1.top = R2.top AndAlso R1.left = R2.left AndAlso R1.right = R2.right AndAlso R1.bottom = R2.bottom
    End Operator
    Public Shared Operator <>(ByVal R1 As RECT, ByVal R2 As RECT) As Boolean
        Return R1.top <> R2.top OrElse R1.left <> R2.left OrElse R1.right <> R2.right OrElse R1.bottom <> R2.bottom
    End Operator
    Public Shared Widening Operator CType(ByVal InRect As RECT) As Rectangle
        Return New Rectangle(InRect.left, InRect.top, InRect.right - InRect.left, InRect.bottom - InRect.top)
    End Operator
    Public Shared Widening Operator CType(ByVal InRect As RECT) As RectangleF
        Return New RectangleF(InRect.left, InRect.top, InRect.right - InRect.left, InRect.bottom - InRect.top)
    End Operator
    Public Shared Widening Operator CType(ByVal InRect As RectangleF) As RECT
        Return New RECT(InRect)
    End Operator
#End Region

#Region "Constructors"
    Public Sub New(ByVal InRect As RECT)
        top = InRect.Y
        left = InRect.X
        right = InRect.X + InRect.Width
        bottom = InRect.Y + InRect.Height
    End Sub
    Public Sub New(ByVal InRect As Rectangle)
        top = InRect.Y
        left = InRect.X
        Try
            right = InRect.X + InRect.Width
            bottom = InRect.Y + InRect.Height
        Catch ex As Exception
            right = InRect.X + 1000
            bottom = InRect.Y + 1000
        End Try
    End Sub
    Public Sub New(ByVal InRect As RectangleF)
        top = CType(InRect.Y, Integer)
        left = CType(InRect.X, Integer)
        right = CType(InRect.X + InRect.Width, Integer)
        bottom = CType(InRect.Y + InRect.Height, Integer)
    End Sub
    Public Sub New(ByVal vector() As System.Drawing.Point)
        Try
            If vector Is Nothing Then
                Return
            Else
                Dim TotalPoints As Integer = vector.Length
                If TotalPoints = 0 Then
                    Return
                Else
                    left = vector(0).X
                    right = vector(0).X
                    top = vector(0).Y
                    bottom = vector(0).Y
                    For iCounter As Integer = 1 To TotalPoints - 1
                        If vector(iCounter).X > right Then
                            right = vector(iCounter).X
                        Else
                            If vector(iCounter).X < left Then
                                left = vector(iCounter).X
                            End If
                        End If
                        If vector(iCounter).Y > bottom Then
                            bottom = vector(iCounter).Y
                        Else
                            If vector(iCounter).Y < top Then
                                top = vector(iCounter).Y
                            End If
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub New(ByVal left As Integer, ByVal top As Integer, ByVal right As Integer, ByVal bottom As Integer)
        Me.top = top
        Me.left = left
        Me.right = right
        Me.bottom = bottom
    End Sub
    Public Sub New(ByVal pointTopLeft As System.Drawing.Point, ByVal size As System.Drawing.Size)
        left = pointTopLeft.X
        top = pointTopLeft.Y
        right = pointTopLeft.X + size.Width
        bottom = pointTopLeft.Y + size.Height
    End Sub
    Public Sub New(ByVal pointTopLeft As System.Drawing.Point, ByVal pointBottomRight As System.Drawing.Point)
        left = pointTopLeft.X
        top = pointTopLeft.Y
        right = pointBottomRight.X
        bottom = pointBottomRight.Y
    End Sub

#End Region

#Region "Properties'"
    Public Property X() As Integer
        Get
            Return Me.left
        End Get
        Set(ByVal value As Integer)
            Me.left = value
            Me.right = Me.left + Me.Width
        End Set
    End Property
    Public Property Y() As Integer
        Get
            Return Me.top
        End Get
        Set(ByVal value As Integer)
            Me.top = value
            Me.bottom = Me.top + Me.Height
        End Set
    End Property
    Public Property Width() As Integer
        Get
            Return right - left
        End Get
        Set(ByVal value As Integer)
            right = left + value
        End Set
    End Property
    Public Property Height() As Integer
        Get
            Return bottom - top
        End Get
        Set(ByVal value As Integer)
            bottom = top + value
        End Set
    End Property
    Public ReadOnly Property CenterPoint() As Point
        Get
            Return New Point((left + right) / 2, (top + bottom) / 2)
        End Get
    End Property
    Public ReadOnly Property Size() As System.Drawing.Size
        Get
            Return New System.Drawing.Size(Width, Height)
        End Get
    End Property
    Public Property TopLeft() As System.Drawing.Point
        Get
            Return New System.Drawing.Point(Me.left, Me.top)
        End Get
        Set(ByVal value As System.Drawing.Point)
            Me.left = value.X
            Me.top = value.Y
        End Set
    End Property
    Public Property TopRight() As System.Drawing.Point
        Get
            Return New System.Drawing.Point(Me.right, Me.top)
        End Get
        Set(ByVal value As System.Drawing.Point)
            Me.right = value.X
            Me.top = value.Y
        End Set
    End Property
    Public Property BottomRight() As System.Drawing.Point
        Get
            Return New System.Drawing.Point(Me.right, Me.bottom)
        End Get
        Set(ByVal value As System.Drawing.Point)
            Me.right = value.X
            Me.bottom = value.Y
        End Set
    End Property
    Public ReadOnly Property BottomCenter() As System.Drawing.Point
        Get
            Return New System.Drawing.Point((Me.left + Me.right) / 2, Me.bottom)
        End Get
    End Property
    Public ReadOnly Property TopCenter() As System.Drawing.Point
        Get
            Return New System.Drawing.Point((Me.left + Me.right) / 2, Me.top)
        End Get
    End Property
    Public ReadOnly Property LeftCenter() As System.Drawing.Point
        Get
            Return New System.Drawing.Point(Me.left, (Me.top + Me.bottom) / 2)
        End Get
    End Property
    Public ReadOnly Property RightCenter() As System.Drawing.Point
        Get
            Return New System.Drawing.Point(Me.right, (Me.top + Me.bottom) / 2)
        End Get
    End Property
    Public Property BottomLeft() As System.Drawing.Point
        Get
            Return New System.Drawing.Point(Me.left, Me.bottom)
        End Get
        Set(ByVal value As System.Drawing.Point)
            Me.left = value.X
            Me.bottom = value.Y
        End Set
    End Property
    Public ReadOnly Property IsZeroSized() As Boolean
        Get
            Return (Me.Height = 0 AndAlso Me.Width = 0)
        End Get
    End Property
    Public ReadOnly Property IsNonZeroSized() As Boolean
        Get
            Return Not IsZeroSized
        End Get
    End Property
    Public ReadOnly Property IsNormalized() As Boolean
        Get
            Return (Me.right >= Me.left) AndAlso (Me.bottom >= Me.top)
        End Get
    End Property
#End Region

#Region "Funzioni pubbliche"

#Region "Routine per la normalizzazione del rettangolo"
    Public Sub AssertIfNotNormalized()
        If IsNormalized() Then
            Return
        Else
            If Not (Me.right >= Me.left) Then
                Debug.Assert(Me.right >= Me.left, "RECT.right e RECT.left sono invertite!")
            End If
            If Not (Me.bottom >= Me.top) Then
                Debug.Assert(Me.bottom >= Me.top, "RECT.bottom e RECT.top sono invertite!")
            End If
        End If
    End Sub
    Public Sub NormalizeRect()
        If Not (Me.right >= Me.left) Then
            Dim tmp As Integer = Me.right
            Me.right = Me.left
            Me.left = tmp
        End If
        If Not (Me.bottom >= Me.top) Then
            Dim tmp As Integer = Me.bottom
            Me.bottom = Me.top
            Me.top = tmp
        End If
    End Sub
#End Region

#Region "Routine per spostare il rettangolo"
    Public Sub Offset(ByVal x As Integer, ByVal y As Integer)
        left = left + x
        top = top + y
        right = right + x
        bottom = bottom + y
    End Sub
    Public Sub Offset(ByVal offs As System.Drawing.Point)
        Offset(offs.X, offs.Y)
    End Sub
#End Region

#Region "Routine per ridimensionare il rettangolo"
    Public Sub Inflate(ByVal size As Size)
        Me.Inflate(size.Width, size.Height)
    End Sub
    Public Sub Inflate(ByVal width As Integer, ByVal height As Integer)
        Me.left = (Me.left - width)
        Me.top = (Me.top - height)
        Me.right = (Me.right + width)
        Me.bottom = (Me.bottom + height)
    End Sub
    Public Sub Inflate(ByVal left As Integer, ByVal top As Integer, ByVal right As Integer, ByVal bottom As Integer)
        Me.left = (Me.left - left)
        Me.top = (Me.top - top)
        Me.right = (Me.right + right)
        Me.bottom = (Me.bottom + bottom)
    End Sub
    Public Function ExpandFromFixedPoint(ByVal zoomMultiplier As Single, ByVal fixedPoint As Point) As RECT
        Dim distanceX As Single = Me.left - fixedPoint.X
        Dim distanceY As Single = Me.top - fixedPoint.Y
        distanceX *= zoomMultiplier
        distanceY *= zoomMultiplier

        Dim newOriginX As Single = fixedPoint.X + distanceX
        Dim newOriginY As Single = fixedPoint.Y + distanceY

        Dim newWidth As Single = zoomMultiplier * Me.Width
        Dim newHeight As Single = zoomMultiplier * Me.Height

        Return New RECT(newOriginX, newOriginY, newOriginX + newWidth, newOriginY + newHeight)
    End Function

#End Region

#Region "Routine di check per punti o RECT contenuti"
    Public Function IsContainedIn(ByRef ARect As RECT) As Boolean
        Try
            ' Check preliminari
            Me.AssertIfNotNormalized()
            ARect.AssertIfNotNormalized()
            ' Controllo le coordinate
            If (Me.bottom <= ARect.bottom) AndAlso (Me.top >= ARect.top) Then
                If (Me.left >= ARect.left) AndAlso (Me.right <= ARect.right) Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
    End Function
    Public Function Contains(ByRef ARect As RECT) As Boolean
        Return ARect.IsContainedIn(Me)
    End Function
    Public Function Contains(ByRef ARect As Rectangle) As Boolean
        Dim ar As New RECT(ARect)
        Return ar.IsContainedIn(Me)
    End Function
    Public Function Contains(ByRef pt As System.Drawing.PointF) As Boolean
        Try
            AssertIfNotNormalized()
            If pt.X > Me.right OrElse pt.X < Me.left Then Return False
            If pt.Y > Me.bottom OrElse pt.Y < Me.top Then Return False
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
    End Function
    Public Function Contains(ByRef pt As System.Drawing.Point) As Boolean
        Try
            AssertIfNotNormalized()
            If pt.X > Me.right OrElse pt.X < Me.left Then Return False
            If pt.Y > Me.bottom OrElse pt.Y < Me.top Then Return False
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
    End Function
#End Region

    Public Function IntersectsWith(ByRef rect As RECT) As Boolean
        AssertIfNotNormalized()
        rect.AssertIfNotNormalized()
        Return Not Intersect(Me, rect).IsZeroSized
    End Function

    Public Overrides Function ToString() As String
        Return String.Format("Left={0} Top={1} Right={2} Bottom={3} [Width={4},Height={5}]", left, top, right, bottom, Width, Height)
    End Function

    Public Function ToPointArray() As System.Drawing.Point()
        Dim _PArr(4) As System.Drawing.Point
        _PArr(0).X = Me.left
        _PArr(0).Y = Me.top
        _PArr(1).X = Me.left
        _PArr(1).Y = Me.bottom
        _PArr(2).X = Me.right
        _PArr(2).Y = Me.bottom
        _PArr(3).X = Me.right
        _PArr(3).Y = Me.top
        _PArr(4).X = Me.left
        _PArr(4).Y = Me.top
        Return _PArr
    End Function

    Public Function ToRectangle() As System.Drawing.Rectangle
        Return New System.Drawing.Rectangle(Me.left, Me.top, Me.Width, Me.Height)
    End Function

#End Region

#Region "Funzioni statiche"
    Public Shared Function Union(ByRef a As RECT, ByRef b As RECT) As RECT
        Dim ra As Rectangle = Rectangle.Union(CType(a, Rectangle), CType(b, Rectangle))
        Return New RECT(ra)
    End Function
    Public Shared Function UnionWithoutZeroSized(ByRef a As RECT, ByRef b As RECT) As RECT
        If a.IsZeroSized Then Return b

        If b.IsZeroSized Then Return a

        Dim ra As Rectangle = Rectangle.Union(CType(a, Rectangle), CType(b, Rectangle))
        Return New RECT(ra)
    End Function
    Public Shared Function Intersect(ByRef a As RECT, ByRef b As RECT) As RECT
        Dim ra As Rectangle = Rectangle.Intersect(CType(a, Rectangle), CType(b, Rectangle))
        Return New RECT(ra)
    End Function
    Public Shared Function IntersectWithoutInvalid(ByVal a As RECT, ByVal b As RECT) As RECT
        If a.IsZeroSized Then
            Return b
        End If
        If b.IsZeroSized Then
            Return a
        End If
        Return Intersect(a, b)
    End Function
    Public Shared Function Inflate(ByVal r As RECT, ByVal x As Integer, ByVal y As Integer) As RECT
        Dim rectangle1 As RECT = r
        rectangle1.Inflate(x, y)
        Return rectangle1
    End Function
    Public Shared Function CoordsBoundaries(ByVal coords() As System.Drawing.Point) As RECT
        Dim retVal As RECT
        If (coords IsNot Nothing) AndAlso (coords.Length > 0) Then
            Try
                retVal.left = coords(0).X
                retVal.right = coords(0).X
                retVal.top = coords(0).Y
                retVal.bottom = coords(0).Y
                For cCounter As Integer = 0 To coords.Length - 1
                    If coords(cCounter).X > retVal.right Then
                        retVal.right = coords(cCounter).X
                    End If
                    If coords(cCounter).X < retVal.left Then
                        retVal.left = coords(cCounter).X
                    End If
                    If coords(cCounter).Y < retVal.top Then
                        retVal.top = coords(cCounter).Y
                    End If
                    If coords(cCounter).Y > retVal.bottom Then
                        retVal.bottom = coords(cCounter).Y
                    End If
                Next
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        Return retVal
    End Function
#End Region

End Structure

<CLSCompliantAttribute(True), Serializable(), StructLayout(LayoutKind.Sequential, Pack:=4)> _
Public Structure SEGMENT

#Region "Members"
    Dim X0 As Integer
    Dim Y0 As Integer
    Dim X1 As Integer
    Dim Y1 As Integer
    Public Property P0() As System.Drawing.Point
        Get
            Return New System.Drawing.Point(X0, Y0)
        End Get
        Set(ByVal value As System.Drawing.Point)
            X0 = value.X
            Y0 = value.Y
        End Set
    End Property
    Public Property P1() As System.Drawing.Point
        Get
            Return New System.Drawing.Point(X1, Y1)
        End Get
        Set(ByVal value As System.Drawing.Point)
            X1 = value.X
            Y1 = value.Y
        End Set
    End Property
#End Region

#Region "Constructors"
    Public Sub New(ByVal aSEGMENT As SEGMENT)
        Try
            Me.X0 = aSEGMENT.X0
            Me.Y0 = aSEGMENT.Y0
            Me.X1 = aSEGMENT.X1
            Me.Y1 = aSEGMENT.Y1
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub New(ByVal X0 As Integer, ByVal Y0 As Integer, ByVal X1 As Integer, ByVal Y1 As Integer)
        Try
            Me.X0 = X0
            Me.Y0 = Y0
            Me.X1 = X1
            Me.Y1 = Y1
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub New(ByVal P0 As System.Drawing.Point, ByVal P1 As System.Drawing.Point)
        Try
            Me.X0 = P0.X
            Me.Y0 = P0.Y
            Me.X1 = P1.X
            Me.Y1 = P1.Y
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
#End Region

#Region "Member functions"
    Public Function ContainsX(ByVal XQuote As Integer) As Boolean
        If (XQuote >= P0.X) AndAlso (XQuote <= P1.X) Then  '(Valido se P1 a sinistra di P0)
            Return True
        Else
            If (XQuote >= P1.X) AndAlso (XQuote <= P0.X) Then  '(Valido se P0 a sinistra di P1)
                Return True
            End If
        End If
        Return False
    End Function
    Public Function MediumPoint() As System.Drawing.Point
        Try
            Dim retVal As System.Drawing.Point
            retVal.X = (Me.X0 + Me.X1) / 2
            retVal.Y = (Me.Y0 + Me.Y1) / 2
            Return retVal
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function
    Public Shared Function SegmentModule(ByVal P0 As System.Drawing.Point, ByVal P1 As System.Drawing.Point) As Double
        Try
            Return System.Math.Sqrt(System.Math.Pow(P1.X - P0.X, 2) + System.Math.Pow(P1.Y - P0.Y, 2))
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Public Function SegmentModule() As Double
        Try
            Return System.Math.Sqrt(System.Math.Pow(X1 - X0, 2) + System.Math.Pow(Y1 - Y0, 2))
        Catch ex As Exception
            Return 0
        End Try
    End Function
    ''' <summary>
    ''' Ritorna la direzione (angolo in radianti) del segmento ...
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SegmentDirection() As Double
        Try
            Dim dblHyp As Double = 0
            Dim dblSin As Double = 0
            Dim RefX As Double = 0
            Dim RefY As Double = 0
            'Traslo il segmento in modo che parta dall'origine ...
            RefX = X1 - X0
            RefY = -(Y1 - Y0) 'Memo: in Windows l'asse Y è invertito ...
            'Riporto a sistema di coordinate standard per
            'applicare formule trigonometriche ...

            If (RefY = 0) Then
                'Segmento orizzontale ...
                If (RefX > 0) Then
                    'Angolo nullo ...
                    Return 0
                Else
                    'Angolo piatto ...
                    Return System.Math.PI
                End If
            End If

            If (RefX = 0) Then
                'Segmento verticale ...
                If (RefY > 0) Then
                    Return System.Math.PI / 2
                Else
                    Return -System.Math.PI / 2
                End If
            End If

            'Se sono arrivato fino a qui, l'angolo non è un multiplo di Pi/2 ...
            If (RefX > 0) Then
                If (RefY > 0) Then
                    'Primo quadrante ....
                    dblHyp = System.Math.Sqrt((RefX * RefX + RefY * RefY))    'Ipotenusa ...
                    dblSin = RefY / dblHyp
                    Return System.Math.Atan(dblSin / System.Math.Sqrt(-dblSin * dblSin + 1))
                Else
                    'Quarto quadrante ...
                    RefY = -RefY
                    dblHyp = System.Math.Sqrt((RefX * RefX + RefY * RefY))    'Ipotenusa ...
                    dblSin = RefY / dblHyp
                    Return (2 * System.Math.PI) - System.Math.Atan(dblSin / System.Math.Sqrt(-dblSin * dblSin + 1))
                End If
            Else
                If (RefY > 0) Then
                    'Secondo quadrante ...
                    RefX = -RefX
                    dblHyp = System.Math.Sqrt((RefX * RefX + RefY * RefY))  'Ipotenusa ...
                    dblSin = CDbl(RefY) / dblHyp
                    Return -System.Math.Atan(dblSin / System.Math.Sqrt(-dblSin * dblSin + 1)) + System.Math.PI
                Else
                    'Terzo quadrante ...
                    RefX = -RefX
                    RefY = -RefY
                    dblHyp = System.Math.Sqrt((RefX * RefX + RefY * RefY))  'Ipotenusa ...
                    dblSin = RefY / dblHyp
                    Return System.Math.Atan(dblSin / System.Math.Sqrt(-dblSin * dblSin + 1)) + System.Math.PI
                End If
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
#End Region

#Region "Operators"
    Public Shared Operator =(ByVal S1 As SEGMENT, ByVal S2 As SEGMENT) As Boolean
        If (S1.X0 = S2.X0) AndAlso (S1.X1 = S2.X1) AndAlso (S1.Y0 = S2.Y0) AndAlso (S1.Y1 = S2.Y1) Then
            Return True
        Else
            Return False
        End If
    End Operator
    Public Shared Operator <>(ByVal S1 As SEGMENT, ByVal S2 As SEGMENT) As Boolean
        If (S1.X0 <> S2.X0) Or (S1.X1 <> S2.X1) Or (S1.Y0 <> S2.Y0) Or (S1.Y1 <> S2.Y1) Then
            Return True
        Else
            Return False
        End If
    End Operator
#End Region


End Structure

