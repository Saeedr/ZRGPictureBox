Imports System.Drawing.Drawing2D


Partial Public Class ZRGPictureBoxControl

    Public Class SelectionBoxElement

#Region "Variabili private"
        Public TopLeftCorner As System.Drawing.Point = System.Drawing.Point.Empty
        Public BottomRightCorner As System.Drawing.Point = RECT.InvalidPoint
        Public KeepAspectRatio As Boolean = False
        Public LinkedPictureBox As ZRGPictureBoxControl
        Private Shared myBoxPenAreaSelection As New Pen(Color.FromArgb(200, Color.Black))
        Private Shared myBoxPenSingleClick As New Pen(Color.FromArgb(200, Color.Red))
        Private Shared myBoxBrush As New SolidBrush(Color.FromArgb(40, Color.CadetBlue))
#End Region

#Region "Proprieta'"
        Public ReadOnly Property IsInvalid() As Boolean
            Get
                Return BottomRightCorner = RECT.InvalidPoint OrElse TopLeftCorner = RECT.InvalidPoint
            End Get
        End Property
        ''' <summary>
        ''' Ritorna la dimensione dell'area che va selezionata nel caso di selezione tramite singolo click.
        ''' </summary>
        Private ReadOnly Property PointSelectAreaSize() As Integer
            Get
                ' Calcolo l'area da tenere attorno al punto, tengo 15 pixel in tutto
                Return LinkedPictureBox.GraphicInfo.ToLogicalDimension(15.0!)
            End Get
        End Property
        Private ReadOnly Property SingleClickRectangle() As RECT
            Get
                Dim halfAreaSize As Integer = PointSelectAreaSize / 2
                Dim r As New RECT(TopLeftCorner.X - halfAreaSize, TopLeftCorner.Y - halfAreaSize, TopLeftCorner.X + halfAreaSize, TopLeftCorner.Y + halfAreaSize)
                r.NormalizeRect()
                Return r
            End Get
        End Property
        Public ReadOnly Property IsCreatedFromSinglePoint() As Boolean
            Get
                ' Check se il rettangolo ha entrambe le coordinate valide
                If IsInvalid Then
                    Return False
                End If
                ' Check se le coordinate sono uguali
                If (TopLeftCorner = BottomRightCorner) Then
                    Return True
                End If
                ' Se il "rettangolo da singolo click" contiene il secondo punto del box,
                ' allora il box di selezione e' stato creato tramite un singolo click
                Return SingleClickRectangle.Contains(BottomRightCorner)
            End Get
        End Property
#End Region

#Region "Operatori"
        Public Shared Widening Operator CType(ByVal box As SelectionBoxElement) As RECT
            If box.IsInvalid Then
                Return New RECT()
            End If
            If box.IsCreatedFromSinglePoint Then
                Return box.SingleClickRectangle
            Else
                Return box.RectFromPoints(box.TopLeftCorner, box.BottomRightCorner)
            End If
        End Operator
#End Region

#Region "Costruttori"
        Public Sub New(ByVal picBox As ZRGPictureBoxControl)
            LinkedPictureBox = picBox
        End Sub

#End Region

#Region "Funzioni private"
        Private Function RectFromPoints(ByVal FirstCorner As System.Drawing.Point, ByVal SecondCorner As System.Drawing.Point) As RECT
            Try
                If FirstCorner = RECT.InvalidPoint OrElse SecondCorner = RECT.InvalidPoint Then
                    Return New RECT()
                End If

                If KeepAspectRatio Then
                    Dim Sign As Integer
                    If (Math.Abs((SecondCorner.X - FirstCorner.X) / LinkedPictureBox.Width)) > Math.Abs(((SecondCorner.Y - FirstCorner.Y) / LinkedPictureBox.Height)) Then
                        If SecondCorner.Y > FirstCorner.Y Then Sign = 1 Else Sign = -1
                        SecondCorner.Y = FirstCorner.Y + Math.Abs((SecondCorner.X - FirstCorner.X) * (LinkedPictureBox.Height / LinkedPictureBox.Width)) * Sign
                    Else
                        If SecondCorner.X > FirstCorner.X Then Sign = 1 Else Sign = -1
                        SecondCorner.X = FirstCorner.X + Math.Abs((SecondCorner.Y - FirstCorner.Y) * (LinkedPictureBox.Width / LinkedPictureBox.Height)) * Sign
                    End If
                End If

                Dim r As New RECT(FirstCorner.X, FirstCorner.Y, SecondCorner.X, SecondCorner.Y)
                r.NormalizeRect()

                Return r
            Catch e As Exception
                MsgBox(e.Message)
                Return New RECT()
            End Try
        End Function

#End Region

#Region "Funzioni pubbliche"
        Public Sub Reset()
            TopLeftCorner = RECT.InvalidPoint
            BottomRightCorner = RECT.InvalidPoint
        End Sub
        Public Sub Draw(ByVal GR As Graphics, Optional ByVal usePhysicalCoords As Boolean = True)
            ' Check se questo box e' valido
            If Me.IsInvalid Then
                Exit Sub
            End If

            ' Trovo il rettangolo da invalidare
            Dim r As RECT = Me

            ' Se serve, converto in cordinate fisiche
            If usePhysicalCoords Then
                r = LinkedPictureBox.GraphicInfo.ToPhysicalRect(r)
            End If

            ' Check se il rettangolo ottenuto e' valido
            If r.IsZeroSized Then
                Exit Sub
            End If

            ' Disegno il rettangolo
            If Me.IsCreatedFromSinglePoint Then
                GR.DrawRectangle(myBoxPenSingleClick, r)
            Else
                GR.FillRectangle(myBoxBrush, r)
                GR.DrawRectangle(myBoxPenAreaSelection, r)
            End If
        End Sub
        Public Sub Invalidate()
            Dim r As RECT = Me
            ' NOTA: Alla Invalidate() vanno passate coordinate fisiche, non logiche
            r = LinkedPictureBox.GraphicInfo.ToPhysicalRect(r)
            r.Inflate(1, 1)
            ' Effettuo il ridisegno della PictureBox associata
            LinkedPictureBox.Invalidate(r)
        End Sub
#End Region

    End Class

End Class
