Partial Public Class ZRGPictureBoxControl
    Friend Class CrossCursor

        ''' <summary>
        ''' Cross cursor default size
        ''' </summary>
        Public Shared ReadOnly DefaultSize As Size = New Size(20, 20)


#Region "Variabili private"

        ''' <summary>
        ''' Controllo su cui disegnare il cursore
        ''' </summary>
        Private myPictureBox As ZRGPictureBoxControl

        ''' <summary>
        ''' Dimensione del cursore a croce
        ''' </summary>
        Private mySize As Size = DefaultSize

        ''' <summary>
        ''' Flag che indica se la croce va disegnata in modalita' "schermo pieno"
        ''' </summary>
        Private myFullPictureBoxCross As Boolean = False

        ''' <summary>
        ''' Colore con cui disegnare la croce
        ''' </summary>
        Private myColor As Color = Drawing.Color.Black

        ''' <summary>
        ''' Posizione su cui viene disegnata la croce [coordinate logiche]
        ''' </summary>
        Private myCrossPosition As System.Drawing.Point = RECT.InvalidPoint

        ''' <summary>
        ''' Rettangolo contenente i quattro punti corrispondenti all'ultima croce disegnata
        ''' </summary>
        Private myLastCrossTopPoint As System.Drawing.Point
        Private myLastCrossLeftPoint As System.Drawing.Point
        Private myLastCrossRightPoint As System.Drawing.Point
        Private myLastCrossBottomPoint As System.Drawing.Point

        ''' <summary>
        ''' Box in cui vengono disegnate le coordinate, viene disegnata solo la parte di cursore che non cade sopra ad esso
        ''' </summary>
        Private myCoordinatesBox As CoordinatesBox = Nothing

#End Region

#Region "Proprieta'"

        ''' <summary>
        ''' PictureBoxControl associato a questa istanza delle classe
        ''' </summary>
        Public Property PictureBoxControl() As ZRGPictureBoxControl
            Get
                Return myPictureBox
            End Get
            Private Set(ByVal value As ZRGPictureBoxControl)
                myPictureBox = value
            End Set
        End Property

        ''' <summary>
        ''' Dimensione del cursore a croce
        ''' </summary>
        Public Property Size() As Size
            Get
                Return mySize
            End Get
            Set(ByVal Value As Size)
                mySize = Value
            End Set
        End Property

        ''' <summary>
        ''' Colore del cursore a croce
        ''' </summary>
        Public Property Color() As Color
            Get
                Return myColor
            End Get
            Set(ByVal Value As Color)
                myColor = Value
            End Set
        End Property
        ''' <summary>
        ''' Box in cui vengono disegnate le coordinate, viene disegnata solo la parte di cursore che non cade sopra ad esso
        ''' </summary>
        Friend Property CoordinatesBox() As CoordinatesBox
            Get
                Return myCoordinatesBox
            End Get
            Set(ByVal value As CoordinatesBox)
                myCoordinatesBox = value
            End Set
        End Property

#End Region

#Region "Costruttori"

        ''' <summary>
        ''' Costruttore dato un controllo su cui disegnare il cursore
        ''' </summary>
        Public Sub New(ByVal picPictureBox As ZRGPictureBoxControl)
            myPictureBox = picPictureBox
        End Sub

#End Region

#Region "Funzioni pubbliche"
        ''' <summary>
        ''' Cancella la posizione su cui viene disegnata la croce. 
        ''' </summary>
        Public Sub ResetCrossPosition()
            myCrossPosition = RECT.InvalidPoint
        End Sub

        ''' <summary>
        ''' Posizione su cui viene disegnata la croce [coordinate logiche]
        ''' </summary>
        Public Property CrossPosition() As System.Drawing.Point
            Get
                Return myCrossPosition
            End Get
            Set(ByVal value As System.Drawing.Point)
                myCrossPosition = value
            End Set
        End Property

        ''' <summary>
        ''' Disegna la croce nella posizione data da CrossPosition
        ''' </summary>
        Friend Sub DrawCross(ByVal GR As Graphics)
            DrawCross(GR, CrossPosition)
        End Sub

        ''' <summary>
        ''' Disegna la croce nella posizione specificata
        ''' </summary>
        Friend Sub DrawCross(ByVal GR As Graphics, ByVal LogicalCoord As System.Drawing.Point)
            Try

                ' Check se ho un controllo associato valido
                If myPictureBox Is Nothing Then
                    Exit Sub
                End If

                ' Check se la coordinata passatami e' valida
                If LogicalCoord = RECT.InvalidPoint Then
                    Exit Sub
                End If

                ' Posizione in cui disegnare la croce [coordinate fisiche della pictureBox]
                Dim physicalCrossCoords As Point = PictureBoxControl.GraphicInfo.ToPhysicalPoint(LogicalCoord)

                ' Minimi e massimi valori permessi per i bracci della croce
                Dim minCrossValue As Point = Point.Empty
                Dim maxCrossValue As Point = New Point(myPictureBox.Width, myPictureBox.Height)

                ' Se ho un box delle coordinate, faccio in modo di non disegnarci sopra
                If myCoordinatesBox IsNot Nothing Then
                    ' Controllo che la croce (non a schermo pieno) non cada completamente dentro al box delle coordinate
                    ' Se e' completamente dentro, non serve disegnare la croce
                    If Not myFullPictureBoxCross AndAlso myCoordinatesBox.DrawingRect.Contains(physicalCrossCoords) Then
                        Exit Sub
                    End If

                    If physicalCrossCoords.X > myCoordinatesBox.DrawingRect.X Then
                        maxCrossValue.Y -= myCoordinatesBox.DrawingRect.Height
                    End If
                    If physicalCrossCoords.Y > myCoordinatesBox.DrawingRect.Y Then
                        maxCrossValue.X -= myCoordinatesBox.DrawingRect.Width
                    End If
                End If

                ' Due valori che utilizzo spesso
                Dim maxCrossValueX As Integer = maxCrossValue.X '- 2
                Dim maxCrossValueY As Integer = maxCrossValue.Y '- 2

                ' Calcolo la posizione della nuova croce
                If myFullPictureBoxCross Then
                    ' Linea orizzontale
                    myLastCrossLeftPoint.X = minCrossValue.X
                    myLastCrossRightPoint.X = maxCrossValue.X
                    myLastCrossLeftPoint.Y = physicalCrossCoords.Y
                    myLastCrossRightPoint.Y = physicalCrossCoords.Y
                    ' Linea verticale
                    myLastCrossTopPoint.Y = minCrossValue.Y
                    myLastCrossBottomPoint.Y = maxCrossValue.Y
                    myLastCrossTopPoint.X = physicalCrossCoords.X
                    myLastCrossBottomPoint.X = physicalCrossCoords.X
                Else
                    ' Linea orizzontale
                    myLastCrossLeftPoint.X = physicalCrossCoords.X - mySize.Width / 2
                    myLastCrossRightPoint.X = physicalCrossCoords.X + mySize.Width / 2
                    myLastCrossLeftPoint.Y = physicalCrossCoords.Y
                    myLastCrossRightPoint.Y = physicalCrossCoords.Y
                    ' Linea verticale
                    myLastCrossTopPoint.Y = physicalCrossCoords.Y - mySize.Height / 2
                    myLastCrossBottomPoint.Y = physicalCrossCoords.Y + mySize.Height / 2
                    myLastCrossTopPoint.X = physicalCrossCoords.X
                    myLastCrossBottomPoint.X = physicalCrossCoords.X
                End If

                ' Controllo che la croce non debordi dalla PictureBox
                ' Va fatto anche nel caso della croce a pieno schermo, perche' potrebbe debordare
                ' nel caso in cui la pictureBox non occupa tutto lo spazio disponibile nell'applicazione.
                ' In questo caso io andrei a disegnare sopra gli altri controlli dell'applicazione
                If myLastCrossRightPoint.X > maxCrossValueX Then myLastCrossRightPoint.X = maxCrossValueX
                If myLastCrossRightPoint.Y > maxCrossValueY Then myLastCrossRightPoint.Y = maxCrossValueY
                If myLastCrossRightPoint.Y < minCrossValue.Y Then myLastCrossRightPoint.Y = minCrossValue.Y
                If myLastCrossBottomPoint.Y > maxCrossValueY Then myLastCrossBottomPoint.Y = maxCrossValueY
                If myLastCrossBottomPoint.X > maxCrossValueX Then myLastCrossBottomPoint.X = maxCrossValueX
                If myLastCrossBottomPoint.X < minCrossValue.X Then myLastCrossBottomPoint.X = minCrossValue.X
                If myLastCrossLeftPoint.X < minCrossValue.X Then myLastCrossLeftPoint.X = minCrossValue.X
                If myLastCrossLeftPoint.Y > maxCrossValueY Then myLastCrossLeftPoint.Y = maxCrossValueY
                If myLastCrossLeftPoint.Y < minCrossValue.Y Then myLastCrossLeftPoint.Y = minCrossValue.Y
                If myLastCrossTopPoint.Y < minCrossValue.Y Then myLastCrossTopPoint.Y = minCrossValue.Y
                If myLastCrossTopPoint.X > maxCrossValueX Then myLastCrossTopPoint.X = maxCrossValueX
                If myLastCrossTopPoint.X < minCrossValue.X Then myLastCrossTopPoint.X = minCrossValue.X

                Using crossPen As New Pen(myColor)
                    GR.DrawLine(crossPen, myLastCrossLeftPoint, myLastCrossRightPoint)
                    GR.DrawLine(crossPen, myLastCrossTopPoint, myLastCrossBottomPoint)
                End Using

            Catch ex As Exception
                MsgBox(ex.Message) 'MsgBox(ex.Message)
            End Try
        End Sub

#End Region

    End Class

End Class
