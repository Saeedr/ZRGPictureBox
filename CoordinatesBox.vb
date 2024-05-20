Partial Public Class ZRGPictureBoxControl

    Friend Class CoordinatesBox

        Private myPictureBoxControl As ZRGPictureBoxControl
        Private myDrawingRect As Rectangle = Rectangle.Empty
        Private myLastCoordToDraw As Point = New Point(Integer.MaxValue, Integer.MaxValue)

#Region "Proprieta'"
        Public Property PictureBoxControl() As ZRGPictureBoxControl
            Get
                Return myPictureBoxControl
            End Get
            Private Set(ByVal value As ZRGPictureBoxControl)
                myPictureBoxControl = value
            End Set
        End Property
        Public ReadOnly Property UnitOfMeasure() As MeasureSystem.enUniMis
            Get
                Return PictureBoxControl.UnitOfMeasure
            End Get
        End Property
        Public ReadOnly Property DrawingRect() As Rectangle
            Get
                Return myDrawingRect
            End Get
        End Property
        Private ReadOnly Property UnitOfMeasureFactor() As Single
            Get
                Return MeasureSystem.CustomUnitToMicron(1, UnitOfMeasure)
            End Get
        End Property
        Private ReadOnly Property UnitOfMeasureString() As String
            Get
                Return MeasureSystem.UniMisDescription(UnitOfMeasure)
            End Get
        End Property
#End Region

#Region "Costruttori"
        Public Sub New(ByVal pictureBox As ZRGPictureBoxControl)
            myPictureBoxControl = pictureBox
        End Sub
#End Region

#Region "Funzioni pubbliche"
        Public Sub DrawCoordinateInfo(ByVal GR As Graphics, ByVal CoordToDraw As Point, Optional ByVal PixelCoordMode As Boolean = False)
            Try
                If myPictureBoxControl Is Nothing Then
                    Exit Sub
                End If
                If GR Is Nothing Then
                    Exit Sub
                End If
                If CoordToDraw.X = Integer.MaxValue OrElse CoordToDraw.Y = Integer.MaxValue Then
                    Exit Sub
                End If
                myLastCoordToDraw = CoordToDraw

                Static textFont As Font = New Drawing.Font("Arial narrow", 8)

                Static borderSize As Integer = Math.Ceiling(GR.MeasureString("_", textFont).Width / 2)
                Dim _umsf As Single = UnitOfMeasureFactor
                If PixelCoordMode Then
                    _umsf = 1
                End If
                Dim xValue As Single = CoordToDraw.X / _umsf
                Dim yValue As Single = CoordToDraw.Y / _umsf
                Dim textToDraw As String
                If PixelCoordMode Then
                    textToDraw = "X=" + xValue.ToString("0000.00") + ", Y=" + yValue.ToString("0000.00")
                Else
                    If UnitOfMeasure <> MeasureSystem.enUniMis.micron Then
                        textToDraw = "X=" + xValue.ToString("0000.00") + ", Y=" + yValue.ToString("0000.00") + UnitOfMeasureString
                    Else
                        ' Se l'uita' di misura e' micron, non ho cifre dopo la virgola
                        textToDraw = "X=" + xValue.ToString("0000") + ", Y=" + yValue.ToString("0000") + UnitOfMeasureString
                    End If
                End If

                Dim textBox As SizeF = GR.MeasureString(textToDraw, textFont)

                ' Se il box cambia dimensioni, invalido la picturebox, cosi' sembra sempre che sia pulito
                ' Serve nel caso in cui:
                '  - si passa da "X=100,Y=100" a "X=99,Y=99", per cui rimarrebbe una parte del box precedente "scoperta"
                '  - si cambia unita' di misura a runtime, quindi il box viene ridimensionato
                Static oldTextBox As SizeF = textBox
                If oldTextBox <> textBox Then
                    ' Aggiorno le dimensioni precedenti del box
                    oldTextBox = textBox
                End If

                ' Aggiorno le coordinate del rettangolo di sfondo
                ' NOTA: Uso il ClientRectangle.Width al posto di Width perche' cosi' tengo conto delle eventuali scrollbar.
                myDrawingRect.X = myPictureBoxControl.ClientRectangle.Width - textBox.Width - borderSize
                myDrawingRect.Y = myPictureBoxControl.ClientRectangle.Height - textBox.Height - borderSize
                myDrawingRect.Width = textBox.Width + borderSize
                myDrawingRect.Height = textBox.Height + borderSize

                ' Se le scrollbar sono visibili, devo fare in modo che sia visibile anche il bordo inferiore/destro del rettangolo
                If myPictureBoxControl.HScroll Then
                    myDrawingRect.Height -= 1
                End If
                If myPictureBoxControl.VScroll Then
                    myDrawingRect.Width -= 1
                End If

                ' Disegno il rettangolo di sfondo
                GR.FillRectangle(Brushes.White, myDrawingRect)
                GR.DrawRectangle(Pens.Black, myDrawingRect)

                ' Disegno la stringa di testo, l'aggiunta di borderSize/2 serve a centrare il testo nel rettangolo di sfondo 
                GR.DrawString(textToDraw, textFont, Brushes.Black, myDrawingRect.X + borderSize / 2, myDrawingRect.Y + borderSize / 2)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Sub

#End Region

    End Class

End Class
