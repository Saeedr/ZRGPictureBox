Public Class cCommonCursors
    <System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint:="DestroyIcon")> _
    Private Shared Function DestroyIcon(<System.Runtime.InteropServices.InAttribute()> ByVal hIcon As System.IntPtr) As <System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.Bool)> Boolean
    End Function
    Private myInternalIcon As Icon = Nothing
    Private myCustomCursor As System.Windows.Forms.Cursor = Nothing
    Public Enum enCursorType As Integer
        Zoom = 0
        Edit = 2
    End Enum
    Public Sub New(ByVal CursorType As enCursorType)
        Try
            If CursorType = enCursorType.Zoom Then
                Dim bmp As Bitmap = LoadBmpRes("Zoom-32.png")
                myInternalIcon = Icon.FromHandle(bmp.GetHicon)
                myCustomCursor = New Cursor(myInternalIcon.Handle)
            ElseIf CursorType = enCursorType.Edit Then
                Dim bmp As Bitmap = LoadBmpRes("Edit.png")
                myInternalIcon = Icon.FromHandle(bmp.GetHicon)
                myCustomCursor = New Cursor(myInternalIcon.Handle)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Shared Function LoadBmpRes(ByVal cursorName As String) As System.Drawing.Bitmap
        Try
            Dim thisAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
            Dim assemblyName As String = thisAssembly.GetName.Name
            'cursorName = assemblyName + "." + cursorName
            Static resourcesNames As String() = thisAssembly.GetManifestResourceNames()
            Dim found As Boolean = False
            For Each name As String In resourcesNames
                If name.EndsWith(cursorName) Then
                    Dim file As System.IO.Stream = thisAssembly.GetManifestResourceStream(name)
                    Return New System.Drawing.Bitmap(file)
                End If
            Next
            Return Nothing
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try
    End Function
    Protected Overrides Sub Finalize()
        Try
            If myInternalIcon IsNot Nothing Then
                DestroyIcon(myInternalIcon.Handle)
            End If
        Catch ex As Exception

        End Try
        MyBase.Finalize()
    End Sub
    Public ReadOnly Property CustomCursor As Cursor
        Get
            Return myCustomCursor
        End Get
    End Property
    Private Shared myEditCursorHelper As cCommonCursors
    Public Shared ReadOnly Property EditCursor As Cursor
        Get
            Try
                If myEditCursorHelper Is Nothing Then
                    myEditCursorHelper = New cCommonCursors(cCommonCursors.enCursorType.Edit)
                End If
                Return myEditCursorHelper.CustomCursor
            Catch ex As Exception
                MsgBox(ex.Message)
                Return Cursors.No
            End Try
        End Get
    End Property
    Private Shared myZoomCursorHelper As cCommonCursors
    Public Shared ReadOnly Property ZoomCursor As Cursor
        Get
            Try
                If myZoomCursorHelper Is Nothing Then
                    myZoomCursorHelper = New cCommonCursors(cCommonCursors.enCursorType.Zoom)
                End If
                Return myZoomCursorHelper.CustomCursor
            Catch ex As Exception
                MsgBox(ex.Message)
                Return Cursors.No
            End Try
        End Get
    End Property
End Class
