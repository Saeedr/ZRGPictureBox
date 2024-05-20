<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class cZoomButton
    Inherits System.Windows.Forms.UserControl
    'Inherits DevComponents.DotNetBar.Bar

    'UserControl esegue l'override del metodo Dispose per pulire l'elenco dei componenti.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Richiesto da Progettazione Windows Form
    Private components As System.ComponentModel.IContainer

    'NOTA: la procedura che segue è richiesta da Progettazione Windows Form
    'Può essere modificata in Progettazione Windows Form.  
    'Non modificarla nell'editor del codice.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(cZoomButton))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.btLoad = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.btZoom = New System.Windows.Forms.ToolStripButton()
        Me.btMeasure = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.btView = New System.Windows.Forms.ToolStripDropDownButton()
        Me.btViewRulers = New System.Windows.Forms.ToolStripMenuItem()
        Me.btViewScrollBars = New System.Windows.Forms.ToolStripMenuItem()
        Me.btViewGrid = New System.Windows.Forms.ToolStripMenuItem()
        Me.btUm = New System.Windows.Forms.ToolStripDropDownButton()
        Me.btUmMicron = New System.Windows.Forms.ToolStripMenuItem()
        Me.btUmDmm = New System.Windows.Forms.ToolStripMenuItem()
        Me.btUmMillimeters = New System.Windows.Forms.ToolStripMenuItem()
        Me.btUmInch = New System.Windows.Forms.ToolStripMenuItem()
        Me.btUmMeters = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripLabel1 = New System.Windows.Forms.ToolStripLabel()
        Me.tbPixelSizeMic = New System.Windows.Forms.ToolStripLabel()
        Me.btZoomFit = New System.Windows.Forms.ToolStripButton()
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStrip1
        '
        Me.ToolStrip1.AutoSize = False
        Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ToolStrip1.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btLoad, Me.ToolStripSeparator2, Me.btZoom, Me.btMeasure, Me.ToolStripSeparator1, Me.btView, Me.btUm, Me.ToolStripSeparator3, Me.ToolStripLabel1, Me.tbPixelSizeMic, Me.btZoomFit})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(552, 39)
        Me.ToolStrip1.TabIndex = 78
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'btLoad
        '
        Me.btLoad.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btLoad.Image = CType(resources.GetObject("btLoad.Image"), System.Drawing.Image)
        Me.btLoad.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.btLoad.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btLoad.Name = "btLoad"
        Me.btLoad.Size = New System.Drawing.Size(36, 36)
        Me.btLoad.Text = "ToolStripButton2"
        Me.btLoad.ToolTipText = "Load Image"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 39)
        '
        'btZoom
        '
        Me.btZoom.Checked = True
        Me.btZoom.CheckState = System.Windows.Forms.CheckState.Checked
        Me.btZoom.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btZoom.Image = CType(resources.GetObject("btZoom.Image"), System.Drawing.Image)
        Me.btZoom.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.btZoom.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btZoom.Name = "btZoom"
        Me.btZoom.Size = New System.Drawing.Size(36, 36)
        Me.btZoom.Text = "ToolStripButton1"
        Me.btZoom.ToolTipText = "Zoom mode"
        '
        'btMeasure
        '
        Me.btMeasure.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btMeasure.Image = CType(resources.GetObject("btMeasure.Image"), System.Drawing.Image)
        Me.btMeasure.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.btMeasure.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btMeasure.Name = "btMeasure"
        Me.btMeasure.Size = New System.Drawing.Size(36, 36)
        Me.btMeasure.Text = "ToolStripButton2"
        Me.btMeasure.ToolTipText = "Gauging mode"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 39)
        '
        'btView
        '
        Me.btView.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btView.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btViewRulers, Me.btViewScrollBars, Me.btViewGrid})
        Me.btView.Image = CType(resources.GetObject("btView.Image"), System.Drawing.Image)
        Me.btView.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.btView.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btView.Name = "btView"
        Me.btView.Size = New System.Drawing.Size(45, 36)
        Me.btView.Text = "ToolStripDropDownButton1"
        Me.btView.ToolTipText = "Visible items"
        '
        'btViewRulers
        '
        Me.btViewRulers.Checked = True
        Me.btViewRulers.CheckState = System.Windows.Forms.CheckState.Checked
        Me.btViewRulers.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.btViewRulers.Name = "btViewRulers"
        Me.btViewRulers.Size = New System.Drawing.Size(152, 22)
        Me.btViewRulers.Text = "Rulers"
        '
        'btViewScrollBars
        '
        Me.btViewScrollBars.Checked = True
        Me.btViewScrollBars.CheckState = System.Windows.Forms.CheckState.Checked
        Me.btViewScrollBars.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.btViewScrollBars.Name = "btViewScrollBars"
        Me.btViewScrollBars.Size = New System.Drawing.Size(152, 22)
        Me.btViewScrollBars.Text = "Scroll bars"
        '
        'btViewGrid
        '
        Me.btViewGrid.Checked = True
        Me.btViewGrid.CheckState = System.Windows.Forms.CheckState.Checked
        Me.btViewGrid.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.btViewGrid.Name = "btViewGrid"
        Me.btViewGrid.Size = New System.Drawing.Size(152, 22)
        Me.btViewGrid.Text = "Grid"
        '
        'btUm
        '
        Me.btUm.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btUm.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btUmMicron, Me.btUmDmm, Me.btUmMillimeters, Me.btUmInch, Me.btUmMeters})
        Me.btUm.Image = CType(resources.GetObject("btUm.Image"), System.Drawing.Image)
        Me.btUm.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.btUm.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btUm.Name = "btUm"
        Me.btUm.Size = New System.Drawing.Size(35, 36)
        Me.btUm.Text = "ToolStripDropDownButton1"
        Me.btUm.ToolTipText = "Measure unit"
        '
        'btUmMicron
        '
        Me.btUmMicron.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.btUmMicron.Name = "btUmMicron"
        Me.btUmMicron.Size = New System.Drawing.Size(152, 22)
        Me.btUmMicron.Text = "micron"
        '
        'btUmDmm
        '
        Me.btUmDmm.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.btUmDmm.Name = "btUmDmm"
        Me.btUmDmm.Size = New System.Drawing.Size(152, 22)
        Me.btUmDmm.Text = "mm/10"
        '
        'btUmMillimeters
        '
        Me.btUmMillimeters.Checked = True
        Me.btUmMillimeters.CheckState = System.Windows.Forms.CheckState.Checked
        Me.btUmMillimeters.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.btUmMillimeters.Name = "btUmMillimeters"
        Me.btUmMillimeters.Size = New System.Drawing.Size(152, 22)
        Me.btUmMillimeters.Text = "mm"
        '
        'btUmInch
        '
        Me.btUmInch.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.btUmInch.Name = "btUmInch"
        Me.btUmInch.Size = New System.Drawing.Size(152, 22)
        Me.btUmInch.Text = "inches"
        '
        'btUmMeters
        '
        Me.btUmMeters.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.btUmMeters.Name = "btUmMeters"
        Me.btUmMeters.Size = New System.Drawing.Size(152, 22)
        Me.btUmMeters.Text = "meters"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(6, 39)
        '
        'ToolStripLabel1
        '
        Me.ToolStripLabel1.Name = "ToolStripLabel1"
        Me.ToolStripLabel1.Size = New System.Drawing.Size(105, 36)
        Me.ToolStripLabel1.Text = "Pixel size (micron):"
        '
        'tbPixelSizeMic
        '
        Me.tbPixelSizeMic.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbPixelSizeMic.ForeColor = System.Drawing.Color.DarkBlue
        Me.tbPixelSizeMic.Name = "tbPixelSizeMic"
        Me.tbPixelSizeMic.Size = New System.Drawing.Size(37, 36)
        Me.tbPixelSizeMic.Text = "100"
        Me.tbPixelSizeMic.ToolTipText = "Click to change"
        '
        'btZoomFit
        '
        Me.btZoomFit.Checked = True
        Me.btZoomFit.CheckState = System.Windows.Forms.CheckState.Checked
        Me.btZoomFit.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btZoomFit.Image = CType(resources.GetObject("btZoomFit.Image"), System.Drawing.Image)
        Me.btZoomFit.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.btZoomFit.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btZoomFit.Name = "btZoomFit"
        Me.btZoomFit.Size = New System.Drawing.Size(36, 36)
        Me.btZoomFit.Text = "ToolStripButton1"
        Me.btZoomFit.ToolTipText = "Fit image"
        '
        'cZoomButton
        '
        Me.BackColor = System.Drawing.Color.Transparent
        Me.Controls.Add(Me.ToolStrip1)
        Me.Name = "cZoomButton"
        Me.Size = New System.Drawing.Size(552, 39)
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents btMeasure As System.Windows.Forms.ToolStripButton
    Friend WithEvents btZoom As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btView As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents btViewRulers As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btViewScrollBars As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btViewGrid As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btLoad As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btUm As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents btUmMicron As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btUmMillimeters As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btUmInch As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btUmMeters As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btUmDmm As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripLabel1 As System.Windows.Forms.ToolStripLabel
    Friend WithEvents tbPixelSizeMic As System.Windows.Forms.ToolStripLabel
    Friend WithEvents btZoomFit As System.Windows.Forms.ToolStripButton
End Class
