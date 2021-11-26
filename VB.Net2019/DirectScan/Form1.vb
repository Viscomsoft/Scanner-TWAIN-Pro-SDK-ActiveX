Option Strict Off
Option Explicit On
Friend Class Form1
	Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'For the start-up form, the first instance created is the default instance.
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkdefaultcap As System.Windows.Forms.CheckBox
    Public WithEvents txtcapheight As System.Windows.Forms.TextBox
    Public WithEvents txtcapwidth As System.Windows.Forms.TextBox
    Public WithEvents txtcaptop As System.Windows.Forms.TextBox
    Public WithEvents txtcapleft As System.Windows.Forms.TextBox
    Public WithEvents Label38 As System.Windows.Forms.Label
    Public WithEvents Label37 As System.Windows.Forms.Label
    Public WithEvents Label36 As System.Windows.Forms.Label
    Public WithEvents Label35 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Frame5 As System.Windows.Forms.GroupBox
    Public WithEvents Check4 As System.Windows.Forms.CheckBox
    Public WithEvents Command4 As System.Windows.Forms.Button
    Public WithEvents cbodpi As System.Windows.Forms.ComboBox
    Public WithEvents cbopixeltype As System.Windows.Forms.ComboBox
    Public WithEvents Check2 As System.Windows.Forms.CheckBox
    Public WithEvents Option2 As System.Windows.Forms.RadioButton
    Public WithEvents cboimagesource As System.Windows.Forms.ComboBox
    Public WithEvents Option1 As System.Windows.Forms.RadioButton
    Public WithEvents Label10 As System.Windows.Forms.Label
    Public WithEvents Label9 As System.Windows.Forms.Label
    Public WithEvents _Frame1_1 As System.Windows.Forms.GroupBox
    Public WithEvents Scanner1 As AxSCANNERLib.AxScanner
    Public WithEvents lbltotalpage As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    'Public WithEvents Frame1 As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Friend WithEvents chkenablefeeder As System.Windows.Forms.CheckBox
    Friend WithEvents chkenableduplex As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Public WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents Label43 As System.Windows.Forms.Label
    Friend WithEvents chkautosaveallpages As System.Windows.Forms.CheckBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me._Frame1_1 = New System.Windows.Forms.GroupBox()
        Me.chkenableduplex = New System.Windows.Forms.CheckBox()
        Me.chkenablefeeder = New System.Windows.Forms.CheckBox()
        Me.Frame5 = New System.Windows.Forms.GroupBox()
        Me.chkdefaultcap = New System.Windows.Forms.CheckBox()
        Me.txtcapheight = New System.Windows.Forms.TextBox()
        Me.txtcapwidth = New System.Windows.Forms.TextBox()
        Me.txtcaptop = New System.Windows.Forms.TextBox()
        Me.txtcapleft = New System.Windows.Forms.TextBox()
        Me.Label38 = New System.Windows.Forms.Label()
        Me.Label37 = New System.Windows.Forms.Label()
        Me.Label36 = New System.Windows.Forms.Label()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Check4 = New System.Windows.Forms.CheckBox()
        Me.Command4 = New System.Windows.Forms.Button()
        Me.cbodpi = New System.Windows.Forms.ComboBox()
        Me.cbopixeltype = New System.Windows.Forms.ComboBox()
        Me.Check2 = New System.Windows.Forms.CheckBox()
        Me.Option2 = New System.Windows.Forms.RadioButton()
        Me.cboimagesource = New System.Windows.Forms.ComboBox()
        Me.Option1 = New System.Windows.Forms.RadioButton()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Scanner1 = New AxSCANNERLib.AxScanner()
        Me.lbltotalpage = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.Label43 = New System.Windows.Forms.Label()
        Me.chkautosaveallpages = New System.Windows.Forms.CheckBox()
        Me._Frame1_1.SuspendLayout()
        Me.Frame5.SuspendLayout()
        CType(Me.Scanner1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        '_Frame1_1
        '
        Me._Frame1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Frame1_1.Controls.Add(Me.Label43)
        Me._Frame1_1.Controls.Add(Me.chkautosaveallpages)
        Me._Frame1_1.Controls.Add(Me.chkenableduplex)
        Me._Frame1_1.Controls.Add(Me.chkenablefeeder)
        Me._Frame1_1.Controls.Add(Me.Frame5)
        Me._Frame1_1.Controls.Add(Me.Check4)
        Me._Frame1_1.Controls.Add(Me.Command4)
        Me._Frame1_1.Controls.Add(Me.cbodpi)
        Me._Frame1_1.Controls.Add(Me.cbopixeltype)
        Me._Frame1_1.Controls.Add(Me.Check2)
        Me._Frame1_1.Controls.Add(Me.Option2)
        Me._Frame1_1.Controls.Add(Me.cboimagesource)
        Me._Frame1_1.Controls.Add(Me.Option1)
        Me._Frame1_1.Controls.Add(Me.Label10)
        Me._Frame1_1.Controls.Add(Me.Label9)
        Me._Frame1_1.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Frame1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Frame1_1.Location = New System.Drawing.Point(8, 0)
        Me._Frame1_1.Name = "_Frame1_1"
        Me._Frame1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame1_1.Size = New System.Drawing.Size(545, 199)
        Me._Frame1_1.TabIndex = 13
        Me._Frame1_1.TabStop = False
        Me._Frame1_1.Text = "Image Source "
        '
        'chkenableduplex
        '
        Me.chkenableduplex.Location = New System.Drawing.Point(384, 184)
        Me.chkenableduplex.Name = "chkenableduplex"
        Me.chkenableduplex.Size = New System.Drawing.Size(152, 16)
        Me.chkenableduplex.TabIndex = 118
        Me.chkenableduplex.Text = "Enable Duplex"
        '
        'chkenablefeeder
        '
        Me.chkenablefeeder.Location = New System.Drawing.Point(384, 135)
        Me.chkenablefeeder.Name = "chkenablefeeder"
        Me.chkenablefeeder.Size = New System.Drawing.Size(152, 16)
        Me.chkenablefeeder.TabIndex = 117
        Me.chkenablefeeder.Text = "Enable Feeder"
        '
        'Frame5
        '
        Me.Frame5.BackColor = System.Drawing.SystemColors.Control
        Me.Frame5.Controls.Add(Me.chkdefaultcap)
        Me.Frame5.Controls.Add(Me.txtcapheight)
        Me.Frame5.Controls.Add(Me.txtcapwidth)
        Me.Frame5.Controls.Add(Me.txtcaptop)
        Me.Frame5.Controls.Add(Me.txtcapleft)
        Me.Frame5.Controls.Add(Me.Label38)
        Me.Frame5.Controls.Add(Me.Label37)
        Me.Frame5.Controls.Add(Me.Label36)
        Me.Frame5.Controls.Add(Me.Label35)
        Me.Frame5.Controls.Add(Me.Label1)
        Me.Frame5.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame5.Location = New System.Drawing.Point(8, 64)
        Me.Frame5.Name = "Frame5"
        Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame5.Size = New System.Drawing.Size(369, 65)
        Me.Frame5.TabIndex = 116
        Me.Frame5.TabStop = False
        Me.Frame5.Text = "Capture Area"
        '
        'chkdefaultcap
        '
        Me.chkdefaultcap.BackColor = System.Drawing.SystemColors.Control
        Me.chkdefaultcap.Checked = True
        Me.chkdefaultcap.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkdefaultcap.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkdefaultcap.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkdefaultcap.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkdefaultcap.Location = New System.Drawing.Point(8, 16)
        Me.chkdefaultcap.Name = "chkdefaultcap"
        Me.chkdefaultcap.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkdefaultcap.Size = New System.Drawing.Size(97, 17)
        Me.chkdefaultcap.TabIndex = 125
        Me.chkdefaultcap.Text = "Use Default"
        Me.chkdefaultcap.UseVisualStyleBackColor = False
        '
        'txtcapheight
        '
        Me.txtcapheight.AcceptsReturn = True
        Me.txtcapheight.BackColor = System.Drawing.SystemColors.Window
        Me.txtcapheight.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtcapheight.Enabled = False
        Me.txtcapheight.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtcapheight.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtcapheight.Location = New System.Drawing.Point(280, 40)
        Me.txtcapheight.MaxLength = 0
        Me.txtcapheight.Name = "txtcapheight"
        Me.txtcapheight.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtcapheight.Size = New System.Drawing.Size(33, 20)
        Me.txtcapheight.TabIndex = 124
        Me.txtcapheight.Text = "11.5"
        '
        'txtcapwidth
        '
        Me.txtcapwidth.AcceptsReturn = True
        Me.txtcapwidth.BackColor = System.Drawing.SystemColors.Window
        Me.txtcapwidth.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtcapwidth.Enabled = False
        Me.txtcapwidth.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtcapwidth.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtcapwidth.Location = New System.Drawing.Point(200, 40)
        Me.txtcapwidth.MaxLength = 0
        Me.txtcapwidth.Name = "txtcapwidth"
        Me.txtcapwidth.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtcapwidth.Size = New System.Drawing.Size(33, 20)
        Me.txtcapwidth.TabIndex = 122
        Me.txtcapwidth.Text = "8.23"
        '
        'txtcaptop
        '
        Me.txtcaptop.AcceptsReturn = True
        Me.txtcaptop.BackColor = System.Drawing.SystemColors.Window
        Me.txtcaptop.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtcaptop.Enabled = False
        Me.txtcaptop.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtcaptop.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtcaptop.Location = New System.Drawing.Point(120, 40)
        Me.txtcaptop.MaxLength = 0
        Me.txtcaptop.Name = "txtcaptop"
        Me.txtcaptop.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtcaptop.Size = New System.Drawing.Size(33, 20)
        Me.txtcaptop.TabIndex = 120
        Me.txtcaptop.Text = "0"
        '
        'txtcapleft
        '
        Me.txtcapleft.AcceptsReturn = True
        Me.txtcapleft.BackColor = System.Drawing.SystemColors.Window
        Me.txtcapleft.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtcapleft.Enabled = False
        Me.txtcapleft.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtcapleft.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtcapleft.Location = New System.Drawing.Point(48, 40)
        Me.txtcapleft.MaxLength = 0
        Me.txtcapleft.Name = "txtcapleft"
        Me.txtcapleft.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtcapleft.Size = New System.Drawing.Size(33, 20)
        Me.txtcapleft.TabIndex = 118
        Me.txtcapleft.Text = "0"
        '
        'Label38
        '
        Me.Label38.BackColor = System.Drawing.SystemColors.Control
        Me.Label38.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label38.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label38.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label38.Location = New System.Drawing.Point(320, 40)
        Me.Label38.Name = "Label38"
        Me.Label38.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label38.Size = New System.Drawing.Size(41, 17)
        Me.Label38.TabIndex = 126
        Me.Label38.Text = "Inches"
        '
        'Label37
        '
        Me.Label37.BackColor = System.Drawing.SystemColors.Control
        Me.Label37.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label37.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label37.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label37.Location = New System.Drawing.Point(240, 40)
        Me.Label37.Name = "Label37"
        Me.Label37.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label37.Size = New System.Drawing.Size(33, 17)
        Me.Label37.TabIndex = 123
        Me.Label37.Text = "Height"
        '
        'Label36
        '
        Me.Label36.BackColor = System.Drawing.SystemColors.Control
        Me.Label36.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label36.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label36.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label36.Location = New System.Drawing.Point(160, 40)
        Me.Label36.Name = "Label36"
        Me.Label36.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label36.Size = New System.Drawing.Size(33, 17)
        Me.Label36.TabIndex = 121
        Me.Label36.Text = "Width"
        '
        'Label35
        '
        Me.Label35.BackColor = System.Drawing.SystemColors.Control
        Me.Label35.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label35.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label35.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label35.Location = New System.Drawing.Point(88, 40)
        Me.Label35.Name = "Label35"
        Me.Label35.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label35.Size = New System.Drawing.Size(33, 17)
        Me.Label35.TabIndex = 119
        Me.Label35.Text = "Top"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(8, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(25, 17)
        Me.Label1.TabIndex = 117
        Me.Label1.Text = "Left"
        '
        'Check4
        '
        Me.Check4.BackColor = System.Drawing.SystemColors.Control
        Me.Check4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Check4.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Check4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Check4.Location = New System.Drawing.Point(384, 111)
        Me.Check4.Name = "Check4"
        Me.Check4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Check4.Size = New System.Drawing.Size(153, 25)
        Me.Check4.TabIndex = 86
        Me.Check4.Text = "Clear Image Buffer After Scan"
        Me.Check4.UseVisualStyleBackColor = False
        '
        'Command4
        '
        Me.Command4.BackColor = System.Drawing.SystemColors.Control
        Me.Command4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command4.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command4.Location = New System.Drawing.Point(392, 64)
        Me.Command4.Name = "Command4"
        Me.Command4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command4.Size = New System.Drawing.Size(144, 27)
        Me.Command4.TabIndex = 85
        Me.Command4.Text = "Scan Image"
        Me.Command4.UseVisualStyleBackColor = False
        '
        'cbodpi
        '
        Me.cbodpi.BackColor = System.Drawing.SystemColors.Window
        Me.cbodpi.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbodpi.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbodpi.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbodpi.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cbodpi.Location = New System.Drawing.Point(400, 40)
        Me.cbodpi.Name = "cbodpi"
        Me.cbodpi.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbodpi.Size = New System.Drawing.Size(137, 20)
        Me.cbodpi.TabIndex = 84
        '
        'cbopixeltype
        '
        Me.cbopixeltype.BackColor = System.Drawing.SystemColors.Window
        Me.cbopixeltype.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbopixeltype.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbopixeltype.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbopixeltype.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cbopixeltype.Location = New System.Drawing.Point(400, 16)
        Me.cbopixeltype.Name = "cbopixeltype"
        Me.cbopixeltype.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbopixeltype.Size = New System.Drawing.Size(137, 20)
        Me.cbopixeltype.TabIndex = 82
        '
        'Check2
        '
        Me.Check2.BackColor = System.Drawing.SystemColors.Control
        Me.Check2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Check2.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Check2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Check2.Location = New System.Drawing.Point(384, 95)
        Me.Check2.Name = "Check2"
        Me.Check2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Check2.Size = New System.Drawing.Size(145, 17)
        Me.Check2.TabIndex = 24
        Me.Check2.Text = "Show TWAIN user interface"
        Me.Check2.UseVisualStyleBackColor = False
        '
        'Option2
        '
        Me.Option2.BackColor = System.Drawing.SystemColors.Control
        Me.Option2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Option2.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Option2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Option2.Location = New System.Drawing.Point(8, 40)
        Me.Option2.Name = "Option2"
        Me.Option2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Option2.Size = New System.Drawing.Size(145, 17)
        Me.Option2.TabIndex = 16
        Me.Option2.TabStop = True
        Me.Option2.Text = "Show All TWAIN Source"
        Me.Option2.UseVisualStyleBackColor = False
        '
        'cboimagesource
        '
        Me.cboimagesource.BackColor = System.Drawing.SystemColors.Window
        Me.cboimagesource.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboimagesource.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboimagesource.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboimagesource.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboimagesource.Location = New System.Drawing.Point(104, 16)
        Me.cboimagesource.Name = "cboimagesource"
        Me.cboimagesource.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboimagesource.Size = New System.Drawing.Size(233, 20)
        Me.cboimagesource.TabIndex = 15
        '
        'Option1
        '
        Me.Option1.BackColor = System.Drawing.SystemColors.Control
        Me.Option1.Checked = True
        Me.Option1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Option1.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Option1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Option1.Location = New System.Drawing.Point(8, 16)
        Me.Option1.Name = "Option1"
        Me.Option1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Option1.Size = New System.Drawing.Size(89, 17)
        Me.Option1.TabIndex = 14
        Me.Option1.TabStop = True
        Me.Option1.Text = "Manual Select"
        Me.Option1.UseVisualStyleBackColor = False
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(344, 40)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(25, 17)
        Me.Label10.TabIndex = 83
        Me.Label10.Text = "DPI"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Location = New System.Drawing.Point(344, 16)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(57, 25)
        Me.Label9.TabIndex = 81
        Me.Label9.Text = "Pixel Type"
        '
        'Scanner1
        '
        Me.Scanner1.Enabled = True
        Me.Scanner1.Location = New System.Drawing.Point(9, 193)
        Me.Scanner1.Name = "Scanner1"
        Me.Scanner1.OcxState = CType(resources.GetObject("Scanner1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.Scanner1.Size = New System.Drawing.Size(25, 31)
        Me.Scanner1.TabIndex = 0
        '
        'lbltotalpage
        '
        Me.lbltotalpage.BackColor = System.Drawing.SystemColors.Control
        Me.lbltotalpage.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbltotalpage.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbltotalpage.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lbltotalpage.Location = New System.Drawing.Point(214, 238)
        Me.lbltotalpage.Name = "lbltotalpage"
        Me.lbltotalpage.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbltotalpage.Size = New System.Drawing.Size(57, 17)
        Me.lbltotalpage.TabIndex = 29
        Me.lbltotalpage.Text = "0"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(62, 238)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(73, 17)
        Me.Label3.TabIndex = 28
        Me.Label3.Text = "Total Pages"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Button3)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Location = New System.Drawing.Point(574, 7)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(325, 173)
        Me.GroupBox1.TabIndex = 32
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Scan to Disk"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(37, 119)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(268, 44)
        Me.Button3.TabIndex = 2
        Me.Button3.Text = "Delete All scanned pages"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(37, 69)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(268, 44)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Merge images to Multi Page PDF"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(37, 20)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(268, 43)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Merge images to Multi Page TIF"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(65, 207)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(185, 17)
        Me.Label2.TabIndex = 33
        Me.Label2.Text = "Output Scanned Pages Folder"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(277, 205)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(276, 20)
        Me.TextBox1.TabIndex = 34
        '
        'Label43
        '
        Me.Label43.AutoSize = True
        Me.Label43.Location = New System.Drawing.Point(399, 169)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(81, 12)
        Me.Label43.TabIndex = 122
        Me.Label43.Text = "last page scanned"
        '
        'chkautosaveallpages
        '
        Me.chkautosaveallpages.AutoSize = True
        Me.chkautosaveallpages.Location = New System.Drawing.Point(384, 154)
        Me.chkautosaveallpages.Name = "chkautosaveallpages"
        Me.chkautosaveallpages.Size = New System.Drawing.Size(145, 16)
        Me.chkautosaveallpages.TabIndex = 121
        Me.chkautosaveallpages.Text = "Automate save to PDF when"
        Me.chkautosaveallpages.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(907, 289)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me._Frame1_1)
        Me.Controls.Add(Me.Scanner1)
        Me.Controls.Add(Me.lbltotalpage)
        Me.Controls.Add(Me.Label3)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 30)
        Me.Name = "Form1"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Direct Scan to Disk Sample"
        Me._Frame1_1.ResumeLayout(False)
        Me._Frame1_1.PerformLayout()
        Me.Frame5.ResumeLayout(False)
        Me.Frame5.PerformLayout()
        CType(Me.Scanner1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As Form1
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As Form1
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New Form1()
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(value As Form1)
            m_vb6FormDefInstance = value
        End Set
    End Property
#End Region
    Dim iX As Short
    Dim iY As Short

    Dim clrHashForeColor As Object
    Dim clrHashBackColor As Object
    Dim strApp As String

    Private Structure PictDesc
        Dim cbSizeofStruct As Integer
        Dim picType As Integer
        Dim hImage As Integer
        Dim xExt As Integer
        Dim yExt As Integer
    End Structure

    Private Structure Guid
        Dim Data1 As Integer
        Dim Data2 As Short
        Dim Data3 As Short
        <VBFixedArray(7)> Dim Data4() As Byte

        Public Sub Initialize()
            ReDim Data4(7)
        End Sub
    End Structure

    Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (ByRef lpPictDesc As PictDesc, ByRef riid As Guid, ByVal fPictureOwnsHandle As Integer, ByRef ipic As System.Drawing.Image) As Integer


    Private Sub Check2_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Check2.CheckStateChanged
        If Check2.CheckState = 1 Then
            Scanner1.ShowTwainUI = True
        Else
            Scanner1.ShowTwainUI = False
        End If


    End Sub

    Private Sub Check4_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Check4.CheckStateChanged
        If Check4.CheckState = 1 Then
            Me.Scanner1.ClearImageBuffer = True
        Else
            Me.Scanner1.ClearImageBuffer = False

        End If




    End Sub


    Private Sub chkshowtext_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
    End Sub

    Private Sub cmdoutlineback_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
       

    End Sub

    Private Sub cmdoutlineborder_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
       

    End Sub

    Private Sub cmdtextcolor_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        


    End Sub


    Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

    End Sub

    Private Sub Command10_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

     

    End Sub

    Private Sub Command11_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
       
    End Sub

    Private Sub Command12_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

     

    End Sub

    Private Sub Command13_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
       

    End Sub


    Private Sub Command17_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
    
    End Sub


    Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

    End Sub

    Private Sub Command20_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)


    End Sub



    Private Sub Command23_Click()



    End Sub





    Private Sub Command27_Click()

       
    End Sub

    Private Sub Command28_Click()
      

    End Sub

    Private Sub Command29_Click()

       

    End Sub

    Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

    End Sub


    Private Sub Command31_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
     

    End Sub


    Private Sub Command33_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

    End Sub

    Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command4.Click

        If Option1.Checked = True Then
            Scanner1.SelectImageSourceByIndex(cboimagesource.SelectedIndex)
        Else
            If Scanner1.SelectImageSource() <> True Then
                Exit Sub
            End If


        End If

        If chkenablefeeder.Checked = True Then
            Scanner1.FeederEnabled = True
        Else
            Scanner1.FeederEnabled = False
        End If

        If chkenableduplex.Checked = True Then
            Scanner1.DuplexEnabled = True
        Else
            Scanner1.DuplexEnabled = False
        End If

        Select Case cbodpi.SelectedIndex
            Case 0
                Scanner1.DPI = 96
            Case 1
                Scanner1.DPI = 200
            Case 2
                Scanner1.DPI = 300
            Case 3
                Scanner1.DPI = 600

        End Select



        Select Case cbopixeltype.SelectedIndex
            Case 0
                Scanner1.PixelType = -1
            Case 1
                Scanner1.PixelType = 0

            Case 2
                Scanner1.PixelType = 1

            Case 3
                Scanner1.PixelType = 2
        End Select


        If chkdefaultcap.CheckState = 1 Then
            Scanner1.SetCaptureArea(0, 0, 0, 0)
        Else
            Scanner1.SetCaptureArea(CDbl(txtcapleft.Text), CDbl(txtcaptop.Text), CDbl(txtcapwidth.Text), CDbl(txtcapheight.Text))
        End If


        Scanner1.Scan()




    End Sub


    Private Sub Command5_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

        Scanner1.View = 9
        Scanner1.Focus()

    End Sub

    Private Sub Command6_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

        Scanner1.View = 6
        Scanner1.Focus()

    End Sub

    Private Sub Command7_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Scanner1.View = 5
        Scanner1.Focus()


    End Sub

    Private Sub Command8_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Scanner1.View = 3
        Scanner1.Focus()


    End Sub

    Private Sub Command9_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

        Scanner1.View = 4
        Scanner1.Focus()

    End Sub

    Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim I As Integer
        Dim iCount As Integer
        Dim iFindIndex As Integer


        cbodpi.Items.Add("Onscreen Viewing 96dpi")
        cbodpi.Items.Add("Fax 200dpi")
        cbodpi.Items.Add("OCR Text 300dpi")
        cbodpi.Items.Add("Laser Print Fine 600dpi")
        cbodpi.SelectedIndex = 2

        cbopixeltype.Items.Add("Default")
        cbopixeltype.Items.Add("Gray Color")
        cbopixeltype.Items.Add("Black & White Color")
        cbopixeltype.Items.Add("True Color")

        cbopixeltype.SelectedIndex = 2




        iCount = Me.Scanner1.GetNumImageSources

        For I = 0 To iCount - 1
            cboimagesource.Items.Add(Scanner1.GetImageSourceName(I))
        Next

        If iCount > 0 Then
            cboimagesource.SelectedIndex = 0
        End If


        strApp = Application.ExecutablePath

        iFindIndex = strApp.IndexOf("Examples")

        If iFindIndex = -1 Then
            MessageBox.Show("Cannot found Temp folder, DirectScanInit method cannot called")
            Exit Sub
        End If


        strApp = strApp.Substring(0, iFindIndex) + "Temp"


        Scanner1.DirectScanInit(True, strApp, "tif")

        TextBox1.Text = strApp



    End Sub

    Private Sub Scanner1_DblClick(sender As Object, e As System.EventArgs) Handles Scanner1.DblClick

    End Sub

    Private Sub Scanner1_DirectScanedPage(sender As Object, e As AxSCANNERLib._DScannerEvents_DirectScanedPageEvent) Handles Scanner1.DirectScanedPage

        lbltotalpage.Text = CStr(Me.Scanner1.TotalPage)
    End Sub


    Private Sub Scanner1_EndScan(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)






    End Sub



    Private Sub Text1_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs)
       
    End Sub

    Private Sub txtLeft_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs)
        
    End Sub

    Private Sub txtTop_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs)
      
    End Sub

    Private Sub chkdefaultcap_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkdefaultcap.CheckedChanged
        If chkdefaultcap.Checked Then
            txtcapleft.Enabled = False
            txtcaptop.Enabled = False
            txtcapwidth.Enabled = False
            txtcapheight.Enabled = False

        Else
            txtcapleft.Enabled = True
            txtcaptop.Enabled = True
            txtcapwidth.Enabled = True
            txtcapheight.Enabled = True
        End If

    End Sub





    Public Function Color2Uint32(ByVal clr As Color) As UInt32
        Dim t As Int32
        Dim a() As Byte

        t = ColorTranslator.ToOle(clr)


        a = BitConverter.GetBytes(t)

        Return BitConverter.ToUInt32(a, 0)


    End Function




    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub chkfullpage_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub chkfullpage_Click(ByVal sender As Object, ByVal e As System.EventArgs)
       
    End Sub

    Private Sub Button1_Click_1(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim strFilter As String
        strFilter = "TIF File (*.tif)|*.tif"

        SaveFileDialog1.Filter = strFilter
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then

            Scanner1.DirectMergeTIF(SaveFileDialog1.FileName)

        End If

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Dim strFilter As String
        strFilter = "PDF File (*.pdf)|*.pdf"

        SaveFileDialog1.Filter = strFilter
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then

            Scanner1.DirectMergePDF(SaveFileDialog1.FileName)

        End If
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Scanner1.DirectScanClearAllPages(strApp)
    End Sub

    Private Sub Scanner1_DirectScanedAllPage(sender As System.Object, e As System.EventArgs) Handles Scanner1.DirectScanedAllPage
        Dim strFilter As String
        If chkautosaveallpages.Checked Then

            strFilter = "PDF File (*.pdf)|*.pdf"

            SaveFileDialog1.Filter = strFilter
            If SaveFileDialog1.ShowDialog() = DialogResult.OK Then

                Scanner1.DirectMergePDF(SaveFileDialog1.FileName)

            End If

        End If
    End Sub
End Class