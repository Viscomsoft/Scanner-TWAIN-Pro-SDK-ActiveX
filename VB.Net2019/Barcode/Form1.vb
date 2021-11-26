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
	Public WithEvents Command17 As System.Windows.Forms.Button
    Public WithEvents txtLeft As System.Windows.Forms.TextBox
    Public WithEvents txtTop As System.Windows.Forms.TextBox
    Public WithEvents Text1 As System.Windows.Forms.TextBox
    Public WithEvents Command33 As System.Windows.Forms.Button
    Public WithEvents Text2 As System.Windows.Forms.TextBox
    Public WithEvents Command31 As System.Windows.Forms.Button
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
    Public WithEvents Command12 As System.Windows.Forms.Button
    Public WithEvents Command11 As System.Windows.Forms.Button
    Public WithEvents Command7 As System.Windows.Forms.Button
    Public WithEvents Command6 As System.Windows.Forms.Button
    Public WithEvents Command8 As System.Windows.Forms.Button
    Public WithEvents Command9 As System.Windows.Forms.Button
    Public WithEvents Command10 As System.Windows.Forms.Button
    Public WithEvents Command5 As System.Windows.Forms.Button
    Public WithEvents Command20 As System.Windows.Forms.Button
    Public WithEvents Command3 As System.Windows.Forms.Button
    Public WithEvents Command2 As System.Windows.Forms.Button
    Public WithEvents Command1 As System.Windows.Forms.Button
    Public WithEvents Scanner1 As AxSCANNERLib.AxScanner
    Public WithEvents Label34 As System.Windows.Forms.Label
    Public WithEvents Label33 As System.Windows.Forms.Label
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents Label32 As System.Windows.Forms.Label
    Public WithEvents Label31 As System.Windows.Forms.Label
    Public WithEvents lblcurrentpage As System.Windows.Forms.Label
    Public WithEvents Label8 As System.Windows.Forms.Label
    Public WithEvents lbltotalpage As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    'Public WithEvents Frame1 As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Friend WithEvents chkenablefeeder As System.Windows.Forms.CheckBox
    Friend WithEvents chkenableduplex As System.Windows.Forms.CheckBox
    Friend WithEvents ColorDialog1 As System.Windows.Forms.ColorDialog
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents txtconfid As System.Windows.Forms.TextBox
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents chkscanmore As System.Windows.Forms.CheckBox
    Friend WithEvents txtcustomrotation As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents chkScan90Rotation As System.Windows.Forms.CheckBox
    Friend WithEvents chkScan45CounterRotation As System.Windows.Forms.CheckBox
    Friend WithEvents chkScan45Rotation As System.Windows.Forms.CheckBox
    Friend WithEvents chkscannorotation As System.Windows.Forms.CheckBox
    Friend WithEvents chkScanMultiple As System.Windows.Forms.CheckBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents txtbarcodeheight As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtbarcodewidth As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtbarcodetop As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtbarcodeleft As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents chkfullpage As System.Windows.Forms.CheckBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Command17 = New System.Windows.Forms.Button()
        Me.txtLeft = New System.Windows.Forms.TextBox()
        Me.txtTop = New System.Windows.Forms.TextBox()
        Me.Text1 = New System.Windows.Forms.TextBox()
        Me.Command33 = New System.Windows.Forms.Button()
        Me.Text2 = New System.Windows.Forms.TextBox()
        Me.Command31 = New System.Windows.Forms.Button()
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
        Me.Command12 = New System.Windows.Forms.Button()
        Me.Command11 = New System.Windows.Forms.Button()
        Me.Command7 = New System.Windows.Forms.Button()
        Me.Command6 = New System.Windows.Forms.Button()
        Me.Command8 = New System.Windows.Forms.Button()
        Me.Command9 = New System.Windows.Forms.Button()
        Me.Command10 = New System.Windows.Forms.Button()
        Me.Command5 = New System.Windows.Forms.Button()
        Me.Command20 = New System.Windows.Forms.Button()
        Me.Command3 = New System.Windows.Forms.Button()
        Me.Command2 = New System.Windows.Forms.Button()
        Me.Command1 = New System.Windows.Forms.Button()
        Me.Scanner1 = New AxSCANNERLib.AxScanner()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.lblcurrentpage = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.lbltotalpage = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ColorDialog1 = New System.Windows.Forms.ColorDialog()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.txtconfid = New System.Windows.Forms.TextBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.chkscanmore = New System.Windows.Forms.CheckBox()
        Me.txtcustomrotation = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.chkScan90Rotation = New System.Windows.Forms.CheckBox()
        Me.chkScan45CounterRotation = New System.Windows.Forms.CheckBox()
        Me.chkScan45Rotation = New System.Windows.Forms.CheckBox()
        Me.chkscannorotation = New System.Windows.Forms.CheckBox()
        Me.chkScanMultiple = New System.Windows.Forms.CheckBox()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.txtbarcodeheight = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtbarcodewidth = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtbarcodetop = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtbarcodeleft = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.chkfullpage = New System.Windows.Forms.CheckBox()
        Me._Frame1_1.SuspendLayout()
        Me.Frame5.SuspendLayout()
        CType(Me.Scanner1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Command17
        '
        Me.Command17.BackColor = System.Drawing.SystemColors.Control
        Me.Command17.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command17.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command17.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command17.Location = New System.Drawing.Point(394, 535)
        Me.Command17.Name = "Command17"
        Me.Command17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command17.Size = New System.Drawing.Size(145, 41)
        Me.Command17.TabIndex = 115
        Me.Command17.Text = "Delete Current Page"
        Me.Command17.UseVisualStyleBackColor = False
        '
        'txtLeft
        '
        Me.txtLeft.AcceptsReturn = True
        Me.txtLeft.BackColor = System.Drawing.SystemColors.Window
        Me.txtLeft.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLeft.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLeft.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLeft.Location = New System.Drawing.Point(44, 443)
        Me.txtLeft.MaxLength = 0
        Me.txtLeft.Name = "txtLeft"
        Me.txtLeft.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLeft.Size = New System.Drawing.Size(57, 20)
        Me.txtLeft.TabIndex = 91
        Me.txtLeft.Text = "0"
        Me.txtLeft.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtTop
        '
        Me.txtTop.AcceptsReturn = True
        Me.txtTop.BackColor = System.Drawing.SystemColors.Window
        Me.txtTop.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTop.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTop.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTop.Location = New System.Drawing.Point(44, 467)
        Me.txtTop.MaxLength = 0
        Me.txtTop.Name = "txtTop"
        Me.txtTop.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTop.Size = New System.Drawing.Size(57, 20)
        Me.txtTop.TabIndex = 90
        Me.txtTop.Text = "0"
        Me.txtTop.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Text1
        '
        Me.Text1.AcceptsReturn = True
        Me.Text1.BackColor = System.Drawing.SystemColors.Window
        Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text1.Location = New System.Drawing.Point(12, 419)
        Me.Text1.MaxLength = 0
        Me.Text1.Name = "Text1"
        Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text1.Size = New System.Drawing.Size(65, 20)
        Me.Text1.TabIndex = 87
        Me.Text1.Text = "100"
        Me.Text1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Command33
        '
        Me.Command33.BackColor = System.Drawing.SystemColors.Control
        Me.Command33.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command33.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command33.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command33.Location = New System.Drawing.Point(113, 535)
        Me.Command33.Name = "Command33"
        Me.Command33.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command33.Size = New System.Drawing.Size(257, 41)
        Me.Command33.TabIndex = 27
        Me.Command33.Text = "Update Current Change to Image Buffer"
        Me.Command33.UseVisualStyleBackColor = False
        '
        'Text2
        '
        Me.Text2.AcceptsReturn = True
        Me.Text2.BackColor = System.Drawing.SystemColors.Window
        Me.Text2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text2.Location = New System.Drawing.Point(425, 463)
        Me.Text2.MaxLength = 0
        Me.Text2.Name = "Text2"
        Me.Text2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text2.Size = New System.Drawing.Size(41, 20)
        Me.Text2.TabIndex = 26
        Me.Text2.Text = "0"
        '
        'Command31
        '
        Me.Command31.BackColor = System.Drawing.SystemColors.Control
        Me.Command31.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command31.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command31.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command31.Location = New System.Drawing.Point(481, 463)
        Me.Command31.Name = "Command31"
        Me.Command31.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command31.Size = New System.Drawing.Size(41, 25)
        Me.Command31.TabIndex = 25
        Me.Command31.Text = "Go"
        Me.Command31.UseVisualStyleBackColor = False
        '
        '_Frame1_1
        '
        Me._Frame1_1.BackColor = System.Drawing.SystemColors.Control
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
        Me._Frame1_1.Location = New System.Drawing.Point(539, 12)
        Me._Frame1_1.Name = "_Frame1_1"
        Me._Frame1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame1_1.Size = New System.Drawing.Size(465, 233)
        Me._Frame1_1.TabIndex = 13
        Me._Frame1_1.TabStop = False
        Me._Frame1_1.Text = "Image Source "
        '
        'chkenableduplex
        '
        Me.chkenableduplex.Location = New System.Drawing.Point(301, 152)
        Me.chkenableduplex.Name = "chkenableduplex"
        Me.chkenableduplex.Size = New System.Drawing.Size(152, 16)
        Me.chkenableduplex.TabIndex = 118
        Me.chkenableduplex.Text = "Enable Duplex"
        '
        'chkenablefeeder
        '
        Me.chkenablefeeder.Location = New System.Drawing.Point(301, 136)
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
        Me.Frame5.Location = New System.Drawing.Point(13, 102)
        Me.Frame5.Name = "Frame5"
        Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame5.Size = New System.Drawing.Size(212, 125)
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
        Me.txtcapheight.Location = New System.Drawing.Point(131, 81)
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
        Me.txtcapwidth.Location = New System.Drawing.Point(51, 81)
        Me.txtcapwidth.MaxLength = 0
        Me.txtcapwidth.Name = "txtcapwidth"
        Me.txtcapwidth.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtcapwidth.Size = New System.Drawing.Size(43, 20)
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
        Me.txtcaptop.Location = New System.Drawing.Point(127, 40)
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
        Me.Label38.Location = New System.Drawing.Point(165, 61)
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
        Me.Label37.Location = New System.Drawing.Point(91, 81)
        Me.Label37.Name = "Label37"
        Me.Label37.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label37.Size = New System.Drawing.Size(43, 17)
        Me.Label37.TabIndex = 123
        Me.Label37.Text = "Height"
        '
        'Label36
        '
        Me.Label36.BackColor = System.Drawing.SystemColors.Control
        Me.Label36.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label36.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label36.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label36.Location = New System.Drawing.Point(11, 81)
        Me.Label36.Name = "Label36"
        Me.Label36.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label36.Size = New System.Drawing.Size(34, 20)
        Me.Label36.TabIndex = 121
        Me.Label36.Text = "Width"
        '
        'Label35
        '
        Me.Label35.BackColor = System.Drawing.SystemColors.Control
        Me.Label35.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label35.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label35.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label35.Location = New System.Drawing.Point(91, 40)
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
        Me.Check4.Location = New System.Drawing.Point(301, 112)
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
        Me.Command4.Location = New System.Drawing.Point(309, 67)
        Me.Command4.Name = "Command4"
        Me.Command4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command4.Size = New System.Drawing.Size(145, 26)
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
        Me.cbodpi.Location = New System.Drawing.Point(317, 43)
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
        Me.cbopixeltype.Location = New System.Drawing.Point(317, 19)
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
        Me.Check2.Location = New System.Drawing.Point(301, 96)
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
        Me.Option2.Location = New System.Drawing.Point(13, 73)
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
        Me.cboimagesource.Location = New System.Drawing.Point(13, 42)
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
        Me.Option1.Location = New System.Drawing.Point(12, 19)
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
        Me.Label10.Location = New System.Drawing.Point(261, 43)
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
        Me.Label9.Location = New System.Drawing.Point(261, 19)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(57, 25)
        Me.Label9.TabIndex = 81
        Me.Label9.Text = "Pixel Type"
        '
        'Command12
        '
        Me.Command12.BackColor = System.Drawing.SystemColors.Control
        Me.Command12.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command12.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command12.Location = New System.Drawing.Point(13, 108)
        Me.Command12.Name = "Command12"
        Me.Command12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command12.Size = New System.Drawing.Size(89, 25)
        Me.Command12.TabIndex = 12
        Me.Command12.Text = "25% of Original"
        Me.Command12.UseVisualStyleBackColor = False
        '
        'Command11
        '
        Me.Command11.BackColor = System.Drawing.SystemColors.Control
        Me.Command11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command11.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command11.Location = New System.Drawing.Point(13, 140)
        Me.Command11.Name = "Command11"
        Me.Command11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command11.Size = New System.Drawing.Size(89, 25)
        Me.Command11.TabIndex = 11
        Me.Command11.Text = "33% of Original"
        Me.Command11.UseVisualStyleBackColor = False
        '
        'Command7
        '
        Me.Command7.BackColor = System.Drawing.SystemColors.Control
        Me.Command7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command7.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command7.Location = New System.Drawing.Point(13, 236)
        Me.Command7.Name = "Command7"
        Me.Command7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command7.Size = New System.Drawing.Size(89, 25)
        Me.Command7.TabIndex = 10
        Me.Command7.Text = "100% of Original"
        Me.Command7.UseVisualStyleBackColor = False
        '
        'Command6
        '
        Me.Command6.BackColor = System.Drawing.SystemColors.Control
        Me.Command6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command6.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command6.Location = New System.Drawing.Point(13, 268)
        Me.Command6.Name = "Command6"
        Me.Command6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command6.Size = New System.Drawing.Size(89, 25)
        Me.Command6.TabIndex = 9
        Me.Command6.Text = "150% of Original"
        Me.Command6.UseVisualStyleBackColor = False
        '
        'Command8
        '
        Me.Command8.BackColor = System.Drawing.SystemColors.Control
        Me.Command8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command8.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command8.Location = New System.Drawing.Point(13, 172)
        Me.Command8.Name = "Command8"
        Me.Command8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command8.Size = New System.Drawing.Size(89, 25)
        Me.Command8.TabIndex = 8
        Me.Command8.Text = "50% of Original"
        Me.Command8.UseVisualStyleBackColor = False
        '
        'Command9
        '
        Me.Command9.BackColor = System.Drawing.SystemColors.Control
        Me.Command9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command9.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command9.Location = New System.Drawing.Point(13, 204)
        Me.Command9.Name = "Command9"
        Me.Command9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command9.Size = New System.Drawing.Size(89, 25)
        Me.Command9.TabIndex = 7
        Me.Command9.Text = "75% of Original"
        Me.Command9.UseVisualStyleBackColor = False
        '
        'Command10
        '
        Me.Command10.BackColor = System.Drawing.SystemColors.Control
        Me.Command10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command10.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command10.Location = New System.Drawing.Point(13, 300)
        Me.Command10.Name = "Command10"
        Me.Command10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command10.Size = New System.Drawing.Size(89, 25)
        Me.Command10.TabIndex = 6
        Me.Command10.Text = "200% of Original"
        Me.Command10.UseVisualStyleBackColor = False
        '
        'Command5
        '
        Me.Command5.BackColor = System.Drawing.SystemColors.Control
        Me.Command5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command5.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command5.Location = New System.Drawing.Point(13, 332)
        Me.Command5.Name = "Command5"
        Me.Command5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command5.Size = New System.Drawing.Size(89, 25)
        Me.Command5.TabIndex = 5
        Me.Command5.Text = "Fit to Window"
        Me.Command5.UseVisualStyleBackColor = False
        '
        'Command20
        '
        Me.Command20.BackColor = System.Drawing.SystemColors.Control
        Me.Command20.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command20.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command20.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command20.Location = New System.Drawing.Point(13, 363)
        Me.Command20.Name = "Command20"
        Me.Command20.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command20.Size = New System.Drawing.Size(89, 25)
        Me.Command20.TabIndex = 4
        Me.Command20.Text = "Aspect Ratio"
        Me.Command20.UseVisualStyleBackColor = False
        '
        'Command3
        '
        Me.Command3.BackColor = System.Drawing.SystemColors.Control
        Me.Command3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command3.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command3.Location = New System.Drawing.Point(13, 76)
        Me.Command3.Name = "Command3"
        Me.Command3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command3.Size = New System.Drawing.Size(89, 25)
        Me.Command3.TabIndex = 3
        Me.Command3.Text = "Rotate 270"
        Me.Command3.UseVisualStyleBackColor = False
        '
        'Command2
        '
        Me.Command2.BackColor = System.Drawing.SystemColors.Control
        Me.Command2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command2.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command2.Location = New System.Drawing.Point(13, 44)
        Me.Command2.Name = "Command2"
        Me.Command2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command2.Size = New System.Drawing.Size(89, 25)
        Me.Command2.TabIndex = 2
        Me.Command2.Text = "Rotate 180"
        Me.Command2.UseVisualStyleBackColor = False
        '
        'Command1
        '
        Me.Command1.BackColor = System.Drawing.SystemColors.Control
        Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command1.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command1.Location = New System.Drawing.Point(13, 12)
        Me.Command1.Name = "Command1"
        Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command1.Size = New System.Drawing.Size(89, 25)
        Me.Command1.TabIndex = 1
        Me.Command1.Text = "Rotate 90"
        Me.Command1.UseVisualStyleBackColor = False
        '
        'Scanner1
        '
        Me.Scanner1.Enabled = True
        Me.Scanner1.Location = New System.Drawing.Point(107, 12)
        Me.Scanner1.Name = "Scanner1"
        Me.Scanner1.OcxState = CType(resources.GetObject("Scanner1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.Scanner1.Size = New System.Drawing.Size(426, 427)
        Me.Scanner1.TabIndex = 0
        '
        'Label34
        '
        Me.Label34.BackColor = System.Drawing.SystemColors.Control
        Me.Label34.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label34.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label34.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label34.Location = New System.Drawing.Point(113, 487)
        Me.Label34.Name = "Label34"
        Me.Label34.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label34.Size = New System.Drawing.Size(297, 49)
        Me.Label34.TabIndex = 114
        Me.Label34.Text = "Any change only valid on current view, you need press this button to update image" &
    " buffer when you have multipage image buffer"
        '
        'Label33
        '
        Me.Label33.BackColor = System.Drawing.SystemColors.Control
        Me.Label33.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label33.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label33.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label33.Location = New System.Drawing.Point(353, 463)
        Me.Label33.Name = "Label33"
        Me.Label33.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label33.Size = New System.Drawing.Size(65, 25)
        Me.Label33.TabIndex = 112
        Me.Label33.Text = "Page No"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(12, 443)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(33, 17)
        Me.Label6.TabIndex = 93
        Me.Label6.Text = "Left"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(12, 467)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(33, 17)
        Me.Label7.TabIndex = 92
        Me.Label7.Text = "Top"
        '
        'Label32
        '
        Me.Label32.BackColor = System.Drawing.SystemColors.Control
        Me.Label32.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label32.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label32.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label32.Location = New System.Drawing.Point(12, 403)
        Me.Label32.Name = "Label32"
        Me.Label32.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label32.Size = New System.Drawing.Size(89, 17)
        Me.Label32.TabIndex = 89
        Me.Label32.Text = "User Define"
        '
        'Label31
        '
        Me.Label31.AutoSize = True
        Me.Label31.BackColor = System.Drawing.SystemColors.Control
        Me.Label31.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label31.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label31.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label31.Location = New System.Drawing.Point(84, 419)
        Me.Label31.Name = "Label31"
        Me.Label31.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label31.Size = New System.Drawing.Size(18, 16)
        Me.Label31.TabIndex = 88
        Me.Label31.Text = "%"
        '
        'lblcurrentpage
        '
        Me.lblcurrentpage.BackColor = System.Drawing.SystemColors.Control
        Me.lblcurrentpage.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblcurrentpage.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblcurrentpage.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblcurrentpage.Location = New System.Drawing.Point(201, 463)
        Me.lblcurrentpage.Name = "lblcurrentpage"
        Me.lblcurrentpage.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblcurrentpage.Size = New System.Drawing.Size(41, 17)
        Me.lblcurrentpage.TabIndex = 31
        Me.lblcurrentpage.Text = "0"
        Me.lblcurrentpage.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(249, 463)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(9, 17)
        Me.Label8.TabIndex = 30
        Me.Label8.Text = "/"
        '
        'lbltotalpage
        '
        Me.lbltotalpage.BackColor = System.Drawing.SystemColors.Control
        Me.lbltotalpage.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbltotalpage.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbltotalpage.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lbltotalpage.Location = New System.Drawing.Point(265, 463)
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
        Me.Label3.Location = New System.Drawing.Point(113, 463)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(73, 17)
        Me.Label3.TabIndex = 28
        Me.Label3.Text = "Total Pages"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(381, 594)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(152, 48)
        Me.Button1.TabIndex = 116
        Me.Button1.Text = "Check Current is Blank page (depend on Confidence value)"
        '
        'Label40
        '
        Me.Label40.Location = New System.Drawing.Point(279, 594)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(96, 24)
        Me.Label40.TabIndex = 117
        Me.Label40.Text = "Confidence(0-100)"
        '
        'txtconfid
        '
        Me.txtconfid.Location = New System.Drawing.Point(279, 618)
        Me.txtconfid.Name = "txtconfid"
        Me.txtconfid.Size = New System.Drawing.Size(100, 20)
        Me.txtconfid.TabIndex = 118
        Me.txtconfid.Text = "99"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(116, 582)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(136, 48)
        Me.Button2.TabIndex = 119
        Me.Button2.Text = "Get Confidence of Current page "
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.Button3)
        Me.GroupBox1.Controls.Add(Me.txtbarcodeheight)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.txtbarcodewidth)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.txtbarcodetop)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txtbarcodeleft)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.chkfullpage)
        Me.GroupBox1.Location = New System.Drawing.Point(539, 245)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(495, 535)
        Me.GroupBox1.TabIndex = 120
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Read barcode for specific area"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.chkscanmore)
        Me.GroupBox2.Controls.Add(Me.txtcustomrotation)
        Me.GroupBox2.Controls.Add(Me.Label12)
        Me.GroupBox2.Controls.Add(Me.chkScan90Rotation)
        Me.GroupBox2.Controls.Add(Me.chkScan45CounterRotation)
        Me.GroupBox2.Controls.Add(Me.chkScan45Rotation)
        Me.GroupBox2.Controls.Add(Me.chkscannorotation)
        Me.GroupBox2.Controls.Add(Me.chkScanMultiple)
        Me.GroupBox2.Location = New System.Drawing.Point(0, 120)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(489, 137)
        Me.GroupBox2.TabIndex = 33
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "UnSelected some option, it will scan more faster"
        '
        'chkscanmore
        '
        Me.chkscanmore.AutoSize = True
        Me.chkscanmore.Checked = True
        Me.chkscanmore.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkscanmore.Location = New System.Drawing.Point(220, 112)
        Me.chkscanmore.Name = "chkscanmore"
        Me.chkscanmore.Size = New System.Drawing.Size(128, 18)
        Me.chkscanmore.TabIndex = 7
        Me.chkscanmore.Text = "Scan More Accuracy"
        Me.chkscanmore.UseVisualStyleBackColor = True
        '
        'txtcustomrotation
        '
        Me.txtcustomrotation.Location = New System.Drawing.Point(322, 86)
        Me.txtcustomrotation.Name = "txtcustomrotation"
        Me.txtcustomrotation.Size = New System.Drawing.Size(106, 20)
        Me.txtcustomrotation.TabIndex = 6
        Me.txtcustomrotation.Text = "0"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(225, 90)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(85, 14)
        Me.Label12.TabIndex = 5
        Me.Label12.Text = "Custom Rotation"
        '
        'chkScan90Rotation
        '
        Me.chkScan90Rotation.AutoSize = True
        Me.chkScan90Rotation.Checked = True
        Me.chkScan90Rotation.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkScan90Rotation.Location = New System.Drawing.Point(220, 58)
        Me.chkScan90Rotation.Name = "chkScan90Rotation"
        Me.chkScan90Rotation.Size = New System.Drawing.Size(170, 18)
        Me.chkScan90Rotation.TabIndex = 4
        Me.chkScan90Rotation.Text = "Scan With 90 Degree Rotation"
        Me.chkScan90Rotation.UseVisualStyleBackColor = True
        '
        'chkScan45CounterRotation
        '
        Me.chkScan45CounterRotation.AutoSize = True
        Me.chkScan45CounterRotation.Checked = True
        Me.chkScan45CounterRotation.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkScan45CounterRotation.Location = New System.Drawing.Point(220, 27)
        Me.chkScan45CounterRotation.Name = "chkScan45CounterRotation"
        Me.chkScan45CounterRotation.Size = New System.Drawing.Size(264, 18)
        Me.chkScan45CounterRotation.TabIndex = 3
        Me.chkScan45CounterRotation.Text = "Scan With 45 Degree Counter Clockwise Rotation"
        Me.chkScan45CounterRotation.UseVisualStyleBackColor = True
        '
        'chkScan45Rotation
        '
        Me.chkScan45Rotation.AutoSize = True
        Me.chkScan45Rotation.Checked = True
        Me.chkScan45Rotation.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkScan45Rotation.Location = New System.Drawing.Point(21, 90)
        Me.chkScan45Rotation.Name = "chkScan45Rotation"
        Me.chkScan45Rotation.Size = New System.Drawing.Size(170, 18)
        Me.chkScan45Rotation.TabIndex = 2
        Me.chkScan45Rotation.Text = "Scan With 45 Degree Rotation"
        Me.chkScan45Rotation.UseVisualStyleBackColor = True
        '
        'chkscannorotation
        '
        Me.chkscannorotation.AutoSize = True
        Me.chkscannorotation.Checked = True
        Me.chkscannorotation.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkscannorotation.Location = New System.Drawing.Point(21, 58)
        Me.chkscannorotation.Name = "chkscannorotation"
        Me.chkscannorotation.Size = New System.Drawing.Size(132, 18)
        Me.chkscannorotation.TabIndex = 1
        Me.chkscannorotation.Text = "Scan Without Rotation"
        Me.chkscannorotation.UseVisualStyleBackColor = True
        '
        'chkScanMultiple
        '
        Me.chkScanMultiple.AutoSize = True
        Me.chkScanMultiple.Checked = True
        Me.chkScanMultiple.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkScanMultiple.Location = New System.Drawing.Point(21, 27)
        Me.chkScanMultiple.Name = "chkScanMultiple"
        Me.chkScanMultiple.Size = New System.Drawing.Size(133, 18)
        Me.chkScanMultiple.TabIndex = 0
        Me.chkScanMultiple.Text = "Scan Multiple Barcode"
        Me.chkScanMultiple.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(6, 269)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(96, 24)
        Me.Button3.TabIndex = 9
        Me.Button3.Text = "Read Barcode"
        '
        'txtbarcodeheight
        '
        Me.txtbarcodeheight.Enabled = False
        Me.txtbarcodeheight.Location = New System.Drawing.Point(320, 64)
        Me.txtbarcodeheight.Name = "txtbarcodeheight"
        Me.txtbarcodeheight.Size = New System.Drawing.Size(32, 20)
        Me.txtbarcodeheight.TabIndex = 8
        Me.txtbarcodeheight.Text = "100"
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(264, 64)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(40, 16)
        Me.Label11.TabIndex = 7
        Me.Label11.Text = "Height"
        '
        'txtbarcodewidth
        '
        Me.txtbarcodewidth.Enabled = False
        Me.txtbarcodewidth.Location = New System.Drawing.Point(224, 64)
        Me.txtbarcodewidth.Name = "txtbarcodewidth"
        Me.txtbarcodewidth.Size = New System.Drawing.Size(32, 20)
        Me.txtbarcodewidth.TabIndex = 6
        Me.txtbarcodewidth.Text = "100"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(192, 64)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(32, 16)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "Width"
        '
        'txtbarcodetop
        '
        Me.txtbarcodetop.Enabled = False
        Me.txtbarcodetop.Location = New System.Drawing.Point(144, 64)
        Me.txtbarcodetop.Name = "txtbarcodetop"
        Me.txtbarcodetop.Size = New System.Drawing.Size(32, 20)
        Me.txtbarcodetop.TabIndex = 4
        Me.txtbarcodetop.Text = "0"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(112, 64)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(32, 16)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Top"
        '
        'txtbarcodeleft
        '
        Me.txtbarcodeleft.Enabled = False
        Me.txtbarcodeleft.Location = New System.Drawing.Point(64, 64)
        Me.txtbarcodeleft.Name = "txtbarcodeleft"
        Me.txtbarcodeleft.Size = New System.Drawing.Size(32, 20)
        Me.txtbarcodeleft.TabIndex = 2
        Me.txtbarcodeleft.Text = "0"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(32, 16)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Left"
        '
        'chkfullpage
        '
        Me.chkfullpage.Checked = True
        Me.chkfullpage.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkfullpage.Location = New System.Drawing.Point(16, 32)
        Me.chkfullpage.Name = "chkfullpage"
        Me.chkfullpage.Size = New System.Drawing.Size(80, 16)
        Me.chkfullpage.TabIndex = 0
        Me.chkfullpage.Text = "Full Page"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(1025, 768)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.txtconfid)
        Me.Controls.Add(Me.Label40)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Command17)
        Me.Controls.Add(Me.txtLeft)
        Me.Controls.Add(Me.txtTop)
        Me.Controls.Add(Me.Text1)
        Me.Controls.Add(Me.Text2)
        Me.Controls.Add(Me.Label31)
        Me.Controls.Add(Me.Command33)
        Me.Controls.Add(Me.Command31)
        Me.Controls.Add(Me._Frame1_1)
        Me.Controls.Add(Me.Command12)
        Me.Controls.Add(Me.Command11)
        Me.Controls.Add(Me.Command7)
        Me.Controls.Add(Me.Command6)
        Me.Controls.Add(Me.Command8)
        Me.Controls.Add(Me.Command9)
        Me.Controls.Add(Me.Command10)
        Me.Controls.Add(Me.Command5)
        Me.Controls.Add(Me.Command20)
        Me.Controls.Add(Me.Command3)
        Me.Controls.Add(Me.Command2)
        Me.Controls.Add(Me.Command1)
        Me.Controls.Add(Me.Scanner1)
        Me.Controls.Add(Me.Label34)
        Me.Controls.Add(Me.Label33)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label32)
        Me.Controls.Add(Me.lblcurrentpage)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.lbltotalpage)
        Me.Controls.Add(Me.Label3)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 30)
        Me.Name = "Form1"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Read Barcode from Scanned Image"
        Me._Frame1_1.ResumeLayout(False)
        Me.Frame5.ResumeLayout(False)
        Me.Frame5.PerformLayout()
        CType(Me.Scanner1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
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
        Set(ByVal value As Form1)
            m_vb6FormDefInstance = value
        End Set
    End Property
#End Region
    Dim iX As Short
    Dim iY As Short

 

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

  


    Private Sub cbohashbrushstyle_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

    End Sub

    Private Sub cbotextstyle_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
    End Sub

    Private Sub Check1_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)



    End Sub

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

     Private Sub chkdefaultcap_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkdefaultcap.CheckStateChanged




    End Sub

    Private Sub chkshowtext_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
    End Sub

    Private Sub cmdhashbackcolor_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

    End Sub

    Private Sub cmdhashforecolor_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

    End Sub

    Private Sub cmdoutlineback_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

    End Sub

    Private Sub cmdoutlineborder_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

    End Sub

    Private Sub cmdtextcolor_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)


    End Sub


    Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
        Me.Scanner1.Rotate90()
    End Sub

    Private Sub Command10_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command10.Click

        Scanner1.View = 7
        Scanner1.Focus()

    End Sub

    Private Sub Command11_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command11.Click
        Scanner1.View = 2
        Scanner1.Focus()
    End Sub

    Private Sub Command12_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command12.Click

        Scanner1.View = 1
        Scanner1.Focus()

    End Sub

    Private Sub Command13_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        If ColorDialog1.ShowDialog() Then

            Me.Scanner1.SetBackgroundColor(Color2Uint32(ColorDialog1.Color))

        End If

    End Sub

    Private Sub Command14_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Scanner1.PrintImage(True)
    End Sub

    Private Sub Command15_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
      


    End Sub

    Private Sub Command16_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
       



    End Sub

    Private Sub Command17_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command17.Click
        Scanner1.DeletePage(Scanner1.GetActivePageNo)


        Scanner1.SetActivePageNo(1)
        lbltotalpage.Text = CStr(Me.Scanner1.TotalPage)
        lblcurrentpage.Text = CStr(Me.Scanner1.GetActivePageNo)

    End Sub

    Private Sub Command18_Click()
       


    End Sub

    Private Sub Command19_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
       

    End Sub

    Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
        Me.Scanner1.Rotate180()
    End Sub

    Private Sub Command20_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command20.Click

        Scanner1.View = 10
        Scanner1.Focus()

    End Sub

    Private Sub Command21_Click()
       

    End Sub

    Private Sub Command22_Click()
       

    End Sub

    Private Sub Command23_Click()


        Me.Scanner1.FileLeft = CInt(txtLeft.Text)
        Me.Scanner1.FileTop = CInt(txtTop.Text)

        Me.Scanner1.Redraw()

    End Sub

    Private Sub Command24_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)




    End Sub

    Private Sub Command25_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

    End Sub

    Private Sub Command26_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

    End Sub

    Private Sub Command27_Click()

        Scanner1.View = 8
        Scanner1.ViewSize = CDbl(Str(CDbl(Text1.Text)))

    End Sub

    Private Sub Command28_Click()
        Me.Scanner1.SetScrollBarPos(iX, iY)


        Me.Scanner1.Focus()

    End Sub

    Private Sub Command29_Click()

        iX = Me.Scanner1.GetHorzScrollBarPos
        iY = Me.Scanner1.GetVertScrollBarPos

        Me.Scanner1.Focus()

    End Sub

    Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
        Me.Scanner1.Rotate270()
    End Sub


    Private Sub Command31_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command31.Click
        Scanner1.SetActivePageNo(CShort(Text2.Text))
        lblcurrentpage.Text = CStr(Me.Scanner1.GetActivePageNo)



    End Sub

    Private Sub Command32_Click()
       

    End Sub

    Private Sub Command33_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command33.Click
        Me.Scanner1.ApplyChange()
    End Sub

    Private Sub Command34_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
    End Sub

    Private Sub Command35_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
    End Sub

    Private Sub Command36_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
    End Sub

    Private Sub Command37_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
    End Sub

    Private Sub Command38_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
    End Sub

    Private Sub Command39_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

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

        'do not resize the image, it may affect the barcode quality
        Scanner1.BufferResizeMode = False

        Scanner1.Scan()




    End Sub

    Private Sub Command40_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
       

    End Sub

    Private Sub Command5_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command5.Click

        Scanner1.View = 9
        Scanner1.Focus()

    End Sub

    Private Sub Command6_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command6.Click

        Scanner1.View = 6
        Scanner1.Focus()

    End Sub

    Private Sub Command7_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command7.Click
        Scanner1.View = 5
        Scanner1.Focus()


    End Sub

    Private Sub Command8_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command8.Click
        Scanner1.View = 3
        Scanner1.Focus()


    End Sub

    Private Sub Command9_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command9.Click

        Scanner1.View = 4
        Scanner1.Focus()

    End Sub

    Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim I As Integer
        Dim iCount As Integer



        cbodpi.Items.Add("Onscreen Viewing 96dpi")
        cbodpi.Items.Add("Fax 200dpi")
        cbodpi.Items.Add("OCR Text 300dpi")
        cbodpi.Items.Add("Laser Print Fine 600dpi")
        cbodpi.SelectedIndex = 1

        cbopixeltype.Items.Add("Default")
        cbopixeltype.Items.Add("Gray Color")
        cbopixeltype.Items.Add("Black & White Color")
        cbopixeltype.Items.Add("True Color")

        cbopixeltype.SelectedIndex = 0



        iCount = Me.Scanner1.GetNumImageSources

        For I = 0 To iCount - 1
            cboimagesource.Items.Add(Scanner1.GetImageSourceName(I))
        Next

        If iCount > 0 Then
            cboimagesource.SelectedIndex = 0
        End If






    End Sub

    Private Sub optborderoff_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

    End Sub

    Private Sub optborderon_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)



    End Sub

    Private Sub Option3_Click()

    End Sub

    Private Sub opthighno_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

    End Sub

    Private Sub opthighyes_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

    End Sub

    Private Sub Scanner1_EndScan(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Scanner1.EndScan


        lbltotalpage.Text = CStr(Me.Scanner1.TotalPage)
        lblcurrentpage.Text = CStr(Me.Scanner1.GetActivePageNo)



    End Sub



    Private Sub Text1_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles Text1.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = 13 Then
            Scanner1.View = 8
            Scanner1.ViewSize = Val(CStr(CDbl(Text1.Text) / 100))
        End If
    End Sub

    Private Sub txtLeft_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtLeft.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = 13 Then
            Me.Scanner1.FileLeft = CInt(txtLeft.Text)
            Me.Scanner1.FileTop = CInt(txtTop.Text)
            Me.Scanner1.Redraw()
        End If
    End Sub

    Private Sub txtTop_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtTop.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = 13 Then
            Me.Scanner1.FileLeft = CInt(txtLeft.Text)
            Me.Scanner1.FileTop = CInt(txtTop.Text)
            Me.Scanner1.Redraw()
        End If
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

    Private Sub Check1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
       
    End Sub

    Private Sub Check3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub opthighyes_Click(ByVal sender As Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub optborderon_Click(ByVal sender As Object, ByVal e As System.EventArgs)
     

    End Sub

    Private Sub optborderoff_Click(ByVal sender As Object, ByVal e As System.EventArgs)
       

    End Sub

    Private Sub opthighno_Click(ByVal sender As Object, ByVal e As System.EventArgs)
       

    End Sub



    Public Function Color2Uint32(ByVal clr As Color) As UInt32
        Dim t As Int32
        Dim a() As Byte

        t = ColorTranslator.ToOle(clr)


        a = BitConverter.GetBytes(t)

        Return BitConverter.ToUInt32(a, 0)


    End Function




    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        txtconfid.Text = Scanner1.BlankPageGetConfidence

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Scanner1.BlankPageIsBlank(CDbl(txtconfid.Text)) Then
            MessageBox.Show("This page is Blank page")
        Else
            MessageBox.Show("This page is not a Blank page")
        End If

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim ibarcodeCount As Integer
        Dim str1 As String
        Dim strTmp As String
        Dim i As Integer

        Me.Scanner1.BarCodeReadScanMultiple = chkScanMultiple.Checked
        Me.Scanner1.BarCodeReadScanWithoutRotation = chkscannorotation.Checked
        Me.Scanner1.BarCodeReadScan45Rotation = chkScan45Rotation.Checked
        Me.Scanner1.BarCodeReadScan45CouterRotation = chkScan45CounterRotation.Checked
        Me.Scanner1.BarCodeReadScan90Rotation = chkScan90Rotation.Checked
        Me.Scanner1.BarCodeReadCustomRotation = txtcustomrotation.Text
        Me.Scanner1.BarCodeReadScanAccuracy = chkscanmore.Checked

        If chkfullpage.Checked Then
            ibarcodeCount = Me.Scanner1.BarCodeReadFullPage

        Else
            ibarcodeCount = Me.Scanner1.BarCodeReadByZone(CInt(txtbarcodeleft.Text), CInt(txtbarcodetop.Text), CInt(txtbarcodewidth.Text), CInt(txtbarcodeheight.Text))

        End If

        If ibarcodeCount < 1 Then

            MessageBox.Show("No barcode found, try scan to high DPI or set the specific area for detect the barcode")
            Exit Sub
        End If


        str1 = "Total " + Str(ibarcodeCount) + " BarCode detected" + Chr(13) + Chr(10)

        For i = 0 To ibarcodeCount - 1

            strTmp = Me.Scanner1.BarCodeGetType(i) + " score:" + Str(Me.Scanner1.BarCodeGetScore(i)) + " value:" + Me.Scanner1.BarCodeGetValue(i)

            str1 = str1 + Chr(13) + Chr(10) + strTmp
        Next

        MessageBox.Show(str1)

    End Sub
End Class