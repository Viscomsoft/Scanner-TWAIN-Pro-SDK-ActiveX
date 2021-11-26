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
	Public WithEvents Check3 As System.Windows.Forms.CheckBox
	Public WithEvents Command16 As System.Windows.Forms.Button
	Public WithEvents Command15 As System.Windows.Forms.Button
	Public WithEvents Frame4 As System.Windows.Forms.GroupBox
	Public WithEvents Command13 As System.Windows.Forms.Button
	Public WithEvents Command14 As System.Windows.Forms.Button
    Public WithEvents Command24 As System.Windows.Forms.Button
	Public WithEvents txt_Height As System.Windows.Forms.TextBox
	Public WithEvents txt_Width As System.Windows.Forms.TextBox
	Public WithEvents Command19 As System.Windows.Forms.Button
	Public WithEvents Check1 As System.Windows.Forms.CheckBox
	Public WithEvents cboutputtype As System.Windows.Forms.ComboBox
	Public WithEvents Command40 As System.Windows.Forms.Button
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Frame3 As System.Windows.Forms.GroupBox
	Public WithEvents txtLeft As System.Windows.Forms.TextBox
	Public WithEvents txtTop As System.Windows.Forms.TextBox
	Public WithEvents Text1 As System.Windows.Forms.TextBox
    Public WithEvents Command33 As System.Windows.Forms.Button
    Public WithEvents Text2 As System.Windows.Forms.TextBox
    Public WithEvents Command31 As System.Windows.Forms.Button
    Public WithEvents opthighyes As System.Windows.Forms.RadioButton
    Public WithEvents opthighno As System.Windows.Forms.RadioButton
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents optborderon As System.Windows.Forms.RadioButton
    Public WithEvents optborderoff As System.Windows.Forms.RadioButton
    Public WithEvents _Frame1_0 As System.Windows.Forms.GroupBox
    Public WithEvents Picture1 As System.Windows.Forms.PictureBox
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
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents cbotifcompression As System.Windows.Forms.ComboBox
    Friend WithEvents ColorDialog1 As System.Windows.Forms.ColorDialog
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents txtconfid As System.Windows.Forms.TextBox
    Public WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents chkpanning As System.Windows.Forms.CheckBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Command17 = New System.Windows.Forms.Button()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.cbotifcompression = New System.Windows.Forms.ComboBox()
        Me.Label39 = New System.Windows.Forms.Label()
        Me.Check3 = New System.Windows.Forms.CheckBox()
        Me.Command16 = New System.Windows.Forms.Button()
        Me.Command15 = New System.Windows.Forms.Button()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.Command13 = New System.Windows.Forms.Button()
        Me.Command14 = New System.Windows.Forms.Button()
        Me.Command24 = New System.Windows.Forms.Button()
        Me.txt_Height = New System.Windows.Forms.TextBox()
        Me.txt_Width = New System.Windows.Forms.TextBox()
        Me.Command19 = New System.Windows.Forms.Button()
        Me.Check1 = New System.Windows.Forms.CheckBox()
        Me.cboutputtype = New System.Windows.Forms.ComboBox()
        Me.Command40 = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtLeft = New System.Windows.Forms.TextBox()
        Me.txtTop = New System.Windows.Forms.TextBox()
        Me.Text1 = New System.Windows.Forms.TextBox()
        Me.Command33 = New System.Windows.Forms.Button()
        Me.Text2 = New System.Windows.Forms.TextBox()
        Me.Command31 = New System.Windows.Forms.Button()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.opthighyes = New System.Windows.Forms.RadioButton()
        Me.opthighno = New System.Windows.Forms.RadioButton()
        Me._Frame1_0 = New System.Windows.Forms.GroupBox()
        Me.optborderon = New System.Windows.Forms.RadioButton()
        Me.optborderoff = New System.Windows.Forms.RadioButton()
        Me.Picture1 = New System.Windows.Forms.PictureBox()
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
        Me.chkpanning = New System.Windows.Forms.CheckBox()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Frame4.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me._Frame1_0.SuspendLayout()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me._Frame1_1.SuspendLayout()
        Me.Frame5.SuspendLayout()
        CType(Me.Scanner1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Command17
        '
        Me.Command17.BackColor = System.Drawing.SystemColors.Control
        Me.Command17.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command17.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command17.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command17.Location = New System.Drawing.Point(408, 624)
        Me.Command17.Name = "Command17"
        Me.Command17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command17.Size = New System.Drawing.Size(145, 41)
        Me.Command17.TabIndex = 115
        Me.Command17.Text = "Delete Current Page"
        Me.Command17.UseVisualStyleBackColor = False
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.Button3)
        Me.Frame4.Controls.Add(Me.cbotifcompression)
        Me.Frame4.Controls.Add(Me.Label39)
        Me.Frame4.Controls.Add(Me.Check3)
        Me.Frame4.Controls.Add(Me.Command16)
        Me.Frame4.Controls.Add(Me.Command15)
        Me.Frame4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame4.Location = New System.Drawing.Point(560, 264)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(217, 202)
        Me.Frame4.TabIndex = 109
        Me.Frame4.TabStop = False
        Me.Frame4.Text = "Image Buffers (Total Pages)"
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.SystemColors.Control
        Me.Button3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Button3.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Button3.Location = New System.Drawing.Point(12, 44)
        Me.Button3.Name = "Button3"
        Me.Button3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Button3.Size = New System.Drawing.Size(185, 33)
        Me.Button3.TabIndex = 116
        Me.Button3.Text = "Save Image Buffers to Multi Page Docx"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'cbotifcompression
        '
        Me.cbotifcompression.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbotifcompression.Location = New System.Drawing.Point(112, 85)
        Me.cbotifcompression.Name = "cbotifcompression"
        Me.cbotifcompression.Size = New System.Drawing.Size(96, 22)
        Me.cbotifcompression.TabIndex = 115
        '
        'Label39
        '
        Me.Label39.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label39.Location = New System.Drawing.Point(8, 85)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(96, 16)
        Me.Label39.TabIndex = 114
        Me.Label39.Text = "TIF Compression"
        '
        'Check3
        '
        Me.Check3.BackColor = System.Drawing.SystemColors.Control
        Me.Check3.Checked = True
        Me.Check3.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Check3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Check3.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Check3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Check3.Location = New System.Drawing.Point(16, 24)
        Me.Check3.Name = "Check3"
        Me.Check3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Check3.Size = New System.Drawing.Size(129, 17)
        Me.Check3.TabIndex = 113
        Me.Check3.Text = "Save to Single File"
        Me.Check3.UseVisualStyleBackColor = False
        '
        'Command16
        '
        Me.Command16.BackColor = System.Drawing.SystemColors.Control
        Me.Command16.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command16.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command16.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command16.Location = New System.Drawing.Point(8, 165)
        Me.Command16.Name = "Command16"
        Me.Command16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command16.Size = New System.Drawing.Size(185, 33)
        Me.Command16.TabIndex = 111
        Me.Command16.Text = "Save Image Buffers to Multi Page PDF"
        Me.Command16.UseVisualStyleBackColor = False
        '
        'Command15
        '
        Me.Command15.BackColor = System.Drawing.SystemColors.Control
        Me.Command15.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command15.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command15.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command15.Location = New System.Drawing.Point(8, 125)
        Me.Command15.Name = "Command15"
        Me.Command15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command15.Size = New System.Drawing.Size(185, 33)
        Me.Command15.TabIndex = 110
        Me.Command15.Text = "Save Image Buffers to Multi Page TIF"
        Me.Command15.UseVisualStyleBackColor = False
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.Command13)
        Me.Frame3.Controls.Add(Me.Command14)
        Me.Frame3.Controls.Add(Me.Command24)
        Me.Frame3.Controls.Add(Me.txt_Height)
        Me.Frame3.Controls.Add(Me.txt_Width)
        Me.Frame3.Controls.Add(Me.Command19)
        Me.Frame3.Controls.Add(Me.Check1)
        Me.Frame3.Controls.Add(Me.cboutputtype)
        Me.Frame3.Controls.Add(Me.Command40)
        Me.Frame3.Controls.Add(Me.Label5)
        Me.Frame3.Controls.Add(Me.Label4)
        Me.Frame3.Controls.Add(Me.Label2)
        Me.Frame3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame3.Location = New System.Drawing.Point(787, 272)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(217, 248)
        Me.Frame3.TabIndex = 94
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Current View"
        '
        'Command13
        '
        Me.Command13.BackColor = System.Drawing.SystemColors.Control
        Me.Command13.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command13.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command13.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command13.Location = New System.Drawing.Point(120, 160)
        Me.Command13.Name = "Command13"
        Me.Command13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command13.Size = New System.Drawing.Size(97, 25)
        Me.Command13.TabIndex = 108
        Me.Command13.Text = "Background Color"
        Me.Command13.UseVisualStyleBackColor = False
        '
        'Command14
        '
        Me.Command14.BackColor = System.Drawing.SystemColors.Control
        Me.Command14.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command14.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command14.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command14.Location = New System.Drawing.Point(120, 192)
        Me.Command14.Name = "Command14"
        Me.Command14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command14.Size = New System.Drawing.Size(97, 25)
        Me.Command14.TabIndex = 107
        Me.Command14.Text = "Print Image"
        Me.Command14.UseVisualStyleBackColor = False
        '
        'Command24
        '
        Me.Command24.BackColor = System.Drawing.SystemColors.Control
        Me.Command24.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command24.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command24.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command24.Location = New System.Drawing.Point(16, 160)
        Me.Command24.Name = "Command24"
        Me.Command24.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command24.Size = New System.Drawing.Size(97, 25)
        Me.Command24.TabIndex = 104
        Me.Command24.Text = "Save to Clipboard"
        Me.Command24.UseVisualStyleBackColor = False
        '
        'txt_Height
        '
        Me.txt_Height.AcceptsReturn = True
        Me.txt_Height.BackColor = System.Drawing.SystemColors.Window
        Me.txt_Height.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txt_Height.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_Height.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txt_Height.Location = New System.Drawing.Point(64, 120)
        Me.txt_Height.MaxLength = 0
        Me.txt_Height.Name = "txt_Height"
        Me.txt_Height.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txt_Height.Size = New System.Drawing.Size(49, 20)
        Me.txt_Height.TabIndex = 101
        '
        'txt_Width
        '
        Me.txt_Width.AcceptsReturn = True
        Me.txt_Width.BackColor = System.Drawing.SystemColors.Window
        Me.txt_Width.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txt_Width.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_Width.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txt_Width.Location = New System.Drawing.Point(64, 88)
        Me.txt_Width.MaxLength = 0
        Me.txt_Width.Name = "txt_Width"
        Me.txt_Width.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txt_Width.Size = New System.Drawing.Size(49, 20)
        Me.txt_Width.TabIndex = 100
        '
        'Command19
        '
        Me.Command19.BackColor = System.Drawing.SystemColors.Control
        Me.Command19.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command19.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command19.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command19.Location = New System.Drawing.Point(120, 120)
        Me.Command19.Name = "Command19"
        Me.Command19.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command19.Size = New System.Drawing.Size(89, 25)
        Me.Command19.TabIndex = 99
        Me.Command19.Text = "SaveBySize"
        Me.Command19.UseVisualStyleBackColor = False
        '
        'Check1
        '
        Me.Check1.BackColor = System.Drawing.SystemColors.Control
        Me.Check1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Check1.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Check1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Check1.Location = New System.Drawing.Point(8, 56)
        Me.Check1.Name = "Check1"
        Me.Check1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Check1.Size = New System.Drawing.Size(105, 25)
        Me.Check1.TabIndex = 98
        Me.Check1.Text = "Output GrayScale"
        Me.Check1.UseVisualStyleBackColor = False
        '
        'cboutputtype
        '
        Me.cboutputtype.BackColor = System.Drawing.SystemColors.Window
        Me.cboutputtype.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboutputtype.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboutputtype.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboutputtype.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboutputtype.Location = New System.Drawing.Point(64, 24)
        Me.cboutputtype.Name = "cboutputtype"
        Me.cboutputtype.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboutputtype.Size = New System.Drawing.Size(73, 22)
        Me.cboutputtype.TabIndex = 96
        '
        'Command40
        '
        Me.Command40.BackColor = System.Drawing.SystemColors.Control
        Me.Command40.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command40.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command40.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command40.Location = New System.Drawing.Point(120, 56)
        Me.Command40.Name = "Command40"
        Me.Command40.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command40.Size = New System.Drawing.Size(89, 25)
        Me.Command40.TabIndex = 95
        Me.Command40.Text = "Save"
        Me.Command40.UseVisualStyleBackColor = False
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(8, 120)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(41, 17)
        Me.Label5.TabIndex = 103
        Me.Label5.Text = "Height"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(8, 88)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(41, 17)
        Me.Label4.TabIndex = 102
        Me.Label4.Text = "Width"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(8, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(49, 17)
        Me.Label2.TabIndex = 97
        Me.Label2.Text = "Save To"
        '
        'txtLeft
        '
        Me.txtLeft.AcceptsReturn = True
        Me.txtLeft.BackColor = System.Drawing.SystemColors.Window
        Me.txtLeft.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLeft.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLeft.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLeft.Location = New System.Drawing.Point(40, 600)
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
        Me.txtTop.Location = New System.Drawing.Point(40, 624)
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
        Me.Text1.Location = New System.Drawing.Point(8, 576)
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
        Me.Command33.Location = New System.Drawing.Point(112, 624)
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
        Me.Text2.Location = New System.Drawing.Point(424, 552)
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
        Me.Command31.Location = New System.Drawing.Point(480, 552)
        Me.Command31.Name = "Command31"
        Me.Command31.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command31.Size = New System.Drawing.Size(41, 25)
        Me.Command31.TabIndex = 25
        Me.Command31.Text = "Go"
        Me.Command31.UseVisualStyleBackColor = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.opthighyes)
        Me.Frame2.Controls.Add(Me.opthighno)
        Me.Frame2.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(672, 472)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(105, 49)
        Me.Frame2.TabIndex = 21
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "High Quality"
        '
        'opthighyes
        '
        Me.opthighyes.BackColor = System.Drawing.SystemColors.Control
        Me.opthighyes.Checked = True
        Me.opthighyes.Cursor = System.Windows.Forms.Cursors.Default
        Me.opthighyes.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.opthighyes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.opthighyes.Location = New System.Drawing.Point(8, 16)
        Me.opthighyes.Name = "opthighyes"
        Me.opthighyes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.opthighyes.Size = New System.Drawing.Size(41, 25)
        Me.opthighyes.TabIndex = 23
        Me.opthighyes.TabStop = True
        Me.opthighyes.Text = "Yes"
        Me.opthighyes.UseVisualStyleBackColor = False
        '
        'opthighno
        '
        Me.opthighno.BackColor = System.Drawing.SystemColors.Control
        Me.opthighno.Cursor = System.Windows.Forms.Cursors.Default
        Me.opthighno.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.opthighno.ForeColor = System.Drawing.SystemColors.ControlText
        Me.opthighno.Location = New System.Drawing.Point(56, 16)
        Me.opthighno.Name = "opthighno"
        Me.opthighno.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.opthighno.Size = New System.Drawing.Size(41, 25)
        Me.opthighno.TabIndex = 22
        Me.opthighno.TabStop = True
        Me.opthighno.Text = "No"
        Me.opthighno.UseVisualStyleBackColor = False
        '
        '_Frame1_0
        '
        Me._Frame1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Frame1_0.Controls.Add(Me.optborderon)
        Me._Frame1_0.Controls.Add(Me.optborderoff)
        Me._Frame1_0.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Frame1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Frame1_0.Location = New System.Drawing.Point(568, 472)
        Me._Frame1_0.Name = "_Frame1_0"
        Me._Frame1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame1_0.Size = New System.Drawing.Size(97, 49)
        Me._Frame1_0.TabIndex = 18
        Me._Frame1_0.TabStop = False
        Me._Frame1_0.Text = "Border"
        '
        'optborderon
        '
        Me.optborderon.BackColor = System.Drawing.SystemColors.Control
        Me.optborderon.Checked = True
        Me.optborderon.Cursor = System.Windows.Forms.Cursors.Default
        Me.optborderon.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optborderon.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optborderon.Location = New System.Drawing.Point(8, 16)
        Me.optborderon.Name = "optborderon"
        Me.optborderon.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optborderon.Size = New System.Drawing.Size(41, 25)
        Me.optborderon.TabIndex = 20
        Me.optborderon.TabStop = True
        Me.optborderon.Text = "On"
        Me.optborderon.UseVisualStyleBackColor = False
        '
        'optborderoff
        '
        Me.optborderoff.BackColor = System.Drawing.SystemColors.Control
        Me.optborderoff.Cursor = System.Windows.Forms.Cursors.Default
        Me.optborderoff.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optborderoff.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optborderoff.Location = New System.Drawing.Point(48, 16)
        Me.optborderoff.Name = "optborderoff"
        Me.optborderoff.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optborderoff.Size = New System.Drawing.Size(41, 25)
        Me.optborderoff.TabIndex = 19
        Me.optborderoff.TabStop = True
        Me.optborderoff.Text = "Off"
        Me.optborderoff.UseVisualStyleBackColor = False
        '
        'Picture1
        '
        Me.Picture1.BackColor = System.Drawing.SystemColors.Control
        Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Picture1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Picture1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Picture1.Location = New System.Drawing.Point(560, 528)
        Me.Picture1.Name = "Picture1"
        Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Picture1.Size = New System.Drawing.Size(145, 121)
        Me.Picture1.TabIndex = 17
        Me.Picture1.TabStop = False
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
        Me.Command4.Enabled = False
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
        Me.Command12.Location = New System.Drawing.Point(11, 235)
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
        Me.Command11.Location = New System.Drawing.Point(11, 267)
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
        Me.Command7.Location = New System.Drawing.Point(11, 363)
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
        Me.Command6.Location = New System.Drawing.Point(11, 395)
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
        Me.Command8.Location = New System.Drawing.Point(11, 299)
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
        Me.Command9.Location = New System.Drawing.Point(11, 331)
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
        Me.Command10.Location = New System.Drawing.Point(11, 427)
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
        Me.Command5.Location = New System.Drawing.Point(11, 459)
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
        Me.Command20.Location = New System.Drawing.Point(11, 491)
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
        Me.Command3.Location = New System.Drawing.Point(11, 203)
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
        Me.Command2.Location = New System.Drawing.Point(11, 171)
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
        Me.Command1.Location = New System.Drawing.Point(11, 139)
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
        Me.Scanner1.Location = New System.Drawing.Point(112, 117)
        Me.Scanner1.Name = "Scanner1"
        Me.Scanner1.OcxState = CType(resources.GetObject("Scanner1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.Scanner1.Size = New System.Drawing.Size(425, 427)
        Me.Scanner1.TabIndex = 0
        '
        'Label34
        '
        Me.Label34.BackColor = System.Drawing.SystemColors.Control
        Me.Label34.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label34.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label34.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label34.Location = New System.Drawing.Point(112, 576)
        Me.Label34.Name = "Label34"
        Me.Label34.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label34.Size = New System.Drawing.Size(297, 49)
        Me.Label34.TabIndex = 114
        Me.Label34.Text = "Any change only valid on current view, you need press this button to update image" & _
            " buffer when you have multipage image buffer"
        '
        'Label33
        '
        Me.Label33.BackColor = System.Drawing.SystemColors.Control
        Me.Label33.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label33.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label33.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label33.Location = New System.Drawing.Point(352, 552)
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
        Me.Label6.Location = New System.Drawing.Point(8, 600)
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
        Me.Label7.Location = New System.Drawing.Point(8, 624)
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
        Me.Label32.Location = New System.Drawing.Point(8, 560)
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
        Me.Label31.Location = New System.Drawing.Point(80, 576)
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
        Me.lblcurrentpage.Location = New System.Drawing.Point(200, 552)
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
        Me.Label8.Location = New System.Drawing.Point(248, 552)
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
        Me.lbltotalpage.Location = New System.Drawing.Point(264, 552)
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
        Me.Label3.Location = New System.Drawing.Point(112, 552)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(73, 17)
        Me.Label3.TabIndex = 28
        Me.Label3.Text = "Total Pages"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(408, 680)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(152, 48)
        Me.Button1.TabIndex = 116
        Me.Button1.Text = "Check Current is Blank page (depend on Confidence value)"
        '
        'Label40
        '
        Me.Label40.Location = New System.Drawing.Point(280, 680)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(96, 24)
        Me.Label40.TabIndex = 117
        Me.Label40.Text = "Confidence(0-100)"
        '
        'txtconfid
        '
        Me.txtconfid.Location = New System.Drawing.Point(280, 704)
        Me.txtconfid.Name = "txtconfid"
        Me.txtconfid.Size = New System.Drawing.Size(100, 20)
        Me.txtconfid.TabIndex = 118
        Me.txtconfid.Text = "99"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(112, 680)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(136, 48)
        Me.Button2.TabIndex = 119
        Me.Button2.Text = "Get Confidence of Current page "
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.chkpanning)
        Me.GroupBox1.Controls.Add(Me.Button4)
        Me.GroupBox1.Controls.Add(Me.TextBox1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(521, 99)
        Me.GroupBox1.TabIndex = 120
        Me.GroupBox1.TabStop = False
        '
        'chkpanning
        '
        Me.chkpanning.AutoSize = True
        Me.chkpanning.Checked = True
        Me.chkpanning.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkpanning.Location = New System.Drawing.Point(154, 51)
        Me.chkpanning.Name = "chkpanning"
        Me.chkpanning.Size = New System.Drawing.Size(158, 18)
        Me.chkpanning.TabIndex = 2
        Me.chkpanning.Text = "Enable Panning with Mouse"
        Me.chkpanning.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.Location = New System.Drawing.Point(343, 50)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(165, 42)
        Me.Button4.TabIndex = 1
        Me.Button4.Text = "Select Image"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(18, 18)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(465, 20)
        Me.TextBox1.TabIndex = 0
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(1016, 731)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.txtconfid)
        Me.Controls.Add(Me.Label40)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Command17)
        Me.Controls.Add(Me.Frame4)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.txtLeft)
        Me.Controls.Add(Me.txtTop)
        Me.Controls.Add(Me.Text1)
        Me.Controls.Add(Me.Text2)
        Me.Controls.Add(Me.Label31)
        Me.Controls.Add(Me.Command33)
        Me.Controls.Add(Me.Command31)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me._Frame1_0)
        Me.Controls.Add(Me.Picture1)
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
        Me.Text = "Scan image to existing Image, PDF or TIFF (Load PDF, XPS need purchase Advanced P" & _
            "DF Viewer plugin)"
        Me.Frame4.ResumeLayout(False)
        Me.Frame3.ResumeLayout(False)
        Me.Frame3.PerformLayout()
        Me.Frame2.ResumeLayout(False)
        Me._Frame1_0.ResumeLayout(False)
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).EndInit()
        Me._Frame1_1.ResumeLayout(False)
        Me.Frame5.ResumeLayout(False)
        Me.Frame5.PerformLayout()
        CType(Me.Scanner1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
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

    Dim clrHashForeColor As Object
    Dim clrHashBackColor As Object

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




    Private Sub cbohashbrushstyle_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

    End Sub

    'UPGRADE_WARNING: Event cbotextstyle.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
    Private Sub cbotextstyle_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
    End Sub

    'UPGRADE_WARNING: Event Check1.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
    Private Sub Check1_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Check1.CheckStateChanged



    End Sub

    'UPGRADE_WARNING: Event Check2.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
    Private Sub Check2_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Check2.CheckStateChanged
        If Check2.CheckState = 1 Then
            Scanner1.ShowTwainUI = True
        Else
            Scanner1.ShowTwainUI = False
        End If


    End Sub

    'UPGRADE_WARNING: Event Check4.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
    Private Sub Check4_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Check4.CheckStateChanged
        If Check4.CheckState = 1 Then
            Me.Scanner1.ClearImageBuffer = True
        Else
            Me.Scanner1.ClearImageBuffer = False

        End If




    End Sub

    'UPGRADE_WARNING: Event chkdefaultcap.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
    Private Sub chkdefaultcap_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkdefaultcap.CheckStateChanged




    End Sub

    'UPGRADE_WARNING: Event chkshowtext.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
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

    Private Sub Command13_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command13.Click
        If ColorDialog1.ShowDialog() Then

            Me.Scanner1.SetBackgroundColor(Color2Uint32(ColorDialog1.Color))

        End If

    End Sub

    Private Sub Command14_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command14.Click
        Scanner1.PrintImage(True)
    End Sub

    Private Sub Command15_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command15.Click
        Dim a As Integer

        Dim strFilter As String

        strFilter = "TIFF File (*.tif)|*.tif"

        SaveFileDialog1.Filter = strFilter
        SaveFileDialog1.DefaultExt = "tif"
        Scanner1.View = 5
        Scanner1.TIFCompression = cbotifcompression.SelectedIndex

        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            Scanner1.SaveAllPage2TIF(SaveFileDialog1.FileName, Check3.CheckState, 1)
        End If




    End Sub

    Private Sub Command16_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command16.Click
        Dim a As Integer

        Dim strFilter As String

        strFilter = "PDF File (*.pdf)|*.pdf"

        SaveFileDialog1.Filter = strFilter
        SaveFileDialog1.DefaultExt = "pdf"
        Scanner1.View = 5
        Scanner1.TIFCompression = cbotifcompression.SelectedIndex

        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            Scanner1.SaveAllPage2PDF(SaveFileDialog1.FileName, Check3.CheckState, 1)
        End If





    End Sub

    Private Sub Command17_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command17.Click
        Scanner1.DeletePage(Scanner1.GetActivePageNo)


        Scanner1.SetActivePageNo(1)
        lbltotalpage.Text = CStr(Me.Scanner1.TotalPage)
        lblcurrentpage.Text = CStr(Me.Scanner1.GetActivePageNo)

    End Sub

    Private Sub Command18_Click()
        Dim a As Object
        Dim strType As Object
        Dim strFile As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object strFile. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        strFile = "c:\test"
        'UPGRADE_WARNING: Couldn't resolve default property of object strType. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        strType = "gif"
        'UPGRADE_WARNING: Couldn't resolve default property of object strType. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        'UPGRADE_WARNING: Couldn't resolve default property of object strFile. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        'UPGRADE_WARNING: Couldn't resolve default property of object a. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        a = Me.Scanner1.Save(strFile, strType)

        'UPGRADE_WARNING: Couldn't resolve default property of object a. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        If a = 1 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object strType. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            'UPGRADE_WARNING: Couldn't resolve default property of object strFile. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            MsgBox("Save " + strFile + "." + strType + " Complete")
        Else
            MsgBox("Save fail")
        End If
        Scanner1.Focus()


    End Sub

    Private Sub Command19_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command19.Click
        Dim a As Object
        Dim strType As Object
        Dim strFile As Object

        strFile = "c:\test"
         strType = "jpg"

          a = Me.Scanner1.SaveBySize(strFile, strType, Val(txt_Width.Text), Val(txt_Height.Text))

        'UPGRADE_WARNING: Couldn't resolve default property of object a. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        If a = 1 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object strType. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            'UPGRADE_WARNING: Couldn't resolve default property of object strFile. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            MsgBox("Save " + strFile + "." + strType + " Complete")
        Else
            MsgBox("Save fail")
        End If


    End Sub

    Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
        Me.Scanner1.Rotate180()
    End Sub

    Private Sub Command20_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command20.Click

        Scanner1.View = 10
        Scanner1.Focus()

    End Sub

    Private Sub Command21_Click()
        Dim a As Object
        Dim strType As Object
        Dim strFile As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object strFile. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        strFile = "c:\test"
        'UPGRADE_WARNING: Couldn't resolve default property of object strType. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        strType = "png"
        'UPGRADE_WARNING: Couldn't resolve default property of object strType. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        'UPGRADE_WARNING: Couldn't resolve default property of object strFile. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        'UPGRADE_WARNING: Couldn't resolve default property of object a. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        a = Me.Scanner1.Save(strFile, strType)

        'UPGRADE_WARNING: Couldn't resolve default property of object a. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        If a = 1 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object strType. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            'UPGRADE_WARNING: Couldn't resolve default property of object strFile. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            MsgBox("Save " + strFile + "." + strType + " Complete")
        Else
            MsgBox("Save fail")
        End If
        Scanner1.Focus()


    End Sub

    Private Sub Command22_Click()
        Dim a As Object
        Dim strType As Object
        Dim strFile As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object strFile. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        strFile = "c:\test"
        'UPGRADE_WARNING: Couldn't resolve default property of object strType. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        strType = "pdf"

        'UPGRADE_WARNING: Couldn't resolve default property of object strType. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        'UPGRADE_WARNING: Couldn't resolve default property of object strFile. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        'UPGRADE_WARNING: Couldn't resolve default property of object a. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        a = Me.Scanner1.Save(strFile, strType)

        'UPGRADE_WARNING: Couldn't resolve default property of object a. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        If a = 1 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object strType. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            'UPGRADE_WARNING: Couldn't resolve default property of object strFile. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            MsgBox("Save " + strFile + "." + strType + " Complete")
        Else
            MsgBox("Save fail")
        End If
        Scanner1.Focus()

    End Sub

    Private Sub Command23_Click()


        Me.Scanner1.FileLeft = CInt(txtLeft.Text)
        Me.Scanner1.FileTop = CInt(txtTop.Text)

        Me.Scanner1.Redraw()

    End Sub

    Private Sub Command24_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command24.Click

        Scanner1.Copy2Clipboard()


    End Sub

    Private Sub Command25_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Picture1.Image = Scanner1.Copy2PictureBox
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
        Dim a As Object
        Dim Text5 As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object Text5. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        'UPGRADE_WARNING: Couldn't resolve default property of object a. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        a = Me.Scanner1.SaveAllPage2PDF(Text5, Check3.CheckState, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object a. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        If a Then
            MsgBox("complete")
        End If

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


        Scanner1.Scan()




    End Sub

    Private Sub Command40_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command40.Click
        Dim result As Object
        Dim strType As Object
        Dim strFile As Object


        strFile = "c:\test"

        Select Case cboutputtype.SelectedIndex

            Case 0
                 strType = "bmp"
            Case 1
                'UPGRADE_WARNING: Couldn't resolve default property of object strType. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                strType = "jpg"
            Case 2
                'UPGRADE_WARNING: Couldn't resolve default property of object strType. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                strType = "gif"
            Case 3
                'UPGRADE_WARNING: Couldn't resolve default property of object strType. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                strType = "png"
            Case 4
                'UPGRADE_WARNING: Couldn't resolve default property of object strType. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                strType = "tif"
            Case 5
                'UPGRADE_WARNING: Couldn't resolve default property of object strType. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                strType = "pdf"
        End Select

        result = Me.Scanner1.Save(strFile, strType)

        'UPGRADE_WARNING: Couldn't resolve default property of object result. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        If result = 1 Then
                MsgBox("Save " + strFile + "." + strType + " Complete")
        Else
            MsgBox("Save fail")
        End If



        Scanner1.Focus()


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

        cbotifcompression.Items.Add("LZW")
        cbotifcompression.Items.Add("CITT3")
        cbotifcompression.Items.Add("CITT4")
        cbotifcompression.Items.Add("RLE")
        cbotifcompression.Items.Add("None")
        cbotifcompression.SelectedIndex = 0

        cboutputtype.Items.Add("BMP")
        cboutputtype.Items.Add("JPG")
        cboutputtype.Items.Add("GIF")
        cboutputtype.Items.Add("PNG")
        cboutputtype.Items.Add("TIF")
        cboutputtype.Items.Add("PDF")
        cboutputtype.SelectedIndex = 0


        cbodpi.Items.Add("Onscreen Viewing 96dpi")
        cbodpi.Items.Add("Fax 200dpi")
        cbodpi.Items.Add("OCR Text 300dpi")
        cbodpi.Items.Add("Laser Print Fine 600dpi")
        cbodpi.SelectedIndex = 0

        cbopixeltype.Items.Add("Default")
        cbopixeltype.Items.Add("Gray Color")
        cbopixeltype.Items.Add("Black & White Color")
        cbopixeltype.Items.Add("True Color")

        cbopixeltype.SelectedIndex = 0



        'UPGRADE_WARNING: Couldn't resolve default property of object clrHashForeColor. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        clrHashForeColor = RGB(0, 0, 0)
        'UPGRADE_WARNING: Couldn't resolve default property of object clrHashBackColor. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        clrHashBackColor = RGB(255, 255, 255)

        'UPGRADE_WARNING: Couldn't resolve default property of object iCount. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        iCount = Me.Scanner1.GetNumImageSources

        'UPGRADE_WARNING: Couldn't resolve default property of object iCount. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        For I = 0 To iCount - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            cboimagesource.Items.Add(Scanner1.GetImageSourceName(I))
        Next

        If iCount > 0 Then
            cboimagesource.SelectedIndex = 0
        End If

       

       


    End Sub

    'UPGRADE_WARNING: Event optborderoff.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
    Private Sub optborderoff_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optborderoff.CheckedChanged

    End Sub

    Private Sub optborderon_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optborderon.CheckedChanged



    End Sub

    Private Sub Option3_Click()

    End Sub

    'UPGRADE_WARNING: Event opthighno.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
    Private Sub opthighno_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles opthighno.CheckedChanged

    End Sub

    'UPGRADE_WARNING: Event opthighyes.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
    Private Sub opthighyes_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles opthighyes.CheckedChanged

    End Sub

    Private Sub Scanner1_EndScan(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Scanner1.EndScan

        txt_Width.Text = CStr(Me.Scanner1.FileWidth)
        txt_Height.Text = CStr(Me.Scanner1.FileHeight)



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

    Private Sub Check1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Check1.CheckedChanged
        If Check1.CheckState = 1 Then
            Me.Scanner1.OutputGrayScale = True
        Else
            Me.Scanner1.OutputGrayScale = False

        End If
    End Sub

    Private Sub Check3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Check3.CheckedChanged

    End Sub

    Private Sub opthighyes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles opthighyes.Click

        If Me.opthighyes.Checked = True Then
            Scanner1.HighQuality = True
        Else
            Scanner1.HighQuality = False
        End If



    End Sub

    Private Sub optborderon_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles optborderon.Click
        If Me.optborderon.Checked = True Then
            Scanner1.Border = True
        Else
            Scanner1.Border = False
        End If

    End Sub

    Private Sub optborderoff_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles optborderoff.Click
        If Me.optborderon.Checked = True Then
            Scanner1.Border = True
        Else
            Scanner1.Border = False
        End If

    End Sub

    Private Sub opthighno_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles opthighno.Click
        If Me.opthighyes.Checked = True Then
            Scanner1.HighQuality = True
        Else
            Scanner1.HighQuality = False
        End If


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

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim a As Integer
        Dim strFilter As String

        strFilter = "Docx File (*.docx)|*.docx"

        SaveFileDialog1.Filter = strFilter
        SaveFileDialog1.DefaultExt = "docx"
        Scanner1.View = 5

        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then

            Scanner1.SaveAllPage2Docx(SaveFileDialog1.FileName, Check3.CheckState, 1)

        End If

    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Dim strFilter As String
        strFilter = "PDF (*.pdf)|*.pdf|Open XML Paper Specification (*.xps *.oxps)|*.xps;*.oxps|TIFF (*.tiff)|*.tif|JPEG (*.jpg)|*.jpg|BMP (*.bmp)|*.bmp|PNG (*.png)|*.png|GIF (*.gif)|*.gif"

        If chkpanning.Checked Then
            Scanner1.EnablePanning = True
        Else
            Scanner1.EnablePanning = False

        End If
        OpenFileDialog1.Filter = strFilter

        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then

            TextBox1.Text = OpenFileDialog1.FileName
            Scanner1.LoadImage(OpenFileDialog1.FileName)

            Command4.Enabled = True
            lbltotalpage.Text = CStr(Me.Scanner1.TotalPage)
            lblcurrentpage.Text = 1
        End If
    End Sub
End Class