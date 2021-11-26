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
	Public WithEvents cbofont As System.Windows.Forms.ComboBox
	Public WithEvents cbofontsize As System.Windows.Forms.ComboBox
	Public WithEvents cbofontstyle As System.Windows.Forms.ComboBox
	Public WithEvents txtTextLeft As System.Windows.Forms.TextBox
	Public WithEvents txtTextTop As System.Windows.Forms.TextBox
	Public WithEvents Command39 As System.Windows.Forms.Button
	Public WithEvents txttext1 As System.Windows.Forms.TextBox
	Public WithEvents txttext2 As System.Windows.Forms.TextBox
	Public WithEvents cbotextstyle As System.Windows.Forms.ComboBox
	Public WithEvents cmdtextcolor As System.Windows.Forms.Button
	Public WithEvents cmdoutlineborder As System.Windows.Forms.Button
	Public WithEvents cmdoutlineback As System.Windows.Forms.Button
	Public WithEvents txttextureimage As System.Windows.Forms.TextBox
	Public WithEvents cmdtextureimage As System.Windows.Forms.Button
	Public WithEvents cbohashbrushstyle As System.Windows.Forms.ComboBox
	Public WithEvents cmdhashforecolor As System.Windows.Forms.Button
	Public WithEvents cmdhashbackcolor As System.Windows.Forms.Button
	Public WithEvents chkshowtext As System.Windows.Forms.CheckBox
	Public WithEvents Label11 As System.Windows.Forms.Label
	Public WithEvents Label12 As System.Windows.Forms.Label
	Public WithEvents Label13 As System.Windows.Forms.Label
	Public WithEvents Label14 As System.Windows.Forms.Label
	Public WithEvents Label15 As System.Windows.Forms.Label
	Public WithEvents Label16 As System.Windows.Forms.Label
	Public WithEvents Label17 As System.Windows.Forms.Label
	Public WithEvents Label18 As System.Windows.Forms.Label
	Public WithEvents Label19 As System.Windows.Forms.Label
	Public WithEvents Label20 As System.Windows.Forms.Label
	Public WithEvents Label21 As System.Windows.Forms.Label
	Public WithEvents Label22 As System.Windows.Forms.Label
	Public WithEvents Label23 As System.Windows.Forms.Label
	Public WithEvents Label24 As System.Windows.Forms.Label
	Public WithEvents Label25 As System.Windows.Forms.Label
	Public WithEvents Frame6 As System.Windows.Forms.GroupBox
	Public WithEvents txtSat As System.Windows.Forms.TextBox
	Public WithEvents Command34 As System.Windows.Forms.Button
	Public WithEvents Command35 As System.Windows.Forms.Button
	Public WithEvents txtgamma As System.Windows.Forms.TextBox
	Public WithEvents txthue As System.Windows.Forms.TextBox
	Public WithEvents txtcontrast As System.Windows.Forms.TextBox
	Public WithEvents Command36 As System.Windows.Forms.Button
	Public WithEvents Command38 As System.Windows.Forms.Button
	Public WithEvents txtBrightness As System.Windows.Forms.TextBox
	Public WithEvents Command37 As System.Windows.Forms.Button
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
	Public WithEvents Label26 As System.Windows.Forms.Label
	Public WithEvents Label28 As System.Windows.Forms.Label
	Public WithEvents Label29 As System.Windows.Forms.Label
	Public WithEvents Label30 As System.Windows.Forms.Label
	Public WithEvents Label27 As System.Windows.Forms.Label
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
    Friend WithEvents chkaddPDFAnnotation As System.Windows.Forms.CheckBox
    Friend WithEvents chkpdfusejpegcomp As System.Windows.Forms.CheckBox
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents txtpdfjpegquality As System.Windows.Forms.TextBox
    Friend WithEvents Label41 As System.Windows.Forms.Label
    Friend WithEvents txtconfid As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents chkpdfusejpegcomp2 As System.Windows.Forms.CheckBox
    Friend WithEvents Label42 As System.Windows.Forms.Label
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Public WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Label43 As System.Windows.Forms.Label
    Friend WithEvents chkautosaveallpages As System.Windows.Forms.CheckBox
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents Label44 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtbufferresizescale As System.Windows.Forms.TextBox
    Friend WithEvents Label48 As System.Windows.Forms.Label
    Friend WithEvents Label47 As System.Windows.Forms.Label
    Friend WithEvents Label46 As System.Windows.Forms.Label
    Friend WithEvents txtbufferresizemaxwidth As System.Windows.Forms.TextBox
    Friend WithEvents Label45 As System.Windows.Forms.Label
    Friend WithEvents chkbufferresize As System.Windows.Forms.CheckBox
    Friend WithEvents chkpanning As System.Windows.Forms.CheckBox
    Friend WithEvents txtpdfjpegquality2 As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Command17 = New System.Windows.Forms.Button()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.txtpdfjpegquality2 = New System.Windows.Forms.TextBox()
        Me.Label42 = New System.Windows.Forms.Label()
        Me.chkpdfusejpegcomp2 = New System.Windows.Forms.CheckBox()
        Me.chkaddPDFAnnotation = New System.Windows.Forms.CheckBox()
        Me.cbotifcompression = New System.Windows.Forms.ComboBox()
        Me.Label39 = New System.Windows.Forms.Label()
        Me.Check3 = New System.Windows.Forms.CheckBox()
        Me.Command16 = New System.Windows.Forms.Button()
        Me.Command15 = New System.Windows.Forms.Button()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.txtpdfjpegquality = New System.Windows.Forms.TextBox()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.chkpdfusejpegcomp = New System.Windows.Forms.CheckBox()
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
        Me.Frame6 = New System.Windows.Forms.GroupBox()
        Me.cbofont = New System.Windows.Forms.ComboBox()
        Me.cbofontsize = New System.Windows.Forms.ComboBox()
        Me.cbofontstyle = New System.Windows.Forms.ComboBox()
        Me.txtTextLeft = New System.Windows.Forms.TextBox()
        Me.txtTextTop = New System.Windows.Forms.TextBox()
        Me.Command39 = New System.Windows.Forms.Button()
        Me.txttext1 = New System.Windows.Forms.TextBox()
        Me.txttext2 = New System.Windows.Forms.TextBox()
        Me.cbotextstyle = New System.Windows.Forms.ComboBox()
        Me.cmdtextcolor = New System.Windows.Forms.Button()
        Me.cmdoutlineborder = New System.Windows.Forms.Button()
        Me.cmdoutlineback = New System.Windows.Forms.Button()
        Me.txttextureimage = New System.Windows.Forms.TextBox()
        Me.cmdtextureimage = New System.Windows.Forms.Button()
        Me.cbohashbrushstyle = New System.Windows.Forms.ComboBox()
        Me.cmdhashforecolor = New System.Windows.Forms.Button()
        Me.cmdhashbackcolor = New System.Windows.Forms.Button()
        Me.chkshowtext = New System.Windows.Forms.CheckBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.txtSat = New System.Windows.Forms.TextBox()
        Me.Command34 = New System.Windows.Forms.Button()
        Me.Command35 = New System.Windows.Forms.Button()
        Me.txtgamma = New System.Windows.Forms.TextBox()
        Me.txthue = New System.Windows.Forms.TextBox()
        Me.txtcontrast = New System.Windows.Forms.TextBox()
        Me.Command36 = New System.Windows.Forms.Button()
        Me.Command38 = New System.Windows.Forms.Button()
        Me.txtBrightness = New System.Windows.Forms.TextBox()
        Me.Command37 = New System.Windows.Forms.Button()
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtbufferresizescale = New System.Windows.Forms.TextBox()
        Me.Label48 = New System.Windows.Forms.Label()
        Me.Label47 = New System.Windows.Forms.Label()
        Me.Label46 = New System.Windows.Forms.Label()
        Me.txtbufferresizemaxwidth = New System.Windows.Forms.TextBox()
        Me.Label45 = New System.Windows.Forms.Label()
        Me.chkbufferresize = New System.Windows.Forms.CheckBox()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Label44 = New System.Windows.Forms.Label()
        Me.Label43 = New System.Windows.Forms.Label()
        Me.chkautosaveallpages = New System.Windows.Forms.CheckBox()
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
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.lblcurrentpage = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.lbltotalpage = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ColorDialog1 = New System.Windows.Forms.ColorDialog()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Label41 = New System.Windows.Forms.Label()
        Me.txtconfid = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.chkpanning = New System.Windows.Forms.CheckBox()
        Me.Frame4.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.Frame6.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me._Frame1_0.SuspendLayout()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me._Frame1_1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Frame5.SuspendLayout()
        CType(Me.Scanner1, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.Frame4.Controls.Add(Me.txtpdfjpegquality2)
        Me.Frame4.Controls.Add(Me.Label42)
        Me.Frame4.Controls.Add(Me.chkpdfusejpegcomp2)
        Me.Frame4.Controls.Add(Me.chkaddPDFAnnotation)
        Me.Frame4.Controls.Add(Me.cbotifcompression)
        Me.Frame4.Controls.Add(Me.Label39)
        Me.Frame4.Controls.Add(Me.Check3)
        Me.Frame4.Controls.Add(Me.Command16)
        Me.Frame4.Controls.Add(Me.Command15)
        Me.Frame4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame4.Location = New System.Drawing.Point(663, 268)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(217, 281)
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
        Me.Button3.Location = New System.Drawing.Point(10, 45)
        Me.Button3.Name = "Button3"
        Me.Button3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Button3.Size = New System.Drawing.Size(185, 33)
        Me.Button3.TabIndex = 120
        Me.Button3.Text = "Save Image Buffers to Multi Page Docx"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'txtpdfjpegquality2
        '
        Me.txtpdfjpegquality2.Location = New System.Drawing.Point(120, 207)
        Me.txtpdfjpegquality2.Name = "txtpdfjpegquality2"
        Me.txtpdfjpegquality2.Size = New System.Drawing.Size(64, 20)
        Me.txtpdfjpegquality2.TabIndex = 119
        Me.txtpdfjpegquality2.Text = "80"
        '
        'Label42
        '
        Me.Label42.Location = New System.Drawing.Point(16, 215)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(80, 24)
        Me.Label42.TabIndex = 118
        Me.Label42.Text = "Quality(0-100)"
        '
        'chkpdfusejpegcomp2
        '
        Me.chkpdfusejpegcomp2.Location = New System.Drawing.Point(16, 191)
        Me.chkpdfusejpegcomp2.Name = "chkpdfusejpegcomp2"
        Me.chkpdfusejpegcomp2.Size = New System.Drawing.Size(144, 16)
        Me.chkpdfusejpegcomp2.TabIndex = 117
        Me.chkpdfusejpegcomp2.Text = "use JPEG Compression"
        '
        'chkaddPDFAnnotation
        '
        Me.chkaddPDFAnnotation.Location = New System.Drawing.Point(16, 167)
        Me.chkaddPDFAnnotation.Name = "chkaddPDFAnnotation"
        Me.chkaddPDFAnnotation.Size = New System.Drawing.Size(176, 16)
        Me.chkaddPDFAnnotation.TabIndex = 116
        Me.chkaddPDFAnnotation.Text = "Add PDF annotation in page 1 "
        '
        'cbotifcompression
        '
        Me.cbotifcompression.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbotifcompression.Location = New System.Drawing.Point(112, 87)
        Me.cbotifcompression.Name = "cbotifcompression"
        Me.cbotifcompression.Size = New System.Drawing.Size(96, 22)
        Me.cbotifcompression.TabIndex = 115
        '
        'Label39
        '
        Me.Label39.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label39.Location = New System.Drawing.Point(8, 87)
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
        Me.Command16.Location = New System.Drawing.Point(8, 231)
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
        Me.Command15.Location = New System.Drawing.Point(8, 127)
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
        Me.Frame3.Controls.Add(Me.txtpdfjpegquality)
        Me.Frame3.Controls.Add(Me.Label40)
        Me.Frame3.Controls.Add(Me.chkpdfusejpegcomp)
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
        Me.Frame3.Location = New System.Drawing.Point(657, 4)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(224, 264)
        Me.Frame3.TabIndex = 94
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Current View"
        '
        'txtpdfjpegquality
        '
        Me.txtpdfjpegquality.Enabled = False
        Me.txtpdfjpegquality.Location = New System.Drawing.Point(112, 112)
        Me.txtpdfjpegquality.Name = "txtpdfjpegquality"
        Me.txtpdfjpegquality.Size = New System.Drawing.Size(48, 20)
        Me.txtpdfjpegquality.TabIndex = 111
        Me.txtpdfjpegquality.Text = "80"
        '
        'Label40
        '
        Me.Label40.Location = New System.Drawing.Point(16, 112)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(80, 16)
        Me.Label40.TabIndex = 110
        Me.Label40.Text = "Quality(0-100)"
        '
        'chkpdfusejpegcomp
        '
        Me.chkpdfusejpegcomp.Enabled = False
        Me.chkpdfusejpegcomp.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkpdfusejpegcomp.Location = New System.Drawing.Point(16, 88)
        Me.chkpdfusejpegcomp.Name = "chkpdfusejpegcomp"
        Me.chkpdfusejpegcomp.Size = New System.Drawing.Size(144, 24)
        Me.chkpdfusejpegcomp.TabIndex = 109
        Me.chkpdfusejpegcomp.Text = "PDF use JPEG Compression"
        '
        'Command13
        '
        Me.Command13.BackColor = System.Drawing.SystemColors.Control
        Me.Command13.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command13.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command13.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command13.Location = New System.Drawing.Point(16, 208)
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
        Me.Command14.Location = New System.Drawing.Point(120, 232)
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
        Me.Command24.Location = New System.Drawing.Point(16, 232)
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
        Me.txt_Height.Location = New System.Drawing.Point(64, 176)
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
        Me.txt_Width.Location = New System.Drawing.Point(64, 144)
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
        Me.Command19.Location = New System.Drawing.Point(120, 176)
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
        Me.Check1.Location = New System.Drawing.Point(16, 48)
        Me.Check1.Name = "Check1"
        Me.Check1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Check1.Size = New System.Drawing.Size(96, 25)
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
        Me.cboutputtype.Sorted = True
        Me.cboutputtype.TabIndex = 96
        '
        'Command40
        '
        Me.Command40.BackColor = System.Drawing.SystemColors.Control
        Me.Command40.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command40.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command40.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command40.Location = New System.Drawing.Point(136, 56)
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
        Me.Label5.Location = New System.Drawing.Point(8, 176)
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
        Me.Label4.Location = New System.Drawing.Point(8, 144)
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
        'Frame6
        '
        Me.Frame6.BackColor = System.Drawing.SystemColors.Control
        Me.Frame6.Controls.Add(Me.cbofont)
        Me.Frame6.Controls.Add(Me.cbofontsize)
        Me.Frame6.Controls.Add(Me.cbofontstyle)
        Me.Frame6.Controls.Add(Me.txtTextLeft)
        Me.Frame6.Controls.Add(Me.txtTextTop)
        Me.Frame6.Controls.Add(Me.Command39)
        Me.Frame6.Controls.Add(Me.txttext1)
        Me.Frame6.Controls.Add(Me.txttext2)
        Me.Frame6.Controls.Add(Me.cbotextstyle)
        Me.Frame6.Controls.Add(Me.cmdtextcolor)
        Me.Frame6.Controls.Add(Me.cmdoutlineborder)
        Me.Frame6.Controls.Add(Me.cmdoutlineback)
        Me.Frame6.Controls.Add(Me.txttextureimage)
        Me.Frame6.Controls.Add(Me.cmdtextureimage)
        Me.Frame6.Controls.Add(Me.cbohashbrushstyle)
        Me.Frame6.Controls.Add(Me.cmdhashforecolor)
        Me.Frame6.Controls.Add(Me.cmdhashbackcolor)
        Me.Frame6.Controls.Add(Me.chkshowtext)
        Me.Frame6.Controls.Add(Me.Label11)
        Me.Frame6.Controls.Add(Me.Label12)
        Me.Frame6.Controls.Add(Me.Label13)
        Me.Frame6.Controls.Add(Me.Label14)
        Me.Frame6.Controls.Add(Me.Label15)
        Me.Frame6.Controls.Add(Me.Label16)
        Me.Frame6.Controls.Add(Me.Label17)
        Me.Frame6.Controls.Add(Me.Label18)
        Me.Frame6.Controls.Add(Me.Label19)
        Me.Frame6.Controls.Add(Me.Label20)
        Me.Frame6.Controls.Add(Me.Label21)
        Me.Frame6.Controls.Add(Me.Label22)
        Me.Frame6.Controls.Add(Me.Label23)
        Me.Frame6.Controls.Add(Me.Label24)
        Me.Frame6.Controls.Add(Me.Label25)
        Me.Frame6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame6.Location = New System.Drawing.Point(887, 4)
        Me.Frame6.Name = "Frame6"
        Me.Frame6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame6.Size = New System.Drawing.Size(201, 553)
        Me.Frame6.TabIndex = 47
        Me.Frame6.TabStop = False
        Me.Frame6.Text = "Text"
        '
        'cbofont
        '
        Me.cbofont.BackColor = System.Drawing.SystemColors.Window
        Me.cbofont.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbofont.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbofont.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbofont.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cbofont.Location = New System.Drawing.Point(72, 24)
        Me.cbofont.Name = "cbofont"
        Me.cbofont.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbofont.Size = New System.Drawing.Size(97, 22)
        Me.cbofont.TabIndex = 80
        '
        'cbofontsize
        '
        Me.cbofontsize.BackColor = System.Drawing.SystemColors.Window
        Me.cbofontsize.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbofontsize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbofontsize.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbofontsize.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cbofontsize.Location = New System.Drawing.Point(72, 48)
        Me.cbofontsize.Name = "cbofontsize"
        Me.cbofontsize.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbofontsize.Size = New System.Drawing.Size(97, 22)
        Me.cbofontsize.TabIndex = 64
        '
        'cbofontstyle
        '
        Me.cbofontstyle.BackColor = System.Drawing.SystemColors.Window
        Me.cbofontstyle.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbofontstyle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbofontstyle.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbofontstyle.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cbofontstyle.Location = New System.Drawing.Point(72, 80)
        Me.cbofontstyle.Name = "cbofontstyle"
        Me.cbofontstyle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbofontstyle.Size = New System.Drawing.Size(97, 22)
        Me.cbofontstyle.TabIndex = 63
        '
        'txtTextLeft
        '
        Me.txtTextLeft.AcceptsReturn = True
        Me.txtTextLeft.BackColor = System.Drawing.SystemColors.Window
        Me.txtTextLeft.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTextLeft.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTextLeft.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTextLeft.Location = New System.Drawing.Point(56, 112)
        Me.txtTextLeft.MaxLength = 0
        Me.txtTextLeft.Name = "txtTextLeft"
        Me.txtTextLeft.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTextLeft.Size = New System.Drawing.Size(33, 20)
        Me.txtTextLeft.TabIndex = 62
        Me.txtTextLeft.Text = "0"
        '
        'txtTextTop
        '
        Me.txtTextTop.AcceptsReturn = True
        Me.txtTextTop.BackColor = System.Drawing.SystemColors.Window
        Me.txtTextTop.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTextTop.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTextTop.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTextTop.Location = New System.Drawing.Point(128, 112)
        Me.txtTextTop.MaxLength = 0
        Me.txtTextTop.Name = "txtTextTop"
        Me.txtTextTop.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTextTop.Size = New System.Drawing.Size(33, 20)
        Me.txtTextTop.TabIndex = 61
        Me.txtTextTop.Text = "0"
        '
        'Command39
        '
        Me.Command39.BackColor = System.Drawing.SystemColors.Control
        Me.Command39.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command39.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command39.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command39.Location = New System.Drawing.Point(8, 208)
        Me.Command39.Name = "Command39"
        Me.Command39.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command39.Size = New System.Drawing.Size(81, 25)
        Me.Command39.TabIndex = 60
        Me.Command39.Text = "Draw Text"
        Me.Command39.UseVisualStyleBackColor = False
        '
        'txttext1
        '
        Me.txttext1.AcceptsReturn = True
        Me.txttext1.BackColor = System.Drawing.SystemColors.Window
        Me.txttext1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txttext1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txttext1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txttext1.Location = New System.Drawing.Point(56, 136)
        Me.txttext1.MaxLength = 0
        Me.txttext1.Name = "txttext1"
        Me.txttext1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txttext1.Size = New System.Drawing.Size(137, 20)
        Me.txttext1.TabIndex = 59
        Me.txttext1.Text = "Viscom Software"
        '
        'txttext2
        '
        Me.txttext2.AcceptsReturn = True
        Me.txttext2.BackColor = System.Drawing.SystemColors.Window
        Me.txttext2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txttext2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txttext2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txttext2.Location = New System.Drawing.Point(56, 168)
        Me.txttext2.MaxLength = 0
        Me.txttext2.Name = "txttext2"
        Me.txttext2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txttext2.Size = New System.Drawing.Size(137, 20)
        Me.txttext2.TabIndex = 58
        Me.txttext2.Text = "www.viscomsoft.com"
        '
        'cbotextstyle
        '
        Me.cbotextstyle.BackColor = System.Drawing.SystemColors.Window
        Me.cbotextstyle.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbotextstyle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbotextstyle.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbotextstyle.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cbotextstyle.Location = New System.Drawing.Point(88, 264)
        Me.cbotextstyle.Name = "cbotextstyle"
        Me.cbotextstyle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbotextstyle.Size = New System.Drawing.Size(105, 22)
        Me.cbotextstyle.TabIndex = 57
        '
        'cmdtextcolor
        '
        Me.cmdtextcolor.BackColor = System.Drawing.SystemColors.Control
        Me.cmdtextcolor.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdtextcolor.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdtextcolor.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdtextcolor.Location = New System.Drawing.Point(120, 296)
        Me.cmdtextcolor.Name = "cmdtextcolor"
        Me.cmdtextcolor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdtextcolor.Size = New System.Drawing.Size(73, 17)
        Me.cmdtextcolor.TabIndex = 56
        Me.cmdtextcolor.Text = "Change"
        Me.cmdtextcolor.UseVisualStyleBackColor = False
        '
        'cmdoutlineborder
        '
        Me.cmdoutlineborder.BackColor = System.Drawing.SystemColors.Control
        Me.cmdoutlineborder.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdoutlineborder.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdoutlineborder.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdoutlineborder.Location = New System.Drawing.Point(120, 328)
        Me.cmdoutlineborder.Name = "cmdoutlineborder"
        Me.cmdoutlineborder.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdoutlineborder.Size = New System.Drawing.Size(73, 17)
        Me.cmdoutlineborder.TabIndex = 55
        Me.cmdoutlineborder.Text = "Change"
        Me.cmdoutlineborder.UseVisualStyleBackColor = False
        '
        'cmdoutlineback
        '
        Me.cmdoutlineback.BackColor = System.Drawing.SystemColors.Control
        Me.cmdoutlineback.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdoutlineback.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdoutlineback.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdoutlineback.Location = New System.Drawing.Point(120, 360)
        Me.cmdoutlineback.Name = "cmdoutlineback"
        Me.cmdoutlineback.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdoutlineback.Size = New System.Drawing.Size(73, 17)
        Me.cmdoutlineback.TabIndex = 54
        Me.cmdoutlineback.Text = "Change"
        Me.cmdoutlineback.UseVisualStyleBackColor = False
        '
        'txttextureimage
        '
        Me.txttextureimage.AcceptsReturn = True
        Me.txttextureimage.BackColor = System.Drawing.SystemColors.Window
        Me.txttextureimage.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txttextureimage.Enabled = False
        Me.txttextureimage.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txttextureimage.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txttextureimage.Location = New System.Drawing.Point(8, 504)
        Me.txttextureimage.MaxLength = 0
        Me.txttextureimage.Name = "txttextureimage"
        Me.txttextureimage.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txttextureimage.Size = New System.Drawing.Size(185, 20)
        Me.txttextureimage.TabIndex = 53
        '
        'cmdtextureimage
        '
        Me.cmdtextureimage.BackColor = System.Drawing.SystemColors.Control
        Me.cmdtextureimage.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdtextureimage.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdtextureimage.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdtextureimage.Location = New System.Drawing.Point(8, 528)
        Me.cmdtextureimage.Name = "cmdtextureimage"
        Me.cmdtextureimage.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdtextureimage.Size = New System.Drawing.Size(81, 17)
        Me.cmdtextureimage.TabIndex = 52
        Me.cmdtextureimage.Text = "Select Image"
        Me.cmdtextureimage.UseVisualStyleBackColor = False
        '
        'cbohashbrushstyle
        '
        Me.cbohashbrushstyle.BackColor = System.Drawing.SystemColors.Window
        Me.cbohashbrushstyle.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbohashbrushstyle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbohashbrushstyle.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbohashbrushstyle.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cbohashbrushstyle.Location = New System.Drawing.Point(8, 400)
        Me.cbohashbrushstyle.Name = "cbohashbrushstyle"
        Me.cbohashbrushstyle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbohashbrushstyle.Size = New System.Drawing.Size(137, 22)
        Me.cbohashbrushstyle.TabIndex = 51
        '
        'cmdhashforecolor
        '
        Me.cmdhashforecolor.BackColor = System.Drawing.SystemColors.Control
        Me.cmdhashforecolor.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdhashforecolor.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdhashforecolor.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdhashforecolor.Location = New System.Drawing.Point(120, 432)
        Me.cmdhashforecolor.Name = "cmdhashforecolor"
        Me.cmdhashforecolor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdhashforecolor.Size = New System.Drawing.Size(73, 17)
        Me.cmdhashforecolor.TabIndex = 50
        Me.cmdhashforecolor.Text = "Change"
        Me.cmdhashforecolor.UseVisualStyleBackColor = False
        '
        'cmdhashbackcolor
        '
        Me.cmdhashbackcolor.BackColor = System.Drawing.SystemColors.Control
        Me.cmdhashbackcolor.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdhashbackcolor.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdhashbackcolor.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdhashbackcolor.Location = New System.Drawing.Point(120, 456)
        Me.cmdhashbackcolor.Name = "cmdhashbackcolor"
        Me.cmdhashbackcolor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdhashbackcolor.Size = New System.Drawing.Size(73, 17)
        Me.cmdhashbackcolor.TabIndex = 49
        Me.cmdhashbackcolor.Text = "Change"
        Me.cmdhashbackcolor.UseVisualStyleBackColor = False
        '
        'chkshowtext
        '
        Me.chkshowtext.BackColor = System.Drawing.SystemColors.Control
        Me.chkshowtext.Checked = True
        Me.chkshowtext.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkshowtext.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkshowtext.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkshowtext.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkshowtext.Location = New System.Drawing.Point(8, 240)
        Me.chkshowtext.Name = "chkshowtext"
        Me.chkshowtext.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkshowtext.Size = New System.Drawing.Size(89, 17)
        Me.chkshowtext.TabIndex = 48
        Me.chkshowtext.Text = "Show Text"
        Me.chkshowtext.UseVisualStyleBackColor = False
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label11.Location = New System.Drawing.Point(8, 24)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.Size = New System.Drawing.Size(57, 17)
        Me.Label11.TabIndex = 79
        Me.Label11.Text = "Font Name"
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.SystemColors.Control
        Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label12.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label12.Location = New System.Drawing.Point(8, 48)
        Me.Label12.Name = "Label12"
        Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label12.Size = New System.Drawing.Size(57, 25)
        Me.Label12.TabIndex = 78
        Me.Label12.Text = "Font Size"
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.SystemColors.Control
        Me.Label13.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label13.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label13.Location = New System.Drawing.Point(8, 80)
        Me.Label13.Name = "Label13"
        Me.Label13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label13.Size = New System.Drawing.Size(57, 17)
        Me.Label13.TabIndex = 77
        Me.Label13.Text = "Font Style"
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.SystemColors.Control
        Me.Label14.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label14.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label14.Location = New System.Drawing.Point(8, 112)
        Me.Label14.Name = "Label14"
        Me.Label14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label14.Size = New System.Drawing.Size(41, 17)
        Me.Label14.TabIndex = 76
        Me.Label14.Text = "Left"
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.SystemColors.Control
        Me.Label15.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label15.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label15.Location = New System.Drawing.Point(104, 112)
        Me.Label15.Name = "Label15"
        Me.Label15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label15.Size = New System.Drawing.Size(25, 25)
        Me.Label15.TabIndex = 75
        Me.Label15.Text = "Top"
        '
        'Label16
        '
        Me.Label16.BackColor = System.Drawing.SystemColors.Control
        Me.Label16.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label16.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label16.Location = New System.Drawing.Point(8, 136)
        Me.Label16.Name = "Label16"
        Me.Label16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label16.Size = New System.Drawing.Size(41, 17)
        Me.Label16.TabIndex = 74
        Me.Label16.Text = "Text 1"
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.SystemColors.Control
        Me.Label17.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label17.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label17.Location = New System.Drawing.Point(8, 168)
        Me.Label17.Name = "Label17"
        Me.Label17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label17.Size = New System.Drawing.Size(41, 17)
        Me.Label17.TabIndex = 73
        Me.Label17.Text = "Text 2"
        '
        'Label18
        '
        Me.Label18.BackColor = System.Drawing.SystemColors.Control
        Me.Label18.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label18.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label18.Location = New System.Drawing.Point(8, 264)
        Me.Label18.Name = "Label18"
        Me.Label18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label18.Size = New System.Drawing.Size(57, 25)
        Me.Label18.TabIndex = 72
        Me.Label18.Text = "Text Style"
        '
        'Label19
        '
        Me.Label19.BackColor = System.Drawing.SystemColors.Control
        Me.Label19.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label19.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label19.Location = New System.Drawing.Point(8, 296)
        Me.Label19.Name = "Label19"
        Me.Label19.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label19.Size = New System.Drawing.Size(33, 17)
        Me.Label19.TabIndex = 71
        Me.Label19.Text = "Color"
        '
        'Label20
        '
        Me.Label20.BackColor = System.Drawing.SystemColors.Control
        Me.Label20.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label20.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label20.Location = New System.Drawing.Point(8, 328)
        Me.Label20.Name = "Label20"
        Me.Label20.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label20.Size = New System.Drawing.Size(105, 17)
        Me.Label20.TabIndex = 70
        Me.Label20.Text = "Outline Border Color"
        '
        'Label21
        '
        Me.Label21.BackColor = System.Drawing.SystemColors.Control
        Me.Label21.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label21.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label21.Location = New System.Drawing.Point(8, 360)
        Me.Label21.Name = "Label21"
        Me.Label21.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label21.Size = New System.Drawing.Size(97, 17)
        Me.Label21.TabIndex = 69
        Me.Label21.Text = "Outline Back Color"
        '
        'Label22
        '
        Me.Label22.BackColor = System.Drawing.SystemColors.Control
        Me.Label22.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label22.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label22.Location = New System.Drawing.Point(8, 488)
        Me.Label22.Name = "Label22"
        Me.Label22.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label22.Size = New System.Drawing.Size(161, 17)
        Me.Label22.TabIndex = 68
        Me.Label22.Text = "Texture Brush Image"
        '
        'Label23
        '
        Me.Label23.BackColor = System.Drawing.SystemColors.Control
        Me.Label23.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label23.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label23.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label23.Location = New System.Drawing.Point(8, 384)
        Me.Label23.Name = "Label23"
        Me.Label23.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label23.Size = New System.Drawing.Size(97, 17)
        Me.Label23.TabIndex = 67
        Me.Label23.Text = "Hash Brush"
        '
        'Label24
        '
        Me.Label24.BackColor = System.Drawing.SystemColors.Control
        Me.Label24.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label24.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label24.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label24.Location = New System.Drawing.Point(8, 432)
        Me.Label24.Name = "Label24"
        Me.Label24.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label24.Size = New System.Drawing.Size(89, 17)
        Me.Label24.TabIndex = 66
        Me.Label24.Text = "Hash Fore Color"
        '
        'Label25
        '
        Me.Label25.BackColor = System.Drawing.SystemColors.Control
        Me.Label25.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label25.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label25.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label25.Location = New System.Drawing.Point(8, 456)
        Me.Label25.Name = "Label25"
        Me.Label25.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label25.Size = New System.Drawing.Size(89, 17)
        Me.Label25.TabIndex = 65
        Me.Label25.Text = "Hash Back Color"
        '
        'txtSat
        '
        Me.txtSat.AcceptsReturn = True
        Me.txtSat.BackColor = System.Drawing.SystemColors.Window
        Me.txtSat.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSat.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSat.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSat.Location = New System.Drawing.Point(912, 624)
        Me.txtSat.MaxLength = 0
        Me.txtSat.Name = "txtSat"
        Me.txtSat.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSat.Size = New System.Drawing.Size(41, 20)
        Me.txtSat.TabIndex = 46
        Me.txtSat.Text = "1"
        '
        'Command34
        '
        Me.Command34.BackColor = System.Drawing.SystemColors.Control
        Me.Command34.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command34.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command34.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command34.Location = New System.Drawing.Point(952, 688)
        Me.Command34.Name = "Command34"
        Me.Command34.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command34.Size = New System.Drawing.Size(33, 25)
        Me.Command34.TabIndex = 41
        Me.Command34.Text = "Set"
        Me.Command34.UseVisualStyleBackColor = False
        '
        'Command35
        '
        Me.Command35.BackColor = System.Drawing.SystemColors.Control
        Me.Command35.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command35.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command35.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command35.Location = New System.Drawing.Point(952, 656)
        Me.Command35.Name = "Command35"
        Me.Command35.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command35.Size = New System.Drawing.Size(33, 25)
        Me.Command35.TabIndex = 40
        Me.Command35.Text = "Set"
        Me.Command35.UseVisualStyleBackColor = False
        '
        'txtgamma
        '
        Me.txtgamma.AcceptsReturn = True
        Me.txtgamma.BackColor = System.Drawing.SystemColors.Window
        Me.txtgamma.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtgamma.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtgamma.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtgamma.Location = New System.Drawing.Point(912, 688)
        Me.txtgamma.MaxLength = 0
        Me.txtgamma.Name = "txtgamma"
        Me.txtgamma.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtgamma.Size = New System.Drawing.Size(41, 20)
        Me.txtgamma.TabIndex = 39
        Me.txtgamma.Text = "1"
        '
        'txthue
        '
        Me.txthue.AcceptsReturn = True
        Me.txthue.BackColor = System.Drawing.SystemColors.Window
        Me.txthue.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txthue.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txthue.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txthue.Location = New System.Drawing.Point(912, 656)
        Me.txthue.MaxLength = 0
        Me.txthue.Name = "txthue"
        Me.txthue.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txthue.Size = New System.Drawing.Size(41, 20)
        Me.txthue.TabIndex = 38
        Me.txthue.Text = "0"
        '
        'txtcontrast
        '
        Me.txtcontrast.AcceptsReturn = True
        Me.txtcontrast.BackColor = System.Drawing.SystemColors.Window
        Me.txtcontrast.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtcontrast.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtcontrast.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtcontrast.Location = New System.Drawing.Point(912, 592)
        Me.txtcontrast.MaxLength = 0
        Me.txtcontrast.Name = "txtcontrast"
        Me.txtcontrast.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtcontrast.Size = New System.Drawing.Size(41, 20)
        Me.txtcontrast.TabIndex = 37
        Me.txtcontrast.Text = "1"
        '
        'Command36
        '
        Me.Command36.BackColor = System.Drawing.SystemColors.Control
        Me.Command36.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command36.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command36.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command36.Location = New System.Drawing.Point(952, 592)
        Me.Command36.Name = "Command36"
        Me.Command36.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command36.Size = New System.Drawing.Size(33, 25)
        Me.Command36.TabIndex = 36
        Me.Command36.Text = "Set"
        Me.Command36.UseVisualStyleBackColor = False
        '
        'Command38
        '
        Me.Command38.BackColor = System.Drawing.SystemColors.Control
        Me.Command38.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command38.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command38.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command38.Location = New System.Drawing.Point(952, 624)
        Me.Command38.Name = "Command38"
        Me.Command38.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command38.Size = New System.Drawing.Size(33, 25)
        Me.Command38.TabIndex = 35
        Me.Command38.Text = "Set"
        Me.Command38.UseVisualStyleBackColor = False
        '
        'txtBrightness
        '
        Me.txtBrightness.AcceptsReturn = True
        Me.txtBrightness.BackColor = System.Drawing.SystemColors.Window
        Me.txtBrightness.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtBrightness.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBrightness.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtBrightness.Location = New System.Drawing.Point(912, 560)
        Me.txtBrightness.MaxLength = 0
        Me.txtBrightness.Name = "txtBrightness"
        Me.txtBrightness.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtBrightness.Size = New System.Drawing.Size(41, 18)
        Me.txtBrightness.TabIndex = 33
        Me.txtBrightness.Text = "0"
        '
        'Command37
        '
        Me.Command37.BackColor = System.Drawing.SystemColors.Control
        Me.Command37.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command37.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command37.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command37.Location = New System.Drawing.Point(952, 560)
        Me.Command37.Name = "Command37"
        Me.Command37.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command37.Size = New System.Drawing.Size(33, 25)
        Me.Command37.TabIndex = 32
        Me.Command37.Text = "Set"
        Me.Command37.UseVisualStyleBackColor = False
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
        Me.Text2.Location = New System.Drawing.Point(424, 568)
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
        Me.Command31.Location = New System.Drawing.Point(480, 568)
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
        Me.Frame2.Location = New System.Drawing.Point(8, 687)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(98, 40)
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
        Me._Frame1_0.Location = New System.Drawing.Point(8, 647)
        Me._Frame1_0.Name = "_Frame1_0"
        Me._Frame1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame1_0.Size = New System.Drawing.Size(97, 35)
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
        Me.optborderon.Location = New System.Drawing.Point(8, 9)
        Me.optborderon.Name = "optborderon"
        Me.optborderon.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optborderon.Size = New System.Drawing.Size(35, 25)
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
        Me.optborderoff.Location = New System.Drawing.Point(48, 9)
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
        Me.Picture1.Location = New System.Drawing.Point(593, 584)
        Me.Picture1.Name = "Picture1"
        Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Picture1.Size = New System.Drawing.Size(112, 56)
        Me.Picture1.TabIndex = 17
        Me.Picture1.TabStop = False
        '
        '_Frame1_1
        '
        Me._Frame1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Frame1_1.Controls.Add(Me.chkpanning)
        Me._Frame1_1.Controls.Add(Me.GroupBox1)
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
        Me._Frame1_1.Size = New System.Drawing.Size(643, 168)
        Me._Frame1_1.TabIndex = 13
        Me._Frame1_1.TabStop = False
        Me._Frame1_1.Text = "Image Source "
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtbufferresizescale)
        Me.GroupBox1.Controls.Add(Me.Label48)
        Me.GroupBox1.Controls.Add(Me.Label47)
        Me.GroupBox1.Controls.Add(Me.Label46)
        Me.GroupBox1.Controls.Add(Me.txtbufferresizemaxwidth)
        Me.GroupBox1.Controls.Add(Me.Label45)
        Me.GroupBox1.Controls.Add(Me.chkbufferresize)
        Me.GroupBox1.Location = New System.Drawing.Point(540, 28)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(102, 134)
        Me.GroupBox1.TabIndex = 123
        Me.GroupBox1.TabStop = False
        '
        'txtbufferresizescale
        '
        Me.txtbufferresizescale.Location = New System.Drawing.Point(7, 111)
        Me.txtbufferresizescale.Name = "txtbufferresizescale"
        Me.txtbufferresizescale.Size = New System.Drawing.Size(44, 18)
        Me.txtbufferresizescale.TabIndex = 6
        Me.txtbufferresizescale.Text = "0.5"
        '
        'Label48
        '
        Me.Label48.AutoSize = True
        Me.Label48.Location = New System.Drawing.Point(10, 95)
        Me.Label48.Name = "Label48"
        Me.Label48.Size = New System.Drawing.Size(77, 12)
        Me.Label48.TabIndex = 5
        Me.Label48.Text = "0.5 = resize 50% "
        '
        'Label47
        '
        Me.Label47.AutoSize = True
        Me.Label47.Location = New System.Drawing.Point(9, 79)
        Me.Label47.Name = "Label47"
        Me.Label47.Size = New System.Drawing.Size(57, 12)
        Me.Label47.TabIndex = 4
        Me.Label47.Text = "Scale Image"
        '
        'Label46
        '
        Me.Label46.AutoSize = True
        Me.Label46.Location = New System.Drawing.Point(54, 60)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(24, 12)
        Me.Label46.TabIndex = 3
        Me.Label46.Text = "pixel"
        '
        'txtbufferresizemaxwidth
        '
        Me.txtbufferresizemaxwidth.Location = New System.Drawing.Point(6, 58)
        Me.txtbufferresizemaxwidth.Name = "txtbufferresizemaxwidth"
        Me.txtbufferresizemaxwidth.Size = New System.Drawing.Size(44, 18)
        Me.txtbufferresizemaxwidth.TabIndex = 2
        Me.txtbufferresizemaxwidth.Text = "1200"
        '
        'Label45
        '
        Me.Label45.AutoSize = True
        Me.Label45.Location = New System.Drawing.Point(4, 42)
        Me.Label45.Name = "Label45"
        Me.Label45.Size = New System.Drawing.Size(74, 12)
        Me.Label45.TabIndex = 1
        Me.Label45.Text = "width of image >"
        '
        'chkbufferresize
        '
        Me.chkbufferresize.AutoSize = True
        Me.chkbufferresize.Checked = True
        Me.chkbufferresize.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkbufferresize.Location = New System.Drawing.Point(7, 17)
        Me.chkbufferresize.Name = "chkbufferresize"
        Me.chkbufferresize.Size = New System.Drawing.Size(82, 16)
        Me.chkbufferresize.TabIndex = 0
        Me.chkbufferresize.Text = "Buffer Resize"
        Me.chkbufferresize.UseVisualStyleBackColor = True
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 14
        Me.ListBox1.Location = New System.Drawing.Point(543, 205)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(107, 18)
        Me.ListBox1.TabIndex = 122
        '
        'Label44
        '
        Me.Label44.AutoSize = True
        Me.Label44.Location = New System.Drawing.Point(543, 177)
        Me.Label44.Name = "Label44"
        Me.Label44.Size = New System.Drawing.Size(79, 14)
        Me.Label44.TabIndex = 121
        Me.Label44.Text = "Scanning Error"
        '
        'Label43
        '
        Me.Label43.AutoSize = True
        Me.Label43.Location = New System.Drawing.Point(399, 141)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(81, 12)
        Me.Label43.TabIndex = 120
        Me.Label43.Text = "last page scanned"
        '
        'chkautosaveallpages
        '
        Me.chkautosaveallpages.AutoSize = True
        Me.chkautosaveallpages.Location = New System.Drawing.Point(384, 126)
        Me.chkautosaveallpages.Name = "chkautosaveallpages"
        Me.chkautosaveallpages.Size = New System.Drawing.Size(145, 16)
        Me.chkautosaveallpages.TabIndex = 119
        Me.chkautosaveallpages.Text = "Automate save to PDF when"
        Me.chkautosaveallpages.UseVisualStyleBackColor = True
        '
        'chkenableduplex
        '
        Me.chkenableduplex.Location = New System.Drawing.Point(384, 153)
        Me.chkenableduplex.Name = "chkenableduplex"
        Me.chkenableduplex.Size = New System.Drawing.Size(152, 16)
        Me.chkenableduplex.TabIndex = 118
        Me.chkenableduplex.Text = "Enable Duplex"
        '
        'chkenablefeeder
        '
        Me.chkenablefeeder.Location = New System.Drawing.Point(384, 108)
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
        Me.Check4.Location = New System.Drawing.Point(384, 85)
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
        Me.Command4.Location = New System.Drawing.Point(392, 42)
        Me.Command4.Name = "Command4"
        Me.Command4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command4.Size = New System.Drawing.Size(145, 24)
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
        Me.cbodpi.Location = New System.Drawing.Point(254, 40)
        Me.cbodpi.Name = "cbodpi"
        Me.cbodpi.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cbodpi.Size = New System.Drawing.Size(132, 20)
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
        Me.Check2.Location = New System.Drawing.Point(384, 71)
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
        Me.Label10.Location = New System.Drawing.Point(223, 43)
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
        'Command12
        '
        Me.Command12.BackColor = System.Drawing.SystemColors.Control
        Me.Command12.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command12.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command12.Location = New System.Drawing.Point(8, 272)
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
        Me.Command11.Location = New System.Drawing.Point(8, 304)
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
        Me.Command7.Location = New System.Drawing.Point(8, 400)
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
        Me.Command6.Location = New System.Drawing.Point(8, 432)
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
        Me.Command8.Location = New System.Drawing.Point(8, 336)
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
        Me.Command9.Location = New System.Drawing.Point(8, 368)
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
        Me.Command10.Location = New System.Drawing.Point(8, 464)
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
        Me.Command5.Location = New System.Drawing.Point(8, 496)
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
        Me.Command20.Location = New System.Drawing.Point(8, 528)
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
        Me.Command3.Location = New System.Drawing.Point(8, 240)
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
        Me.Command2.Location = New System.Drawing.Point(8, 208)
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
        Me.Command1.Location = New System.Drawing.Point(8, 176)
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
        Me.Scanner1.Location = New System.Drawing.Point(112, 168)
        Me.Scanner1.Name = "Scanner1"
        Me.Scanner1.OcxState = CType(resources.GetObject("Scanner1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.Scanner1.Size = New System.Drawing.Size(425, 381)
        Me.Scanner1.TabIndex = 0
        '
        'Label34
        '
        Me.Label34.BackColor = System.Drawing.SystemColors.Control
        Me.Label34.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label34.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label34.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label34.Location = New System.Drawing.Point(112, 592)
        Me.Label34.Name = "Label34"
        Me.Label34.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label34.Size = New System.Drawing.Size(440, 32)
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
        Me.Label33.Location = New System.Drawing.Point(352, 568)
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
        'Label26
        '
        Me.Label26.BackColor = System.Drawing.SystemColors.Control
        Me.Label26.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label26.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label26.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label26.Location = New System.Drawing.Point(720, 592)
        Me.Label26.Name = "Label26"
        Me.Label26.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label26.Size = New System.Drawing.Size(185, 17)
        Me.Label26.TabIndex = 45
        Me.Label26.Text = "Contrast Value From 0.1 - 4.0"
        '
        'Label28
        '
        Me.Label28.BackColor = System.Drawing.SystemColors.Control
        Me.Label28.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label28.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label28.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label28.Location = New System.Drawing.Point(720, 624)
        Me.Label28.Name = "Label28"
        Me.Label28.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label28.Size = New System.Drawing.Size(185, 25)
        Me.Label28.TabIndex = 44
        Me.Label28.Text = "Saturation Value From 0.0 - 3.0"
        '
        'Label29
        '
        Me.Label29.BackColor = System.Drawing.SystemColors.Control
        Me.Label29.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label29.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label29.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label29.Location = New System.Drawing.Point(720, 656)
        Me.Label29.Name = "Label29"
        Me.Label29.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label29.Size = New System.Drawing.Size(185, 25)
        Me.Label29.TabIndex = 43
        Me.Label29.Text = "Hue Value From -180 - 180"
        '
        'Label30
        '
        Me.Label30.BackColor = System.Drawing.SystemColors.Control
        Me.Label30.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label30.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label30.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label30.Location = New System.Drawing.Point(720, 688)
        Me.Label30.Name = "Label30"
        Me.Label30.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label30.Size = New System.Drawing.Size(185, 25)
        Me.Label30.TabIndex = 42
        Me.Label30.Text = "Gamma Value From 0.1 - 4.0"
        '
        'Label27
        '
        Me.Label27.BackColor = System.Drawing.SystemColors.Control
        Me.Label27.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label27.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label27.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label27.Location = New System.Drawing.Point(720, 560)
        Me.Label27.Name = "Label27"
        Me.Label27.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label27.Size = New System.Drawing.Size(185, 25)
        Me.Label27.TabIndex = 34
        Me.Label27.Text = "Brightness Value From -1.0 - 1.0"
        '
        'lblcurrentpage
        '
        Me.lblcurrentpage.BackColor = System.Drawing.SystemColors.Control
        Me.lblcurrentpage.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblcurrentpage.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblcurrentpage.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblcurrentpage.Location = New System.Drawing.Point(200, 568)
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
        Me.Label8.Location = New System.Drawing.Point(248, 568)
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
        Me.lbltotalpage.Location = New System.Drawing.Point(264, 568)
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
        Me.Label3.Location = New System.Drawing.Point(112, 568)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(73, 17)
        Me.Label3.TabIndex = 28
        Me.Label3.Text = "Total Pages"
        '
        'Label41
        '
        Me.Label41.Location = New System.Drawing.Point(288, 680)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(104, 16)
        Me.Label41.TabIndex = 116
        Me.Label41.Text = "Confidence(0-100)"
        '
        'txtconfid
        '
        Me.txtconfid.Location = New System.Drawing.Point(288, 704)
        Me.txtconfid.Name = "txtconfid"
        Me.txtconfid.Size = New System.Drawing.Size(80, 20)
        Me.txtconfid.TabIndex = 117
        Me.txtconfid.Text = "99"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(408, 672)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(160, 40)
        Me.Button1.TabIndex = 118
        Me.Button1.Text = "Check Current is Blank page (depend on Confidence value)"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(112, 672)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(144, 56)
        Me.Button2.TabIndex = 119
        Me.Button2.Text = "Get Confidence of Current page "
        '
        'chkpanning
        '
        Me.chkpanning.AutoSize = True
        Me.chkpanning.Checked = True
        Me.chkpanning.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkpanning.Location = New System.Drawing.Point(225, 141)
        Me.chkpanning.Name = "chkpanning"
        Me.chkpanning.Size = New System.Drawing.Size(137, 16)
        Me.chkpanning.TabIndex = 124
        Me.chkpanning.Text = "Enable Panning with Mouse"
        Me.chkpanning.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(1088, 679)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label44)
        Me.Controls.Add(Me.txtconfid)
        Me.Controls.Add(Me.Label41)
        Me.Controls.Add(Me.Command17)
        Me.Controls.Add(Me.Frame4)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.txtLeft)
        Me.Controls.Add(Me.txtTop)
        Me.Controls.Add(Me.Text1)
        Me.Controls.Add(Me.txtSat)
        Me.Controls.Add(Me.txtgamma)
        Me.Controls.Add(Me.txthue)
        Me.Controls.Add(Me.txtcontrast)
        Me.Controls.Add(Me.txtBrightness)
        Me.Controls.Add(Me.Text2)
        Me.Controls.Add(Me.Label31)
        Me.Controls.Add(Me.Frame6)
        Me.Controls.Add(Me.Command34)
        Me.Controls.Add(Me.Command35)
        Me.Controls.Add(Me.Command36)
        Me.Controls.Add(Me.Command38)
        Me.Controls.Add(Me.Command37)
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
        Me.Controls.Add(Me.Label26)
        Me.Controls.Add(Me.Label28)
        Me.Controls.Add(Me.Label29)
        Me.Controls.Add(Me.Label30)
        Me.Controls.Add(Me.Label27)
        Me.Controls.Add(Me.lblcurrentpage)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.lbltotalpage)
        Me.Controls.Add(Me.Label3)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 30)
        Me.Name = "Form1"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Scanner Pro ActiveX Control"
        Me.Frame4.ResumeLayout(False)
        Me.Frame4.PerformLayout()
        Me.Frame3.ResumeLayout(False)
        Me.Frame3.PerformLayout()
        Me.Frame6.ResumeLayout(False)
        Me.Frame6.PerformLayout()
        Me.Frame2.ResumeLayout(False)
        Me._Frame1_0.ResumeLayout(False)
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).EndInit()
        Me._Frame1_1.ResumeLayout(False)
        Me._Frame1_1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.Frame5.ResumeLayout(False)
        Me.Frame5.PerformLayout()
        CType(Me.Scanner1, System.ComponentModel.ISupportInitialize).EndInit()
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

    Private Sub cbofont_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbofont.SelectedIndexChanged
        Scanner1.TextFontName = cbofont.Items(cbofont.SelectedIndex)
    End Sub

    Private Sub cbofontsize_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbofontsize.SelectedIndexChanged
        Scanner1.TextFontSize = Val(cbofontsize.Items(cbofontsize.SelectedIndex))

    End Sub



    Private Sub cbofontstyle_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbofontstyle.SelectedIndexChanged
        Scanner1.TextFontStyle = cbofontstyle.SelectedIndex
    End Sub

    Private Sub cbohashbrushstyle_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbohashbrushstyle.SelectedIndexChanged

    End Sub

    'UPGRADE_WARNING: Event cbotextstyle.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
    Private Sub cbotextstyle_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbotextstyle.SelectedIndexChanged
        Select Case cbotextstyle.SelectedIndex

            Case 0
                cmdtextcolor.Enabled = True
                cmdoutlineborder.Enabled = False
                cmdoutlineback.Enabled = False
                cbohashbrushstyle.Enabled = False
                cmdhashforecolor.Enabled = False
                cmdhashbackcolor.Enabled = False
                cmdtextureimage.Enabled = False
            Case 1
                cmdtextcolor.Enabled = False
                cmdoutlineborder.Enabled = True
                cmdoutlineback.Enabled = False
                cbohashbrushstyle.Enabled = False
                cmdhashforecolor.Enabled = False
                cmdhashbackcolor.Enabled = False
                cmdtextureimage.Enabled = False

            Case 2
                cmdtextcolor.Enabled = False
                cmdoutlineborder.Enabled = True
                cmdoutlineback.Enabled = True
                cbohashbrushstyle.Enabled = False
                cmdhashforecolor.Enabled = False
                cmdhashbackcolor.Enabled = False
                cmdtextureimage.Enabled = False

            Case 3
                cmdtextcolor.Enabled = False
                cmdoutlineborder.Enabled = False
                cmdoutlineback.Enabled = False
                cbohashbrushstyle.Enabled = True
                cmdhashforecolor.Enabled = True
                cmdhashbackcolor.Enabled = True
                cmdtextureimage.Enabled = False

            Case 4

                cmdtextcolor.Enabled = False
                cmdoutlineborder.Enabled = False
                cmdoutlineback.Enabled = False
                cbohashbrushstyle.Enabled = False
                cmdhashforecolor.Enabled = False
                cmdhashbackcolor.Enabled = False
                cmdtextureimage.Enabled = True

                If txttextureimage.Text = "" Then
                    MsgBox("Please select texture brush")
                    Exit Sub
                End If

                Me.Scanner1.SetTextureBrushImage(txttextureimage.Text)


        End Select


        Me.Scanner1.TextStyle = cbotextstyle.SelectedIndex
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
    Private Sub chkshowtext_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkshowtext.CheckStateChanged
    End Sub

    Private Sub cmdhashbackcolor_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdhashbackcolor.Click
        If ColorDialog1.ShowDialog() Then
            clrHashBackColor = ColorDialog1.Color
            Scanner1.SetHashBrushValue(cbohashbrushstyle.SelectedIndex, clrHashForeColor, clrHashBackColor)
        End If

    End Sub

    Private Sub cmdhashforecolor_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdhashforecolor.Click
        If ColorDialog1.ShowDialog() Then
            clrHashForeColor = ColorDialog1.Color
            Scanner1.SetHashBrushValue(cbohashbrushstyle.SelectedIndex, clrHashForeColor, clrHashBackColor)
        End If

    End Sub

    Private Sub cmdoutlineback_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdoutlineback.Click
        If ColorDialog1.ShowDialog() Then
            Scanner1.SetOutlineTextBackColor(Color2Uint32(ColorDialog1.Color))
        End If

    End Sub

    Private Sub cmdoutlineborder_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdoutlineborder.Click
        If ColorDialog1.ShowDialog() Then
            Scanner1.SetOutlineTextBorderColor(Color2Uint32(ColorDialog1.Color))
        End If

    End Sub

    Private Sub cmdtextcolor_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdtextcolor.Click
        If ColorDialog1.ShowDialog() Then
            Me.Scanner1.TextColor = ColorDialog1.Color
        End If


    End Sub

    Private Sub cmdtextureimage_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdtextureimage.Click


        OpenFileDialog1.Filter = "All Files (*.*)|*.*|JPEG (*.jpg)|*.jpg|Bitmap (*.bmp)|*.bmp|TIF (*.tif)|*.tif|Gif (*.gif)|*.gif |PNG (*.png)|*.png|Icon (*.ico)|*.ico"
        If OpenFileDialog1.ShowDialog() Then
            txttextureimage.Text = OpenFileDialog1.FileName
            Me.Scanner1.SetTextureBrushImage(txttextureimage.Text)
        End If

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

        Scanner1.View = 5
        Scanner1.TIFCompression = cbotifcompression.SelectedIndex

        Me.SaveFileDialog1.Filter = "TIF File (*.tif)|*.tif"

        If Me.SaveFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then

            Dim strTmpFile As String
            strTmpFile = SaveFileDialog1.FileName

            a = Me.Scanner1.SaveAllPage2TIF(strTmpFile, Check3.CheckState, 1)

            If a Then
                MessageBox.Show("save " + strTmpFile + " complete")
            End If
        End If

    End Sub

    Sub AddPDFAnnotation()
        Scanner1.PDFInitAnnotation() ' init PDF annotation, you must call it
        Dim FontID As Long
        Dim FontID2 As Long
      

        ' draw rectangle at 1 cm x 5cm, width 20 cm, height 2 cm, borderwidth 0.1cm
        Scanner1.PDFDrawFillRectangle(0, 10000, 50000, 200000, 20000, 1000, 255, 0, 0)

        FontID = Scanner1.PDFAddFont("Arial", True, True)
        Scanner1.PDFEmbedFont(FontID)



        Scanner1.PDFSetTextColor(255, 255, 255)

        ' draw text at 1cm x 4cm
        Scanner1.PDFDrawText(0, 10000, 40000, "Scanner Pro's PDF annotation features", FontID, 20, 0)


        ' draw line from 2cm x 0.9 cm to 20cm x 0.9 cm
        Scanner1.PDFDrawLine(0, 20000, 9000, 200000, 9000, 1000, 0, 0, 0)


        FontID2 = Scanner1.PDFAddFont("Times New Roman", True, True)

        Scanner1.PDFSetTextColor(0, 255, 0)

        ' draw text at  1cm x 6cm and center alignment
        Scanner1.PDFDrawTextAlign(0, 10000, 60000, "www.viscomsoft.com", FontID2, 18, 1)


        Scanner1.PDFSetTextColor(255, 0, 0)

        ' draw text at specificed centimeters 3cm x 20cm
        Scanner1.PDFDrawText(0, 30000, 200000, "Rotated Text 1234567890", FontID, 40, 45)


        Scanner1.PDFDrawRectangle(0, 30000, 150000, 50000, 20000, 1000, 255, 255, 0)

        Scanner1.PDFDrawRectangle(0, 20000, 140000, 30000, 20000, 1000, 0, 255, 0)

        Scanner1.PDFDrawRectangle(0, 50000, 120000, 50000, 20000, 1000, 255, 100, 0)



    End Sub

    Sub PdfSecureSetting()


        Me.Scanner1.PDFAuthor = "my author"
        Me.Scanner1.PDFCreator = "Created from Scanner Pro ActiveX"
        Me.Scanner1.PDFKeyword = "my keyword"
        Me.Scanner1.PDFProducer = "my producer"
        Me.Scanner1.PDFTitle = "my title"
        Me.Scanner1.PDFSubject = "my subject"

        ' use pdf secure must set output pdf/a = true
        Me.Scanner1.PDFOutputPDFA = True
    End Sub

    Private Sub SaveToMultiPagePDF()
        Dim a As Integer

        Scanner1.View = 5


        Scanner1.TIFCompression = cbotifcompression.SelectedIndex

        PdfSecureSetting()

        If Me.chkaddPDFAnnotation.Checked Then

            AddPDFAnnotation()
        End If

        If Me.chkpdfusejpegcomp2.Checked Then
            Scanner1.PDFUseJPEGCompression = True
            Scanner1.PDFJPEGQuality = Val(txtpdfjpegquality2.Text)
        Else
            Scanner1.PDFUseJPEGCompression = False

        End If

        Me.SaveFileDialog1.Filter = "PDF File (*.pdf)|*.pdf"

        If Me.SaveFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then

            Dim strTmpFile As String
            strTmpFile = SaveFileDialog1.FileName
            a = Me.Scanner1.SaveAllPage2PDF(strTmpFile, Check3.CheckState, 1)


            If a Then
                MessageBox.Show("Save " + strTmpFile + " Complete")
            Else
                MessageBox.Show("Save fail")
            End If

        End If

    End Sub
    Private Sub Command16_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command16.Click
        SaveToMultiPagePDF()

    End Sub

    Private Sub Command17_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command17.Click
        Scanner1.DeletePage(Scanner1.GetActivePageNo)


        Scanner1.SetActivePageNo(1)
        lbltotalpage.Text = CStr(Me.Scanner1.TotalPage)
        lblcurrentpage.Text = CStr(Me.Scanner1.GetActivePageNo)

    End Sub

    Private Sub Command18_Click()


    End Sub

    Private Sub Command19_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command19.Click

        Dim result As Integer
        Dim strType As String

        strType = cboutputtype.Text


        PdfSecureSetting()

        If Me.chkaddPDFAnnotation.Checked Then

            AddPDFAnnotation()
        End If

        If Me.chkpdfusejpegcomp.Checked Then
            Scanner1.PDFUseJPEGCompression = True
            Scanner1.PDFJPEGQuality = Val(txtpdfjpegquality.Text)
        Else
            Scanner1.PDFUseJPEGCompression = False

        End If

        Me.SaveFileDialog1.Filter = ""

        If Me.SaveFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then

            Dim strTmpFile As String
            strTmpFile = SaveFileDialog1.FileName

            result = Me.Scanner1.SaveBySize(strTmpFile, strType, Val(txt_Width.Text), Val(txt_Height.Text))


            If result = 1 Then
                MessageBox.Show("Save " + strTmpFile + "." + strType + " Complete")
            Else
                MessageBox.Show("Save fail")
            End If


        End If

        Scanner1.Focus()



    End Sub

    Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
        Me.Scanner1.Rotate180()
    End Sub

    Private Sub Command20_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command20.Click

        Scanner1.View = 10
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


    Private Sub Command33_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command33.Click
        Me.Scanner1.ApplyChange()
    End Sub

    Private Sub Command34_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command34.Click
        Me.Scanner1.Gamma = CDbl(txtgamma.Text)
    End Sub

    Private Sub Command35_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command35.Click
        Me.Scanner1.Hue = CDbl(txthue.Text)
    End Sub

    Private Sub Command36_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command36.Click
        Me.Scanner1.Contrast = CDbl(txtcontrast.Text)
    End Sub

    Private Sub Command37_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command37.Click
        Me.Scanner1.Brightness = CDbl(txtBrightness.Text)
    End Sub

    Private Sub Command38_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command38.Click
        Me.Scanner1.Saturation = CDbl(txtSat.Text)
    End Sub

    Private Sub Command39_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command39.Click
        Scanner1.DrawText(CShort(txtTextLeft.Text), CShort(txtTextTop.Text), txttext1.Text & Chr(13) & Chr(10) & txttext2.Text)

    End Sub

    Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command4.Click

        If chkpanning.Checked Then
            Scanner1.EnablePanning = True
        Else
            Scanner1.EnablePanning = False
        End If

        If chkbufferresize.Checked Then
            Scanner1.BufferResizeMode = True
            Scanner1.BufferResizeMaxWidth = txtbufferresizemaxwidth.Text
            Scanner1.BufferResizeScale = txtbufferresizescale.Text
        Else
            Scanner1.BufferResizeMode = False
        End If

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


        If chkdefaultcap.Checked Then
            Scanner1.SetCaptureArea(0, 0, 0, 0)
        Else
            Scanner1.SetCaptureArea(CDbl(txtcapleft.Text), CDbl(txtcaptop.Text), CDbl(txtcapwidth.Text), CDbl(txtcapheight.Text))
        End If


        Scanner1.Scan()





    End Sub

    Private Sub Command40_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command40.Click
        Dim result As Integer
        Dim strType As String


        strType = cboutputtype.Text


        PdfSecureSetting()

        If Me.chkaddPDFAnnotation.Checked Then

            AddPDFAnnotation()
        End If

        If Me.chkpdfusejpegcomp.Checked Then
            Scanner1.PDFUseJPEGCompression = True
            Scanner1.PDFJPEGQuality = Val(txtpdfjpegquality.Text)
        Else
            Scanner1.PDFUseJPEGCompression = False

        End If

        Me.SaveFileDialog1.Filter = ""

        If Me.SaveFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then

            Dim strTmpFile As String
            strTmpFile = SaveFileDialog1.FileName
            result = Me.Scanner1.Save(strTmpFile, strType)

            If result = 1 Then
                MessageBox.Show("Save " + strTmpFile + "." + strType + " Complete")
            Else
                MessageBox.Show("Save fail")
            End If

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
        cboutputtype.Items.Add("PCX")
        cboutputtype.Items.Add("TGA")
        cboutputtype.Items.Add("JP2")
        cboutputtype.Items.Add("JPC")
        cboutputtype.Items.Add("RAS")
        cboutputtype.Items.Add("PGX")
        cboutputtype.Items.Add("PNM")
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



        clrHashForeColor = RGB(0, 0, 0)
        clrHashBackColor = RGB(255, 255, 255)
        iCount = Me.Scanner1.GetNumImageSources

        For I = 0 To iCount - 1
            cboimagesource.Items.Add(Scanner1.GetImageSourceName(I))
        Next

        If iCount > 0 Then
            cboimagesource.SelectedIndex = 0
        End If

        For I = 8 To 60 Step 2
            cbofontsize.Items.Add(Trim(Str(I)))
        Next
        cbofontsize.SelectedIndex = 2


        cbofont.Items.Add("Arial")
        cbofont.Items.Add("Arial Black")
        cbofont.Items.Add("Comic Sans MS")
        cbofont.Items.Add("Times New Roman")
        cbofont.SelectedIndex = 0


        Me.cbofontstyle.Items.Add("Regular")
        Me.cbofontstyle.Items.Add("Bold")
        Me.cbofontstyle.Items.Add("Italic")
        Me.cbofontstyle.Items.Add("BoldItalic")
        Me.cbofontstyle.Items.Add("Underline")
        Me.cbofontstyle.SelectedIndex = 0

        cbotextstyle.Items.Add("Normal")
        cbotextstyle.Items.Add("Outline")
        cbotextstyle.Items.Add("Filled Outline")
        cbotextstyle.Items.Add("Hash Brush")
        cbotextstyle.Items.Add("Texture Brush")
        cbotextstyle.SelectedIndex = 0


        cbohashbrushstyle.Items.Add("HatchStyleHorizontal ")
        cbohashbrushstyle.Items.Add("HatchStyleVertical ")
        cbohashbrushstyle.Items.Add("HatchStyleForwardDiagonal ")
        cbohashbrushstyle.Items.Add("HatchStyleBackwardDiagonal ")
        cbohashbrushstyle.Items.Add("HatchStyleCross ")
        cbohashbrushstyle.Items.Add("HatchStyleDiagonalCross ")
        cbohashbrushstyle.Items.Add("HatchStyle05Percent ")
        cbohashbrushstyle.Items.Add("HatchStyle10Percent ")
        cbohashbrushstyle.Items.Add("HatchStyle20Percent ")
        cbohashbrushstyle.Items.Add("HatchStyle25Percent ")
        cbohashbrushstyle.Items.Add("HatchStyle30Percent ")
        cbohashbrushstyle.Items.Add("HatchStyle40Percent ")
        cbohashbrushstyle.Items.Add("HatchStyle50Percent ")
        cbohashbrushstyle.Items.Add("HatchStyle60Percent ")
        cbohashbrushstyle.Items.Add("HatchStyle70Percent ")
        cbohashbrushstyle.Items.Add("HatchStyle75Percent ")
        cbohashbrushstyle.Items.Add("HatchStyle80Percent ")
        cbohashbrushstyle.Items.Add("HatchStyle90Percent ")
        cbohashbrushstyle.Items.Add("HatchStyleLightDownwardDiagonal ")
        cbohashbrushstyle.Items.Add("HatchStyleLightUpwardDiagonal ")
        cbohashbrushstyle.Items.Add("HatchStyleDarkDownwardDiagonal ")
        cbohashbrushstyle.Items.Add("HatchStyleDarkUpwardDiagonal ")
        cbohashbrushstyle.Items.Add("HatchStyleWideDownwardDiagonal ")
        cbohashbrushstyle.Items.Add("HatchStyleWideUpwardDiagonal ")
        cbohashbrushstyle.Items.Add("HatchStyleLightVertical ")
        cbohashbrushstyle.Items.Add("HatchStyleLightHorizontal ")
        cbohashbrushstyle.Items.Add("HatchStyleNarrowVertical ")
        cbohashbrushstyle.Items.Add("HatchStyleNarrowHorizontal ")
        cbohashbrushstyle.Items.Add("HatchStyleDarkVertical ")
        cbohashbrushstyle.Items.Add("HatchStyleDarkHorizontal ")
        cbohashbrushstyle.Items.Add("HatchStyleDashedDownwardDiagonal ")
        cbohashbrushstyle.Items.Add("HatchStyleDashedUpwardDiagonal ")
        cbohashbrushstyle.Items.Add("HatchStyleDashedHorizontal ")
        cbohashbrushstyle.Items.Add("HatchStyleDashedVertical ")
        cbohashbrushstyle.Items.Add("HatchStyleSmallConfetti ")
        cbohashbrushstyle.Items.Add("HatchStyleLargeConfetti ")
        cbohashbrushstyle.Items.Add("HatchStyleZigZag ")
        cbohashbrushstyle.Items.Add("HatchStyleWave ")
        cbohashbrushstyle.Items.Add("HatchStyleDiagonalBrick ")
        cbohashbrushstyle.Items.Add("HatchStyleHorizontalBrick ")
        cbohashbrushstyle.Items.Add("HatchStyleWeave ")
        cbohashbrushstyle.Items.Add("HatchStylePlaid ")
        cbohashbrushstyle.Items.Add("HatchStyleDivot ")
        cbohashbrushstyle.Items.Add("HatchStyleDottedGrid ")
        cbohashbrushstyle.Items.Add("HatchStyleDottedDiamond ")
        cbohashbrushstyle.Items.Add("HatchStyleShingle ")
        cbohashbrushstyle.Items.Add("HatchStyleTrellis ")
        cbohashbrushstyle.Items.Add("HatchStyleSphere ")
        cbohashbrushstyle.Items.Add("HatchStyleSmallGrid ")
        cbohashbrushstyle.Items.Add("HatchStyleSmallCheckerBoard ")
        cbohashbrushstyle.Items.Add("HatchStyleLargeCheckerBoard ")
        cbohashbrushstyle.Items.Add("HatchStyleOutlinedDiamond ")
        cbohashbrushstyle.Items.Add("HatchStyleSolidDiamond ")
        cbohashbrushstyle.SelectedIndex = 0



    End Sub

    Private Sub Option3_Click()

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



    Private Sub chkshowtext_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkshowtext.Click
        If chkshowtext.Checked Then
            Me.Scanner1.ShowText = True
        Else
            Me.Scanner1.ShowText = False
        End If

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

    Private Sub cbohashbrushstyle_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbohashbrushstyle.Click
        Scanner1.SetHashBrushValue(cbohashbrushstyle.SelectedIndex, clrHashForeColor, clrHashBackColor)

    End Sub

    Public Function Color2Uint32(ByVal clr As Color) As UInt32
        Dim t As Int32
        Dim a() As Byte

        t = ColorTranslator.ToOle(clr)


        a = BitConverter.GetBytes(t)

        Return BitConverter.ToUInt32(a, 0)


    End Function




    Private Sub cboutputtype_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboutputtype.SelectedIndexChanged
        If cboutputtype.Text = "PDF" Then
            chkpdfusejpegcomp.Enabled = True
            txtpdfjpegquality.Enabled = True
        Else
            chkpdfusejpegcomp.Enabled = False
            txtpdfjpegquality.Enabled = False

        End If
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Scanner1.BlankPageIsBlank(CDbl(txtconfid.Text)) Then
            MessageBox.Show("This page is Blank page")
        Else
            MessageBox.Show("This page is not a Blank page")
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        txtconfid.Text = Scanner1.BlankPageGetConfidence

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim a As Integer

        Scanner1.View = 5

        If Check3.Checked Then
            Me.Scanner1.DocxAddChar(0, "T", 0, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, False, False, False)
            Me.Scanner1.DocxAddChar(0, "h", 1, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, False, False, False)
            Me.Scanner1.DocxAddChar(0, "i", 2, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, False, False, False)
            Me.Scanner1.DocxAddChar(0, "s", 3, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, False, False, False)
            Me.Scanner1.DocxAddChar(0, " ", 4, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, False, False, False)
            Me.Scanner1.DocxAddChar(0, "i", 5, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, False, False, False)
            Me.Scanner1.DocxAddChar(0, "s", 6, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, False, False, False)
            Me.Scanner1.DocxAddChar(0, " ", 7, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, False, False, False)
            Me.Scanner1.DocxAddChar(0, "t", 8, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, False, False, False)
            Me.Scanner1.DocxAddChar(0, "e", 9, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, False, False, False)
            Me.Scanner1.DocxAddChar(0, "x", 10, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, False, False, False)
            Me.Scanner1.DocxAddChar(0, "t", 11, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, False, False, False)
            Me.Scanner1.DocxAddChar(0, "1", 12, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, False, False, False)
        End If


        Me.SaveFileDialog1.Filter = "Ms Word Docx File (*.docx)|*.docx"

        If Me.SaveFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then

            Dim strTmpFile As String
            strTmpFile = SaveFileDialog1.FileName

            a = Me.Scanner1.SaveAllPage2Docx(strTmpFile, Check3.CheckState, 1)

            If a Then
                MessageBox.Show("save " + strTmpFile + " complete")
            End If
        End If

    End Sub

    Private Sub Scanner1_EndAllScan(sender As System.Object, e As System.EventArgs) Handles Scanner1.EndAllScan
        If chkautosaveallpages.Checked Then
            SaveToMultiPagePDF()
        End If
    End Sub

    Private Sub Scanner1_ScanningError(sender As System.Object, e As System.EventArgs) Handles Scanner1.ScanningError
        ListBox1.Items.Add("Error occur, may be Paper Jam")
    End Sub
End Class