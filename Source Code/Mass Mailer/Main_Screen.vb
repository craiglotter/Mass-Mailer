Imports Microsoft.Win32
Imports System.IO



Public Class Main_Screen
    Inherits System.Windows.Forms.Form

    Dim WithEvents Worker1 As Worker
    Public Delegate Sub WorkerhHandler(ByVal Result As String)
    Public Delegate Sub WorkerProgresshHandler(ByVal uniqueemails As Long, ByVal totalemails As Long, ByVal emailssent As Long)


    Private splash_loader As Splash_Screen
    Public dataloaded As Boolean = False

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerHandler
        AddHandler Worker1.WorkerProgress, AddressOf WorkerProgressHandler
    End Sub

    Public Sub New(ByVal splash As Splash_Screen)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        splash_loader = splash
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerHandler
        AddHandler Worker1.WorkerProgress, AddressOf WorkerProgressHandler
    End Sub
    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents PictureBox5 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox4 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents OFD As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtTo As System.Windows.Forms.TextBox
    Friend WithEvents txtFromDisplayName As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents txtFrom As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents txtSMTPServer As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents lstAttachment As System.Windows.Forms.ListBox
    Friend WithEvents chkFormat As System.Windows.Forms.CheckBox
    Friend WithEvents txtMessage As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents txtSubject As System.Windows.Forms.TextBox
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents txtToInputFile As System.Windows.Forms.TextBox
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents lblUniqueEmailAddresses As System.Windows.Forms.Label
    Friend WithEvents lblProcessEnd As System.Windows.Forms.Label
    Friend WithEvents lblEmailAddressesFound As System.Windows.Forms.Label
    Friend WithEvents lblProcessStart As System.Windows.Forms.Label
    Friend WithEvents lblProgramLaunched As System.Windows.Forms.Label
    Friend WithEvents lblTotalEmailsSent As System.Windows.Forms.Label
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents PictureBox6 As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Main_Screen))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.PictureBox6 = New System.Windows.Forms.PictureBox
        Me.lblTotalEmailsSent = New System.Windows.Forms.Label
        Me.Label26 = New System.Windows.Forms.Label
        Me.lblUniqueEmailAddresses = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.Label23 = New System.Windows.Forms.Label
        Me.lblProcessEnd = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.lblEmailAddressesFound = New System.Windows.Forms.Label
        Me.lblProcessStart = New System.Windows.Forms.Label
        Me.lblProgramLaunched = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.PictureBox5 = New System.Windows.Forms.PictureBox
        Me.PictureBox4 = New System.Windows.Forms.PictureBox
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.Label8 = New System.Windows.Forms.Label
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Label9 = New System.Windows.Forms.Label
        Me.Button4 = New System.Windows.Forms.Button
        Me.OFD = New System.Windows.Forms.OpenFileDialog
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Button2 = New System.Windows.Forms.Button
        Me.Label22 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.txtToInputFile = New System.Windows.Forms.TextBox
        Me.chkFormat = New System.Windows.Forms.CheckBox
        Me.txtMessage = New System.Windows.Forms.TextBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.txtSubject = New System.Windows.Forms.TextBox
        Me.Label25 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.btnRemove = New System.Windows.Forms.Button
        Me.btnAdd = New System.Windows.Forms.Button
        Me.lstAttachment = New System.Windows.Forms.ListBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.txtTo = New System.Windows.Forms.TextBox
        Me.txtFromDisplayName = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.txtFrom = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.txtSMTPServer = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Goldenrod
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.PictureBox6)
        Me.Panel1.Controls.Add(Me.lblTotalEmailsSent)
        Me.Panel1.Controls.Add(Me.Label26)
        Me.Panel1.Controls.Add(Me.lblUniqueEmailAddresses)
        Me.Panel1.Controls.Add(Me.Button3)
        Me.Panel1.Controls.Add(Me.Label23)
        Me.Panel1.Controls.Add(Me.lblProcessEnd)
        Me.Panel1.Controls.Add(Me.Label15)
        Me.Panel1.Controls.Add(Me.lblEmailAddressesFound)
        Me.Panel1.Controls.Add(Me.lblProcessStart)
        Me.Panel1.Controls.Add(Me.lblProgramLaunched)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.PictureBox5)
        Me.Panel1.Controls.Add(Me.PictureBox4)
        Me.Panel1.Controls.Add(Me.PictureBox3)
        Me.Panel1.Controls.Add(Me.PictureBox2)
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Location = New System.Drawing.Point(8, 480)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(560, 152)
        Me.Panel1.TabIndex = 11
        '
        'PictureBox6
        '
        Me.PictureBox6.Image = CType(resources.GetObject("PictureBox6.Image"), System.Drawing.Image)
        Me.PictureBox6.Location = New System.Drawing.Point(376, 8)
        Me.PictureBox6.Name = "PictureBox6"
        Me.PictureBox6.Size = New System.Drawing.Size(160, 96)
        Me.PictureBox6.TabIndex = 58
        Me.PictureBox6.TabStop = False
        '
        'lblTotalEmailsSent
        '
        Me.lblTotalEmailsSent.BackColor = System.Drawing.Color.Transparent
        Me.lblTotalEmailsSent.ForeColor = System.Drawing.Color.White
        Me.lblTotalEmailsSent.Location = New System.Drawing.Point(176, 72)
        Me.lblTotalEmailsSent.Name = "lblTotalEmailsSent"
        Me.lblTotalEmailsSent.Size = New System.Drawing.Size(256, 16)
        Me.lblTotalEmailsSent.TabIndex = 57
        Me.lblTotalEmailsSent.Text = "0"
        '
        'Label26
        '
        Me.Label26.BackColor = System.Drawing.Color.Transparent
        Me.Label26.ForeColor = System.Drawing.Color.White
        Me.Label26.Location = New System.Drawing.Point(16, 72)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(376, 16)
        Me.Label26.TabIndex = 56
        Me.Label26.Text = "Total Emails Sent:"
        '
        'lblUniqueEmailAddresses
        '
        Me.lblUniqueEmailAddresses.BackColor = System.Drawing.Color.Transparent
        Me.lblUniqueEmailAddresses.ForeColor = System.Drawing.Color.White
        Me.lblUniqueEmailAddresses.Location = New System.Drawing.Point(176, 56)
        Me.lblUniqueEmailAddresses.Name = "lblUniqueEmailAddresses"
        Me.lblUniqueEmailAddresses.Size = New System.Drawing.Size(256, 16)
        Me.lblUniqueEmailAddresses.TabIndex = 55
        Me.lblUniqueEmailAddresses.Text = "0"
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.Gold
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.ForeColor = System.Drawing.Color.Black
        Me.Button3.Location = New System.Drawing.Point(16, 112)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(96, 23)
        Me.Button3.TabIndex = 54
        Me.Button3.Text = "Send Message"
        '
        'Label23
        '
        Me.Label23.BackColor = System.Drawing.Color.Transparent
        Me.Label23.ForeColor = System.Drawing.Color.White
        Me.Label23.Location = New System.Drawing.Point(16, 56)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(376, 16)
        Me.Label23.TabIndex = 44
        Me.Label23.Text = "Unique Email Addresses:"
        '
        'lblProcessEnd
        '
        Me.lblProcessEnd.BackColor = System.Drawing.Color.Transparent
        Me.lblProcessEnd.ForeColor = System.Drawing.Color.White
        Me.lblProcessEnd.Location = New System.Drawing.Point(176, 88)
        Me.lblProcessEnd.Name = "lblProcessEnd"
        Me.lblProcessEnd.Size = New System.Drawing.Size(256, 16)
        Me.lblProcessEnd.TabIndex = 41
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.Color.Transparent
        Me.Label15.ForeColor = System.Drawing.Color.White
        Me.Label15.Location = New System.Drawing.Point(16, 88)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(392, 16)
        Me.Label15.TabIndex = 37
        Me.Label15.Text = "Process End:"
        '
        'lblEmailAddressesFound
        '
        Me.lblEmailAddressesFound.BackColor = System.Drawing.Color.Transparent
        Me.lblEmailAddressesFound.ForeColor = System.Drawing.Color.White
        Me.lblEmailAddressesFound.Location = New System.Drawing.Point(176, 40)
        Me.lblEmailAddressesFound.Name = "lblEmailAddressesFound"
        Me.lblEmailAddressesFound.Size = New System.Drawing.Size(256, 16)
        Me.lblEmailAddressesFound.TabIndex = 35
        Me.lblEmailAddressesFound.Text = "0"
        '
        'lblProcessStart
        '
        Me.lblProcessStart.BackColor = System.Drawing.Color.Transparent
        Me.lblProcessStart.ForeColor = System.Drawing.Color.White
        Me.lblProcessStart.Location = New System.Drawing.Point(176, 24)
        Me.lblProcessStart.Name = "lblProcessStart"
        Me.lblProcessStart.Size = New System.Drawing.Size(256, 16)
        Me.lblProcessStart.TabIndex = 34
        '
        'lblProgramLaunched
        '
        Me.lblProgramLaunched.BackColor = System.Drawing.Color.Transparent
        Me.lblProgramLaunched.ForeColor = System.Drawing.Color.White
        Me.lblProgramLaunched.Location = New System.Drawing.Point(176, 8)
        Me.lblProgramLaunched.Name = "lblProgramLaunched"
        Me.lblProgramLaunched.Size = New System.Drawing.Size(256, 16)
        Me.lblProgramLaunched.TabIndex = 33
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.White
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Gray
        Me.Label7.Location = New System.Drawing.Point(264, 120)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(112, 16)
        Me.Label7.TabIndex = 32
        Me.Label7.Text = "Resting..."
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.ForeColor = System.Drawing.Color.White
        Me.Label6.Location = New System.Drawing.Point(16, 40)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(376, 16)
        Me.Label6.TabIndex = 31
        Me.Label6.Text = "Email Addresses Found:"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.ForeColor = System.Drawing.Color.White
        Me.Label5.Location = New System.Drawing.Point(16, 24)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(376, 16)
        Me.Label5.TabIndex = 30
        Me.Label5.Text = "Process Start:"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(16, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(376, 16)
        Me.Label3.TabIndex = 29
        Me.Label3.Text = "Program Launched:"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.LimeGreen
        Me.Label1.Location = New System.Drawing.Point(472, 120)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 16)
        Me.Label1.TabIndex = 27
        Me.Label1.Text = "Waiting"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PictureBox5
        '
        Me.PictureBox5.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox5.BackgroundImage = CType(resources.GetObject("PictureBox5.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox5.ForeColor = System.Drawing.Color.Black
        Me.PictureBox5.Location = New System.Drawing.Point(448, 120)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox5.TabIndex = 26
        Me.PictureBox5.TabStop = False
        '
        'PictureBox4
        '
        Me.PictureBox4.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox4.BackgroundImage = CType(resources.GetObject("PictureBox4.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox4.ForeColor = System.Drawing.Color.Black
        Me.PictureBox4.Location = New System.Drawing.Point(432, 120)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox4.TabIndex = 25
        Me.PictureBox4.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox3.BackgroundImage = CType(resources.GetObject("PictureBox3.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox3.ForeColor = System.Drawing.Color.Black
        Me.PictureBox3.Location = New System.Drawing.Point(416, 120)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox3.TabIndex = 24
        Me.PictureBox3.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox2.BackgroundImage = CType(resources.GetObject("PictureBox2.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox2.ForeColor = System.Drawing.Color.Black
        Me.PictureBox2.Location = New System.Drawing.Point(400, 120)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox2.TabIndex = 23
        Me.PictureBox2.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox1.BackgroundImage = CType(resources.GetObject("PictureBox1.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox1.ForeColor = System.Drawing.Color.Black
        Me.PictureBox1.Location = New System.Drawing.Point(384, 120)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox1.TabIndex = 22
        Me.PictureBox1.TabStop = False
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.White
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(256, 112)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(288, 32)
        Me.Label2.TabIndex = 28
        '
        'ImageList1
        '
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'Timer2
        '
        Me.Timer2.Interval = 1000
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.PapayaWhip
        Me.Label8.Location = New System.Drawing.Point(440, 16)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(112, 16)
        Me.Label8.TabIndex = 33
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.ToolTip1.SetToolTip(Me.Label8, "Current System Time")
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.ContextMenu = Me.ContextMenu1
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "Resting..."
        Me.NotifyIcon1.Visible = True
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem2, Me.MenuItem3})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "Display Main Screen"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.Text = "-"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 2
        Me.MenuItem3.Text = "Exit Application"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.Transparent
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Black
        Me.Label9.Location = New System.Drawing.Point(448, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(109, 8)
        Me.Label9.TabIndex = 54
        Me.Label9.Text = "BUILD 20051109.1"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.DarkSeaGreen
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button4.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.ForeColor = System.Drawing.Color.White
        Me.Button4.Location = New System.Drawing.Point(464, 48)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(88, 16)
        Me.Button4.TabIndex = 59
        Me.Button4.Text = "KILL PROCESSES"
        '
        'OFD
        '
        Me.OFD.DefaultExt = "*.*"
        Me.OFD.InitialDirectory = "c:\"
        Me.OFD.Multiselect = True
        Me.OFD.Title = "Select the file to send as an attachment"
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.Goldenrod
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.Label10)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.Button2)
        Me.Panel2.Controls.Add(Me.Label22)
        Me.Panel2.Controls.Add(Me.Button1)
        Me.Panel2.Controls.Add(Me.txtToInputFile)
        Me.Panel2.Controls.Add(Me.chkFormat)
        Me.Panel2.Controls.Add(Me.txtMessage)
        Me.Panel2.Controls.Add(Me.Label21)
        Me.Panel2.Controls.Add(Me.txtSubject)
        Me.Panel2.Controls.Add(Me.Label25)
        Me.Panel2.Controls.Add(Me.Label13)
        Me.Panel2.Controls.Add(Me.btnRemove)
        Me.Panel2.Controls.Add(Me.btnAdd)
        Me.Panel2.Controls.Add(Me.lstAttachment)
        Me.Panel2.Controls.Add(Me.Label14)
        Me.Panel2.Controls.Add(Me.txtTo)
        Me.Panel2.Controls.Add(Me.txtFromDisplayName)
        Me.Panel2.Controls.Add(Me.Label16)
        Me.Panel2.Controls.Add(Me.txtFrom)
        Me.Panel2.Controls.Add(Me.Label17)
        Me.Panel2.Controls.Add(Me.txtSMTPServer)
        Me.Panel2.Controls.Add(Me.Label18)
        Me.Panel2.Controls.Add(Me.Label20)
        Me.Panel2.Controls.Add(Me.Label9)
        Me.Panel2.Controls.Add(Me.Button4)
        Me.Panel2.Controls.Add(Me.Label8)
        Me.Panel2.Location = New System.Drawing.Point(8, 8)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(560, 464)
        Me.Panel2.TabIndex = 81
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.Transparent
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.Gold
        Me.Label10.Location = New System.Drawing.Point(448, 144)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(32, 8)
        Me.Label10.TabIndex = 94
        Me.Label10.Text = "CLEAR"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(0, Byte), CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(32, 120)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(16, 8)
        Me.Label4.TabIndex = 93
        Me.Label4.Text = "OR"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.LimeGreen
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.White
        Me.Button2.Location = New System.Drawing.Point(464, 64)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(88, 16)
        Me.Button2.TabIndex = 92
        Me.Button2.Text = "EXAMPLE USAGE"
        '
        'Label22
        '
        Me.Label22.Location = New System.Drawing.Point(8, 136)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(104, 16)
        Me.Label22.TabIndex = 91
        Me.Label22.Text = "To  (Input Text File)"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.DarkGoldenrod
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Location = New System.Drawing.Point(368, 136)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 20)
        Me.Button1.TabIndex = 90
        Me.Button1.Text = "Browse"
        '
        'txtToInputFile
        '
        Me.txtToInputFile.BackColor = System.Drawing.Color.LemonChiffon
        Me.txtToInputFile.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtToInputFile.ForeColor = System.Drawing.Color.Sienna
        Me.txtToInputFile.Location = New System.Drawing.Point(128, 136)
        Me.txtToInputFile.Name = "txtToInputFile"
        Me.txtToInputFile.ReadOnly = True
        Me.txtToInputFile.Size = New System.Drawing.Size(232, 20)
        Me.txtToInputFile.TabIndex = 89
        Me.txtToInputFile.Text = ""
        '
        'chkFormat
        '
        Me.chkFormat.BackColor = System.Drawing.Color.Goldenrod
        Me.chkFormat.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.chkFormat.ForeColor = System.Drawing.Color.White
        Me.chkFormat.Location = New System.Drawing.Point(400, 432)
        Me.chkFormat.Name = "chkFormat"
        Me.chkFormat.TabIndex = 88
        Me.chkFormat.Text = "Send AS HTML"
        '
        'txtMessage
        '
        Me.txtMessage.BackColor = System.Drawing.Color.LemonChiffon
        Me.txtMessage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMessage.ForeColor = System.Drawing.Color.Sienna
        Me.txtMessage.Location = New System.Drawing.Point(128, 288)
        Me.txtMessage.Multiline = True
        Me.txtMessage.Name = "txtMessage"
        Me.txtMessage.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtMessage.Size = New System.Drawing.Size(368, 144)
        Me.txtMessage.TabIndex = 87
        Me.txtMessage.Text = ""
        '
        'Label21
        '
        Me.Label21.BackColor = System.Drawing.Color.Transparent
        Me.Label21.ForeColor = System.Drawing.Color.White
        Me.Label21.Location = New System.Drawing.Point(8, 288)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(104, 16)
        Me.Label21.TabIndex = 86
        Me.Label21.Text = "Message"
        '
        'txtSubject
        '
        Me.txtSubject.BackColor = System.Drawing.Color.LemonChiffon
        Me.txtSubject.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSubject.ForeColor = System.Drawing.Color.Sienna
        Me.txtSubject.Location = New System.Drawing.Point(128, 256)
        Me.txtSubject.Name = "txtSubject"
        Me.txtSubject.Size = New System.Drawing.Size(368, 20)
        Me.txtSubject.TabIndex = 85
        Me.txtSubject.Text = ""
        '
        'Label25
        '
        Me.Label25.BackColor = System.Drawing.Color.Transparent
        Me.Label25.ForeColor = System.Drawing.Color.White
        Me.Label25.Location = New System.Drawing.Point(8, 256)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(104, 16)
        Me.Label25.TabIndex = 84
        Me.Label25.Text = "Subject"
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.Color.Transparent
        Me.Label13.ForeColor = System.Drawing.Color.White
        Me.Label13.Location = New System.Drawing.Point(8, 168)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(72, 32)
        Me.Label13.TabIndex = 83
        Me.Label13.Text = "Attachments"
        '
        'btnRemove
        '
        Me.btnRemove.BackColor = System.Drawing.Color.Peru
        Me.btnRemove.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnRemove.ForeColor = System.Drawing.Color.White
        Me.btnRemove.Location = New System.Drawing.Point(368, 200)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(120, 20)
        Me.btnRemove.TabIndex = 82
        Me.btnRemove.Text = "Remove attachment"
        '
        'btnAdd
        '
        Me.btnAdd.BackColor = System.Drawing.Color.Peru
        Me.btnAdd.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAdd.ForeColor = System.Drawing.Color.White
        Me.btnAdd.Location = New System.Drawing.Point(368, 176)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(120, 20)
        Me.btnAdd.TabIndex = 81
        Me.btnAdd.Text = "Add attachment"
        '
        'lstAttachment
        '
        Me.lstAttachment.BackColor = System.Drawing.Color.LemonChiffon
        Me.lstAttachment.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstAttachment.ForeColor = System.Drawing.Color.Sienna
        Me.lstAttachment.Location = New System.Drawing.Point(128, 168)
        Me.lstAttachment.Name = "lstAttachment"
        Me.lstAttachment.Size = New System.Drawing.Size(232, 80)
        Me.lstAttachment.TabIndex = 80
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(8, 104)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(72, 16)
        Me.Label14.TabIndex = 78
        Me.Label14.Text = "To"
        '
        'txtTo
        '
        Me.txtTo.BackColor = System.Drawing.Color.LemonChiffon
        Me.txtTo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTo.ForeColor = System.Drawing.Color.Sienna
        Me.txtTo.Location = New System.Drawing.Point(128, 104)
        Me.txtTo.Name = "txtTo"
        Me.txtTo.Size = New System.Drawing.Size(232, 20)
        Me.txtTo.TabIndex = 79
        Me.txtTo.Text = ""
        '
        'txtFromDisplayName
        '
        Me.txtFromDisplayName.BackColor = System.Drawing.Color.LemonChiffon
        Me.txtFromDisplayName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFromDisplayName.ForeColor = System.Drawing.Color.Sienna
        Me.txtFromDisplayName.Location = New System.Drawing.Point(128, 72)
        Me.txtFromDisplayName.Name = "txtFromDisplayName"
        Me.txtFromDisplayName.Size = New System.Drawing.Size(232, 20)
        Me.txtFromDisplayName.TabIndex = 77
        Me.txtFromDisplayName.Text = ""
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(8, 72)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(104, 32)
        Me.Label16.TabIndex = 75
        Me.Label16.Text = "From Display name"
        '
        'txtFrom
        '
        Me.txtFrom.BackColor = System.Drawing.Color.LemonChiffon
        Me.txtFrom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFrom.ForeColor = System.Drawing.Color.Sienna
        Me.txtFrom.Location = New System.Drawing.Point(128, 40)
        Me.txtFrom.Name = "txtFrom"
        Me.txtFrom.Size = New System.Drawing.Size(232, 20)
        Me.txtFrom.TabIndex = 74
        Me.txtFrom.Text = ""
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(8, 40)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(72, 16)
        Me.Label17.TabIndex = 73
        Me.Label17.Text = "From"
        '
        'txtSMTPServer
        '
        Me.txtSMTPServer.BackColor = System.Drawing.Color.LemonChiffon
        Me.txtSMTPServer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSMTPServer.ForeColor = System.Drawing.Color.Sienna
        Me.txtSMTPServer.Location = New System.Drawing.Point(128, 8)
        Me.txtSMTPServer.Name = "txtSMTPServer"
        Me.txtSMTPServer.Size = New System.Drawing.Size(280, 20)
        Me.txtSMTPServer.TabIndex = 72
        Me.txtSMTPServer.Text = ""
        '
        'Label18
        '
        Me.Label18.Location = New System.Drawing.Point(8, 8)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(72, 16)
        Me.Label18.TabIndex = 71
        Me.Label18.Text = "SMTP Server"
        '
        'Label20
        '
        Me.Label20.Location = New System.Drawing.Point(16, 72)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(104, 32)
        Me.Label20.TabIndex = 76
        Me.Label20.Text = "From Display name"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.DefaultExt = "txt"
        Me.OpenFileDialog1.Filter = "Text Files|*.txt"
        Me.OpenFileDialog1.Title = "Select the File Containing the email addresses"
        '
        'Main_Screen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.DarkGoldenrod
        Me.ClientSize = New System.Drawing.Size(578, 640)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.ForeColor = System.Drawing.Color.White
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(584, 672)
        Me.MinimumSize = New System.Drawing.Size(584, 672)
        Me.Name = "Main_Screen"
        Me.ShowInTaskbar = False
        Me.Text = "Mass Mailer 1.0"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private current_light As Integer = 0
    Private current_colour As Integer = 0
    Private currently_working As Boolean = False




    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message("The Application encountered the following problem: " & vbCrLf & identifier_msg & ":" & ex.ToString)
                Display_Message1.ShowDialog()
                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & identifier_msg & ":" & ex.ToString)
                filewriter.Flush()
                filewriter.Close()
            End If
        Catch exc As Exception
            MsgBox("An error occurred in Mass Mailer's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Public Sub Load_Registry_Values()
        Try
            Dim configflag As Boolean
            configflag = False
            Dim str As String
            Dim keyflag1 As Boolean = False
            Dim oReg As RegistryKey = Registry.LocalMachine
            Dim keys() As String = oReg.GetSubKeyNames()
            System.Array.Sort(keys)

            For Each str In keys
                If str.Equals("Software\Mass Mailer") = True Then
                    keyflag1 = True
                    Exit For
                End If
            Next str

            If keyflag1 = False Then
                oReg.CreateSubKey("Software\Mass Mailer")
            End If

            keyflag1 = False

            Dim oKey As RegistryKey = oReg.OpenSubKey("Software\Mass Mailer", True)

            str = oKey.GetValue("txtSMTPServer")
            If Not IsNothing(str) And Not (str = "") Then
                txtSMTPServer.Text = str
            Else
                configflag = True
                oKey.SetValue("txtSMTPServer", "")
                txtSMTPServer.Text = ("")
            End If

            oKey.Close()
            oReg.Close()

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Save_Registry_Values()
        Try
            Dim oReg As RegistryKey = Registry.LocalMachine
            Dim oKey As RegistryKey = oReg.OpenSubKey("Software\Mass Mailer", True)

            oKey.SetValue("txtSMTPServer", txtSMTPServer.Text)

            oKey.Close()
            oReg.Close()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub run_green_lights()
        Try
            Label1.ForeColor = Color.LimeGreen
            Label1.Text = "Waiting"
            Label7.Text = "Resting..."

            current_light = current_light - 1
            If current_light < 1 Then
                current_light = 5
            End If
            current_colour = 0
            PictureBox1.Image = ImageList1.Images(1)
            PictureBox2.Image = ImageList1.Images(1)
            PictureBox3.Image = ImageList1.Images(1)
            PictureBox4.Image = ImageList1.Images(1)
            PictureBox5.Image = ImageList1.Images(1)

            Select Case current_light
                Case 0

                    PictureBox1.Image = ImageList1.Images(0)
                Case 1

                    PictureBox2.Image = ImageList1.Images(0)
                Case 2

                    PictureBox3.Image = ImageList1.Images(0)
                Case 3

                    PictureBox4.Image = ImageList1.Images(0)
                Case 4

                    PictureBox5.Image = ImageList1.Images(0)
                Case 5

                    PictureBox1.Image = ImageList1.Images(0)
            End Select

            current_light = current_light + 1
            If current_light > 5 Then
                current_light = 1
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub run_orange_lights()
        Try
            Label1.ForeColor = Color.DarkOrange
            Label1.Text = "Working"

            current_light = current_light - 1
            If current_light < 1 Then
                current_light = 5
            End If
            current_colour = 1
            PictureBox1.Image = ImageList1.Images(3)
            PictureBox2.Image = ImageList1.Images(3)
            PictureBox3.Image = ImageList1.Images(3)
            PictureBox4.Image = ImageList1.Images(3)
            PictureBox5.Image = ImageList1.Images(3)
            Select Case current_light
                Case 0
                    PictureBox1.Image = ImageList1.Images(2)
                Case 1
                    PictureBox2.Image = ImageList1.Images(2)
                Case 2
                    PictureBox3.Image = ImageList1.Images(2)
                Case 3
                    PictureBox4.Image = ImageList1.Images(2)
                Case 4
                    PictureBox5.Image = ImageList1.Images(2)
                Case 5
                    PictureBox1.Image = ImageList1.Images(2)
            End Select

            current_light = current_light + 1
            If current_light > 5 Then
                current_light = 1
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub run_lights()
        Try
            If current_colour = 1 Then
                Select Case current_light
                    Case 0
                        PictureBox5.Image = ImageList1.Images(3)
                        PictureBox1.Image = ImageList1.Images(2)
                    Case 1
                        PictureBox1.Image = ImageList1.Images(3)
                        PictureBox2.Image = ImageList1.Images(2)
                    Case 2
                        PictureBox2.Image = ImageList1.Images(3)
                        PictureBox3.Image = ImageList1.Images(2)
                    Case 3
                        PictureBox3.Image = ImageList1.Images(3)
                        PictureBox4.Image = ImageList1.Images(2)
                    Case 4
                        PictureBox4.Image = ImageList1.Images(3)
                        PictureBox5.Image = ImageList1.Images(2)
                    Case 5
                        PictureBox5.Image = ImageList1.Images(3)
                        PictureBox1.Image = ImageList1.Images(2)
                End Select
            Else
                Select Case current_light
                    Case 0
                        PictureBox5.Image = ImageList1.Images(1)
                        PictureBox1.Image = ImageList1.Images(0)
                    Case 1
                        PictureBox1.Image = ImageList1.Images(1)
                        PictureBox2.Image = ImageList1.Images(0)
                    Case 2
                        PictureBox2.Image = ImageList1.Images(1)
                        PictureBox3.Image = ImageList1.Images(0)
                    Case 3
                        PictureBox3.Image = ImageList1.Images(1)
                        PictureBox4.Image = ImageList1.Images(0)
                    Case 4
                        PictureBox4.Image = ImageList1.Images(1)
                        PictureBox5.Image = ImageList1.Images(0)
                    Case 5
                        PictureBox5.Image = ImageList1.Images(1)
                        PictureBox1.Image = ImageList1.Images(0)
                End Select

            End If

            current_light = current_light + 1
            If current_light > 5 Then
                current_light = 1
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Try
            'Show open dialogue box to select the files to attach
        Dim Counter As Integer
        OFD.CheckFileExists = True
        OFD.Title = "Select file(s) to attach"
        OFD.ShowDialog()
        For Counter = 0 To UBound(OFD.FileNames)
            lstAttachment.Items.Add(OFD.FileNames(Counter))
            Next
        Catch ex As Exception
            Error_Handler(ex, "Add Attachments")
        End Try
    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        Try
            'Remove the attachments
            If lstAttachment.SelectedIndex > -1 Then
                lstAttachment.Items.RemoveAt(lstAttachment.SelectedIndex)
            End If
        Catch ex As Exception
            Error_Handler(ex, "Removing Attachments")
        End Try
    End Sub


    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Try
            run_lights()
            Label8.Text = Format(Now(), "dd/MM/yyyy hh:mm:ss tt")
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Label8.Text = Format(Now(), "dd/MM/yyyy hh:mm:ss tt")
            lblProgramLaunched.Text = Format(Now(), "dd/MM/yyyy hh:mm:ss tt")

            Load_Registry_Values()
            Timer2.Start()
            dataloaded = True
            splash_loader.Visible = False
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub exit_application()
        Try
            Save_Registry_Values()
            If Worker1.WorkerThread Is Nothing = False Then
                Worker1.WorkerThread.Abort()
                Worker1.Dispose()
            End If
            NotifyIcon1.Dispose()
            Application.Exit()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Main_Screen_closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Try
            exit_application()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub


    Public Sub WorkerHandler(ByVal Result As String)
        Try
            currently_working = False
            lblProcessEnd.Text = Format(Now(), "dd/MM/yyyy hh:mm:ss tt")
            NotifyIcon1.Text = "Resting... "
            run_green_lights()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerProgressHandler(ByVal uniqueemails As Long, ByVal totalemails As Long, ByVal emailssent As Long)
        Try
            lblUniqueEmailAddresses.Text = uniqueemails
            lblTotalEmailsSent.Text = emailssent
            lblEmailAddressesFound.Text = totalemails
            
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub run_worker()
        run_orange_lights()
        lblUniqueEmailAddresses.Text = 0
        lblTotalEmailsSent.Text = 0
        lblEmailAddressesFound.Text = 0
        lblProcessStart.Text = Format(Now(), "dd/MM/yyyy hh:mm:ss tt")


        Worker1.txtSMTPServer = txtSMTPServer.Text
        Worker1.txtFrom = txtFrom.Text
        Worker1.txtFromDisplayName = txtFromDisplayName.Text
        Worker1.txtTo = txtTo.Text
        Worker1.txtToInputFile = txtToInputFile.Text
        Dim str As String
        Worker1.lstAttachment.Clear()
        For Each str In lstAttachment.Items
            Worker1.lstAttachment.Add(str)
        Next
        Worker1.txtSubject = txtSubject.Text
        Worker1.txtMessage = txtMessage.Text
        Worker1.chkformat = chkFormat.Checked

        Worker1.ChooseThreads(1)
        currently_working = True
    End Sub



    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Try
            Me.Close()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub show_application()
        Try
            Me.Opacity = 1

            Me.BringToFront()
            Me.Refresh()
            Me.WindowState = FormWindowState.Normal

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub NotifyIcon1_dblclick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.DoubleClick
        show_application()
    End Sub
    Private Sub NotifyIcon1_snglclick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.Click
        show_application()
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        show_application()
    End Sub

    Private Sub Main_Screen_resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        Try

            If Me.WindowState = FormWindowState.Minimized Then
                NotifyIcon1.Visible = True
                Me.Opacity = 0
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub force_check()
        Try
            Label7.Text = "Busy Working..."
            NotifyIcon1.Text = "Compiling Emails..."
            If currently_working = False Then
                run_worker()
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        force_check()
    End Sub



    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Try
            If Worker1.WorkerThread Is Nothing = False Then
                Worker1.WorkerThread.Abort()
                Worker1.Dispose()
            End If
        Catch ex As Exception
            Error_Handler(ex)
        Finally
            WorkerHandler("Killed")
        End Try
    End Sub



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Dim result As DialogResult = OpenFileDialog1.ShowDialog
            If result = DialogResult.OK Then
                txtToInputFile.Text = OpenFileDialog1.FileName
            End If
        Catch ex As Exception
            Error_Handler(ex, "Input Text File Browser")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            txtSMTPServer.Text = "mail.uct.ac.za"
            txtFrom.Text = "webmaster@commerce.uct.ac.za"
            txtFromDisplayName.Text = "Webmaster Test Account"
            txtTo.Text = "clotter@commerce.uct.ac.za"
            txtToInputFile.Text = "C:\Addresses.txt"
            lstAttachment.Items.Clear()
            lstAttachment.Items.Add("C:\Addresses.txt")
            txtSubject.Text = "Testing Application"
            txtMessage.Text = "<h1>Hello</h1><p><font color=""0000FF"">Testing Application</font></p>"
            chkFormat.Checked = True
        Catch ex As Exception
            Error_Handler(ex, "Example Usage")
        End Try
    End Sub

    Private Sub Label10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label10.Click
        txtToInputFile.Text = ""
    End Sub
End Class
