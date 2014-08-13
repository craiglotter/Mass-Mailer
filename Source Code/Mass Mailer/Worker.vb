Imports System.IO
Imports System.Web.Mail

Public Class Worker

    Inherits System.ComponentModel.Component

    ' Declares the variables you will use to hold your thread objects.

    Public WorkerThread As System.Threading.Thread


 

    Public txtSMTPServer As String = ""
    Public txtFrom As String = ""
    Public txtFromDisplayName As String = ""
    Public txtTo As String = ""
    Public txtToInputFile As String = ""
    Public lstAttachment As ArrayList
    Public txtSubject As String = ""
    Public txtMessage As String = ""
    Public chkformat As Boolean = False
    


    Private uniqueemails As Long = 0
    Private totalemails As Long = 0
    Private emailssent As Long = 0

    Public result As String = ""

    Public Event WorkerComplete(ByVal Result As String)
    Public Event WorkerProgress(ByVal uniqueemails As Long, ByVal totalemails As Long, ByVal emailssent As Long)



#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)
    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        lstAttachment = New ArrayList
    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

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



    Public Sub ChooseThreads(ByVal threadNumber As Integer)
        Try
            ' Determines which thread to start based on the value it receives.
            Select Case threadNumber
                Case 1
                    ' Sets the thread using the AddressOf the subroutine where
                    ' the thread will start.
                    WorkerThread = New System.Threading.Thread(AddressOf WorkerExecute)
                    ' Starts the thread.
                    WorkerThread.Start()

            End Select
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerExecute()
        Try

            uniqueemails = 0
            totalemails = 0
            emailssent = 0

        
            Dim Counter As Integer

            'Validate the data

            If txtSMTPServer = "" Then
                Dim disp As Display_Message = New Display_Message("Enter the SMTP server info.")
                disp.ShowDialog()
                result = "Failure"
                RaiseEvent WorkerComplete(result)
                Exit Sub
            End If

            If txtFrom = "" Then
                Dim disp As Display_Message = New Display_Message("Enter the From email address")
                disp.ShowDialog()
                result = "Failure"
                RaiseEvent WorkerComplete(result)
                Exit Sub
            End If

            If txtTo = "" And txtToInputFile = "" Then
                Dim disp As Display_Message = New Display_Message("Enter the Recipient email address or input file to parse")
                disp.ShowDialog()
                result = "Failure"
                RaiseEvent WorkerComplete(result)
                Exit Sub
            End If

            If txtSubject = "" Then
                Dim disp As Display_Message = New Display_Message("Enter the Email subject")
                disp.ShowDialog()
                result = "Failure"
                RaiseEvent WorkerComplete(result)
                Exit Sub
            End If


            Dim addresses As ArrayList = New ArrayList
            If ((Not txtTo Is Nothing) And (Not txtTo = "")) Then


                addresses.Add(txtTo)
                totalemails = totalemails + 1
                RaiseEvent WorkerProgress(uniqueemails, totalemails, emailssent)


            End If

            If Not txtToInputFile = "" Then


            Dim ginfo As FileInfo = New FileInfo(txtToInputFile)
            If ginfo.Exists = False Then
                ginfo = Nothing
                Dim disp As Display_Message = New Display_Message("The specified address input file cannot be located on this system. (" & txtToInputFile & ")")
                disp.ShowDialog()
                result = "Failure"
                RaiseEvent WorkerComplete(result)
                Exit Sub
            Else

                Dim filereader As StreamReader = New StreamReader(ginfo.FullName)
                Dim lineread As String
                While filereader.Peek > -1
                    lineread = filereader.ReadLine
                    If lineread.Trim.IndexOf("@") > -1 Then
                        addresses.Add(lineread.Trim)
                        totalemails = totalemails + 1
                        RaiseEvent WorkerProgress(uniqueemails, totalemails, emailssent)
                    End If
                End While
                filereader.Close()
            End If
                ginfo = Nothing
            End If

            addresses.Sort()
            uniqueemails = addresses.Count
            RaiseEvent WorkerProgress(uniqueemails, totalemails, emailssent)
            Dim cc As String = ""
            Dim dd As String = ""
            Dim il As Long

            For il = addresses.Count - 1 To 0 Step -1
                dd = addresses.Item(il)
                If dd = cc Then
                    addresses.RemoveAt(il)
                    uniqueemails = uniqueemails - 1
                    RaiseEvent WorkerProgress(uniqueemails, totalemails, emailssent)
                Else
                    cc = dd
                End If
            Next

            Dim gosend1 As String = ""
            'Dim filewriter As StreamWriter = New StreamWriter("C:\addresstogo.txt", True)
            'Dim counter1 As Integer = 0
            'For Each gosend1 In addresses
            '    counter1 = counter1 + 1
            '    If counter1 > 252 Then
            '        filewriter.WriteLine(gosend1)
            '    End If
            'Next
            'filewriter.Close()

            Dim res As DialogResult = MsgBox(uniqueemails & " unique email addresses have been harvested. Do you still wish to send out your email to all of these addresses?", MsgBoxStyle.OKCancel, "Send out Email?")
            If res = DialogResult.OK Then


                Dim gosend As String = ""
                For Each gosend In addresses
                    Try
                        emailssent = emailssent + 1
                        RaiseEvent WorkerProgress(uniqueemails, totalemails, emailssent)

                        Dim obj As System.Web.Mail.SmtpMail     ' Variable which will send the mail
                        Dim Attachment As System.Web.Mail.MailAttachment    'Variable to store the attachments
                        Dim Mailmsg As New System.Web.Mail.MailMessage      'Variable to create the message to send


                        'Set the properties

                        obj.SmtpServer = txtSMTPServer
                        'Multiple recepients can be specified using ; as the delimeter

                        Mailmsg.From = "\" & txtFromDisplayName & "\ <" & txtFrom & ">"

                        'Specify the body format
                        If chkformat = True Then
                            Mailmsg.BodyFormat = MailFormat.Html   'Send the mail in HTML Format
                        Else
                            Mailmsg.BodyFormat = MailFormat.Text
                        End If

                        'If you want you can add a reply to header 
                        Mailmsg.Headers.Add("Reply-To", txtFrom)
                        'custom headersare added like this
                        'Mailmsg.Headers.Add("Manoj", "TestHeader")
                        Mailmsg.Subject = txtSubject
                        For Counter = 0 To lstAttachment.Count - 1
                            Dim f As FileInfo = New FileInfo(lstAttachment.Item(Counter))
                            If f.Exists = False Then
                                f = Nothing
                                Dim disp As Display_Message = New Display_Message("The specified attachment cannot be located on this system. (" & lstAttachment.Item(Counter) & ")")
                                disp.ShowDialog()
                                result = "Failure"
                                RaiseEvent WorkerComplete(result)
                                Exit Sub
                            End If
                            Attachment = New MailAttachment(lstAttachment.Item(Counter))
                            Mailmsg.Attachments.Add(Attachment)
                            f = Nothing
                        Next

                        Mailmsg.Body = txtMessage
                        Mailmsg.To = gosend
                        obj.Send(Mailmsg)

                        Attachment = Nothing
                        Mailmsg = Nothing
                        obj = Nothing

                    Catch ex As Exception
                        Error_Handler(ex, "Sending Message")
                    End Try

                Next
            End If

            result = "Success"
            RaiseEvent WorkerComplete(result)
        Catch ex As Exception
            result = "Failure"
            RaiseEvent WorkerComplete(result)
        End Try

        WorkerThread.Abort()
    End Sub



    




End Class
