VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mailer - by ag"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   6240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txt_email_to 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1200
      TabIndex        =   3
      Top             =   960
      Width           =   4815
   End
   Begin VB.TextBox txt_subject 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1200
      TabIndex        =   4
      Top             =   1320
      Width           =   4815
   End
   Begin VB.TextBox txt_attach 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1200
      TabIndex        =   6
      Top             =   1680
      Width           =   4815
   End
   Begin VB.TextBox txt_email_from 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Top             =   600
      Width           =   4815
   End
   Begin VB.TextBox txt_smtp_server 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   4095
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   5760
      Width           =   5895
      Begin VB.TextBox txt_status 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Message text "
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   5895
      Begin VB.TextBox txt_message_text 
         Appearance      =   0  'Flat
         Height          =   3255
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Attach"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   1005
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Subject"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   1005
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Email to"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   1005
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Email from"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "SMTP server"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   1005
   End
   Begin VB.Menu mnu_send 
      Caption         =   "Send e-mail"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const err_SMTP = "No SMTP server"
Const err_FROM = "No Email from"
Const err_TO = "No Email to"
Const err_SUBJECT = "No subject"

Dim response As String

Sub wait_for(winsock_answare As String)
    Do While Left(response, 3) <> winsock_answare
        DoEvents
    Loop
    response = ""
End Sub

Function find_date() As String
    Dim temp As String
    Dim fd_day As String
    Dim fd_month As String
    Dim fd_time As String
    
    fd_day = Format(Date, "Dddd")
    Select Case fd_day
        Case "éåí øàùåï": fd_day = "Sun, "
        Case "éåí ùðé": fd_day = "Mon, "
        Case "éåí ùìéùé": fd_day = "Tue, "
        Case "éåí øáéòé": fd_day = "Wed, "
        Case "éåí çîéùé": fd_day = "Thu, "
        Case "éåí ùéùé": fd_day = "Fri, "
        Case "éåí ùáú": fd_day = "Sat, "
    End Select
    fd_month = Month(Date)
    Select Case fd_month
        Case 1: fd_month = "Jan "
        Case 2: fd_month = "Feb "
        Case 3: fd_month = "Mar "
        Case 4: fd_month = "Apr "
        Case 5: fd_month = "May "
        Case 6: fd_month = "Jun "
        Case 7: fd_month = "Jul "
        Case 8: fd_month = "Aug "
        Case 9: fd_month = "Sep "
        Case 10: fd_month = "Oct "
        Case 11: fd_month = "Nov "
        Case 12: fd_month = "Dec "
    End Select
    fd_time = Format(Time) & " +0200"
    temp = fd_day & Day(Format(Date)) & " " & fd_month & Year(Format(Date, "dd/mm/yyyy")) & " " & fd_time
    find_date = temp
End Function

Function attach_file(attach_str As String) As String
    Dim s As Integer
    Dim temp As String
    
    s = InStr(1, attach_str, "\")
    temp = attach_str
    Do While s > 0
        temp = Mid(temp, s + 1, Len(temp))
        s = InStr(1, temp, "\")
    Loop
    attach_file = temp
End Function

Function encode_the_file(attach_str As String) As String
    Dim blocksize As Long
    Dim buffer As String
    Dim s As String
    Dim i As Long
    Dim temp As String
    
    Open attach_str For Binary Access Read As #1
        blocksize = 3
        Do While Not EOF(1)
            buffer = Space(blocksize)
            Get 1, , buffer
            s = s & base64_encode_string(buffer)
            DoEvents
        Loop
    Close #1
    For i = 1 To Len(s) Step 76
        temp = temp & Mid(s, i, 76) & vbCrLf
    Next i
    temp = Mid(temp, 1, Len(temp) - 2)
    encode_the_file = temp
End Function

Sub send_email(email_to As String, email_from As String, subject As String, message_text As String, attach As String)
    Const boundary = "Hapoel_Tel_Aviv"
    
    Dim se_body As String
    Dim se_date As String
    Dim se_from As String
    Dim se_to As String
    Dim se_mime As String
    Dim se_content_type As String
    Dim se_content_type_message As String
    Dim se_content_type_attach As String
    Dim x_mailer As String
    Dim x_oem As String
     
    se_date = "Date: " & find_date
    se_from = "From: " & email_from
    se_to = "To: " & email_to
    subject = "Subject: " & subject
    
    se_mime = "MIME-Version: 1.0"
    se_content_type = "Content-Type: multipart/mixed;" & vbCrLf _
        & vbTab & "boundary = " & """" & boundary & """"
    
    x_oem = "X-OEM: zubin"
    x_mailer = "X-Mailer: " & """" & "Mailer" & """" & " - by ag v1.0"
    
    se_content_type_message = "This is a multi-part message in MIME format." & vbCrLf _
        & "--" & boundary & vbCrLf _
        & "Content-Type: text/plain;" & vbCrLf _
        & vbTab & "charset=" & """" & "iso-8859-1" & """" & vbCrLf _
        & "Content-Transfer-Encoding: 7bit"
        
    If Len(txt_attach.Text) > 0 Then
        se_content_type_attach = "--" & boundary & vbCrLf _
            & "Content-Type: application/octet-stream;" & vbCrLf _
            & vbTab & "name=" & attach_file(txt_attach.Text) & vbCrLf _
            & "Content-Transfer-Encoding: base64" & vbCrLf _
            & "Content-Disposition: attachment;" & vbCrLf _
            & vbTab & "filename=" & attach_file(txt_attach.Text) & vbCrLf _
            & vbCrLf _
            & encode_the_file(txt_attach.Text)
    End If
    
    se_body = se_from & vbCrLf _
        & se_to & vbCrLf _
        & subject & vbCrLf _
        & se_date & vbCrLf _
        & se_mime & vbCrLf _
        & x_oem & vbCrLf _
        & x_mailer & vbCrLf _
        & se_content_type & vbCrLf _
        & vbCrLf _
        & se_content_type_message & vbCrLf _
        & vbCrLf _
        & message_text & vbCrLf _
        & vbCrLf _
        & se_content_type_attach & vbCrLf _
        & "." & vbCrLf
    
    txt_status.Text = "Sending message..." & vbCrLf & txt_status.Text & vbCrLf
    Winsock1.SendData "HELO " & Left(email_from, InStr(1, email_from, "@") - 1) & vbCrLf
    wait_for "250"
    Winsock1.SendData "MAIL FROM: " & email_from & vbCrLf
    wait_for "250"
    Winsock1.SendData "RCPT TO: " & email_to & vbCrLf
    wait_for "250"
    Winsock1.SendData "DATA" & vbCrLf
    wait_for "354"
    Winsock1.SendData se_body
    wait_for "250"
    Winsock1.SendData "QUIT" & vbCrLf
    wait_for "221"
    txt_status.Text = "Message sent." & vbCrLf & txt_status.Text & vbCrLf
    Winsock1.Close
    DoEvents
End Sub

Sub connect_to_smtp_server(smtp_server As String)
    Winsock1.LocalPort = 0
    Winsock1.RemoteHost = txt_smtp_server
    Winsock1.RemotePort = 25
    Winsock1.Connect
End Sub

Sub init_me()
    txt_status.Text = "Ready." & vbCrLf
    response = ""
End Sub

Function form_errors() As Boolean
    Dim temp As Boolean
    
    temp = False
    If Len(txt_subject.Text) = 0 Then
        txt_status.Text = "Error: " & err_SUBJECT & "." & vbCrLf & txt_status.Text & vbCrLf
        temp = True
    End If
    If Len(txt_email_to.Text) = 0 Then
        txt_status.Text = "Error: " & err_TO & "." & vbCrLf & txt_status.Text & vbCrLf
        temp = True
    End If
    If Len(txt_email_from.Text) = 0 Then
        txt_status.Text = "Error: " & err_FROM & "." & vbCrLf & txt_status.Text & vbCrLf
        temp = True
    End If
    If Len(txt_smtp_server.Text) = 0 Then
        txt_status.Text = "Error: " & err_SMTP & "." & vbCrLf & txt_status.Text & vbCrLf
        temp = True
    End If
    form_errors = temp
End Function

Private Sub Form_Load()
    init_me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Winsock1.Close
    Unload Me
End Sub

Private Sub mnu_send_Click()
    If form_errors = False Then
        connect_to_smtp_server txt_smtp_server
    End If
End Sub

Private Sub Winsock1_Connect()
    txt_status.Text = "Connected to: " & txt_smtp_server & "." & vbCrLf & txt_status.Text & vbCrLf
    send_email txt_email_to, txt_email_from, txt_subject, txt_message_text, txt_attach
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Winsock1.GetData response
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    txt_status.Text = "Error: " & Description & "." & vbCrLf & txt_status.Text & vbCrLf
End Sub
