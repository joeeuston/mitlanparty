<%
' change to address of your own SMTP server
strHost = "smtp-au.server-mail.com"
If Request("Send") <> "" Then
   Set Mail = Server.CreateObject("Persits.MailSender")
   ' enter valid SMTP host
   Mail.Host = strHost

   Mail.From = "lanparty@melbourneit.com" ' From address
   Mail.AddAddress "joeeuston@gmail.com"

   ' message subject
   Mail.Subject = "lanparty registration"
   ' message body
   Mail.Body = Request("emailAddress")
   strErr = ""
   bSuccess = False
   On Error Resume Next ' catch errors
   Mail.Send	' send message
   If Err <> 0 Then ' error occurred
      strErr = Err.Description
   else
      bSuccess = True
   End If
End If
%>