<%@ language="VBScript" %>
<%
  Option Explicit

  Const lngMaxFormBytes = 200

  Dim objASPError, blnErrorWritten, strServername, strServerIP, strRemoteIP
  Dim strMethod, lngPos, datNow, strQueryString, strURL

  If Response.Buffer Then
    Response.Clear
    Response.Status = "500 Internal Server Error"
    Response.ContentType = "text/html"
    Response.Expires = 0
  End If

  Set objASPError = Server.GetLastError
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">

<html dir=ltr>

<head>
	<meta http-equiv="Content-Type" content="text/html;charset=UTF-8" />
	<meta name="ROBOTS" content="NOINDEX" />
	<title>The page cannot be displayed</title>
	<style>
	a:link			{font:8pt/11pt verdana; color:FF0000}
	a:visited		{font:8pt/11pt verdana; color:#4e4e4e}
	</style>
</head>

<script> 
function Homepage(){
<!--
// in real bits, urls get returned to our script like this:
// res://shdocvw.dll/http_404.htm#http://www.DocURL.com/bar.htm 

	//For testing use DocURL = "res://shdocvw.dll/http_404.htm#https://www.microsoft.com/bar.htm"
	DocURL=document.URL;
	
	//this is where the http or https will be, as found by searching for :// but skipping the res://
	protocolIndex=DocURL.indexOf("://",4);
	
	//this finds the ending slash for the domain server 
	serverIndex=DocURL.indexOf("/",protocolIndex + 3);

	//for the href, we need a valid URL to the domain. We search for the # symbol to find the begining 
	//of the true URL, and add 1 to skip it - this is the BeginURL value. We use serverIndex as the end marker.
	//urlresult=DocURL.substring(protocolIndex - 4,serverIndex);
	BeginURL=DocURL.indexOf("#",1) + 1;
	urlresult=DocURL.substring(BeginURL,serverIndex);
		
	//for display, we need to skip after http://, and go to the next slash
	displayresult=DocURL.substring(protocolIndex + 3 ,serverIndex);
	InsertElementAnchor(urlresult, displayresult);
}

function HtmlEncode(text)
{
    return text.replace(/&/g, '&amp').replace(/'/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function TagAttrib(name, value)
{
    return ' '+name+'="'+HtmlEncode(value)+'"';
}

function PrintTag(tagName, needCloseTag, attrib, inner){
    document.write( '<' + tagName + attrib + '>' + HtmlEncode(inner) );
    if (needCloseTag) document.write( '</' + tagName +'>' );
}

function URI(href)
{
    IEVer = window.navigator.appVersion;
    IEVer = IEVer.substr( IEVer.indexOf('MSIE') + 5, 3 );

    return (IEVer.charAt(1)=='.' && IEVer >= '5.5') ?
        encodeURI(href) :
        escape(href).replace(/%3A/g, ':').replace(/%3B/g, ';');
}

function InsertElementAnchor(href, text)
{
    PrintTag('A', true, TagAttrib('HREF', URI(href)), text);
}

//-->
</script>

<body bgcolor="FFFFFF">

<table width="410" cellpadding="3" cellspacing="5" ID="Table1">

  <tr>    
    <td align="left" valign="middle" width="360">
	<h1 style="COLOR:000000; FONT: 13pt/15pt verdana"><!--Problem-->Error: The page cannot be displayed</h1>
    </td>
  </tr>
  
  <tr>
    <td width="400" colspan="2">
	<font style="COLOR:000000; FONT: 8pt/11pt verdana">There is a problem with the page you are trying to reach and it 
		cannot be displayed. This error message is being forwarded to PPRI so that the cause of the error can be determined
		and the problem can be fixed.</font></td>
  </tr>
  
  <tr>
    <td width="400" colspan="2">
	<font style="COLOR:000000; FONT: 8pt/11pt verdana">

	<hr color="#C0C0C0" noshade>
	
    <p>Please try the following:</p>

	<ul>
      <li id="instructionsText1">Click the 
      <a href="javascript:location.reload()">
      Refresh</a> button, or try again later.<br>
      </li>
	  
      <li>Open the 
	  
	  <script>
	  <!--
	  if (!((window.navigator.userAgent.indexOf("MSIE") > 0) && (window.navigator.appVersion.charAt(0) == "2")))
	  {
	  	 Homepage();
	  }
	  //-->
	  </script>

	  home page, and then look for links to the information you want. </li>
    </ul>
<%	If Request.ServerVariables("REMOTE_ADDR")="127.0.0.1" or Left(Request.ServerVariables("REMOTE_ADDR"),12) = "128.194.173." Then%>	
    <h2 style="font:8pt/11pt verdana; color:000000">HTTP 500.100 - Internal Server
    Error - ASP error<br>
    Internet Information Services</h2>

	<hr color="#C0C0C0" noshade>
	
	<p>Technical Information (for support personnel)</p>

<ul>
<li>Error Type:<br>
<%
  Dim bakCodepage
  on error resume next
	  bakCodepage = Session.Codepage
	  Session.Codepage = 1252
  on error goto 0
  Response.Write Server.HTMLEncode(objASPError.Category)
  If objASPError.ASPCode > "" Then Response.Write Server.HTMLEncode(", " & objASPError.ASPCode)
  Response.Write Server.HTMLEncode(" (0x" & Hex(objASPError.Number) & ")" ) & "<br>"

  If objASPError.ASPDescription > "" Then 
		Response.Write Server.HTMLEncode(objASPError.ASPDescription) & "<br>"

  elseIf (objASPError.Description > "") Then 
		 Response.Write Server.HTMLEncode(objASPError.Description) & "<br>" 
  end if



  blnErrorWritten = False

  ' Only show the Source if it is available and the request is from the same machine as IIS
  If objASPError.Source > "" Then
    strServername = LCase(Request.ServerVariables("SERVER_NAME"))
    strServerIP = Request.ServerVariables("LOCAL_ADDR")
    strRemoteIP =  Request.ServerVariables("REMOTE_ADDR")
    If (strServername = "localhost" Or strServerIP = strRemoteIP) And objASPError.File <> "?" Then
      Response.Write Server.HTMLEncode(objASPError.File)
      If objASPError.Line > 0 Then Response.Write ", line " & objASPError.Line
      If objASPError.Column > 0 Then Response.Write ", column " & objASPError.Column
      Response.Write "<br>"
      Response.Write "<font style=""COLOR:000000; FONT: 8pt/11pt courier new""><b>"
      Response.Write Server.HTMLEncode(objASPError.Source) & "<br>"
      If objASPError.Column > 0 Then Response.Write String((objASPError.Column - 1), "-") & "^<br>"
      Response.Write "</b></font>"
      blnErrorWritten = True
    End If
  End If

  If Not blnErrorWritten And objASPError.File <> "?" Then
    Response.Write "<b>" & Server.HTMLEncode(  objASPError.File)
    If objASPError.Line > 0 Then Response.Write Server.HTMLEncode(", line " & objASPError.Line)
    If objASPError.Column > 0 Then Response.Write ", column " & objASPError.Column
    Response.Write "</b><br>"
  End If
%>
</li>
<p>
<li>Browser Type:<br>
<%= Server.HTMLEncode(Request.ServerVariables("HTTP_USER_AGENT")) %>
</li>
<p>
<li>Page:<br>
<%
  strMethod = Request.ServerVariables("REQUEST_METHOD")

  Response.Write strMethod & " "

  If strMethod = "POST" Then
    Response.Write Request.TotalBytes & " bytes to "
  End If

  Response.Write Request.ServerVariables("SCRIPT_NAME")

  lngPos = InStr(Request.QueryString, "|")

  If lngPos > 1 Then
    Response.Write "?" & Server.HTMLEncode(Left(Request.QueryString, (lngPos - 1)))
  End If

  Response.Write "</li>"

  If strMethod = "POST" Then
    Response.Write "<p><li>POST Data:<br>"
    If Request.TotalBytes > lngMaxFormBytes Then
       Response.Write Server.HTMLEncode(Left(Request.Form, lngMaxFormBytes)) & " . . ."
    Else
      Response.Write Server.HTMLEncode(Request.Form)
    End If
    Response.Write "</li>"
  End If

%>
<p>
<li>Time:<br>
<%
  datNow = Now()

  Response.Write Server.HTMLEncode(FormatDateTime(datNow, 1) & ", " & FormatDateTime(datNow, 3))
  on error resume next
	  Session.Codepage = bakCodepage 
  on error goto 0
%>
</li>
</p>
<p>
<li>More information:<br>
 <%  strQueryString = "prd=iis&sbp=&pver=5.0&ID=500;100&cat=" & Server.URLEncode(objASPError.Category) & _
    "&os=&over=&hrd=&Opt1=" & Server.URLEncode(objASPError.ASPCode)  & "&Opt2=" & Server.URLEncode(objASPError.Number) & _
    "&Opt3=" & Server.URLEncode(objASPError.Description) 
       strURL = "http://www.microsoft.com/ContentRedirect.asp?" & _
    strQueryString
%>
<a href="<%= strURL %>">Microsoft Support</a>
</li>
</p>

    </font></td>
  </tr>
</table>
<%	Else %><b>A server error has occurred. A detailed error message has been forwarded to the developer for review.</b>  
<%	End If %>
</body>
</html>
<%
'on error resume next
'********************************
'Send error message to webmaster
'********************************
	dim ObjMail, Sender, Recipient, Recipient2, Subject, Body, strItem, strItemKey
	Body = "<table border=0>" & vbCrLf
	Body = Body & "<tr><td>Date/Time: </td><td>" & Now() & "</td></tr>" & vbCrLf
	Body = Body & "<tr><td>Site: </td><td>http://" & Request.ServerVariables("SERVER_NAME") &"</td></tr>" & vbCrLf
	Body = Body & "<tr><td>Page: </td><td>http://" & Request.ServerVariables("PATH_TRANSLATED") &"</td></tr>" & vbCrLf
	Body = Body & "<tr><td>Error Category: </td><td>" & objASPError.Category & "</td></tr>" & vbCrLf
	Body = Body & "<tr><td>Error Number: </td><td>" & objASPError.Number & " (0x" & Hex(objASPError.Number) & ")" &"</td></tr>" & vbCrLf
	If Len(Trim(objASPError.ASPDescription)) > 0 Then
		Body = Body & "<tr><td>ASP Error Description: </td><td>" & objASPError.ASPDescription & "</td></tr>" & vbCrLf
	End If
	If Len(Trim(objASPError.Description)) > 0 Then
		Body = Body & "<tr><td>Error Description: </td><td>" & objASPError.Description & "</td></tr>" & vbCrLf
	End If
	Body = Body & "<tr><td>Error Location: </td><td>" & objASPError.File & ", line " & objASPError.Line & ", column " & objASPError.Column & "</td></tr>" & vbCrLf
	If Len(Trim(objASPError.Source)) > 0 Then
		Body = Body & "<tr><td>Source: </td><td>" & objASPError.Source & "</td></tr>" & vbCrLf
	End If
	If Len(Request.ServerVariables("QUERY_STRING")) > 0 Then
		Body = Body & "<tr><td>QueryString: </td><td>" & Request.ServerVariables("QUERY_STRING") & "</td></tr>" & vbCrLf
	End If
	If Request.QueryString.Count > 0 then
		Body = Body & vbCrLf & "<tr><td><b>QueryString:</b></td><td></td></tr>" & vbCrLF
	For Each strItem in Request.QueryString
		Body = Body & "<tr><td>    " & strItem & ": </td><td>" & Request.QueryString(strItem) & "</td></tr>" & vbCrLf
	Next
		Body = Body & vbCrLf
	End If
	  	
	If Request.Form.Count > 0 then
		Body = Body & vbCrLf & "<tr><td><b>Form Variables:</b></td><td>" & "</td></tr>" & vbCrLF
		For Each strItem in Request.Form
			Body = Body & "<tr><td>    " & strItem & ": </td><td>" & Request.Form(strItem) & "</td></tr>" & vbCrLf
		Next
		Body = Body & vbCrLf
	End If

	if Session.Contents.Count > 0 Then
		Body = Body & vbCrLf & "<tr><td><b>Session Variables:</b></td><td>" & vbCrLf
		For each strItem in Session.Contents
			Body = Body & "<tr><td>    " & strItem & ": </td><td>" & Session.Contents(strItem) & "</td></tr>" & vbCrLf
		Next
	End If
	
	If Application.Contents.Count > 0 Then
		Body = Body & "<tr><td><b>Application Variables:<b></td></tr>" & vbCrLf
		For Each strItem in Application.Contents
			Body = Body & "<tr><td>    " & strItem & ": </td><td>" & Application.Contents(strItem) & "</td></tr>" & vbCrLf
		Next
	End If
	
	If Request.Cookies.Count > 0 Then
		Body = Body & vbCrLf & "<tr><td><b>Cookies:</b> (" & Request.Cookies.Count & ")</td></tr>" & vbCrLF
		For Each strItem in Request.Cookies
			If Request.Cookies(strItem).HasKeys Then
				For Each strItemKey in Request.Cookies(strItem)
					Body = Body & "<tr><td colspan=""2"">    " & strItem & "(" & strItemKey & "): " & Request.Cookies(strItem)(strItemKey) &"</td></tr>" &  vbCrLf
				Next
			Else
				Body = Body & "<tr><td colspan=""2"">    " & strItem & ": </td><td>" & Request.Cookies(strItem) & "</td></tr>" & vbCrLf
			End If
		Next
		Body = Body & vbCrLf
	End If
	  
	If Len(Trim(Request.ServerVariables("HTTP_REFERER"))) > 0 Then
		Body = Body & vbCrLf & "<tr><td>Referer: </td><td>" & Request.ServerVariables("HTTP_REFERER") & "</td></tr>" & vbCrLF
	End If
		  
	Body = Body & "<tr><td>Remote IP: </td><td>" & Request.ServerVariables("REMOTE_ADDR") & "</td></tr>" & vbCrLf
	Body = Body & "<tr><td>Browser: </td><td>" & Request.ServerVariables("HTTP_USER_AGENT") & "</td></tr>" & vbCrLf
	Body = Body & "</table>" & vbCrLF
	Sender = "No-Reply@ppri.tamu.edu"

	If Application("Instance")="Production" Then
		Recipient = "mvcpa@tamu.edu"
		Recipient2 = "JimVanBeek@tamu.edu"
	ElseIf Application("Instance")="Test" Then
		Recipient = "mvcpa@tamu.edu"
		Recipient2 = ""
	ElseIf Application("Instance")="Development - Jim"
		Recipient = "JimVanBeek@tamu.edu"
		Recipient2 = ""
	Else
		Recipient = "gsingla@tamu.edu"
		Recipient2 = ""
	End If
	Subject = Request.ServerVariables("SERVER_NAME") & " Error 500:100 from " & Request.ServerVariables("REMOTE_ADDR")
	SendMail Sender, Recipient, Recipient2, Subject, Body	  
'********************************
on error goto 0 

Function SendMail(vSender, vRecipient, vRecipient2, vSubject, vBody)

	Dim Debug
	
	Debug = False
	
	'Messaging - build transport configuration
	Dim iMsg
	Dim iConf
	Dim Flds
	Dim strHTML

	Const cdoSendUsingPickup = 1	'Use local SMTP service using pickup directory
	Const cdoSendUsingPort = 2		'Use network SMTP service

	set iMsg = CreateObject("CDO.Message")
	set iConf = CreateObject("CDO.Configuration")

	Set Flds = iConf.Fields
	With Flds
		'Local SMTP service using pickup directory
		'.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
		'.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = "c:\inetpub\mailroot\pickup"
		
		'Network SMTP service
		.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "relay.tamu.edu"
	
		.Update
	End With

	'Messaging - build HTML
	strHTML = "<html lang=""en-us"">"
	strHTML = strHTML & "<head></head>"
	strHTML = strHTML & "<body>"
	strHTML = strHTML & vBody
	strHTML = strHTML & "</body>"
	strHTML = strHTML & "</html>"

	If debug = True then
		'vRecipient = "mvcpa@tamu.edu"
		vRecipient = "jim@vanbeek.us"
	End if
		
	'Messaging - apply seetings to message
	With iMsg
		Set .Configuration = iConf
		.To = vRecipient
		'Messaging - determine/assign carbon copy
		If Len(vRecipient2)>0 Then
			.CC = vRecipient2
		End if
		.From = vSender
		.Subject = vSubject
		.HTMLBody = strHTML
		If debug=False then 
			.Send
		End if
	End With

	If debug = True then
		%><br><br><%
		response.write("From: " & iMsg.From)%><br><%
		response.write("To: " & iMsg.To)%><br><%
		response.write("CC: " & iMsg.CC)%><br><%
		response.write("Subject: " & iMsg.Subject)%><br><%
		response.write("Body:" & vbCrLf & iMsg.HTMLBody)
		response.flush
	End if

	'Cleanup variables
	Set iMsg = Nothing
	Set iConf = Nothing
	Set Flds = Nothing

End Function
%>

