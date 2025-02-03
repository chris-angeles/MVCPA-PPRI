<%
Dim UserSystemID, MVCPAAdministrator, MVCPAGrantCoordinator, MVCPAAuditor, MVCPAAdministrativeAssistant, MVCPAScorer, MVCPAViewer

' Check to be see if user is logged in. If they are logged in, the System_ID Session variable will be set.
' If they are not logged in, they will just get public system id of -1.
If isempty(Session("SystemID")) Then
	UserSystemID = -1
	userid = "public"
	Administrative = False
	member = false
	last_login = ""
	Session("system_id") = System_ID
	Session("user_id") = "public"
	Session("County_ID") = 0
	Session("county") = ""
	Session("name") = "Public User"
	Session("last_login") = ""
	Session("Developer") = False
	Session("MVCPAAdministrator") = False
	Session("MVCPAGrantCoordinator") = False
	Session("MVCPAAdministrativeAssistant") = False
	Session("MVCPAAuditor") = False
	Session("MVCPAViewer") = False
	Session("MVCPAScorer") = False
	Session("MVCPAStaff") = False
Else
	UserSystemID = Session("SystemID")
	Developer = Session("Developer")
	MVCPAAdministrator = Session("MVCPAAdministrator")
	MVCPAGrantCoordinator = Session("MVCPAGrantCoordinator")
	MVCPAAuditor = Session("MVCPAAuditor")
	MVCPAAdministrativeAssistant = Session("MVCPAAdministrativeAssistant")
	MVCPAScorer = Session("MVCPAScorer")
	MVCPAViewer = Session("MVCPAViewer")
	MVCPAStaff = Session("MVCPAStaff")
End If
%>