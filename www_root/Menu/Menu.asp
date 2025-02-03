	<div id="sidenavigation">
	<ul>
	<li><a href="../Home/Default.asp" title="Return to Home Page">Home Page</a></li>
	<li>&nbsp;</li>
	<li>Grantees</li>
	<li><a href="../Grantees/GranteeReport.asp" title="Grant Report with list of grants" target="_blank">Grantee Report</a></li>
	<li>Grants</li>
	<li><a href="../Grants/GrantReport.asp" title="Grant Report with list of grants" target="_blank">Grant Report</a></li>
	<li>My Account</li>
	<li><a href="../User/UpdateProfile.asp" title="Reset Password">Update Profile</a></li>
	<li><a href="../User/ChangePassword.asp" title="Change Password">Change Password</a></li>
	<%	If Developer=True Or MVCPAAdministrator=True Or MVCPAGrantCoordinator=True Or MVCPAAdministrativeAssistant=True Then %>
	<li>Administrator</li>
	<li><a href="../User/UpdateUser1.asp" title="Add or Update a user">Add/Update User</a></li>
	<li><a href="../Admin/LoginLog.asp" title="Add or Update a user" target="_blank">Login Log</a></li>
	<%	End If %>
	<li>&nbsp;</li>
	<li><a href="../Logout.asp" title="Reset Password">Logout</a></li>
	</ul>
	</div>
