<% 
If Request("status") = "login" Then
	UName = LCase(Request("user"))
	PWord = LCase(Request("pass"))
	Set Connect = Server.CreateObject("ADODB.Connection")
	Connect.Open "BBBSAforms", "sa","12sist12"
	Set GetUser = Connect.Execute("SELECT * FROM tbl_UserLogins WHERE password='" & PWord & "'")
	If Trim(PWord) = "" Then
		Allow = "bad"
	End If
	If Trim(UName) = "" Then
		Allow = "bad"
	End If
	
	If (GetUser.BOF And GetUser.EOF) Then
		Allow = "bad"
	End If
	
	If	(Not  Allow = "bad") Then
		GetUser.MoveFirst
		Do While Not GetUser.EOF
			If Trim((GetUser("password")) = Trim(PWord)) And (Trim(GetUser("Username")) = Trim(UName)) 	Then
				Allow = "good"
				aidn = Trim(GetUser("AgencyID"))
				staffFormAccess = Trim(GetUser("StaffFormAccess"))
				readonly = Trim(GetUser("ReadOnly"))
				Admin = Trim(GetUser("Admin"))
				' MEH START
				UserLoginID = Trim(GetUser("UserLoginID"))
				'MEH END
				Exit Do	
			Else
				Allow = "bad"
			End If
			GetUser.MoveNext
		Loop
	End If

End If
 
If Allow = "good" Then
	Session("login") = UName
	' MEH START
	Session("UserLoginID") = UserLoginID
	' MEH END
	Session("AgencyIDN") = aidn
	if (Admin = 1) then
		Session("Admin") = true
	else
		Session("Admin") = false
	end if
	if (staffFormAccess = 1) then
		Session("staffFormAccess") = true
	else
		Session("staffFormAccess") = false
	end if
	
	if (readonly = 1) then
		Session("readonly") = true
	else
		Session("readonly") = false
	end if
	
	Response.Redirect("index.asp")
End If

If Request("go") = "expired" Then
	Allow = "expired"
End If
 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Online Agency Forms</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">

<% '<!--#include file="../includes/top_nav_forms_agency.inc"--><!-- include file has </head> and <body> tags --><br>     %>
<!--#include file="../includes/surveytitle.inc"-->

<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
<td valign="top" width="220"><img src="../includes/images/photos_slinky.jpg" alt="" width="220" height="477" border="0"></td>
<td valign="top" width="100%">


<br><br>
<form method="post" action="login.asp">
<font class="formIndex">
<% If Allow = "bad" Then %>
<b>Incorrect Username or Password.&nbsp;</b><br>
<P>
<% End If %>
<% If Allow = "expired" Then %>
<b>Your Login Has <u>Expired</u>, Please Log in Again.</b><br>
<P>
<% End If %>
<b>Please Log In</b><br><br>
<input type="hidden" name="status" value="login">
<table border=0>
<tr>
<td valign=top align=right><font class="formIndex">Username:</font></td>
<td valign=top><input type="text" name="user" size=15 class="formMain"></td>
</tr>
<tr>
<td valign=top align=right><font class="formIndex">Password:</font></td>
<td valign=top><input type="password" size=15 name="pass" class="formMain"></td>
</tr>
<tr>
<td>&nbsp;</td>
<td valign=top><input type="submit" value="Login" class="formMainBold"></td>
</tr>
</table>
</form>

</font>
<br>
<!--#include file="../includes/contact_info.inc"-->
<br>


<P>
</td>
</tr>
</table>

</body>
</html>
