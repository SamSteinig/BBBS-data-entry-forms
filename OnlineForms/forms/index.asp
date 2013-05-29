
<% 
If Len(Session("login")) = 0 Then
	Response.Redirect("login.asp")
End IF
 %>




<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Online Agency Forms</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
	


<% '<!--#include file="../includes/top_nav_forms_agency.inc"--><!-- include file has </head> and <body> tags --><br>     %>
<!--#include file="../includes/surveytitle.inc"-->

<%

RefURL = Request.ServerVariables("HTTP_REFERRER")

%>
<p>
Referring URL: <%=RefURL%>	
</p>

<table width=100% cellpadding="0" cellspacing="0" border="0">
<tr>
	<td width="220" valign="top">
	<img src="../includes/images/photos_wheelbarrow.jpg" alt="" width="220" height="477" border="0">
	<br><a href="FormAdminLogin.asp">...</a>
	</td>
	<td width="100%" valign="top">
		<br><br>
		<font class="MainCentered"><i>Welcome!</i><br><br>Here you can quickly and easily complete your yearly<br> and monthly survey forms. Please choose the type of form<br> you would like to complete from the options above.</font>
		
		
	</td>
</tr>
</table>

</body>

</html>
