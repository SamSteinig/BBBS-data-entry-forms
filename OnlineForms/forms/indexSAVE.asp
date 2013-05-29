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

<!--#include file="../includes/top_nav_forms_agency.inc"--><!-- include file has </head> and <body> tags --><br>     
<div align="center">
<center>
<font class="formIndex">
<a href="monthly.asp">Monthly Forms</a><br>
<a href="yearly.asp">Yearly Forms</a><br><br></font>
<font class="MainCentered"><i>Welcome!</i> Here, you can quickly and easily complete your yearly<br> and monthly forms for BBBS. Please choose the type of form<br> you would like to complete from the options above.</font>
<br>

<!--#include file="../includes/contact_info.inc"-->
<br>
</center>
</div>
</body>

</html>
