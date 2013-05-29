<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>Form Admin</title>
</head>

<!--#include file="../Includes/NAD_BE.asp" -->
<!--#include file="../includes/surveytitle.inc"-->

<% 	
Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"
query = "SELECT * FROM tbl_FormNames order by FormType, FormName"
Set GetFormNames = Con.Execute(query)
%>

<body>
<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">

<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
	<td width="220" valign="top"><img src="../includes/images/photos_football.jpg" alt="" width="220" height="477" border="0"></td>
	<td valign="top" class="formMain">

		<table width="250">
		<form name="Login" method="post" action="FormAdmin.asp">
		
		<tr>
		<br>
		<td class="formMain">Enter Password</td>
		
		</tr>
		<tr>
		
		<td VALIGN="TOP">
			<input type="password" name="password" width="10">
		</td>
		
		<td>
			<input type="submit" value="Go" class="formMain" align="left">
		</td>
		
		
		</tr>
		
		</form>
		
		</table>
	</td>
</tr>
</table>	


<%
GetFormNames.Close()
set GetFormNames = nothing
%>

</body>
</html>
