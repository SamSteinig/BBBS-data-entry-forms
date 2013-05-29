<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>Form Admin</title>
</head>

<!--#include file="../Includes/NAD_BE.asp" -->
<!--#include file="../includes/surveytitle.inc"-->

<%

dim Password
Password = request.form("Password")

%>


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
		<% if Password <> "mondo" then %>
			<br><br><i>Incorrect Password.  <a href="FormAdminLogin.asp">Try Again.</a></i>
		<% else %>
			<table width="250">
			<form name="FormChooser" method="post" action="FormStatusEdit.asp">
			
			<tr>
			<br>
			<td class="formMain">Choose Form</td>
			
			</tr>
			<tr>
			
			<td VALIGN="TOP">
			<select NAME="FormName">
				  <option value=""></option>
				  <%
				  While (NOT GetFormNames.EOF)
				  %>
				  <option value="<%=(GetFormNames.Fields.Item("FormName").Value)%>"><%=(GetFormNames.Fields.Item("FormType").Value)%>&nbsp;--&nbsp;<%=(GetFormNames.Fields.Item("FormName").Value)%></option>
				  <%
				  GetFormNames.MoveNext()
				  Wend
				  %>
				</select>
			</td>
			
			<td>
				<input type="submit" value="Go" class="formMain" align="left">
			</td>
			
			
			</tr>
			
			</form>
			
			</table>
		<% end if %>
	</td>
</tr>
</table>	


<%
GetFormNames.Close()
set GetFormNames = nothing
%>

</body>
</html>
