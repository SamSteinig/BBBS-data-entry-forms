<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>Form Admin</title>
</head>

<body>
<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">

<!--#include file="../Includes/NAD_BE.asp" -->
<!--#include file="../includes/surveytitle.inc"-->


<% ' Update The Record

dim ID
ID = request("ID")

FormName = request.form("FormName")
Status = request.form("Status")
Message = request.form("Message")

Set SaveForm = Server.CreateObject("ADODB.Recordset")
SaveForm.ActiveConnection = Connstr

  sql="UPDATE tbl_FormStatus SET "
  sql=sql & "Status='" & Status & "',"
  sql=sql & "Message='" & Message & "'" 
  sql=sql & " WHERE FormName='" & FormName & "'"
  
SaveForm.Source = SQL
SaveForm.Open()

%>


<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
	<td width="220" valign="top"><img src="../includes/images/photos_football.jpg" alt="" width="220" height="477" border="0"></td>
	<td valign="top" class="formMain">
		<table width="600" border="0" cellpadding="3" cellspacing="3">
		<tr>
			<td colspan="2" class="formMain">
			<br><br><i>The <%=FormName%> form has been updated.</i><hr>
			</td>	
		</tr>
		<tr>
			<td class="formMainBold" valign="top" width="10%">
			Status:
			</td>
			<td class="formMain" valign="top" align="left">
			<%=Status%>
			</td>
	
		</tr>
		<tr>
			<td class="formMainBold" valign="top" width="10%">
			Message:
			</td>	
			<td class="formMain" valign="top" align="left">
			<% if Status = "Up" then %>
				n/a
			<% else %>
				<%=Message%>
			<% end if %>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>

</body>
</html>
