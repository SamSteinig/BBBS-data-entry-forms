<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>Form Admin</title>
</head>

<!--#include file="../includes/surveytitle.inc"-->
<body>
<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
<table width="700">

<% 
Dim FormName
FormName = request("FormName")
%>

<% 	
Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"
query = "SELECT * FROM tbl_FormStatus WHERE FormName = '" & FormName & "'"
Set GetFormNames = Con.Execute(query)
%>

<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
	<td width="220" valign="top"><img src="../includes/images/photos_football.jpg" alt="" width="220" height="477" border="0"></td>
	<td valign="top" class="FormMain">
		
		<form name="FormChooser" method="post" action="FormAdminSave.asp?FormName='<%=FormName%>'">
		<br><br>
		<table width="650">
		<tr>
			<td class="formMainbold">Form Name</td>
			<td class="formMainbold">Status</td>
			<td class="formMainbold">Message</td>
			<td class="formMainbold">&nbsp;</td>
		</tr>
		<%
		While (NOT GetFormNames.EOF)
		%>

		<tr>
			<td class="formMain" valign="top">
				<%=(GetFormNames.Fields.Item("FormName").Value)%><br>
			</td>
			<td valign="top" class="formMain">
				<select size="1" class="formMain" name="Status">
					<option value="Up" class="formMain" <% If (GetFormNames.Fields.Item("Status").Value) = "Up" then %> selected<% End If %>>Up</option>								
					<option value="Down" class="formMain" <% If (GetFormNames.Fields.Item("Status").Value) = "Down" then %> selected<% End If %>>Down</option>		
				</select>
			</td>	
			<td valign="top" class="formMain">
				<TEXTAREA class="formMain" NAME="Message" ROWS=5 COLS=75 align="left" >
				<%=(GetFormNames.Fields.Item("Message").Value)%>
				</TEXTAREA>	
			</td>	
			<td valign="top" align="left" class="formMain">
				<% formName = GetFormNames.Fields.Item("FormName").Value %>
				<input type="hidden" name="FormName" value = "<%=FormName%>">
				<input type="submit" value="Save" class="formMain" align="left">
		
			</td>
		</tr>
		<%
		GetFormNames.MoveNext()
		Wend
		%>		
		
		</form>
		
		<%
		GetFormNames.Close()
		set GetFormNames = nothing
		%>
		
		</table>
	</td>
</tr>
</table>








</body>
</html>
