
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Performance</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">

<% '<!--#include file="../includes/top_nav_forms_yearly.inc"--><!-- include file has </head> and <body> tags --><br>      %>
<!--#include file="../includes/surveytitle.inc"-->
<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
<td width="220" valign="top"><img src="../includes/images/photos_fishing.jpg" alt="" width="220" height="477" border="0"></td>
<td width="100%" valign="top">
<br>

<% 	
Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"
query = "SELECT * FROM tbl_frmPerformanceBaseline WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))
Set GetPerformance = Con.Execute(query)
 %>	

		<table width="400" border="1" cellspacing="0" cellpadding="3" bordercolordark="#003063">
		<form name="frmPerformanceBaseline" action="PerformanceBaseline_edit.asp?y=<%= Request("y") %>" method="post">
		<!--#include file="../includes/form_stamp.asp"-->
		<input type="hidden" name="status" value="editOld">
			<tr>
				<td colspan="7" class="formHeader">PERFORMANCE</td>
			</tr>
			<tr>
				<td colspan="7" class="formMainBold">Created: <%= GetPerformance("CreateDate") %><br>
		<% form = "Performance" %> 
		<% gid = GetPerformance("PerformanceBaselineID") %>
		<%= GetPerformance("PerformanceBaselineID") %>
		<!--#include file="../includes/lastmodified_stamp.asp"-->
				</td>
			</tr>
			<tr>
				<td align="center" valign="middle" class="formMain">&nbsp;</td>
				<td align="center" valign="middle" class="formMain">Community Based</td>
				<td align="center" valign="middle" class="formMain">School Based</td>
				<td align="center" valign="middle" class="formMain">Other Site Based</td>
				<td align="center" valign="middle" class="formMain">Group Mentoring</td>
				<td align="center" valign="middle" class="formMain">Special Programs: Mentoring</td>
				<td align="center" valign="middle" class="formMain">Special Programs: Non-Mentoring</td>
			</tr>
			<tr>
				<td align="center" valign="middle" class="formMain">
				OPEN/ACTIVE&nbsp;matches<br>on&nbsp;the&nbsp;last&nbsp;day&nbsp;of<br><b><%= Request("y") %></b></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("OpenMatchesCommunityBased") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("OpenMatchesSchoolBased") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("OpenMatchesOtherSiteBased") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("OpenMatchesGroupMentoring") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("OpenMatchesSpecialProgramsMentoring") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("OpenMatchesSpecialProgramsNonMentoring") %></td>
			</tr>
			<tr>
				<td align="center" valign="middle" class="formMain">
				Matches&nbsp;CLOSED&nbsp;during<br><b><%= Request("y") %></b></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("ClosedMatchesCommunityBased") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("ClosedMatchesSchoolBased") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("ClosedMatchesOtherSiteBased") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("ClosedMatchesGroupMentoring") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("ClosedMatchesSpecialProgramsMentoring") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("ClosedMatchesSpecialProgramsNonMentoring") %></td>
			</tr>
				<tr>
					<td colspan="7" class="formHeader"><input type="submit" value="Edit Form" class="formMainBold"></td>
				</tr>
				
				<tr>
					<td colspan="7"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
				</tr>
		</table>
		

<% 
GetPerformance.Close
Set GetPerformance = Nothing
Con.Close
Set Con = Nothing
 %>
</form>
</td>
</tr>
</table>
</body>
</html>
