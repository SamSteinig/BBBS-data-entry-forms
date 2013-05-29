
<!--#include file="../includes/session_stamp.asp"-->

<% 

' check to see if user has rights to view and edit this form


If Request("status") = "addNew" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Set RST = Server.CreateObject("ADODB.Recordset")
	Con.Open "BBBSAforms", "sa","12sist12"
	RST.Open "SELECT * FROM tbl_frmSDMClosureMetrics", Con, 1, 3
	RST.AddNew
	RST("AgencyID") = Request("AgencyIDN")
	RST("MatchID") = Request("frmSDMClosureMetricsMatchID")
	RST("MatchStartDate") = Request("frmSDMClosureMetricsMatchStartDate")
	RST("MatchEndDate") = Request("frmSDMClosureMetricsMatchEndDate")
	RST("MatchType") = Request("frmSDMClosureMetricsMatchType")
	RST("CreateDate") = Now
	RST.Update
	Set RST = Nothing
	form = "SDMClosureMetrics"
	modtype = "new"
	%>
	<!--#include file="../includes/modify_stamp.asp"-->
	<%	
	Con.Close
	Set Con = Nothing
	say = "add"
ElseIf Request("status") = "deleteRow" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Set RST = Server.CreateObject("ADODB.Recordset")
	Con.Open "BBBSAforms", "sa","12sist12"
	RST.Open "SELECT * FROM tbl_frmSDMClosureMetrics WHERE SDMClosureMetricsID=" & Int(Request("row")), Con, 1, 3
	jMod = RST("SDMClosureMetricsID")
	RST.Delete
	RST.Update
	Set RST = Nothing
	form = "SDMClosureMetrics"
	modtype = "delete"
	%>
	<!--#include file="../includes/modify_stamp.asp"-->
	<%	
	Con.Close
	Set Con = Nothing
	say = "delete"
ElseIf Request("status") = "editRow" Then
	say = "edit"
ElseIf Request("status") = "editSave" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Set RST = Server.CreateObject("ADODB.Recordset")
	Con.Open "BBBSAforms", "sa","12sist12"
	RST.Open "SELECT * FROM tbl_frmSDMClosureMetrics WHERE AgencyID='" & Session("AgencyIDN") & "' AND SDMClosureMetricsID=" & Int(Request("row")), Con, 1, 3
	RST("MatchID") = Request("frmSDMClosureMetricsMatchID")
	RST("MatchStartDate") = Request("frmSDMClosureMetricsMatchStartDate")
	RST("MatchEndDate") = Request("frmSDMClosureMetricsMatchEndDate")
	RST("MatchType") = Request("frmSDMClosureMetricsMatchType")
	jMod = RST("SDMClosureMetricsID")
	RST.Update
	Set RST = Nothing
	form = "SDMClosureMetrics"
	modtype = "edit"
	%>
	<!--#include file="../includes/modify_stamp.asp"-->
	<%	
	Con.Close
	Set Con = Nothing
	say = "add"
ElseIf Request("status") = "done" Then
	say = "thanks"
ElseIf Request("status") = "newSDMClosureMetrics" Then
	say = "form"
Else
	say = "form"
End If
 %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>SDMClosureMetrics</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
	
	


<script type="text/javascript" src="calendarDateInput.js">

/***********************************************
* Jason's Date Input Calendar- By Jason Moon http://calendar.moonscript.com/dateinput.cfm
* Script featured on and available at http://www.dynamicdrive.com
* Keep this notice intact for use.
***********************************************/

</script>





<%

Dim SortField
SortField = request("SortField")

Dim SortDirection
SortDirection = request("SortDirection")

%>


<!--#include file="../includes/surveytitle.inc"-->

<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
<td width="220" valign="top"><img src="../includes/images/photos_wheelbarrow.jpg" alt="" width="220" height="477" border="0"></td>
<td valign="top">
<br>

<% If say = "thanks" Then %>

<font class="formMain"><BR><BR>
<strong>Thank you!</strong> Your information has been saved in the BBBS database.<br>
To choose another form, please select the form type from the choices above.
<br><br>
<i>Please note: These changes will not be reflected in the <strong>Agency Profile</strong> (in the My Agency Page and the Agency Directory) for 24 hours.</i>
</font>

<br>
<!--#include file="../includes/contact_info.inc"-->
<br>



<% ElseIf say <> "thanks" Then  %>

<form name="frmSDMClosureMetrics" action="SDMClosureMetrics_edit.asp?y=0" method="post">
<input type="hidden" name="SortField" value=<%=SortField%>>
<input type="hidden" name="SortDirection" value=<%=SortDirection%>>
<!--#include file="../includes/form_stamp.asp"-->
<% 
If say = "edit" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmSDMClosureMetrics WHERE AgencyID='" & Session("AgencyIDN") & "' AND SDMClosureMetricsID=" & Int(Request("row"))
	Set GetSDMClosureMetrics = Con.Execute(query)
 %>
<input type="hidden" name="status" value="editSave">

<input type="hidden" name="row" value="<%= Request("row") %>">
<% Else %>
<input type="hidden" name="status" value="addNew">



<%
End If
 %>

<table border="1" cellpadding="2" cellspacing="0" bordercolordark="003063" width="640">
	<tr>
		<td colspan="2" align="center" class="formSubhead">BBBS - Monthly Performance</td>
	</tr>
	<tr>
		<td colspan="2" class="formHeader">SDM Closure Metrics</td>
	</tr>
	<tr>
		<td colspan="2" align="center" valign="top" class="formMain">Please enter the following SDM Closure Metrics for the match into the fields below.<br>Click "Save This Entry" when you have completed each. Saved information will appear in a grid below.</td>
	</tr>


	<tr>
		<td class="formMain">Match ID:</td>
		<td align="right" valign="top" class="formMain"><input type="text" size="20" value="<% If say = "edit" Then %><%= GetSDMClosureMetrics("MatchID") %><% Else  %>0<% End If %>" class="formMain" name="frmSDMClosureMetricsMatchID" ></td>
	</tr>
	
	<tr>
		<td class="formMain">Match Open Date:</td>
		<td align="right" valign="top" class="formMain">
		<script>
		DateInput('frmSDMClosureMetricsMatchStartDate', true, 'MM/DD/YYYY'<% If say = "edit" Then %>,'<%=GetSDMClosureMetrics("MatchStartDate")%>'<%end if%>)
		</script>
		</td>
	</tr>
	
	<tr>
		<td class="formMain">Match Close Date:</td>
		<td align="right" valign="top" class="formMain">
		<script>
		DateInput('frmSDMClosureMetricsMatchEndDate', true, 'MM/DD/YYYY'<% If say = "edit" Then %>,'<%=GetSDMClosureMetrics("MatchEndDate")%>'<%end if%>)
		</script>
		</td>
	</tr>	
	
	<tr>
		<td class="formMain">Match Type:</td>
		<td align="right" valign="top" class="formMain">
		<select size="1" class="formMain" name="frmSDMClosureMetricsMatchType">				
			<option value="1" class="formMain"<% If say = "edit" Then %><% If GetSDMClosureMetrics("MatchType") = 1 Then %> selected<% End If %><% End If %>>1 - Community</option>				
			<option value="2" class="formMain"<% If say = "edit" Then %><% If GetSDMClosureMetrics("MatchType") = 2 Then %> selected<% End If %><% End If %>>2 - School</option>							
			<option value="3" class="formMain"<% If say = "edit" Then %><% If GetSDMClosureMetrics("MatchType") = 3 Then %> selected<% End If %><% End If %>>3 - Other Site&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>						
		</select>			
		
		</td>
	
	</tr>
	


	<% If say = "edit" Then %>
	<tr>
		<td colspan="2" class="formHeader"><input type="submit" value="Save This Record" class="formMainBold"></td>
	</tr>
	<% Else %>
	<tr>
		<td colspan="2" class="formHeader"><input type="submit" value="Save This New Record" class="formMainBold"></td>
	</tr>
	<% End If %>

</form>

<% End If %>


<% 
If say <> "thanks" Then
 %>
	
		<form name="frmSDMClosureMetricsDone" action="SDMClosureMetrics_complete.asp" method="post">
		<input type="hidden" name="SortField" value=<%=SortField%>>
		<input type="hidden" name="SortDirection" value=<%=SortDirection%>>
			<tr>
                <td colspan="8" class="formMain"><img src="../images/spacer.gif" width="1" height="5" alt="" border="0"></td>
       		</tr>
			<tr>
                <td colspan="2" class="formHeader"><input type="submit" value="Go Back to SDM Closure Metrics Main Screen" class="formMainBold"></td>
       		</tr>
		</form>
			<tr>
				<td colspan="8"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
			</tr>

		</table><br>

<br>
<P>
<% 
End If
 %>
 
</td>
</tr>
</table>

</body>
</html>
