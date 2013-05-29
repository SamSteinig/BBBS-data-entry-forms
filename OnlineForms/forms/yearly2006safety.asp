<!--#include file="../includes/session_stamp.asp"-->
<% 

Dim AssessmentExpired
AssessmentExpired = Request("AssessmentExpired")

Dim StaffLevel

Dim ReadOnlyLevel
If Session("ReadOnly") then
	ReadOnlyLevel=1
Else
	ReadOnlyLevel=0
End If

If Session("StaffFormAccess") then
	StaffLevel="Privileged"
Else
	StaffLevel="Shared"
End if
	
If Request("status") = "bounce" Then
	y = Request("year")
	f = Request("forms")
	Redim x(8)
'	x(1) = "SDMInformation"
	x(2) = "Income"
	x(3) = "Expenses"
	x(4) = "BoardMembers"
	if Session("staffFormAccess") then
		x(5) = "Staff"
	end if
	x(6) = "SelfAssessment"
	
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
		query = "SELECT " & x(f) & "ID FROM tbl_frm" & x(f) & " WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(y)
		Set GetData = Con.Execute(query)
		If (GetData.EOF OR GetData.BOF) Then
			'show blank form
			Response.Redirect(x(f) & "_edit.asp?y=" & y)
		Else
			'show complete form w/ edit button
			z = x(f) & "ID"
			id = GetData(z)
			Response.Redirect(x(f) & "_complete.asp?y=" & y & "&id=" & id)
		End If
		GetData.Close
		Set GetData = Nothing	
	Con.Close
	Set Con = Nothing
End If

' Yearly Assessment Form Selections

If Request("status") = "BounceAssessment" Then
	y = Request("year")
	f = Request("forms")
	Redim x(8)
	x(1) = "SelfAssessment"
	x(2) = "SelfAssessment"
	dim section
	if f = 1 then
		section = "Operational"
	else
		section = "Program"
	end if
	
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
		query = "SELECT " & x(f) & "ID FROM tbl_frm" & x(f) & " WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(y)
	'	query = "SELECT " & x(f) & "ID FROM tbl_frm" & x(f) & " WHERE AgencyID='9999' AND Year=" & Int(y)	
		Set GetData = Con.Execute(query)
		If (GetData.EOF OR GetData.BOF) Then
'			show blank form
			Response.Redirect(x(f) & "_edit.asp?y=" & y & "&section=" & section)
'			AssessmentExpired="Yes"
'			Response.Redirect("yearly.asp?AssessmentExpired=Yes")
			
		Else
			'show complete form w/ edit button
			z = x(f) & "ID"
			id = GetData(z)
			Response.Redirect(x(f) & "_complete.asp?y=" & y & "&id=" & id & "&section=" & section)
		End If
		GetData.Close
		Set GetData = Nothing	
	Con.Close
	Set Con = Nothing
End If




 %>
 
<!-- Popup Window Script -->
<SCRIPT LANGUAGE = "JavaScript">

<!-- Begin
function NewWindow(mypage, myname, w, h) {
var winl = (screen.width - w) / 2;
var wint = (screen.height - h) / 2;
winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',resizable, scrollbars'
win = window.open(mypage, myname, winprops)
if (parseInt(navigator.appVersion) >= 4) { win.window.focus(); }
}

//  End -->

</SCRIPT> 
 


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Annual Agency Information Forms (AAI)</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">

<% ' <!--#include file="../includes/top_nav_forms_yearly.inc"--><!-- include file has </head> and <body> tags --><br>     %>
<!--#include file="../includes/surveytitle.inc"-->

<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
<td width="220" valign="top">
<img src="../includes/images/photos_baseball.jpg" alt="" width="220" height="477" border="0">
<br><a href="FormAdminLogin.asp">...</a></td>
<td width="100%" valign="top">

<br><br>
<font class="formIndex">
Annual Agency Information (AAI)</font>
<% if StaffLevel = "Shared" or ReadOnlyLevel = 1 then %>
	<p><span class="formMain" ><em><strong>Please Note: </strong>The Staff Form is not available.  Only users with the "Full Access" password (Agency ED's/CEO's) may access the Staff Form.</em></span></p>
<% End If %>




<table width = 250 cellpadding="3" cellspacing="2" border="1" bordercolor="#800080">
<tr>
<form method="post" action="yearly.asp">
<input type="hidden" name="status" value="bounce">
<td align="left" bgcolor="#c0c0c0">
<select name="forms" size=1 class="formMain">
<!-- <option value="1" class="formMain">SDM Information -->
<!-- <option value="2" class="formMain">Revenue -->
<option value="3" class="formMain">Expenses
<option value="4" class="formMain">Board Members
<% if Session("staffFormAccess") then %>
<option value="5" class="formMain">Staff
<% end if %>

<% 'if latestmonth < 10 then %>
<!-- <option value="7" class="formMain">End Of Year Performance -->
<% ' end if %>

</select>&nbsp;
</td>
<td align="left" bgcolor="#c0c0c0">
<select name="year" size=1 class="formMain">
<% 
' y = 2003
' ydisplay = 2003
y = 2004
ydisplay = 2004
If Year(Now) > (y+1) Then
	ydisplay = (Int(Year(Now))+1) - 2
	Do Until y = (Int(Year(Now))+1) - 1
 %>
<option value="<%= ydisplay %>" class="formMain"><%= ydisplay %>
<% 
		ydisplay = (ydisplay - 1)
		y = (y + 1)
	Loop
Else
 %>
<option value="<%= y %>" class="formMain"><%= y %>
<% 
End if
 %>
</select>&nbsp;
</td>
<td align="left" bgcolor="#c0c0c0">
<input type="submit" value="Go" class="formMainBold">
</td>
</form>
</tr>

</table>

<br><br>
<font class="formIndex">
Annual Self-Assessment</font>
<br>


<table width = 300 cellpadding="3" cellspacing="2" border="1" bordercolor="#800080">

<form method="post" action="yearly.asp">
<input type="hidden" name="status" value="BounceAssessment">
<td align="left" bgcolor="#c0c0c0">
<select name="forms" size=1 class="formMain">
<option value="1" class="formMain">Operational Standards
<option value="2" class="formMain">Program Standards

</select>&nbsp;
</td>
<td align="left" bgcolor="#c0c0c0">
<select name="year" size=1 class="formMain">
<% 
 y = 2006
 ydisplay = 2005
  If Year(Now) > (y+1) Then
 	ydisplay = (Int(Year(Now))+1) - 2
	Do Until y = (Int(Year(Now))+1) - 1
 %>
<option value="<%= ydisplay %>" class="formMain"><%= ydisplay %>
<% 
		ydisplay = (ydisplay - 1)
		y = (y + 1)
	Loop
Else
 %>
<option value="<%= y %>" class="formMain"><%= y %>
<% 
End if
 %>
</select>&nbsp;
</td>
<td align="left" bgcolor="#c0c0c0">
<input type="submit" value="Go" class="formMainBold">
</td>
</form>
</tr>
<tr>
<td colspan="3" class="formsubhead"  bgcolor="#c0c0c0">Please make sure that you complete <em>BOTH</em> the Operational Standards and Program Standards forms.</td>
</tr>

</table>

<%' if AssessmentExpired = "Yes" then %>


<% 'end if %>



<br>
<span class="formMain">
<!-- Changes have been made to the yearly forms. <a href="../helpfiles/surveyhelp.asp?HelpID=yearly1" onclick="NewWindow(this.href,'name','700','400','yes');return false;">Click Here</a> for an explanation. -->
</span>

<br>
<!--#include file="../includes/contact_info.inc"-->
<br>

<P>

</td>
</tr>
</table>

</body>
</html>
