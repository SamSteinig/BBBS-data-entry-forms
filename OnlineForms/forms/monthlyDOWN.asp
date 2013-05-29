<!--#include file="../includes/session_stamp.asp"-->
<% 
If Request("status") = "bounce" Then
	z = Split(Request("month"),"-")
	m = z(0)
	y = z(1)
	f = Request("forms")
	Redim x(1)
	x(1) = "Performance"
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
		query = "SELECT " & x(f) & "ID FROM tbl_frm" & x(f) & " WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(y) & " AND Month=" & Int(m)
		Set GetData = Con.Execute(query)
		If (GetData.EOF OR GetData.BOF) Then
			'show blank form
			Response.Redirect(x(f) & "_edit.asp?y=" & y & "&m=" & m)
		Else
			'show complete form w/ edit button
			z = x(f) & "ID"
			id = GetData(z)
			Response.Redirect(x(f) & "_complete.asp?y=" & y & "&m=" & m & "&id=" & id)
		End If
		GetData.Close
		Set GetData = Nothing	
	Con.Close
	Set Con = Nothing
End If


Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"

yearquery = "SELECT Max(Year) as MaxYear from tbl_frmPerformance WHERE AgencyID='" & Session("AgencyIDN") &"'"
set getyear = con.execute(yearquery)

monthquery = "SELECT Max(Month) as MaxMonth from tbl_frmPerformance WHERE AgencyID='" & Session("AgencyIDN") & "' and year='" & getyear("MaxYear") & "'"
' maxmonth = 6
set getmonth = con.execute(monthquery)

 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Monthly Forms</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">

<% ' <!--#include file="../includes/top_nav_forms_monthly.inc"--><!-- include file has </head> and <body> tags --><br>     %>
<!--#include file="../includes/surveytitle.inc"-->

<table width="100%" cellspacing="0" cellpadding="0" border="0">
<tr>
<td width="220" valign="top"><img src="../includes/images/photos_football.jpg" alt="" width="220" height="477" border="0"></td>
<td width="100%" valign="top">

<br><br>
<font class="formIndex">
Monthly Forms</font>
<br><br>
<em>Monthly Performance Reporting is currently down in preparation for launching of new forms. Please check back again soon.<br><br>Thank you for your patience.</em>

<br>
<!--#include file="../includes/contact_info.inc"-->
<br>

<P>
</td>
</tr>
</table>

</body>
</html>
