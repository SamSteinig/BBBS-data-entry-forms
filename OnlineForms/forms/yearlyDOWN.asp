<!--#include file="../includes/session_stamp.asp"-->
<% 
If Request("status") = "bounce" Then
	y = Request("year")
	f = Request("forms")
	Redim x(8)
	x(1) = "BoardMembers"
	x(2) = "Expenses"
	x(3) = "GeneralInformation"
	x(4) = "Income"
	x(5) = "SpecialPopulations"
	x(6) = "SpecialPrograms"
	x(7) = "PerformanceBaseline"
	if Session("staffFormAccess") then
		x(8) = "Staff"
	end if
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
 %>
 
<!-- This was Monthly Performance Code that I don't think is needed here  --Sam Steinig 12/02/2002
Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"
SQL = "sp_getLatestPerformanceEntry '" & Session("AgencyIDN") & "'"
Set getLatestPerformanceEntry = Con.Execute(SQL)

datecounter = cdate(trim(getlatestperformanceentry("LatestDate"))) 
latestmonth = month(datecounter)

getLatestPerformanceEntry.Close
Set getLatestPerformanceEntry = Nothing
Con.Close
Set Con = Nothing

-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Yearly Forms</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">

<% ' <!--#include file="../includes/top_nav_forms_yearly.inc"--><!-- include file has </head> and <body> tags --><br>     %>
<!--#include file="../includes/surveytitle.inc"-->

<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
<td width="220" valign="top"><img src="../includes/images/photos_baseball.jpg" alt="" width="220" height="477" border="0"></td>
<td width="100%" valign="top">

<br><br>
<FONT class="formIndex">2003 Yearly Forms are currently unavailable.<br>Please check back on or after January 5, 2004.  <BR>Thank you.</FONT>

</table>

</body>
</html>
