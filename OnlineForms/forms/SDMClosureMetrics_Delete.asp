<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>SDM Closure Metrics</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">

<body>

<!--#include file="../includes/surveytitle.inc"-->

<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
<td width="220" valign="top"><img src="../includes/images/photos_wheelbarrow.jpg" alt="" width="220" height="477" border="0"></td>
<td valign="top">


<%

	Order=request("Order")

	Set Con = Server.CreateObject("ADODB.Connection")
	Set RST = Server.CreateObject("ADODB.Recordset")
	Con.Open "BBBSAforms", "sa","12sist12"
	RST.Open "SELECT * FROM tbl_frmSDMClosureMetrics WHERE SDMClosureMetricsID=" & Int(Request("row")), Con, 1, 3
	jMod = RST("SDMClosureMetricsID")
	
	RST.Delete
	RST.Update
	Set RST = Nothing
	form = "ParentSatPostEnroll"
	modtype = "delete"
	%>
	<!--#include file="../includes/modify_stamp.asp"-->




<%	
	Con.Close
	Set Con = Nothing
	say = "delete"


%>


<%

Dim SortField
SortField = request("SortField")

Dim SortDirection
SortDirection = request("SortDirection")

%>

<br><br>
<table width="660" cellpadding="0" cellspacing="0" border="1">

	<tr>
	<td class="formHeaderSmall" colspan="3">Delete Successful</td>
	</tr>

<tr>
	<td class=formMainBold align="center" valign="top">
	<form name="frmSDMClosureMetrics" action="SDMClosureMetrics_Complete.asp?AgencyID=<%=session("AgencyIDN")%>&SortField=<%=SortField%>&SortDirection=<%=SortDirection%>" method="post">
	<br><br>Closure Metrics Record for Match ID&nbsp;<b><i><%=request("MatchID")%></i></b> Has Been Deleted.<br><input type="submit" value="Continue" class="formMainBold">
	</form>
	</td>
</tr>

</table>

</td>
</tr>
</table>

</body>
</html>
