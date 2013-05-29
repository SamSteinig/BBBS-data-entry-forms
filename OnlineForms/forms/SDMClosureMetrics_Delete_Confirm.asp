<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>SDM Closure Metrics</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">

<body>

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
<td valign="top" class=formMainBold>
<br><br>
<table border="1" cellpadding="2" cellspacing="0" bordercolordark="003063" width="660">
	<tr>
	<td class="formHeaderSmall">Confirm Delete</td>
	</tr>
	<tr>
		<td valign="top" align="center" class=formMainBold>
			<br>
			You are about to delete Match ID: <i><%=Request("MatchID")%></i>

			<form name="frmSDMClosureMetrics_Delete_Confirm" action="SDMClosureMetrics_delete.asp?row=<%=Int(Request("row"))%>&AgencyID=<%=Session(AgencyIDN)%>&MatchID=<%=Request("MatchID")%>&SortField=<%=SortField%>&SortDirection=<%=SortDirection%>" method="post">
				<input type="submit" value="Delete" class="formMainBold">
			</form>

			<form name="frmSDMClosureMetricsCancel" action="SDMClosureMetrics_Complete.asp?AgencyID=<%=session("AgencyIDN")%>&SortField=<%=SortField%>&SortDirection=<%=SortDirection%>" method="post">
			<input type="submit" value="Cancel Delete" class="formMainBold">
			</form>	
		</td>
	</tr>

</td>
</tr>
</table>

</body>
</html>
