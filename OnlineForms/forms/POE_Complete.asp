<% dim printform
printform="No"

AgencyID = request("AgencyID")
%> 

<SCRIPT LANGUAGE = "JavaScript">

<!-- Begin
function NewWindow(mypage, myname, w, h) {
var winl = (screen.width - w) / 2;
var wint = (screen.height - h) / 2;
winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',resizable'
win = window.open(mypage, myname, winprops)
if (parseInt(navigator.appVersion) >= 4) { win.window.focus(); }
}
//  End -->

</SCRIPT> 

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>POE</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">

<%
Dim Order
Order="MatchID"
if request("Order")<> "" then 
	Order="MatchID"
else
	Order="DateAssessmentDone"
end if
%>

<table width="660" cellpadding="0" cellspacing="0" border="0">

<tr>
<td align="right" valign="top">
<A class = "formmain" href="staff_complete_printable.asp?y=<%=Int(Request("y"))%>&printform='Yes'" onclick="NewWindow(this.href,'name','625','1000','yes');return false;"><img src="../images/print_icon.gif" alt="" width="34" height="34" border="0">Print This Form</a>
</td>
</tr>

<tr>
<td valign="top" colspan="3">
<!--#include file="POE_complete_data.asp"-->
</td>
</tr>
</table>

</body>
</html>
