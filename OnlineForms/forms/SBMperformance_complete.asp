

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>SBM Performance</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">

<%'<!--#include file="../includes/top_nav`_forms_monthly.inc"--><!-- include file has </head> and <body> tags --><br>%>

<!--#include file="../includes/surveytitle.inc"-->





<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
<td width="220" valign="top"><img src="../includes/images/photos_football.jpg" alt="" width="220" height="477" border="0"></td>
<td valign="top">
<form name="frmSBMPerformance" action="SBMPerformance_edit.asp?y=<%= Request("y") %>&m=<%= Request("m") %>" method="post">
<!--#include file="../includes/form_stamp.asp"-->
<input type="hidden" name="status" value="editOld">


<% 	
Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"
query = "SELECT * FROM tbl_frmSBMPerformance WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y")) & " AND Month=" & Int(Request("m"))
Set GetPerformance = Con.Execute(query)
 %>	
 
<% dim printform
printform="No"
%>

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


<br>

<A class = "formmain" href="SBMperformance_complete_printable.asp?y=<%= Request("y") %>&m=<%= Request("m")%>&printform='Yes'&id=<%= GetPerformance("SBMPerformanceID")%>" onclick="NewWindow(this.href,'name','600','400','yes');return false;"><img src="../images/print_icon.gif" alt="" width="34" height="34" border="0">Print This Form</a>


<br>		
		
		<!--#include file="SBMperformance_complete_data.asp"-->	
	
</td>
</tr>
</table>
		
<% 
GetPerformance.Close
Set GetPerformance = Nothing
Con.Close
Set Con = Nothing
 %>
</form>

</body>
</html>
