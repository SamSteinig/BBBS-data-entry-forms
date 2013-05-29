

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>DOEPerformance</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">

<%'<!--#include file="../includes/top_nav_forms_monthly.inc"--><!-- include file has </head> and <body> tags --><br>%>

<!--#include file="../includes/surveytitle.inc"-->




<% 	
Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"
query = "SELECT * FROM tbl_frmDOEPerformance WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y")) & " AND Month=" & Int(Request("m"))
Set GetDOEPerformance = Con.Execute(query)
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

<%
' Pull Previous Month's Match Info


Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"

if Int(Request("m")) = 1 then
	query = "SELECT * FROM tbl_frmDOEPerformance WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))-1 & " AND Month=" & Int(Request("m"))+11
	PrevMonth = 12
	PrevYear = Int(Request("y"))-1
else	
	query = "SELECT * FROM tbl_frmDOEPerformance WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y")) & " AND Month=" & Int(Request("m"))-1
	PrevMonth = Int(Request("m")) - 1
	PrevYear = Int(Request("y"))
end if

Set GetPrev = Con.Execute(query)
if (GetPrev.eof) then
	PrevOpenSchool = 0
else
	PrevOpenSchool = GetPrev("OpenMatchesSchoolBased")
End if



GetPrev.Close
Set GetPrev = Nothing %>





<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
<td width="220" valign="top"><img src="../includes/images/photos_football.jpg" alt="" width="220" height="477" border="0"></td>
<td valign="top">
<form name="frmDOEPerformance" action="DOEPerformance_edit.asp?y=<%= Request("y") %>&m=<%= Request("m") %>&PrevOpenComm=<%=PrevOpenComm%>&PrevOpenSchool=<%=PrevOpenSchool%>&PrevOpenOther=<%=PrevOpenOther%>" method="post">
<!--#include file="../includes/form_stamp.asp"-->
<input type="hidden" name="status" value="editOld">


<br>

<A class = "formmain" href="DOEPerformance_complete_printable.asp?y=<%= Request("y") %>&m=<%= Request("m")%>&printform='Yes'&id=<%= GetDOEPerformance("DOEPerformanceID")%>" onclick="NewWindow(this.href,'name','600','400','yes');return false;"><img src="../images/print_icon.gif" alt="" width="34" height="34" border="0">Print This Form</a>


<br>		
		
		<!--#include file="DOEPerformance_complete_data.asp"-->	
	
</td>
</tr>
</table>
		
<% 
GetDOEPerformance.Close
Set GetDOEPerformance = Nothing
Con.Close
Set Con = Nothing
 %>
</form>

</body>
</html>
