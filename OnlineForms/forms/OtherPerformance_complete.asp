

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Performance</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">

<%'<!--#include file="../includes/top_nav_forms_monthly.inc"--><!-- include file has </head> and <body> tags --><br>%>

<!--#include file="../includes/surveytitle.inc"-->


<%

' Pull Previous Month's Match Info


Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"

if Int(Request("m")) = 1 then
	query = "SELECT * FROM tbl_frmPerformance WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))-1 & " AND Month=" & Int(Request("m"))+11
	PrevMonth = 12
	PrevYear = Int(Request("y"))-1
else	
	query = "SELECT * FROM tbl_frmPerformance WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y")) & " AND Month=" & Int(Request("m"))-1
	PrevMonth = Int(Request("m")) - 1
	PrevYear = Int(Request("y"))
end if

Set GetPrev = Con.Execute(query)

PrevOpenComm = GetPrev("OpenMatchesCommunityBased")
PrevOpenSchool = GetPrev("OpenMatchesSchoolBased")
PrevOpenOther = GetPrev("OpenMatchesOtherSiteBased")
PrevOpenGroup = GetPrev("OpenMatchesGroupMentoring")
PrevOpenSpecMent = GetPrev("OpenMatchesSpecialProgramsMentoring")
PrevOpenSpecNonMent = GetPrev("OpenMatchesSpecialProgramsNonMentoring")

GetPrev.Close
Set GetPrev = Nothing %>


<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
<td width="220" valign="top"><img src="../includes/images/photos_football.jpg" alt="" width="220" height="477" border="0"></td>
<td valign="top">
<form name="frmPerformance" action="OtherPerformance_edit.asp?y=<%= Request("y") %>&m=<%= Request("m") %>&PrevOpenComm=<%=PrevOpenComm%>&PrevOpenSchool=<%=PrevOpenSchool%>&PrevOpenOther=<%=PrevOpenOther%>&PrevOpenGroup=<%=PrevOpenGroup%>&PrevOpenSpecMent=<%=PrevOpenSpecMent%>&PrevOpenSpecNonMent=<%=PrevOpenSpecNonMent%>" method="post">
<!--#include file="../includes/form_stamp.asp"-->
<input type="hidden" name="status" value="editOld">


<% 

' Check for SBM Agency

Set SBMCon = Server.CreateObject("ADODB.Connection")
SBMCon.Open "BBBSAforms","sa","12sist12"
query = "SELECT SBM FROM tbl_AgencyInfo WHERE AgencyID = '" & Session("AgencyIDN") & "' and SBM = -1  " 
Set SBMQuery = SBMCon.Execute(query)
if (SBMquery.eof) then
	SBMAgency = 0
else
	SBMAgency = 1
End if
	
SBMQuery.Close
Set SBMQuery = Nothing
SBMCon.Close
Set SBMCon = Nothing


' Check for Faith-Based / Incarcerated Agency
Dim FBIAgency
Set FBICon = Server.CreateObject("ADODB.Connection")
FBICon.Open "BBBSAforms","sa","12sist12"
query = "SELECT FBI FROM tbl_AgencyInfo WHERE AgencyID = '" & Session("AgencyIDN") & "' and FBI = -1 "
Set FBIQuery = FBICon.Execute(query)
if (FBIQuery.eof) then
	FBIAgency = 0
else
	FBIAgency = 1
End If

FBIQuery.Close
Set FBIQuery = Nothing
FBICon.Close
Set FBICon = Nothing

' Check for SDM Agency
DIM SDMPilot
Set SDMCon = Server.CreateObject("ADODB.Connection")
SDMCon.Open "BBBSAForms","sa","12sist12"
query = "SELECT SDMPilot FROM tbl_AgencyInfo WHERE AgencyID = '" & Session("AgencyIDN") & "' and SDMPilot = -1 "
Set SDMQuery = SDMCon.Execute(query)
if (SDMquery.eof) then 
	SDMPilot = 0
else
	SDMPilot = 1
End if

SDMQuery.Close
Set SDMQuery = Nothing
SDMCon.Close
Set SDMCon = Nothing




%>


<% 	
Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"
query = "SELECT * FROM tbl_frmPerformance WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y")) & " AND Month=" & Int(Request("m"))
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

<A class = "formmain" href="OtherPerformance_complete_printable.asp?y=<%= Request("y") %>&m=<%= Request("m")%>&printform='Yes'&id=<%= GetPerformance("PerformanceID")%>" onclick="NewWindow(this.href,'name','600','400','yes');return false;"><img src="../images/print_icon.gif" alt="" width="34" height="34" border="0">Print This Form</a>


<br>		
		
		<!--#include file="OtherPerformance_complete_data.asp"-->	
	
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
