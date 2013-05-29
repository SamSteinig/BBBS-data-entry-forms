

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


<!-- Pull YEAR TO DATE Match and Revenue Info -->

<% 
dim CommunityYTD
CommunityYTD = 0
dim SchoolYTD
SchoolYTD = 0
dim OtherYTD
OtherYTD = 0
dim RevenueYTD
RevenueYTD = 0


 	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmPerformance WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))& "ORDER BY month"
	Set GetYTD = Con.Execute(query)	

	' Get First Month's Data

	RevenueYTD = GetYTD("Revenue")	
	GetYTD.MoveNext()
	
	
	count = 0
	for count = 1 to Int(Request("m")) - 1
			RevenueYTD = RevenueYTD + GetYTD("Revenue")
			GetYTD.MoveNext()	
	next	
	
	GetYTD.Close
	Set GetYTD = Nothing
	


 	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	
	query = "select OpenCBAtEnd+NewCBMatches as 'CommunityYTD', OpenSBAtEnd+NewSBMatches as 'SiteYTD', OpenOSBAtEnd+NewOSBMatches as 'OSBYTD' from (select  (select OpenMatchesCommunityBased from tbl_frmPerformance p2  where p2.agencyid = p.agencyid and p2.year = (p.year - 1) and p2.month = 12) as 'OpenCBAtEnd', (select OpenMatchesOtherSiteBased from tbl_frmPerformance p2 where p2.agencyid = p.agencyid and p2.year = (p.year - 1) and p2.month = 12) as 'OpenOSBAtEnd', (select OpenMatchesSchoolBased from tbl_frmPerformance p2 where p2.agencyid = p.agencyid and p2.year = (p.year - 1) and p2.month = 12) as 'OpenSBAtEnd', SUM(NewMatchesCommunityBased) as 'NewCBMatches', SUM(NewMatchesSchoolBased) as 'NewSBMatches', SUM(NewMatchesSiteBasedNonSchool) as 'NewOSBMatches' from tbl_frmPerformance p WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y")) & "AND Month<=" & Int(Request("m"))& " group by agencyid, year) a"

	Set GetYTD = Con.Execute(query)

	CommunityYTD = GetYTD("CommunityYTD")
	SchoolYTD = GetYTD("SiteYTD")
	OtherYTD = GetYTD("OSBYTD")
	
	GetYTD.Close
	Set GetYTD = Nothing	
			
	%>


<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
<td width="220" valign="top"><img src="../includes/images/photos_football.jpg" alt="" width="220" height="477" border="0"></td>
<td valign="top">
<form name="frmPerformance" action="Performance_edit.asp?y=<%= Request("y") %>&m=<%= Request("m") %>&PrevOpenComm=<%=PrevOpenComm%>&PrevOpenSchool=<%=PrevOpenSchool%>&PrevOpenOther=<%=PrevOpenOther%>&PrevOpenGroup=<%=PrevOpenGroup%>&PrevOpenSpecMent=<%=PrevOpenSpecMent%>&PrevOpenSpecNonMent=<%=PrevOpenSpecNonMent%>" method="post">
<!--#include file="../includes/form_stamp.asp"-->
<input type="hidden" name="status" value="editOld">
<input type="hidden" name="CommunityYTD" value="<%=CommunityYTD%>">
<input type="hidden" name="SchoolYTD" value="<%=SchoolYTD%>">
<input type="hidden" name="OtherYTD" value="<%=OtherYTD%>">
<input type="hidden" name="RevenueYTD" value="<%=RevenueYTD%>">


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

<A class = "formmain" href="performance_complete_printable.asp?y=<%= Request("y") %>&m=<%= Request("m")%>&printform='Yes'&id=<%= GetPerformance("PerformanceID")%>" onclick="NewWindow(this.href,'name','600','400','yes');return false;"><img src="../images/print_icon.gif" alt="" width="34" height="34" border="0">Print This Form</a>


<br>		
		
		<!--#include file="performance_complete_data.asp"-->	
	
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
