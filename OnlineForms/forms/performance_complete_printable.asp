

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Performance</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">



<!--#include file="../includes/form_stamp.asp"-->

	
	
	
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


<% 	
Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"
query = "SELECT * FROM tbl_frmPerformance WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y")) & " AND Month=" & Int(Request("m"))
Set GetPerformance = Con.Execute(query)
 %>	
 
<script language="JavaScript">

function CloseThisWindow() {
	window.close()
	}

function PrintThisWindow() {
	window.print()
	}
	
function GoBack() {
	top.history.back()
	}


</script>


 
<br>	

<table width="75%" border="0" cellspacing="0" cellpadding="3" bordercolordark="#003063">
<tr>
	<TD align="left">
	<form>
		<input type="button" Value="Send to Printer" onClick="PrintThisWindow()">
	</form>
	</TD>
	
	<TD align="right">
	<form>
		<input type="button" Value="Back to Monthly Performance" onClick="CloseThisWindow()">
	</form>	
	</TD>
</tr>
</table>


		<!--#include file="performance_complete_data.asp"-->	

<% 
GetPerformance.Close
Set GetPerformance = Nothing
Con.Close
Set Con = Nothing
 %>
</form>

</body>
</html>
