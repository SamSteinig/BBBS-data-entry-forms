<! --#include file="../includes/session_stamp.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>General Information</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">




<% 	

dim AgencyID
dim CompYear

AgencyID = Session("AgencyIDN")

if agencyid = "" then

	AgencyID = Request.QueryString("Id")
	if len(agencyid) = 3 then agencyID = "0" & AgencyID
	if len(agencyid) = 2 then agencyid = "00" & AgencyID
	if len(agencyid) = 1 then agencyid = "000" & Agencyid
end if


Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"
SQL = "SELECT m.FK_Agency_ID, d.AgencyName, d.AgencyCity, d.AgencyState, m.Compliance_Year, m.Fee_Calculation_Form_Submited, m.CEO_Position_Open, " &_
	"m.CEO_Position_Open_Date, m.Core_Matches_LastYear_Submited, m.Audit_Report_Submited, m.InsuranceCert_Submited, m.InsuranceCert_Effective_Date, m.InsuranceCert_Expiration_Date, " &_
	"m.Survey_Board_Submited, m.Survey_Expances_Submited, m.Survey_Staff_Submited, m.Fee_Payments_Current, m.Survey_Benefits, m.Survey_Forecast " &_
	"FROM tbl_frmMinCompliance m inner join tblDemogs d on m.FK_Agency_ID = d.AgencyID " &_
	"WHERE FK_Agency_ID='" & agencyid & "' AND Compliance_Year=" & Int(Request("y"))
 'query = "SELECT * FROM tbl_frmMinCompliance WHERE FK_Agency_ID='" & agencyid & "' AND Compliance_Year=" & Int(Request("y"))
 'query1 = "SELECT d.AgencyName, d.AgencyState FROM tblDemogs d, tbl_frmMinCompliance m WHERE  d.AgencyID = m.FK_Agency_ID"

If Month(Date()) > 7 Then
	CompYear = Request("y") + 1
Else CompYear = Request("y")
End if

SQL1 = "SELECT SelfAssessment_Operational_Completed, SelfAssessment_Program_Completed FROM tbl_frmMinCompliance " &_
		"WHERE FK_Agency_ID='" & agencyid & "' AND Compliance_Year=" & Int(CompYear)
		
' query = "SELECT * FROM tbl_frmSelfAssessment WHERE AgencyID='9999' AND Year=" & Int(Request("y"))
Set GetMinCompliance = Con.Execute(SQL)
Set GetSelfAss = Con.Execute(SQL1)
'Set GetAgencyData = Con.Execute(query1)
 %>	 
 

 
<SCRIPT LANGUAGE = "JavaScript">

function CloseThisWindow() {
	window.close()
	}

function PrintThisWindow() {
	window.print()
	}
	
function GoBack() {
	top.history.back()
	}



</SCRIPT>

 

<br>

<table width="650" border="0" cellspacing="0" cellpadding="3" bordercolordark="#003063">
<tr>
	<TD align="left">
	<form>
		<input type="button" Value="Send to Printer" onClick="PrintThisWindow()">
	</form>
	</TD>
	
	<TD align="right">
	<form>
		<input type="button" Value="Close Window" onClick="CloseThisWindow()">
	</form>	
	</TD>
</tr>

</table>

<% section = request(section)%>

<!--#include file="MinCompliance_complete_data.asp"-->

<br>

<br>
	

<% 
GetMinCompliance.Close
GetSelfAss.Close
Set GetMinCompliance = Nothing
Set GetSelfAss = Nothing
Con.Close
Set Con = Nothing
 %>
	<p>&nbsp;</p>

</body>
</html>
