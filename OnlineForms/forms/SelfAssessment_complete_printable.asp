<! --#include file="../includes/session_stamp.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>General Information</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">




<% 	

dim AgencyID

AgencyID = Session("AgencyIDN")

if agencyid = "" then

	AgencyID = Request.QueryString("Id")
	if len(agencyid) = 3 then agencyID = "0" & AgencyID
	if len(agencyid) = 2 then agencyid = "00" & AgencyID
	if len(agencyid) = 1 then agencyid = "000" & Agencyid
end if






Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"
 query = "SELECT * FROM tbl_frmSelfAssessment WHERE AgencyID='" & agencyid & "' AND Year=" & Int(Request("y"))
 

' query = "SELECT * FROM tbl_frmSelfAssessment WHERE AgencyID='9999' AND Year=" & Int(Request("y"))
Set GetSelfAssessment = Con.Execute(query)
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

<table width="600" border="0" cellspacing="0" cellpadding="3" bordercolordark="#003063">
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

<!--#include file="SelfAssessment_complete_data.asp"-->

<br>

<br>
	

<% 
GetSelfAssessment.Close
Set GetSelfAssessment = Nothing
Con.Close
Set Con = Nothing
 %>
	<p>&nbsp;</p>

</body>
</html>
