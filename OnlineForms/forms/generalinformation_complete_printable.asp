<!--#include file="../includes/session_stamp.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>SDM Information</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">


<% 	
Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"
query = "SELECT * FROM tbl_frmSDMInformation WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))
Set GetSDMInformation = Con.Execute(query)
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
		<input type="button" Value="Back to SDM Information Form" onClick="CloseThisWindow()">
	</form>	
	</TD>
</tr>
</table>

<!--#include file="SDMinformation_complete_data.asp"-->

<br>

<br>
	

<% 
GetSDMInformation.Close
Set GetSDMInformation = Nothing
Con.Close
Set Con = Nothing
 %>
	<p>&nbsp;</p>

</body>
</html>
