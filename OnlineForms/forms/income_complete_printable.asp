<!--#include file="../includes/session_stamp.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Revenue</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">

<% 	
Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"
query = "SELECT * FROM tbl_frmIncome WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))
Set GetIncome = Con.Execute(query)
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


<table width="400" border="0" cellspacing="0" cellpadding="3" bordercolordark="#003063">
<tr>
	<TD align="left">
	<form>
		<input type="button" Value="Send to Printer" onClick="PrintThisWindow()">
	</form>
	</TD>
	
	<TD align="right">
	<form>
		<input type="button" Value="Back to Income Form" onClick="CloseThisWindow()">
	</form>	
	</TD>
</tr>

</table>

<!--#include file="income_complete_data.asp"-->


<% 
GetIncome.Close
Set GetIncome = Nothing
Con.Close
Set Con = Nothing
 %>
</form>




</body>
</html>
