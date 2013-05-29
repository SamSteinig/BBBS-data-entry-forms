
<!--#include file="../includes/session_stamp.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Budget Forecast</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">

<center>
<% 	
Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"
query = "SELECT * FROM tbl_frmBudgetForecast WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))
Set GetBudget = Con.Execute(query)
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

<table width="600" border="0" cellspacing="0" cellpadding="3" bordercolordark="#003063" ID="Table1">
<tr>
	<TD align="left">
	<form ID="Form1">
		<input type="button" Value="Send to Printer" onClick="PrintThisWindow()" ID="Button1" NAME="Button1">
	</form>
	</TD>
	
	<TD align="right">
	<form ID="Form2">
		<input type="button" Value="Back to Budget Form" onClick="CloseThisWindow()" ID="Button2" NAME="Button2">
	</form>	
	</TD>
</tr>

</table>
<!--#include file="budget_complete_data.asp"-->

<% 
GetBudget.Close
Set GetBudget = Nothing
Con.Close
Set Con = Nothing
 %>



</center>


</body>
</html>
