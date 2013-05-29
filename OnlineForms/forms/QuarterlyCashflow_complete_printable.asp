<!--#include file="../includes/session_stamp.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Quarterly Balance Sheet</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
<br>

<% 	
Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"
query = "SELECT * FROM tbl_frmQuarterlyCashflow WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y")) & " AND Quarter=" & Int(Request("q"))
Set GetQCF = Con.Execute(query)
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

<table width="575" border="0" cellspacing="0" cellpadding="3" bordercolordark="#003063">
<tr>
	<TD align="left">
	<form id=form1 name=form1>
		<input type="button" Value="Send to Printer" onClick="PrintThisWindow()" id=button1 name=button1>
	</form>
	</TD>
	
	<TD align="right">
	<form id=form2 name=form2>
		<input type="button" Value="Close" onClick="CloseThisWindow()" id=button2 name=button2>
	</form>	
	</TD>
</tr>
 
 <!--#include file="QuarterlyCashflow_complete_data.asp"-->

<% 
GetQCF.Close
Set GetQCF = Nothing
Con.Close
Set Con = Nothing
 %>
</body>
</html>
