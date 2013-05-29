<!--#include file="../includes/session_stamp.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Special Programs</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">	

<% 	
Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"
query = "SELECT * FROM tbl_frmSpecialPrograms WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))
Set GetSpecialPrograms = Con.Execute(query)
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
 

<table width="625" border="0" cellspacing="0" cellpadding="3" bordercolordark="#003063">
<tr>
	<TD align="left">
	<form>
		<input type="button" Value="Send to Printer" onClick="PrintThisWindow()">
	</form>
	</TD>
	
	<TD align="right">
	<form>
		<input type="button" Value="Back to Special Programs Form" onClick="CloseThisWindow()">
	</form>	
	</TD>
</tr>

</table>

<br> 
<!--#include file="specialprograms_complete_data.asp"-->


</form>
<% 
GetSpecialPrograms.Close
Set GetSpecialPrograms = Nothing
Con.Close
Set Con = Nothing
 %>
<p></p>
<p></p>
</body>
</html>
