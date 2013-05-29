
<!--#include file="../includes/session_stamp.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Board Members</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
	
<br>

<% 	
Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"
query = "SELECT * FROM tbl_frmBoardMembers WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))
Set GetBoardMembers = Con.Execute(query)
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
	<form>
		<input type="button" Value="Send to Printer" onClick="PrintThisWindow()">
	</form>
	</TD>
	
	<TD align="right">
	<form>
		<input type="button" Value="Back to Board Members Form" onClick="CloseThisWindow()">
	</form>	
	</TD>
</tr>
 
 
 <!--#include file="boardmembers_complete_data.asp"-->
			

<% 
GetBoardMembers.Close
Set GetBoardMembers = Nothing
Con.Close
Set Con = Nothing
 %>

	

</body>
</html>
