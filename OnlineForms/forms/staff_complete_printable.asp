

<!--#include file="../includes/session_stamp.asp"-->

<% 

' check to see if user has rights to view and edit this form
if not Session("staffFormAccess") then %> 

	<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

		<html>
		<head>
			<title>Untitled</title>
		<link rel="stylesheet" type="text/css" href="../../bbbsa.css">
</head>
		
		<body>
			<p align="center"><br><br><b>You do not have access to view this form.<br><br><br>
			<a href="javascript: history.back()">back</a></p>
		</body>
		</html> <%
	
	response.end
end if
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



<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Staff</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">


<table width="600" border="0" cellspacing="0" cellpadding="3" bordercolordark="#003063">
<tr>
	<TD align="left">
	<form>
		<input type="button" Value="Send to Printer" onClick="PrintThisWindow()">
	</form>
	</TD>
	
	<TD align="right">
	<form>
		<input type="button" Value="Back to Staff Form" onClick="CloseThisWindow()">
	</form>	
	</TD>
</tr>

<BR>
<!--#include file="staff_complete_data.asp"-->

</td>
</tr>
</table>

</body>
</html>
