

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

<% dim printform
printform="No"
%> 

<SCRIPT LANGUAGE = "JavaScript">

<!-- Begin
function NewWindow(mypage, myname, w, h) {
var winl = (screen.width - w) / 2;
var wint = (screen.height - h) / 2;
winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',resizable'
win = window.open(mypage, myname, winprops)
if (parseInt(navigator.appVersion) >= 4) { win.window.focus(); }
}
//  End -->

</SCRIPT> 

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Staff</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">

<% '<!--#include file="../includes/top_nav_forms_yearly.inc"--><!-- include file has </head> and <body> tags --><br>     %>
<!--#include file="../includes/surveytitle.inc"-->

<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
<td width="220" valign="top"><img src="../includes/images/photos_wheelbarrow.jpg" alt="" width="220" height="477" border="0"></td>
<td valign="top">


<BR>
<A class = "formmain" href="staff_complete_printable.asp?y=<%=Int(Request("y"))%>&printform='Yes'" onclick="NewWindow(this.href,'name','625','1000','yes');return false;"><img src="../images/print_icon.gif" alt="" width="34" height="34" border="0">Print This Form</a>
<!--#include file="staff_complete_data.asp"-->

</td>
</tr>
</table>

</body>
</html>
