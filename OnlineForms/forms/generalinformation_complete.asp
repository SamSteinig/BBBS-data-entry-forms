<!--#include file="../includes/session_stamp.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>General Information</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
<% '<!--#include file="../includes/top_nav_forms_yearly.inc"--><!-- include file has </head> and <body> tags --><br>      %>
<!--#include file="../includes/surveytitle.inc"-->

<% 	
Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"
query = "SELECT * FROM tbl_frmGeneralInformation WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))
Set GetGeneralInformation = Con.Execute(query)
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

 
<table border="0" width="100%" cellpadding="0" cellspacing="0">
<tr>
<td width="220" valign="top"><img src="../includes/images/photos_fishing.jpg" alt="" width="220" height="477" border="0"></td>
<td width="100%" valign="top">
<br>
<A class = "formmain" href="generalinformation_complete_printable.asp?y=<%=Int(Request("y"))%>&printform='Yes'" onclick="NewWindow(this.href,'name','625','800','yes');return false;"><img src="../images/print_icon.gif" alt="" width="34" height="34" border="0">Print This Form</a>
<!--#include file="generalinformation_complete_data.asp"-->

<br>

<br>
	

</td>
</tr>
</table>

<% 
GetGeneralInformation.Close
Set GetGeneralInformation = Nothing
Con.Close
Set Con = Nothing
 %>
	<p>&nbsp;</p>

</body>
</html>
