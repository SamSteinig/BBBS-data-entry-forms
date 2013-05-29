
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>Quarterly Balance Sheet</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">

<!--#include file="../includes/surveytitle.inc"-->

<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr><td width="220" valign="top"><img src="../includes/images/photos_football.jpg" alt="" width="220" height="477" border="0"></td>
	<td valign="top" class="formMain">

<!-- Check Form Status -->

<% 	
Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"
query = "SELECT * FROM tbl_FormStatus WHERE FormName='Finance'"
Set GetFormStatus = Con.Execute(query)
%>	

<% if (GetFormStatus("Status").Value) = "Down" then %>
	<p><br><br><br>
	<i><font color="red"><b>
	<%= (GetFormStatus("Message").Value) %>
	</p></i></font></b>

<% else %>	

	<form name="frmQuarterlyCashflow" action="QuarterlyCashflow_edit.asp?id=<%=Request("id")%>&y=<%= Request("y") %>&q=<%= Request("q") %>" method="post">
	<!--#include file="../includes/form_stamp.asp"-->
	<input type="hidden" name="status" value="editOld">
	
	<% 	
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmQuarterlyCashflow WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y")) & " AND Quarter=" & Int(Request("q"))
	Set GetQCF = Con.Execute(query)
	 %>	
	 
	<% dim printform
	printform="No"
	%>
	
	<SCRIPT LANGUAGE = "JavaScript">
	
	<!-- Begin
	function NewWindow(mypage, myname, w, h) {
	var winl = (screen.width - w) / 2;
	var wint = (screen.height - h) / 2;
	winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',resizable, scrollbars'
	win = window.open(mypage, myname, winprops)
	if (parseInt(navigator.appVersion) >= 4) { win.window.focus(); }
	}
	//  End -->
	
	</SCRIPT>
	
	<br>
	
	<A class = "formmain" href="QuarterlyCashflow_complete_printable.asp?y=<%= Request("y") %>&q=<%= Request("q")%>&printform='Yes'&id=<%= GetQCF("QuarterlyCashflowID")%>" onclick="NewWindow(this.href,'name','600','400','yes');return false;"><img src="../images/print_icon.gif" alt="" width="34" height="34" border="0">Print This Form</a>
	<br>		
			<!--#include file="QuarterlyCashflow_complete_data.asp"-->	
	</td>
	</tr>
	</table>
			
	<% 
	GetQCF.Close
	Set GetQCF = Nothing
	Con.Close
	Set Con = Nothing
	 %>
	</form>
<% 
GetFormStatus.Close
Set GetFormStatus = Nothing

end if %>

</body>
</html>
