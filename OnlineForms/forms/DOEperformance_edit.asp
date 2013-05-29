<!--#include file="../includes/NAD_BE.asp" -->

<% 


' Get Initial Values for fields to check for propagation

If Request("status") = "editOld" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmDOEPerformance WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y")) & " AND Month=" & Int(Request("m"))
	Set GetInitialValues = Con.Execute(query)
	
	OpenMatchesSchoolBasedInitial = GetInitialValues("OpenMatchesSchoolBased")

	GetInitialValues.Close
	Set GetInitialValues = Nothing
	
End If




If Request("status") = "addNew" Then

	
' Check for duplicate records
	
	Set DupCon = Server.CreateObject("ADODB.Connection")
	DupCon.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT Count(*) As NumberOfEntries FROM tbl_frmDOEPerformance WHERE AgencyID = '" & Request("AgencyIDN") & "' and Year = " & Request("Year")	& " and Month = " & Request("Month")
	Set DuplicateRecord = DupCon.Execute(query)
	numberOfExisting = DuplicateRecord("NumberOfEntries")
	DuplicateRecord.Close
	Set DuplicateRecord = Nothing
	DupCon.Close
	Set DupCon = Nothing
	
	
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	
	If(numberOfExisting = 0) Then
		Set RST = Server.CreateObject("ADODB.Recordset")
		RST.Open "SELECT * FROM tbl_frmDOEPerformance", Con, 1, 3
		RST.AddNew
		RST("AgencyID") = Request("AgencyIDN")
		RST("Year") = Request("year")
		RST("Month") = Request("month")
		

		RST("OpenMatchesSchoolBased") = Request("frmDOEPerformanceOpenMatchesSchoolBased")
		RST("NewMatchesSchoolBased") = Request("frmDOEPerformanceNewMatchesSchoolBased")
		RST("ClosedMatchesSchoolBased") = Request("frmDOEPerformanceClosedMatchesSchoolBased")
		
		RST("CreateDate") = Now
		RST.Update
		RST.Close
		Set RST = Nothing
		form = "DOEPerformance"
		modtype = "new"
		m = Request("month")
		%>
		<!--#include file="../includes/modify_stamp.asp"-->
		<%	
		Con.Close
		Set Con = Nothing
		say = "thanks"
	Else
		say = "previouslyEdited"
		Con.Close
		Set Con = Nothing
	End If
		
ElseIf Request("status") = "editSave" Then


	Set Con = Server.CreateObject("ADODB.Connection")
	Set RST = Server.CreateObject("ADODB.Recordset")
	Con.Open "BBBSAforms", "sa","12sist12"
	RST.Open "SELECT * FROM tbl_frmDOEPerformance WHERE agencyID='" & Request("AgencyIDN") & "' AND Year=" & Int(Request("year")) & " AND Month=" & Int(Request("month")), Con, 1, 3
	
	OpenMatchesSchoolBasedInitial = RST("OpenMatchesSchoolBased")	
	
	RST("OpenMatchesSchoolBased") = Request("frmDOEPerformanceOpenMatchesSchoolBased")
	RST("ClosedMatchesSchoolBased") = Request("frmDOEPerformanceClosedMatchesSchoolBased")
	RST("NewMatchesSchoolBased") = Request("frmDOEPerformanceNewMatchesSchoolBased")
		
	jMod = RST("DOEPerformanceID") %>
	
	
	
	
	<%
	RST.Update
	RST.Close
	Set RST = Nothing
	form = "DOEPerformance"
	modtype = "edit"
	m = Request("month")
	%>
	<!--#include file="../includes/modify_stamp.asp"-->
	<%	
	Con.Close
	Set Con = Nothing
	say = "thanks" %>
	
	<!-- Run Propagation Routine -->

	<%	
	
	set Con = Server.CreateObject("ADODB.Connection")
		Con.Open "BBBSAforms", "sa", "12sist12"

		SchoolDifference = int(trim(Request("frmDOEPerformanceOpenMatchesSchoolBased"))) - OpenMatchesSchoolBasedInitial
		sql_SchoolDifference = "sp_PropagatePerformanceChanges '" & Session("AgencyIDN") & "'," & "'" & right("00" & Int(Request("month")),2) & "/01/" & Int(Request("year")) & "','>','tbl_frmDOEPerformance','OpenMatchesSchoolBased','+','" & SchoolDifference & "'"
	
		if SchoolDifference <> 0 then		
			Set rs = Con.Execute(sql_SchoolDifference)
		end if

		
	Con.Close
	Set Con = Nothing
	
	
		
	
	
	
ElseIf Request("status") = "editOld" Then
	say = "edit"
Else
	say = "form"
End If
 %>


<% dim HelpId
HelpId = 0
%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>DOE Grant Performance</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
<script language="javascript">
<!--	

function checkForIntegerCommas(valueToCheck)
{
	var myRegularExpression = /^[0-9]+(,[0-9]{3})*$/;  // Checks for integer with or without commas
	if(!(myRegularExpression.test(valueToCheck)))
	{
		alert("Please make sure you have entered a whole number with no spaces.\n Do not leave this field blank."); 
	} 
}


function addUpOpenSchool()
{
	var box1 = Number(document.frmDOEPerformance.frmDOEPerformancePrevOpenSchool.value)
	var box2 = Number(document.frmDOEPerformance.frmDOEPerformanceClosedMatchesSchoolBased.value)
	var box3 = Number(document.frmDOEPerformance.frmDOEPerformanceNewMatchesSchoolBased.value)	

	var boxtotal = box1 - box2 + box3
	document.frmDOEPerformance.frmDOEPerformanceOpenMatchesSchoolBased.value = boxtotal
	
}




function validateForm()
{	
	
 	var onlyInteger = /^[0-9]+(,[0-9]{3})*$/;

	var PrevOpenSchool = new Number(frmDOEPerformance.frmDOEPerformancePrevOpenSchool.value)
	var CurSchoolOpen = new Number(frmDOEPerformance.frmDOEPerformanceOpenMatchesSchoolBased.value)
	var CurSchoolClosed = new Number(frmDOEPerformance.frmDOEPerformanceClosedMatchesSchoolBased.value)	
	var CurSchoolTotal = CurSchoolOpen + CurSchoolClosed
	

	if(document.frmDOEPerformance.frmDOEPerformanceOpenMatchesSchoolBased.value == "")
		{alert("Please complete all form fields");document.frmDOEPerformance.frmDOEPerformanceOpenMatchesSchoolBased.focus();}

	else if(document.frmDOEPerformance.frmDOEPerformanceClosedMatchesSchoolBased.value == "")
		{alert("Please complete all form fields");document.frmDOEPerformance.frmDOEPerformanceClosedMatchesSchoolBased.focus();}
		
	else if(document.frmDOEPerformance.frmDOEPerformanceOpenMatchesSchoolBased.value < 0)
		{alert("Open School-Based Matches at the end of the month cannot be less than zero.");document.frmDOEPerformance.frmDOEPerformanceOpenMatchesSchoolBased.focus();}
	else if(!(onlyInteger.test(document.frmDOEPerformance.frmDOEPerformanceOpenMatchesSchoolBased.value)))
		{alert(document.frmDOEPerformance.frmDOEPerformanceOpenMatchesSchoolBased.value + " is an invalid number");document.frmDOEPerformance.frmDOEPerformanceOpenMatchesSchoolBased.focus();}

	else if(document.frmDOEPerformance.frmDOEPerformanceClosedMatchesSchoolBased.value < 0)
		{alert("Closed School-Based matches cannot be less than zero.");document.frmDOEPerformance.frmDOEPerformanceClosedMatchesSchoolBased.focus();}
	else if(!(onlyInteger.test(document.frmDOEPerformance.frmDOEPerformanceClosedMatchesSchoolBased.value)))
		{alert(document.frmDOEPerformance.frmDOEPerformanceClosedMatchesSchoolBased.value + " is an invalid number");document.frmDOEPerformance.frmDOEPerformanceClosedMatchesSchoolBased.focus();}

	else
		document.frmDOEPerformance.submit();	
}		

function getNextElement (field) 
{
	var form = field.form;
  	for (var e = 0; e < form.elements.length; e++)
    	if (field == form.elements[e])
      	break;
  	return form.elements[++e % form.elements.length];
}

//-->	
</script>

<!-- Popup Window Script -->
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

<% If say="form" Then %>
	<body onLoad="addUpOpenSchool()">
<% else %>
	<body>
<% end if %>

<% ' <!--#include file="../includes/top_nav_forms_monthly.inc"--><!-- include file has </head> and <body> tags --><br>      %>
<!--#include file="../includes/surveytitle.inc"-->


<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
<td width="220" valign="top"><img src="../includes/images/photos_football.jpg" alt="" width="220" height="477" border="0"></td> 
<td valign="top" align="left">

<% If say = "thanks" Then %>

<font class="formMain">
<br><br>
<strong>Thank you!</strong> Your information has been saved in the BBBS database.<br>
To choose another form, please select the form type from the choices above.
<br><br>
<i>Please note: These changes will not be reflected in the <strong>Agency Profile</strong> (in the My Agency Page and the Agency Directory) for 24 hours.</i>
</font>
<br>
<!--#include file="../includes/contact_info.inc"-->
<br>


<% ElseIf say <> "thanks" Then  %>


<form name="frmDOEPerformance" action="DOEPerformance_edit.asp" method="post"> <!-- onsubmit="return submitFormValidate(this)"> -->
<!--#include file="../includes/form_stamp.asp"-->

<% 
If say = "edit" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmDOEPerformance WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y")) & " AND Month=" & Int(Request("m"))
	Set GetPerformance = Con.Execute(query)
	
		
	
 %>
<input type="hidden" name="status" value="editSave">
<% Else %>
<input type="hidden" name="status" value="addNew">
<%
End If
 %>
 
<%
If say = "previouslyEdited" Then
%>
<p class="formMain"><br>We're sorry, but this form was previously completed. To make changes please <a href="monthly.asp">reselect</a> the 
appropriate form and year and update the existing information.</p>
<%
Response.End
End If 
%> 
<br>



		<table width="550" border="1" cellspacing="0" cellpadding="3" bordercolordark="#003063" >
		<tr>
			<td colspan="7" class="formHeader">DOE GRANT PERFORMANCE - <%= MonthName(Request("m"), False) & " " & Request("y") %></td>
		</tr>

		<tr>
			<td colspan="7" class="formMain"><font color="#ff0000"><div align="center"><strong>Please Note: </strong>After entering your information, you <strong>must</strong> click on the "Save" button at the bottom of the form and wait for the "Thank You" screen or your changes will be lost.</div></font></td>
		</tr>
		
			<tr>
				<td>&nbsp;</td>

				<td align="center" valign="middle" class="formMain">School Based</td>

				
			</tr>
			<tr>
			
			<!-- Open Matches at the Beginning of the Month -->
			<% if m="1" and y="2005" then %>

			<% else %>
			<tr>
				<td align="center" valign="middle" class="formMain">OPEN/ACTIVE&nbsp;matches<br>on&nbsp;the&nbsp;<strong>FIRST</strong>&nbsp;day&nbsp;of<br><b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
				<td align="center" valign="middle" class="formMain" bgcolor="#c0c0c0">
					<input type="text" class="formMain" size="5" value="<%=Request("PrevOpenSchool")%>" readonly onFocus="addUpOpenSchool();"><br><span class="formSubHead">calculated by system</span>
				</td>
			</tr>		
			<% end if %>

			<!-- Matches Closed During the Month -->

			<tr>
				<td align="center" valign="middle" class="formMain">Matches&nbsp;CLOSED&nbsp;during<br><b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesSchoolBased") %><% Else %>0<% End If %>" name="frmDOEPerformanceClosedMatchesSchoolBased" tabindex="2" onchange="checkForIntegerCommas(this.value); addUpOpenSchool();">
				</td>
			</tr>

	
			<!-- New Matches Opened During the Month -->		
			<tr>
				<% if m="1" and y="2005" then %>
					<td align="center" valign="middle" class="formMain"><b><font color="red">ONE-TIME Baseline Entry for January 2005:<br></font></b>Enter any DOE matches that existed<br><em>prior</em> to January 2005 <strong>PLUS</strong><br>any new DOE matches created <strong>DURING</strong> January 2005.</td>												
				<% else %>
					<td align="center" valign="middle" class="formMain">NEW&nbsp;matches opened<br>during&nbsp;<b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>				
				<% end if %>
				
				

				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("NewMatchesSchoolBased") %><% Else %>0<% End If %>" name="frmDOEPerformanceNewMatchesSchoolBased" tabindex="8" onchange="checkForIntegerCommas(this.value); addUpOpenSchool();">
				</td>

			</tr>					
			
			
			</tr>			
			
			<!-- Populate Initial Values of Open Matches for Propagation Comparison -->
			<input type="hidden" name="frmDOEPerformanceOpenMatchesSchoolBasedBegin" value="<%=OpenMatchesSchoolBasedBegin%>">						
			

			<!-- Open Matches on the Last Day of the Month -->
			<tr>
			
				<td align="center" valign="middle" class="formMain">OPEN/ACTIVE&nbsp;matches<br>on&nbsp;the&nbsp;<strong>LAST</strong>&nbsp;day&nbsp;of<br><b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>

				
				<td align="center" valign="middle" bgcolor="#c0c0c0" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesSchoolBased") %><% Else %>0<% End If %>" name="frmDOEPerformanceOpenMatchesSchoolBased" onFocus="addUpOpenSchool();" onchange="addUpOpenSchool();" readonly><br><span class="formSubHead">calculated by system</span>				
				</td>
				
			</tr>
			



<!-- ADD PREVIOUS MONTH'S MATCH FIELDS TO FORM FOR COMPARISON -->

<input type="hidden" class="formMain" size="5" maxlength="10" value="<%=Request("PrevOpenSchool")%>" name="frmDOEPerformancePrevOpenSchool" onchange="checkForIntegerCommas(this.value);">	


	<tr>
		<td colspan="7" class="formHeader">
		<input type="button" value="Save Form" class="formMainBold" onclick="validateForm();  addUpOpenSchool();  return false;"> 
		</td>
	</tr>
	<tr>
		<td colspan="7"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
	</tr>
	</table>

</td>
</tr>
</table>

									
<% 
If say = "edit" Then
	GetPerformance.Close
	Set GetPerformance = Nothing

	Con.Close
	Set Con = Nothing


	
End If


 %>


</form>
<% End If %>
</body>
</html>
