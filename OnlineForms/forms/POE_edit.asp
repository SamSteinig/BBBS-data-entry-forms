
<%

Order=request("Order")

%>Order:<%=Order%><%

If Request("status") = "addNew" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Set RST = Server.CreateObject("ADODB.Recordset")
	Con.Open "BBBSAforms", "sa","12sist12"
	RST.Open "SELECT * FROM tbl_frmPOE", Con, 1, 3
	RST.AddNew
	RST("AgencyID") = Session("AgencyIDN")
	RST("MatchID") = Request("frmPOEMatchID")
	RST("Source") = Request("frmPOESource")
	RST("ProgramType") = Request("frmPOEProgramType")
	RST("DateAssessmentDone") = Request("frmPOEDateAssessmentDone")
	RST("MatchLength") = Request("frmPOEMatchLength")
	RST("Age") = Request("frmPOEAge")
	RST("Gender") = Request("frmPOEGender")
	RST("Ethnicity") = Request("frmPOEEthnicity")
	RST("SelfConfidence") = Request("frmPOESelfConfidence")
	RST("ExpressFeelings") = Request("frmPOEExpressFeelings")
	RST("MakeDecisions") = Request("frmPOEMakeDecisions")
	RST("InterestsHobbies") = Request("frmPOEInterestsHobbies")
	RST("Hygiene") = Request("frmPOEHygiene")
	RST("SenseOfFuture") = Request("frmPOESenseOfFuture")
	RST("CommunityResources") = Request("frmPOECommunityResources")
	RST("SchoolResources") = Request("frmPOESchoolResources")
	RST("AcademicPerformance") = Request("frmPOEAcademicPerformance")
	RST("AttitudeTowardSchool") = Request("frmPOEAttitudeTowardSchool")
	RST("SchoolPreparedness") = Request("frmPOESchoolPreparedness")
	RST("ClassParticipation") = Request("frmPOEClassParticipation")
	RST("ClassroomBehavior") = Request("frmPOEClassroomBehavior")
	RST("AvoidDelinquency") = Request("frmPOEAvoidDelinquency")
	RST("AvoidSubstanceAbuse") = Request("frmPOEAvoidSubstanceAbuse")
	RST("AvoidEarlyParenting") = Request("frmPOEAvoidEarlyParenting")
	RST("ShowsTrust") = Request("frmPOEShowsTrust")
	RST("RespectsOtherCultures") = Request("frmPOERespectsOtherCultures")
	RST("RelationshipWithFamily") = Request("frmPOERelationshipWithFamily")
	RST("RelationshipWithPeers") = Request("frmPOERelationshipWithPeers")
	RST("RelationshipWithOtherAdults") = Request("frmPOERelationshipWithOtherAdults")
	RST("SubjectImprovement") = Request("frmPOESubjectImprovement")
	RST("NumberOfSubjects") = Request("frmPOENumberOfSubjects")
	RST("CreateDate") = Now
	
	RST.Update
	Set RST = Nothing
	form = "POE"
	modtype = "new"
	%>
	<!-- include file="../includes/modify_stamp.asp"-->
	<%	
	Con.Close
	Set Con = Nothing
	say = "add"
	
ElseIf Request("status") = "deleteRow" Then

	Set Con = Server.CreateObject("ADODB.Connection")
	Set RST = Server.CreateObject("ADODB.Recordset")
	Con.Open "BBBSAforms", "sa","12sist12"
	RST.Open "SELECT * FROM tbl_frmPOE WHERE POEID=" & Int(Request("row")), Con, 1, 3
	jMod = RST("POEID")
	RST.Delete
	RST.Update
	Set RST = Nothing
	form = "POE"
	modtype = "delete"
	%>
	<!--include file="../includes/modify_stamp.asp"-->
	<%	
	Con.Close
	Set Con = Nothing
	say = "delete"

ElseIf Request("status") = "editRow" Then

	say = "edit"

ElseIf Request("status") = "editSave" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Set RST = Server.CreateObject("ADODB.Recordset")
	Con.Open "BBBSAforms", "sa","12sist12"
	RST.Open "SELECT * FROM tbl_frmPOE WHERE AgencyID='" & Session("AgencyIDN") & "' AND POEID=" & Int(Request("row")), Con, 1, 3
	RST("AgencyID") = Session("AgencyIDN")
	RST("MatchID") = Request("frmPOEMatchID")
	RST("Source") = Request("frmPOESource")
	RST("ProgramType") = Request("frmPOEProgramType")
	RST("DateAssessmentDone") = Request("frmPOEDateAssessmentDone")
	RST("MatchLength") = Request("frmPOEMatchLength")
	RST("Age") = Request("frmPOEAge")
	RST("Gender") = Request("frmPOEGender")
	RST("Ethnicity") = Request("frmPOEEthnicity")
	RST("SelfConfidence") = Request("frmPOESelfConfidence")
	RST("ExpressFeelings") = Request("frmPOEExpressFeelings")
	RST("MakeDecisions") = Request("frmPOEMakeDecisions")
	RST("InterestsHobbies") = Request("frmPOEInterestsHobbies")
	RST("Hygiene") = Request("frmPOEHygiene")
	RST("SenseOfFuture") = Request("frmPOESenseOfFuture")
	RST("CommunityResources") = Request("frmPOECommunityResources")
	RST("SchoolResources") = Request("frmPOESchoolResources")
	RST("AcademicPerformance") = Request("frmPOEAcademicPerformance")
	RST("AttitudeTowardSchool") = Request("frmPOEAttitudeTowardSchool")
	RST("SchoolPreparedness") = Request("frmPOESchoolPreparedness")
	RST("ClassParticipation") = Request("frmPOEClassParticipation")
	RST("ClassroomBehavior") = Request("frmPOEClassroomBehavior")
	RST("AvoidDelinquency") = Request("frmPOEAvoidDelinquency")
	RST("AvoidSubstanceAbuse") = Request("frmPOEAvoidSubstanceAbuse")
	RST("AvoidEarlyParenting") = Request("frmPOEAvoidEarlyParenting")
	RST("ShowsTrust") = Request("frmPOEShowsTrust")
	RST("RespectsOtherCultures") = Request("frmPOERespectsOtherCultures")
	RST("RelationshipWithFamily") = Request("frmPOERelationshipWithFamily")
	RST("RelationshipWithPeers") = Request("frmPOERelationshipWithPeers")
	RST("RelationshipWithOtherAdults") = Request("frmPOERelationshipWithOtherAdults")
	RST("SubjectImprovement") = Request("frmPOESubjectImprovement")
	RST("NumberOfSubjects") = Request("frmPOENumberOfSubjects")

	jMod = RST("POEID")

	RST.Update
	Set RST = Nothing
	form = "POE"
	modtype = "edit"
	%>
	<!-- include file="../includes/modify_stamp.asp"-->
	<%	
	Con.Close
	Set Con = Nothing
	say = "add"
ElseIf Request("status") = "done" Then
	say = "thanks"
ElseIf Request("status") = "newPOE" Then
	say = "form"
Else
	say = "form"
End If
 %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>POE</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
	
	
<script language="javascript">
<!--
function checkForInteger(valueToCheck)
{
	var myRegularExpression = /^[0-9]+(,[0-9]{3})*$/;  // Checks for integer with or without commas
	var replaceWhiteSpace = /\s/; // searches for any whitespace character
	var formField = valueToCheck; // passed in as parameter 1
	var newFormField = valueToCheck.replace(replaceWhiteSpace, ""); // remove any whitespace from the form entry and replace it with nothing
	var bContainsNonNumbers = myRegularExpression.test(newFormField); // check newFormField variable to see if it contains any nonnumeric character
	
	if(!bContainsNonNumbers)
	{
		alert("Please make sure you have entered a whole number.\n We cannot process letters or words."); 
	} 
}

// -->
</script>


<script language = "Javascript">
/**
 * DHTML date validation script. Courtesy of SmartWebby.com (http://www.smartwebby.com/dhtml/)
 */
// Declaring valid date character, minimum year and maximum year
var dtCh= "/";
var minYear=1900;
var maxYear=2100;

function isInteger(s){
	var i;
    for (i = 0; i < s.length; i++){   
        // Check that current character is number.
        var c = s.charAt(i);
        if (((c < "0") || (c > "9"))) return false;
    }
    // All characters are numbers.
    return true;
}

function stripCharsInBag(s, bag){
	var i;
    var returnString = "";
    // Search through string's characters one by one.
    // If character is not in bag, append to returnString.
    for (i = 0; i < s.length; i++){   
        var c = s.charAt(i);
        if (bag.indexOf(c) == -1) returnString += c;
    }
    return returnString;
}

function daysInFebruary (year){
	// February has 29 days in any year evenly divisible by four,
    // EXCEPT for centurial years which are not also divisible by 400.
    return (((year % 4 == 0) && ( (!(year % 100 == 0)) || (year % 400 == 0))) ? 29 : 28 );
}
function DaysArray(n) {
	for (var i = 1; i <= n; i++) {
		this[i] = 31
		if (i==4 || i==6 || i==9 || i==11) {this[i] = 30}
		if (i==2) {this[i] = 29}
   } 
   return this
}

function isDate(dtStr){
	var daysInMonth = DaysArray(12)
	var pos1=dtStr.indexOf(dtCh)
	var pos2=dtStr.indexOf(dtCh,pos1+1)
	var strMonth=dtStr.substring(0,pos1)
	var strDay=dtStr.substring(pos1+1,pos2)
	var strYear=dtStr.substring(pos2+1)
	strYr=strYear
	if (strDay.charAt(0)=="0" && strDay.length>1) strDay=strDay.substring(1)
	if (strMonth.charAt(0)=="0" && strMonth.length>1) strMonth=strMonth.substring(1)
	for (var i = 1; i <= 3; i++) {
		if (strYr.charAt(0)=="0" && strYr.length>1) strYr=strYr.substring(1)
	}
	month=parseInt(strMonth)
	day=parseInt(strDay)
	year=parseInt(strYr)
	if (pos1==-1 || pos2==-1){
		alert("The date format should be : mm/dd/yyyy")
		return false
	}
	if (strMonth.length<1 || month<1 || month>12){
		alert("Please enter a valid month")
		return false
	}
	if (strDay.length<1 || day<1 || day>31 || (month==2 && day>daysInFebruary(year)) || day > daysInMonth[month]){
		alert("Please enter a valid day")
		return false
	}
	if (strYear.length != 4 || year==0 || year<minYear || year>maxYear){
		alert("Please enter a valid 4 digit year between "+minYear+" and "+maxYear)
		return false
	}
	if (dtStr.indexOf(dtCh,pos2+1)!=-1 || isInteger(stripCharsInBag(dtStr, dtCh))==false){
		alert("Please enter a valid date")
		return false
	}
return true
}

function ValidateForm(){
	var dt=document.frmPOEEdit.frmPOEDateAssessmentDone
	if (isDate(dt.value)==false){
		dt.focus()
		return false
	}
    return true
 }

</script>






<% If say = "thanks" Then %>

<font class="formMain"><BR><BR>
<strong>Thank you!</strong> Your information has been saved in the BBBS database.<br>
To choose another form, please select the form type from the choices above.
<br><br>
<i>Please note: These changes will not be reflected in the <strong>Agency Profile</strong> (in the My Agency Page and the Agency Directory) for 24 hours.</i>
</font>

<br>
<!--#include file="../includes/contact_info.inc"-->
<br>



<% ElseIf say <> "thanks" Then  %>

<form name="frmPOEEdit" action="POE_edit.asp?Order=<%=Order%>" method="post" onSubmit="return ValidateForm()">


<!--#include file="../includes/form_stamp.asp"-->
<% 
If say = "edit" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmPOE WHERE AgencyID='" & Session("AgencyIDN") & "' AND POEID=" & Int(Request("row"))
%>query:<%=query%><%	
	
'	query = "SELECT * FROM tbl_frmPOE WHERE POEID=" & Int(Request("row"))

	Set GetPOE = Con.Execute(query)
 %>
<input type="hidden" name="status" value="editSave">

<input type="hidden" name="row" value="<%= Request("row") %>">

<% Else %>

<input type="hidden" name="status" value="addNew">
<%
End If
 %>


<table border="1" cellpadding="2" cellspacing="0" bordercolordark="003063" width="600">		
	<tr>
		<td colspan="12" align="center" class="formHeader">POE Data Entry</td>
	</tr>


	<tr>
		<td colspan="11" align="center" valign="top" class="formMain">Please enter the following information on each POE match into the fields below.<br>Click "Save This Entry" when you have completed each. Saved information will appear in a grid below.</td>
	</tr>

	<tr>
		<td colspan="12" align="center" class="formSubhead">
		Code Key for Confidence, Competence, and Caring Questions:<br>
		1 = Much Worse; 2 = A little Worse; 3 = No Change; 4 = A Little Better; 5 = Much Better; 6 = Don't Know; 7 = Not a Problem
		</td>
	</tr>	


	<tr>
		<td class="formMainBold">Match ID</td>
		<td class="formMainBold">Source</td>
		<td class="formMainBold">Program Type</td>
		<td class="formMainBold">Date Assessment Done<br><em>(mm/dd/yyyy)</em></td>
		<td class="formMainBold">Match Length (Months)</td>
		<td class="formMainBold">Age</td>
		<td class="formMainBold">Gender</td>
		<td class="formMainBold">Ethnicity</td>
		<td class="formMainBold">Self Confidence</td>
		<td class="formMainBold">Express Feelings</td>				
		<td class="formMainBold">Make Decisions</td>					
	</tr>
	
	<tr>

		<td class="formMain" align="center">
			<input type="text" size="7" maxlength="10" value="<% If say = "edit" Then %><%= GetPOE("MatchID") %><% Else  %>0<% End If %>" class="formMain" name="frmPOEMatchID" onchange="checkForInteger(this.value)">
		</td>

		
		
		</td>


		<td class="formMain" align="center">
		<select size="1" class="formMain" name="frmPOESource">		
			<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("Source") = 1 Then %> selected<% End If %><% End If %>>1 - Volunteer</option>		
			<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("Source") = 2 Then %> selected<% End If %><% End If %>>2 - Parent</option>					
			<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("Source") = 3 Then %> selected<% End If %><% End If %>>3 - Teacher</option>				
		</select>
		</td>
		


		
		<td class="formMain" align="center">
		<select size="1" class="formMain" name="frmPOEProgramType">				
			<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("ProgramType") = 1 Then %> selected<% End If %><% End If %>>1 - Community</option>				
			<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("ProgramType") = 2 Then %> selected<% End If %><% End If %>>2 - School</option>							
			<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("ProgramType") = 3 Then %> selected<% End If %><% End If %>>3 - Other Site</option>						
		</select>		
		</td>

		<td class="formMain" align="center">
		<INPUT type="Text" name="frmPOEDateAssessmentDone" maxlength="10" size="15" value="<% If say = "edit" Then %><%= GetPOE("DateAssessmentDone") %><% Else %>01/01/1900<% End If %>" onchange="ValidateForm();">


		
		</td>
		
		<td class="formMain" align="center">
			<INPUT type="Text" name="frmPOEMatchLength" size="3" maxlength="3" value="<% If say = "edit" Then %><%= GetPOE("MatchLength") %><% Else %>0<% End If %>" >		
		</td>		
		<td class="formMain" align="center">
			<INPUT type="Text" name="frmPOEAge" size="2" maxlength="2" value="<% If say = "edit" Then %><%= GetPOE("Age") %><% Else %>0<% End If %>" >				
		</td>
		
		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOEGender">				
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("Gender") = 1 Then %> selected<% End If %><% End If %>>1 - Male</option>				
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("Gender") = 2 Then %> selected<% End If %><% End If %>>2 - Female</option>							
			</select>			
		</td>
		
		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOEEthnicity">				
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("Ethnicity") = 1 Then %> selected<% End If %><% End If %>>1 - White</option>				
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("Ethnicity") = 2 Then %> selected<% End If %><% End If %>>2 - Black</option>
				<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("Ethnicity") = 3 Then %> selected<% End If %><% End If %>>3 - Hispanic</option>				
				<option value="4" class="formMain"<% If say = "edit" Then %><% If GetPOE("Ethnicity") = 4 Then %> selected<% End If %><% End If %>>4 - Asian</option>				
				<option value="5" class="formMain"<% If say = "edit" Then %><% If GetPOE("Ethnicity") = 5 Then %> selected<% End If %><% End If %>>5 - Native American</option>								
				<option value="6" class="formMain"<% If say = "edit" Then %><% If GetPOE("Ethnicity") = 6 Then %> selected<% End If %><% End If %>>6 - Other</option>				
			</select>					
		</td>
		
		
		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOESelfConfidence">
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("SelfConfidence") = 1 Then %> selected<% End If %><% End If %>>1 - Much Worse</option>							
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("SelfConfidence") = 2 Then %> selected<% End If %><% End If %>>2 - A Little Worse</option>
				<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("SelfConfidence") = 3 Then %> selected<% End If %><% End If %>>3 - No Change</option>
				<option value="4" class="formMain"<% If say = "edit" Then %><% If GetPOE("SelfConfidence") = 4 Then %> selected<% End If %><% End If %>>4 - A Little Better</option>				
				<option value="5" class="formMain"<% If say = "edit" Then %><% If GetPOE("SelfConfidence") = 5 Then %> selected<% End If %><% End If %>>5 - Much Better</option>
				<option value="6" class="formMain"<% If say = "edit" Then %><% If GetPOE("SelfConfidence") = 6 Then %> selected<% End If %><% End If %>>6 - Don't Know</option>
				<option value="7" class="formMain"<% If say = "edit" Then %><% If GetPOE("SelfConfidence") = 7 Then %> selected<% End If %><% End If %>>7 - Not a Problem</option>
			</select>		
		</td>
		
		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOEExpressFeelings">
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("ExpressFeelings") = 1 Then %> selected<% End If %><% End If %>>1 - Much Worse</option>							
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("ExpressFeelings") = 2 Then %> selected<% End If %><% End If %>>2 - A Little Worse</option>
				<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("ExpressFeelings") = 3 Then %> selected<% End If %><% End If %>>3 - No Change</option>
				<option value="4" class="formMain"<% If say = "edit" Then %><% If GetPOE("ExpressFeelings") = 4 Then %> selected<% End If %><% End If %>>4 - A Little Better</option>				
				<option value="5" class="formMain"<% If say = "edit" Then %><% If GetPOE("ExpressFeelings") = 5 Then %> selected<% End If %><% End If %>>5 - Much Better</option>
				<option value="6" class="formMain"<% If say = "edit" Then %><% If GetPOE("ExpressFeelings") = 6 Then %> selected<% End If %><% End If %>>6 - Don't Know</option>
				<option value="7" class="formMain"<% If say = "edit" Then %><% If GetPOE("ExpressFeelings") = 7 Then %> selected<% End If %><% End If %>>7 - Not a Problem</option>
			</select>		
		</td>
		
		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOEMakeDecisions">
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("MakeDecisions") = 1 Then %> selected<% End If %><% End If %>>1 - Much Worse</option>							
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("MakeDecisions") = 2 Then %> selected<% End If %><% End If %>>2 - A Little Worse</option>
				<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("MakeDecisions") = 3 Then %> selected<% End If %><% End If %>>3 - No Change</option>
				<option value="4" class="formMain"<% If say = "edit" Then %><% If GetPOE("MakeDecisions") = 4 Then %> selected<% End If %><% End If %>>4 - A Little Better</option>				
				<option value="5" class="formMain"<% If say = "edit" Then %><% If GetPOE("MakeDecisions") = 5 Then %> selected<% End If %><% End If %>>5 - Much Better</option>
				<option value="6" class="formMain"<% If say = "edit" Then %><% If GetPOE("MakeDecisions") = 6 Then %> selected<% End If %><% End If %>>6 - Don't Know</option>
				<option value="7" class="formMain"<% If say = "edit" Then %><% If GetPOE("MakeDecisions") = 7 Then %> selected<% End If %><% End If %>>7 - Not a Problem</option>
			</select>		
		</td>		
	</tr>	

	<tr>
		<td rowspan="4">&nbsp;</td>
		<td class="formMainBold">Interests / Hobbies</td>						
		<td class="formMainBold">Hygiene</td>
		<td class="formMainBold">Sense of Future</td>
		<td class="formMainBold">Community Resources</td>
		<td class="formMainBold">School Resources</td>				
		<td class="formMainBold">Academic Performance</td>				
		<td class="formMainBold">Attitude Toward School</td>
		<td class="formMainBold">School Preparedness</td>								
		<td class="formMainBold">Class Participation</td>
		<td class="formMainBold">Classroom Behavior</td>					
	</tr>				
		

	<tr>

		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOEInterestsHobbies">
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("InterestsHobbies") = 1 Then %> selected<% End If %><% End If %>>1 - Much Worse</option>							
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("InterestsHobbies") = 2 Then %> selected<% End If %><% End If %>>2 - A Little Worse</option>
				<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("InterestsHobbies") = 3 Then %> selected<% End If %><% End If %>>3 - No Change</option>
				<option value="4" class="formMain"<% If say = "edit" Then %><% If GetPOE("InterestsHobbies") = 4 Then %> selected<% End If %><% End If %>>4 - A Little Better</option>				
				<option value="5" class="formMain"<% If say = "edit" Then %><% If GetPOE("InterestsHobbies") = 5 Then %> selected<% End If %><% End If %>>5 - Much Better</option>
				<option value="6" class="formMain"<% If say = "edit" Then %><% If GetPOE("InterestsHobbies") = 6 Then %> selected<% End If %><% End If %>>6 - Don't Know</option>
				<option value="7" class="formMain"<% If say = "edit" Then %><% If GetPOE("InterestsHobbies") = 7 Then %> selected<% End If %><% End If %>>7 - Not a Problem</option>
			</select>		
		</td>		
		
		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOEHygiene">
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("Hygiene") = 1 Then %> selected<% End If %><% End If %>>1 - Much Worse</option>							
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("Hygiene") = 2 Then %> selected<% End If %><% End If %>>2 - A Little Worse</option>
				<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("Hygiene") = 3 Then %> selected<% End If %><% End If %>>3 - No Change</option>
				<option value="4" class="formMain"<% If say = "edit" Then %><% If GetPOE("Hygiene") = 4 Then %> selected<% End If %><% End If %>>4 - A Little Better</option>				
				<option value="5" class="formMain"<% If say = "edit" Then %><% If GetPOE("Hygiene") = 5 Then %> selected<% End If %><% End If %>>5 - Much Better</option>
				<option value="6" class="formMain"<% If say = "edit" Then %><% If GetPOE("Hygiene") = 6 Then %> selected<% End If %><% End If %>>6 - Don't Know</option>
				<option value="7" class="formMain"<% If say = "edit" Then %><% If GetPOE("Hygiene") = 7 Then %> selected<% End If %><% End If %>>7 - Not a Problem</option>
			</select>		
		</td>		
			
		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOESenseOfFuture">
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("SenseOfFuture") = 1 Then %> selected<% End If %><% End If %>>1 - Much Worse</option>							
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("SenseOfFuture") = 2 Then %> selected<% End If %><% End If %>>2 - A Little Worse</option>
				<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("SenseOfFuture") = 3 Then %> selected<% End If %><% End If %>>3 - No Change</option>
				<option value="4" class="formMain"<% If say = "edit" Then %><% If GetPOE("SenseOfFuture") = 4 Then %> selected<% End If %><% End If %>>4 - A Little Better</option>				
				<option value="5" class="formMain"<% If say = "edit" Then %><% If GetPOE("SenseOfFuture") = 5 Then %> selected<% End If %><% End If %>>5 - Much Better</option>
				<option value="6" class="formMain"<% If say = "edit" Then %><% If GetPOE("SenseOfFuture") = 6 Then %> selected<% End If %><% End If %>>6 - Don't Know</option>
				<option value="7" class="formMain"<% If say = "edit" Then %><% If GetPOE("SenseOfFuture") = 7 Then %> selected<% End If %><% End If %>>7 - Not a Problem</option>
			</select>		
		</td>				

		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOECommunityResources">
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("CommunityResources") = 1 Then %> selected<% End If %><% End If %>>1 - Much Worse</option>							
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("CommunityResources") = 2 Then %> selected<% End If %><% End If %>>2 - A Little Worse</option>
				<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("CommunityResources") = 3 Then %> selected<% End If %><% End If %>>3 - No Change</option>
				<option value="4" class="formMain"<% If say = "edit" Then %><% If GetPOE("CommunityResources") = 4 Then %> selected<% End If %><% End If %>>4 - A Little Better</option>				
				<option value="5" class="formMain"<% If say = "edit" Then %><% If GetPOE("CommunityResources") = 5 Then %> selected<% End If %><% End If %>>5 - Much Better</option>
				<option value="6" class="formMain"<% If say = "edit" Then %><% If GetPOE("CommunityResources") = 6 Then %> selected<% End If %><% End If %>>6 - Don't Know</option>
				<option value="7" class="formMain"<% If say = "edit" Then %><% If GetPOE("CommunityResources") = 7 Then %> selected<% End If %><% End If %>>7 - Not a Problem</option>
			</select>		
		</td>				
		
		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOESchoolResources">
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("SchoolResources") = 1 Then %> selected<% End If %><% End If %>>1 - Much Worse</option>							
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("SchoolResources") = 2 Then %> selected<% End If %><% End If %>>2 - A Little Worse</option>
				<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("SchoolResources") = 3 Then %> selected<% End If %><% End If %>>3 - No Change</option>
				<option value="4" class="formMain"<% If say = "edit" Then %><% If GetPOE("SchoolResources") = 4 Then %> selected<% End If %><% End If %>>4 - A Little Better</option>				
				<option value="5" class="formMain"<% If say = "edit" Then %><% If GetPOE("SchoolResources") = 5 Then %> selected<% End If %><% End If %>>5 - Much Better</option>
				<option value="6" class="formMain"<% If say = "edit" Then %><% If GetPOE("SchoolResources") = 6 Then %> selected<% End If %><% End If %>>6 - Don't Know</option>
				<option value="7" class="formMain"<% If say = "edit" Then %><% If GetPOE("SchoolResources") = 7 Then %> selected<% End If %><% End If %>>7 - Not a Problem</option>
			</select>			
		</td>	

		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOEAcademicPerformance">
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("AcademicPerformance") = 1 Then %> selected<% End If %><% End If %>>1 - Much Worse</option>							
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("AcademicPerformance") = 2 Then %> selected<% End If %><% End If %>>2 - A Little Worse</option>
				<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("AcademicPerformance") = 3 Then %> selected<% End If %><% End If %>>3 - No Change</option>
				<option value="4" class="formMain"<% If say = "edit" Then %><% If GetPOE("AcademicPerformance") = 4 Then %> selected<% End If %><% End If %>>4 - A Little Better</option>				
				<option value="5" class="formMain"<% If say = "edit" Then %><% If GetPOE("AcademicPerformance") = 5 Then %> selected<% End If %><% End If %>>5 - Much Better</option>
				<option value="6" class="formMain"<% If say = "edit" Then %><% If GetPOE("AcademicPerformance") = 6 Then %> selected<% End If %><% End If %>>6 - Don't Know</option>
				<option value="7" class="formMain"<% If say = "edit" Then %><% If GetPOE("AcademicPerformance") = 7 Then %> selected<% End If %><% End If %>>7 - Not a Problem</option>
			</select>		
		</td>	
		
		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOEAttitudeTowardSchool">
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("AttitudeTowardSchool") = 1 Then %> selected<% End If %><% End If %>>1 - Much Worse</option>							
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("AttitudeTowardSchool") = 2 Then %> selected<% End If %><% End If %>>2 - A Little Worse</option>
				<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("AttitudeTowardSchool") = 3 Then %> selected<% End If %><% End If %>>3 - No Change</option>
				<option value="4" class="formMain"<% If say = "edit" Then %><% If GetPOE("AttitudeTowardSchool") = 4 Then %> selected<% End If %><% End If %>>4 - A Little Better</option>				
				<option value="5" class="formMain"<% If say = "edit" Then %><% If GetPOE("AttitudeTowardSchool") = 5 Then %> selected<% End If %><% End If %>>5 - Much Better</option>
				<option value="6" class="formMain"<% If say = "edit" Then %><% If GetPOE("AttitudeTowardSchool") = 6 Then %> selected<% End If %><% End If %>>6 - Don't Know</option>
				<option value="7" class="formMain"<% If say = "edit" Then %><% If GetPOE("AttitudeTowardSchool") = 7 Then %> selected<% End If %><% End If %>>7 - Not a Problem</option>
			</select>				
		</td>
			
		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOESchoolPreparedness">
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("SchoolPreparedness") = 1 Then %> selected<% End If %><% End If %>>1 - Much Worse</option>							
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("SchoolPreparedness") = 2 Then %> selected<% End If %><% End If %>>2 - A Little Worse</option>
				<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("SchoolPreparedness") = 3 Then %> selected<% End If %><% End If %>>3 - No Change</option>
				<option value="4" class="formMain"<% If say = "edit" Then %><% If GetPOE("SchoolPreparedness") = 4 Then %> selected<% End If %><% End If %>>4 - A Little Better</option>				
				<option value="5" class="formMain"<% If say = "edit" Then %><% If GetPOE("SchoolPreparedness") = 5 Then %> selected<% End If %><% End If %>>5 - Much Better</option>
				<option value="6" class="formMain"<% If say = "edit" Then %><% If GetPOE("SchoolPreparedness") = 6 Then %> selected<% End If %><% End If %>>6 - Don't Know</option>
				<option value="7" class="formMain"<% If say = "edit" Then %><% If GetPOE("SchoolPreparedness") = 7 Then %> selected<% End If %><% End If %>>7 - Not a Problem</option>
			</select>				
		</td>	
		
		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOEClassParticipation">
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("ClassParticipation") = 1 Then %> selected<% End If %><% End If %>>1 - Much Worse</option>							
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("ClassParticipation") = 2 Then %> selected<% End If %><% End If %>>2 - A Little Worse</option>
				<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("ClassParticipation") = 3 Then %> selected<% End If %><% End If %>>3 - No Change</option>
				<option value="4" class="formMain"<% If say = "edit" Then %><% If GetPOE("ClassParticipation") = 4 Then %> selected<% End If %><% End If %>>4 - A Little Better</option>				
				<option value="5" class="formMain"<% If say = "edit" Then %><% If GetPOE("ClassParticipation") = 5 Then %> selected<% End If %><% End If %>>5 - Much Better</option>
				<option value="6" class="formMain"<% If say = "edit" Then %><% If GetPOE("ClassParticipation") = 6 Then %> selected<% End If %><% End If %>>6 - Don't Know</option>
				<option value="7" class="formMain"<% If say = "edit" Then %><% If GetPOE("ClassParticipation") = 7 Then %> selected<% End If %><% End If %>>7 - Not a Problem</option>
			</select>			
		</td>	
		
		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOEClassroomBehavior">
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("ClassroomBehavior") = 1 Then %> selected<% End If %><% End If %>>1 - Much Worse</option>							
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("ClassroomBehavior") = 2 Then %> selected<% End If %><% End If %>>2 - A Little Worse</option>
				<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("ClassroomBehavior") = 3 Then %> selected<% End If %><% End If %>>3 - No Change</option>
				<option value="4" class="formMain"<% If say = "edit" Then %><% If GetPOE("ClassroomBehavior") = 4 Then %> selected<% End If %><% End If %>>4 - A Little Better</option>				
				<option value="5" class="formMain"<% If say = "edit" Then %><% If GetPOE("ClassroomBehavior") = 5 Then %> selected<% End If %><% End If %>>5 - Much Better</option>
				<option value="6" class="formMain"<% If say = "edit" Then %><% If GetPOE("ClassroomBehavior") = 6 Then %> selected<% End If %><% End If %>>6 - Don't Know</option>
				<option value="7" class="formMain"<% If say = "edit" Then %><% If GetPOE("ClassroomBehavior") = 7 Then %> selected<% End If %><% End If %>>7 - Not a Problem</option>
			</select>					
		</td>			
	</tr>


	<tr>
		<td class="formMainBold">Avoid Delinquency</td>				
		<td class="formMainBold">Avoid Substance Abuse</td>				
		<td class="formMainBold">Avoid Early Parenting</td>				
		<td class="formMainBold">Shows Trust</td>
		<td class="formMainBold">Respects Other Cultures</td>
		<td class="formMainBold">Relationship With Family</td>
		<td class="formMainBold">Relationship With Peers</td>
		<td class="formMainBold">Relationship With Other Adults</td>																				
		<td class="formMainBold">Subject Improvement</td>				
		<td class="formMainBold">Number of Subjects</td>				
	</tr> 
	
	<tr>
		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOEAvoidDelinquency">
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("AvoidDelinquency") = 1 Then %> selected<% End If %><% End If %>>1 - Much Worse</option>							
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("AvoidDelinquency") = 2 Then %> selected<% End If %><% End If %>>2 - A Little Worse</option>
				<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("AvoidDelinquency") = 3 Then %> selected<% End If %><% End If %>>3 - No Change</option>
				<option value="4" class="formMain"<% If say = "edit" Then %><% If GetPOE("AvoidDelinquency") = 4 Then %> selected<% End If %><% End If %>>4 - A Little Better</option>				
				<option value="5" class="formMain"<% If say = "edit" Then %><% If GetPOE("AvoidDelinquency") = 5 Then %> selected<% End If %><% End If %>>5 - Much Better</option>
				<option value="6" class="formMain"<% If say = "edit" Then %><% If GetPOE("AvoidDelinquency") = 6 Then %> selected<% End If %><% End If %>>6 - Don't Know</option>
				<option value="7" class="formMain"<% If say = "edit" Then %><% If GetPOE("AvoidDelinquency") = 7 Then %> selected<% End If %><% End If %>>7 - Not a Problem</option>
			</select>					
		</td>			

		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOEAvoidSubstanceAbuse">
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("AvoidSubstanceAbuse") = 1 Then %> selected<% End If %><% End If %>>1 - Much Worse</option>							
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("AvoidSubstanceAbuse") = 2 Then %> selected<% End If %><% End If %>>2 - A Little Worse</option>
				<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("AvoidSubstanceAbuse") = 3 Then %> selected<% End If %><% End If %>>3 - No Change</option>
				<option value="4" class="formMain"<% If say = "edit" Then %><% If GetPOE("AvoidSubstanceAbuse") = 4 Then %> selected<% End If %><% End If %>>4 - A Little Better</option>				
				<option value="5" class="formMain"<% If say = "edit" Then %><% If GetPOE("AvoidSubstanceAbuse") = 5 Then %> selected<% End If %><% End If %>>5 - Much Better</option>
				<option value="6" class="formMain"<% If say = "edit" Then %><% If GetPOE("AvoidSubstanceAbuse") = 6 Then %> selected<% End If %><% End If %>>6 - Don't Know</option>
				<option value="7" class="formMain"<% If say = "edit" Then %><% If GetPOE("AvoidSubstanceAbuse") = 7 Then %> selected<% End If %><% End If %>>7 - Not a Problem</option>
			</select>					
		</td>				
		
		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOEAvoidEarlyParenting">
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("AvoidEarlyParenting") = 1 Then %> selected<% End If %><% End If %>>1 - Much Worse</option>							
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("AvoidEarlyParenting") = 2 Then %> selected<% End If %><% End If %>>2 - A Little Worse</option>
				<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("AvoidEarlyParenting") = 3 Then %> selected<% End If %><% End If %>>3 - No Change</option>
				<option value="4" class="formMain"<% If say = "edit" Then %><% If GetPOE("AvoidEarlyParenting") = 4 Then %> selected<% End If %><% End If %>>4 - A Little Better</option>				
				<option value="5" class="formMain"<% If say = "edit" Then %><% If GetPOE("AvoidEarlyParenting") = 5 Then %> selected<% End If %><% End If %>>5 - Much Better</option>
				<option value="6" class="formMain"<% If say = "edit" Then %><% If GetPOE("AvoidEarlyParenting") = 6 Then %> selected<% End If %><% End If %>>6 - Don't Know</option>
				<option value="7" class="formMain"<% If say = "edit" Then %><% If GetPOE("AvoidEarlyParenting") = 7 Then %> selected<% End If %><% End If %>>7 - Not a Problem</option>
			</select>			
		</td>
		
		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOEShowsTrust">
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("ShowsTrust") = 1 Then %> selected<% End If %><% End If %>>1 - Much Worse</option>							
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("ShowsTrust") = 2 Then %> selected<% End If %><% End If %>>2 - A Little Worse</option>
				<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("ShowsTrust") = 3 Then %> selected<% End If %><% End If %>>3 - No Change</option>
				<option value="4" class="formMain"<% If say = "edit" Then %><% If GetPOE("ShowsTrust") = 4 Then %> selected<% End If %><% End If %>>4 - A Little Better</option>				
				<option value="5" class="formMain"<% If say = "edit" Then %><% If GetPOE("ShowsTrust") = 5 Then %> selected<% End If %><% End If %>>5 - Much Better</option>
				<option value="6" class="formMain"<% If say = "edit" Then %><% If GetPOE("ShowsTrust") = 6 Then %> selected<% End If %><% End If %>>6 - Don't Know</option>
				<option value="7" class="formMain"<% If say = "edit" Then %><% If GetPOE("ShowsTrust") = 7 Then %> selected<% End If %><% End If %>>7 - Not a Problem</option>
			</select>			
		</td>
		
		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOERespectsOtherCultures">
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("RespectsOtherCultures") = 1 Then %> selected<% End If %><% End If %>>1 - Much Worse</option>							
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("RespectsOtherCultures") = 2 Then %> selected<% End If %><% End If %>>2 - A Little Worse</option>
				<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("RespectsOtherCultures") = 3 Then %> selected<% End If %><% End If %>>3 - No Change</option>
				<option value="4" class="formMain"<% If say = "edit" Then %><% If GetPOE("RespectsOtherCultures") = 4 Then %> selected<% End If %><% End If %>>4 - A Little Better</option>				
				<option value="5" class="formMain"<% If say = "edit" Then %><% If GetPOE("RespectsOtherCultures") = 5 Then %> selected<% End If %><% End If %>>5 - Much Better</option>
				<option value="6" class="formMain"<% If say = "edit" Then %><% If GetPOE("RespectsOtherCultures") = 6 Then %> selected<% End If %><% End If %>>6 - Don't Know</option>
				<option value="7" class="formMain"<% If say = "edit" Then %><% If GetPOE("RespectsOtherCultures") = 7 Then %> selected<% End If %><% End If %>>7 - Not a Problem</option>
			</select>			
		</td>
		
		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOERelationshipWithFamily">
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("RelationshipWithFamily") = 1 Then %> selected<% End If %><% End If %>>1 - Much Worse</option>							
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("RelationshipWithFamily") = 2 Then %> selected<% End If %><% End If %>>2 - A Little Worse</option>
				<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("RelationshipWithFamily") = 3 Then %> selected<% End If %><% End If %>>3 - No Change</option>
				<option value="4" class="formMain"<% If say = "edit" Then %><% If GetPOE("RelationshipWithFamily") = 4 Then %> selected<% End If %><% End If %>>4 - A Little Better</option>				
				<option value="5" class="formMain"<% If say = "edit" Then %><% If GetPOE("RelationshipWithFamily") = 5 Then %> selected<% End If %><% End If %>>5 - Much Better</option>
				<option value="6" class="formMain"<% If say = "edit" Then %><% If GetPOE("RelationshipWithFamily") = 6 Then %> selected<% End If %><% End If %>>6 - Don't Know</option>
				<option value="7" class="formMain"<% If say = "edit" Then %><% If GetPOE("RelationshipWithFamily") = 7 Then %> selected<% End If %><% End If %>>7 - Not a Problem</option>
			</select>			
		</td>
		
		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOERelationshipWithPeers">
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("RelationshipWithPeers") = 1 Then %> selected<% End If %><% End If %>>1 - Much Worse</option>							
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("RelationshipWithPeers") = 2 Then %> selected<% End If %><% End If %>>2 - A Little Worse</option>
				<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("RelationshipWithPeers") = 3 Then %> selected<% End If %><% End If %>>3 - No Change</option>
				<option value="4" class="formMain"<% If say = "edit" Then %><% If GetPOE("RelationshipWithPeers") = 4 Then %> selected<% End If %><% End If %>>4 - A Little Better</option>				
				<option value="5" class="formMain"<% If say = "edit" Then %><% If GetPOE("RelationshipWithPeers") = 5 Then %> selected<% End If %><% End If %>>5 - Much Better</option>
				<option value="6" class="formMain"<% If say = "edit" Then %><% If GetPOE("RelationshipWithPeers") = 6 Then %> selected<% End If %><% End If %>>6 - Don't Know</option>
				<option value="7" class="formMain"<% If say = "edit" Then %><% If GetPOE("RelationshipWithPeers") = 7 Then %> selected<% End If %><% End If %>>7 - Not a Problem</option>
			</select>		
		</td>				
		
		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOERelationshipWithOtherAdults">
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("RelationshipWithOtherAdults") = 1 Then %> selected<% End If %><% End If %>>1 - Much Worse</option>							
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("RelationshipWithOtherAdults") = 2 Then %> selected<% End If %><% End If %>>2 - A Little Worse</option>
				<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("RelationshipWithOtherAdults") = 3 Then %> selected<% End If %><% End If %>>3 - No Change</option>
				<option value="4" class="formMain"<% If say = "edit" Then %><% If GetPOE("RelationshipWithOtherAdults") = 4 Then %> selected<% End If %><% End If %>>4 - A Little Better</option>				
				<option value="5" class="formMain"<% If say = "edit" Then %><% If GetPOE("RelationshipWithOtherAdults") = 5 Then %> selected<% End If %><% End If %>>5 - Much Better</option>
				<option value="6" class="formMain"<% If say = "edit" Then %><% If GetPOE("RelationshipWithOtherAdults") = 6 Then %> selected<% End If %><% End If %>>6 - Don't Know</option>
				<option value="7" class="formMain"<% If say = "edit" Then %><% If GetPOE("RelationshipWithOtherAdults") = 7 Then %> selected<% End If %><% End If %>>7 - Not a Problem</option>
			</select>		
		</td>				
		
		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOESubjectImprovement">
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("SubjectImprovement") = 1 Then %> selected<% End If %><% End If %>>1 - Much Worse</option>							
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("SubjectImprovement") = 2 Then %> selected<% End If %><% End If %>>2 - A Little Worse</option>
				<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("SubjectImprovement") = 3 Then %> selected<% End If %><% End If %>>3 - No Change</option>
				<option value="4" class="formMain"<% If say = "edit" Then %><% If GetPOE("SubjectImprovement") = 4 Then %> selected<% End If %><% End If %>>4 - A Little Better</option>				
				<option value="5" class="formMain"<% If say = "edit" Then %><% If GetPOE("SubjectImprovement") = 5 Then %> selected<% End If %><% End If %>>5 - Much Better</option>
				<option value="6" class="formMain"<% If say = "edit" Then %><% If GetPOE("SubjectImprovement") = 6 Then %> selected<% End If %><% End If %>>6 - Don't Know</option>
				<option value="7" class="formMain"<% If say = "edit" Then %><% If GetPOE("SubjectImprovement") = 7 Then %> selected<% End If %><% End If %>>7 - Not a Problem</option>
			</select>		
		</td>	
		
		<td class="formMain" align="center">
			<select size="1" class="formMain" name="frmPOENumberOfSubjects">
				<option value="1" class="formMain"<% If say = "edit" Then %><% If GetPOE("NumberOfSubjects") = 1 Then %> selected<% End If %><% End If %>>1 - Much Worse</option>							
				<option value="2" class="formMain"<% If say = "edit" Then %><% If GetPOE("NumberOfSubjects") = 2 Then %> selected<% End If %><% End If %>>2 - A Little Worse</option>
				<option value="3" class="formMain"<% If say = "edit" Then %><% If GetPOE("NumberOfSubjects") = 3 Then %> selected<% End If %><% End If %>>3 - No Change</option>
				<option value="4" class="formMain"<% If say = "edit" Then %><% If GetPOE("NumberOfSubjects") = 4 Then %> selected<% End If %><% End If %>>4 - A Little Better</option>				
				<option value="5" class="formMain"<% If say = "edit" Then %><% If GetPOE("NumberOfSubjects") = 5 Then %> selected<% End If %><% End If %>>5 - Much Better</option>
				<option value="6" class="formMain"<% If say = "edit" Then %><% If GetPOE("NumberOfSubjects") = 6 Then %> selected<% End If %><% End If %>>6 - Don't Know</option>
				<option value="7" class="formMain"<% If say = "edit" Then %><% If GetPOE("NumberOfSubjects") = 7 Then %> selected<% End If %><% End If %>>7 - Not a Problem</option>
			</select>		
		</td>							

	</tr>


	<% If say = "edit" Then %>
	<tr>
		<td colspan="11" class="formHeader"><input type="submit" value="Save and Add a New Record" class="formMainBold"></td>
	</tr>
	<% Else %>
	<tr>
		<td colspan="11" class="formHeader"><input type="submit" value="Save This Entry and Add a New Record" class="formMainBold"></td>
	</tr>
	<% End If %>
</table>
</form>
<% End If %>


<% 
If say <> "thanks" Then
 %>
<script language="JavaScript">
<!-- 
	function confirmDelete(row)
	{
		if (confirm("Are you sure you want to delete this record?"))
		{
			location.href = "POE_edit.asp?status=deleteRow&row=" + row + "&y=<%= Request("y") %>";
			// alert("Record deleted.");
		}
		else
		{
		return false;
		}
	}		
// -->
</script>	




<!-- RESULTS TABLE STARTS HERE -->
<!-- <form name="frmPOE" action="POE_edit.asp?status=done" method="post"> -->
<form name="frmPOE" action="POE_Complete.asp?AgencyID=<%=session("AgencyIDN")%>&Order=<%=Order%>" method="post">
			<tr>
				<td><a href="POE_complete_data.asp?AgencyID=<%=session("AgencyIDN")%>&Order=<%=Order%>">go back</a></td>
                <td colspan="8" class="formHeader"><input type="submit" value="Exit" class="formMainBold"></td>
       		</tr>
			<tr>
				<td colspan="8"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
			</tr>
</form>
		</table><br>

<br>
<P>
<% 
End If
 %>
 


</body>
</html>
