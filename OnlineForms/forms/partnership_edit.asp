<%@LANGUAGE="VBSCRIPT"%>

<!--#include file="../includes/NAD_BE.asp" -->


<head>
	<title>Partnership Questionnaire</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
	
</head>


<form name="frmPartnership" action="Partnership_edit.asp" method="post">
<table width="550" border="1" cellspacing="0" cellpadding="3" bordercolordark="#003063" >
<tr>
	<td colspan="7" class="formHeader">PARTNERSHIP QUESTIONNAIRE</td>
</tr>


<!-- Active Matches -->
<tr>
	<td align="center" colspan="7" class="formmain">If Applicable, Enter the number of <strong>Active Matches</strong> with the following organizations:</td>
</tr>

<tr>
	<td>&nbsp;</td>
	<td align="center" class="formMain">Community<br>Based</td>
	<td align="center" class="formMain">School<br>Based</td>	
	<td align="center" class="formMain">Other<br>Site Based</td>	
	<td align="center" class="formMain">Not<br>Partnering</td>	
	<td align="center" colspan="2" class="formMain"><em>If Not Partnering,</em> interested <br>in forming a partnership?</td>
</tr>

<!-- Alpha Phi Alpha -->
<tr>
	<td align="left" class="formmain">Alpha Phi Alpha</td>
	
	<!-- Alpha Community Based -->
	<td align="center" class="formmain">
		<input type="text"  class="formMain" size="5" maxlength="10" value="0" name="frmPartnershipsAlphaCommunityBased" tabindex="1" onchange="checkForIntegerCommas(this.value);">
	</td>
	
	<!-- Alpha School Based -->
	<td align="center" class="formmain">
		<input type="text"  class="formMain" size="5" maxlength="10" value="0" name="frmPartnershipsAlphaSchoolBased" tabindex="1" onchange="checkForIntegerCommas(this.value);">
	</td>	
	
	<!-- Alpha Other Site Based -->
	<td align="center" class="formmain">
		<input type="text"  class="formMain" size="5" maxlength="10" value="0" name="frmPartnershipsAlphaOtherSiteBased" tabindex="1" onchange="checkForIntegerCommas(this.value);">
	</td>		
	
	<!-- Alpha Not Partnering -->
	<td align="center" class="formmain">
		<input type="Checkbox" name="frmPartnershipsAlphaNotPartnering" value="1">
	</td>	
	
	<!-- Alpha Interest -->
	<td align="center" colspan="2" class="formmain">
		<input type="radio" name="frmPartnershipsAlphainterest" value="Yes" checked>Yes
		<input type="radio" name="frmPartnershipsAlphainterest" value="No">No
	</td>
	
</tr>

<!-- Lions Club -->
<tr>
	<td align="left" class="formmain">Lions Club</td>	
	
	<!-- Lions Community Based -->
	<td align="center" class="formmain">
		<input type="text"  class="formMain" size="5" maxlength="10" value="0" name="frmPartnershipsLionsCommunityBased" tabindex="1" onchange="checkForIntegerCommas(this.value);">
	</td>
	
	<!-- Lions School Based -->
	<td align="center" class="formmain">
		<input type="text"  class="formMain" size="5" maxlength="10" value="0" name="frmPartnershipsLionsSchoolBased" tabindex="1" onchange="checkForIntegerCommas(this.value);">
	</td>
	
	<!-- Lions Other Site Based -->
	<td align="center" class="formmain">
		<input type="text"  class="formMain" size="5" maxlength="10" value="0" name="frmPartnershipsLionsOtherSiteBased" tabindex="1" onchange="checkForIntegerCommas(this.value);">
	</td>	
	
	<!-- Lions Not Partnering -->
	<td align="center" class="formmain">
		<input type="Checkbox" name="frmpartnershipsLionsNotPartnering" value="1">
	</td>

	<!-- Lions Interest -->
	<td align="center" colspan="2" class="formmain">
		<input type="radio" name="frmpartnershipslionsinterest" value="Yes" checked>Yes
		<input type="radio" name="frmpartnershipslionsinterest" value="No">No
	</td>	
			
</tr>

<!-- Rotary Club -->
<tr>
	<td align="left" class="formmain">Rotary Club</td>	
	
	<!-- Rotary Community Based -->
	<td align="center" class="formmain">
		<input type="text"  class="formMain" size="5" maxlength="10" value="0" name="frmPartnershipsRotaryCommunityBased" tabindex="1" onchange="checkForIntegerCommas(this.value);">
	</td>
	
	<!-- Rotary School Based -->
	<td align="center" class="formmain">
		<input type="text"  class="formMain" size="5" maxlength="10" value="0" name="frmPartnershipsRotarySchoolBased" tabindex="1" onchange="checkForIntegerCommas(this.value);">
	</td>
	
	<!-- Rotary Other Site Based -->
	<td align="center" class="formmain">
		<input type="text"  class="formMain" size="5" maxlength="10" value="0" name="frmPartnershipsRotaryOtherSiteBased" tabindex="1" onchange="checkForIntegerCommas(this.value);">
	</td>	
	
	<!-- Rotary Not Partnering -->
	<td align="center" class="formmain">
		<input type="Checkbox" name="frmPartnershipsRotaryNotPartnering" value="1">
	</td>	
	
	<!-- Rotary Interest -->
	<td align="center" colspan="2" class="formmain">
		<input type="radio" name="frmpartnershipsRotaryinterest" value="Yes" checked>Yes
		<input type="radio" name="frmpartnershipsRotaryinterest" value="No">No
	</td>		
	
</tr>

<!-- Kiwanis Club -->
<tr>
	<td align="left" class="formmain">Kiwanis Club</td>	
	
	<!-- Kiwanis Community Based -->
	<td align="center" class="formmain">
		<input type="text"  class="formMain" size="5" maxlength="10" value="0" name="frmPartnershipsKiwanisCommunityBased" tabindex="1" onchange="checkForIntegerCommas(this.value);">
	</td>
	
	<!-- Kiwanis School Based -->
	<td align="center" class="formmain">
		<input type="text"  class="formMain" size="5" maxlength="10" value="0" name="frmPartnershipsKiwanisSchoolBased" tabindex="1" onchange="checkForIntegerCommas(this.value);">
	</td>
	
	<!-- Kiwanis Other Site Based -->
	<td align="center" class="formmain">
		<input type="text"  class="formMain" size="5" maxlength="10" value="0" name="frmPartnershipsKiwanisOtherSiteBased" tabindex="1" onchange="checkForIntegerCommas(this.value);">
	</td>	
	
	<!-- Kiwanis Not Partnering -->
	<td align="center" class="formmain">
		<input type="Checkbox" name="frmPartnershipsKiwanisNotPartnering" value="1">
	</td>		
	
	<!-- Kiwanis Interest -->
	<td align="center" colspan="2" class="formmain">
		<input type="radio" name="frmpartnershipsKiwanisinterest" value="Yes" checked>Yes
		<input type="radio" name="frmpartnershipsKiwanisinterest" value="No">No
	</td>		
	
</tr>


<!-- Optimist Club -->
<tr>
	<td align="left" class="formmain">Optimist Club</td>	
	
	<!-- Optimist Community Based -->
	<td align="center" class="formmain">
		<input type="text"  class="formMain" size="5" maxlength="10" value="0" name="frmPartnershipsOptimistCommunityBased" tabindex="1" onchange="checkForIntegerCommas(this.value);">
	</td>
	
	<!-- Optimist School Based -->
	<td align="center" class="formmain">
		<input type="text"  class="formMain" size="5" maxlength="10" value="0" name="frmPartnershipsOptimistSchoolBased" tabindex="1" onchange="checkForIntegerCommas(this.value);">
	</td>
	
	<!-- Optimist Other Site Based -->
	<td align="center" class="formmain">
		<input type="text"  class="formMain" size="5" maxlength="10" value="0" name="frmPartnershipsOptimistOtherSiteBased" tabindex="1" onchange="checkForIntegerCommas(this.value);">
	</td>	
	
	<!-- Optimist Not Partnering -->
	<td align="center" class="formmain">
		<input type="Checkbox" name="frmPartnershipsOptimistNotPartnering" value="1">
	</td>		
	
	<!-- Optimist Interest -->
	<td align="center" colspan="2" class="formmain">
		<input type="radio" name="frmpartnershipsOptimistinterest" value="Yes" checked>Yes
		<input type="radio" name="frmpartnershipsOptimistinterest" value="No">No
	</td>		
	
</tr>


<!-- AARP Club -->
<tr>
	<td align="left" class="formmain">AARP</td>	
	
	<!-- AARP Community Based -->
	<td align="center" class="formmain">
		<input type="text"  class="formMain" size="5" maxlength="10" value="0" name="frmPartnershipsAARPCommunityBased" tabindex="1" onchange="checkForIntegerCommas(this.value);">
	</td>
	
	<!-- AARP School Based -->
	<td align="center" class="formmain">
		<input type="text"  class="formMain" size="5" maxlength="10" value="0" name="frmPartnershipsAARPSchoolBased" tabindex="1" onchange="checkForIntegerCommas(this.value);">
	</td>
	
	<!-- AARP Other Site Based -->
	<td align="center" class="formmain">
		<input type="text"  class="formMain" size="5" maxlength="10" value="0" name="frmPartnershipsAARPOtherSiteBased" tabindex="1" onchange="checkForIntegerCommas(this.value);">
	</td>	
	
	<!-- AARP Not Partnering -->
	<td align="center" class="formmain">
		<input type="Checkbox" name="frmPartnershipsAARPNotPartnering" value="1">
	</td>		
	
	<!-- AARP Interest -->
	<td align="center" colspan="2" class="formmain">
		<input type="radio" name="frmpartnershipsAARPinterest" value="Yes" checked>Yes
		<input type="radio" name="frmpartnershipsAARPinterest" value="No">No
	</td>		
	
</tr>

<!-- Partnership Rating -->
<tr>
	<td colspan="7" class="formHeaderMedium">PARTNERSHIP RATING</td>	
</tr>

<tr>
	<td colspan="7" align="center" class="formMain">Rate the Nature of the Partnership from 1 to 5 - based on level of interaction, with 5 being the highest -  or select 'Not Applicable'</td>	
</tr>

<!-- Alpha Rating -->
<tr>
	<td class="formMain" colspan="2">Alpha Phi Alpha</td>

	<td class="formMain" colspan="5" align="left">	
	<select name="frmPartnershipAlphaRating">
	<option value="0">Not Applicable</option>
	<option value="1" selected>1 - Informal</option>
	<option value="2">2</option>
	<option value="3">3</option>
	<option value="4">4</option>
	<option value="5">5 - Formal</option>	
	</select>
	</td>
</tr>

<!-- Lions Club Rating -->
<tr>
	<td class="formMain" colspan="2">Lions Club</td>
	<td class="formMain" colspan="5" align="left">	
	<select name="frmPartnershipLionsRating">
	<option value="0">Not Applicable</option>
	<option value="1" selected>1 - Informal</option>
	<option value="2">2</option>
	<option value="3">3</option>
	<option value="4">4</option>
	<option value="5">5 - Formal</option>	
	</select>
	</td>	
</tr>

<!-- Rotary Club Rating -->
<tr>
	<td class="formMain" colspan="2">Rotary Club</td>
	<td class="formMain" colspan="5" align="left">	
	<select name="frmPartnershipRotaryRating">
	<option value="0">Not Applicable</option>
	<option value="1" selected>1 - Informal</option>
	<option value="2">2</option>
	<option value="3">3</option>
	<option value="4">4</option>
	<option value="5">5 - Formal</option>	
	</select>
	</td>	
</tr>

<!-- Kiwanis Club Rating -->
<tr>
	<td class="formMain" colspan="2">Kiwanis Club</td>
	<td class="formMain" colspan="5" align="left">	
	<select name="frmPartnershipKiwanisRating">
	<option value="0">Not Applicable</option>
	<option value="1" selected>1 - Informal</option>
	<option value="2">2</option>
	<option value="3">3</option>
	<option value="4">4</option>
	<option value="5">5 - Formal</option>	
	</select>
	</td>	
</tr>

<!-- Optimist Club Rating -->
<tr>
	<td class="formMain" colspan="2">Optimist Club</td>
	<td class="formMain" colspan="5" align="left">	
	<select name="frmPartnershipOptimistRating">
	<option value="0">Not Applicable</option>
	<option value="1" selected>1 - Informal</option>
	<option value="2">2</option>
	<option value="3">3</option>
	<option value="4">4</option>
	<option value="5">5 - Formal</option>	
	</select>
	</td>	
</tr>

<!-- AARP Rating -->
<tr>
	<td class="formMain" colspan="2">AARP</td>
	<td class="formMain" colspan="5" align="left">	
	<select name="frmPartnershipAARPRating">
	<option value="0">Not Applicable</option>
	<option value="1" selected>1 - Informal</option>
	<option value="2">2</option>
	<option value="3">3</option>
	<option value="4">4</option>
	<option value="5">5 - Formal</option>	
	</select>
	</td>	
</tr>

<!-- Alpha Phi Alpha Partnership -->
<tr>
	<td colspan="7" class="formHeaderMedium">ALPHA PHI ALPHA PARTNERSHIP</td>	
</tr>

<tr>
	<td colspan="7" align="center" class="formMain">I am partnering with the Alphas in the following ways (check all that apply):</td>
</tr>

<tr>
	<td colspan="7" align="left" class="formMain">
	<input type="Checkbox" name="frmPartnershipFunding" value="1">Funding: Alpha chapter supports BBBS funding efforts<br>
	<input type="Checkbox" name="frmPartnershipProgramInitiative" value="1">Program Initiative: Chapter has activities with children on waiting list<br>
	<input type="Checkbox" name="frmPartnershipLeadershipInvolvement" value="1">Leadership Involvement: Alpha serves on board, provides agency with professional skills and resources (serves as volunteer)
	
	</td>
</tr>


<!-- Chapter Locations -->

<%
set StateChoices = Server.CreateObject("ADODB.Recordset")
StateChoices.ActiveConnection = ConnStr
StateChoices.Source = "SELECT DISTINCT StateSpelledOut,StateAbbreviation FROM tblAGLUST order by StateSpelledOut"
StateChoices.CursorType = 0
StateChoices.CursorLocation = 2
StateChoices.Open()
%>



<tr>
	<td colspan="7" align="center" class="formMain">Please enter the name and location of your local Alpha Phi Alpha Chapter(s):</td>
</tr>

<tr>
	<td align="left" class="formMain">Undergraduate Chapter</td>
	<td align="left" colspan="6" class="formMain">
	Name:&nbsp;<input type="text" name="frmPartnershipUndergradChapterName" size="50" value=" "><br>
	City:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="frmPartnershipUndergradChapterCity" size="40" value=" "><br>
	State:&nbsp;&nbsp;
	<select NAME="frmPartnershipUndergradChapterState">
	  <option value=""></option>
	  <%
	  While (NOT StateChoices.EOF)
	  %>
	  <option value="<%=(StateChoices.Fields.Item("StateAbbreviation").Value)%>"><%=(StateChoices.Fields.Item("StateSpelledOut").Value)%></option>
	  <%
	   StateChoices.MoveNext()
	  Wend
	  %>
	</select>
	</td>	
</tr>

<% StateChoices.MoveFirst() %>
<tr>
	<td align="left" class="formMain">Alumni Chapter</td>
	<td align="left" colspan="6" class="formMain">
	Name:&nbsp;<input type="text" name="frmPartnershipAlumniChapterName" size="50" value=" "><br>
	City:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="frmPartnershipAlumniChapterCity" size="40" value=" "><br>
	State:&nbsp;&nbsp;
	<select NAME="frmPartnershipAlumniChapterState">
	  <option value=""></option>
	  <%
	  While (NOT StateChoices.EOF)
	  %>
	  <option value="<%=(StateChoices.Fields.Item("StateAbbreviation").Value)%>"><%=(StateChoices.Fields.Item("StateSpelledOut").Value)%></option>
	  <%
	   StateChoices.MoveNext()
	  Wend
	  %>
	</select>
	</td>	
</tr>


<tr>
<td class="formHeader" colspan="7">
<input type="submit" value="Save" class="formMainBold">
</td>
</tr>
</table>

</form>
