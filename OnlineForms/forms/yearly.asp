<!--#include file="../includes/session_stamp.asp"-->
<% 

Dim AssessmentExpired
AssessmentExpired = Request("AssessmentExpired")


Dim StaffLevel

Dim ReadOnlyLevel
If Session("ReadOnly") then
	ReadOnlyLevel=1
Else
	ReadOnlyLevel=0
End If

If Session("StaffFormAccess") then
	StaffLevel="Privileged"
Else
	StaffLevel="Shared"
End if

Dim NoDataEntered 
NoDataEntered=0

' Check for AIM Agency
DIM AIMAgency
Set AIMCon = Server.CreateObject("ADODB.Connection")
AIMCon.Open "BBBSAForms","sa","12sist12"
query = "SELECT AIMReconciled from tblDemogs Where AgencyID = " & int(Session("AgencyIDN"))

Set AIMQuery = AIMCon.Execute(query)

if (AIMQuery.eof) or (AIMQuery.bof) then
	AIMAgency = "n"
else
	if AIMQuery("AIMReconciled") = 1 then
		AIMAgency = "y"
	else
		AIMAgency = "n"
	end if	
end if 

' Check to see if core form was submitted. If so, redirect to appropriate form.

If Request("CoreStatus") = "Bounce"  or Request("SDMStatus") = "Bounce"  or Request("SBMStatus") = "Bounce"  or Request("FinanceStatus") = "Bounce" or Request("DOEStatus") = "Bounce" or Request("MCPStatus") = "Bounce" Then
	z = Split(Request("month"),"-")
	m = z(0)
	y = z(1)
	
	If Request("CoreStatus") = "Bounce" Then
		f = Request("CoreForms")
		Redim x(1)
		x(1) = "Performance"
		Else
			If Request("FinanceStatus") = "Bounce" Then
				f = Request("FinanceForms")
				Redim x(4)
				x(4) = "FinancePerformance"
	End If
End If	
	
Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa", "12sist12"
		query = "SELECT " & x(f) & "ID FROM tbl_frm" & x(f) & " WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(y) & " AND Month=" & Int(m)
		Set GetData = Con.Execute(query)
		If (GetData.EOF OR GetData.BOF) Then
			'show blank form in new month only if selecting core performance
			if ReadonlyLevel = 0 then
													
				' Pull Previous Month's Match Info only if a new record is being created, and only if performance is being selected.  If editing existing record, previous month's match info
				' is pulled in the performance_complete.asp page
					

				if x(f) = "Performance" or x(f) = "DOEPerformance"  or x(f) = "MCPPerformance" then
					Set Con = Server.CreateObject("ADODB.Connection")
					Con.Open "BBBSAforms", "sa","12sist12"
					
					if m = 1 then
						query = "SELECT * FROM tbl_frm" & x(f) & " WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & y-1 & " AND Month=" & m+11
						PrevMonth = 12
						PrevYear = y-1
					else	
						query = "SELECT * FROM tbl_frm" & x(f) & " WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & y & " AND Month=" & m-1
						PrevMonth = m - 1
						PrevYear = y
					end if
					
					Set GetPrev = Con.Execute(query)
					
					if GetPrev.BOF and GetPrev.EOF then
						PrevOpenComm = 0
						PrevOpenSchool = 0
						PrevOpenOther = 0
						PrevOpenGroup = 0
						PrevOpenSpecMent = 0
						PrevOpenSpecNonMent = 0
					else
					
					' Check to see which core performance form is selected.  Different forms use different fields
					
						Select Case x(f)
							Case "Performance"
								PrevOpenComm = GetPrev("OpenMatchesCommunityBased")
								PrevOpenSchool = GetPrev("OpenMatchesSchoolBased")
								PrevOpenOther = GetPrev("OpenMatchesOtherSiteBased")						
								PrevOpenGroup = GetPrev("OpenMatchesGroupMentoring")
								PrevOpenSpecMent = GetPrev("OpenMatchesSpecialProgramsMentoring")
								PrevOpenSpecNonMent = GetPrev("OpenMatchesSpecialProgramsNonMentoring")			
							Case "DOEPerformance"				
								PrevOpenSchool = GetPrev("OpenMatchesSchoolBased")									
							Case "MCPPerformance"
								PrevOpenComm = GetPrev("OpenMatchesCommunityBased")
								PrevOpenSchool = GetPrev("OpenMatchesSchoolBased")
								PrevOpenOther = GetPrev("OpenMatchesOtherSiteBased")						
						End Select							
					
					end if
					
					
					GetPrev.Close
					Set GetPrev = Nothing
				
				
				
					' End Pull Previous Month's Match Info
					
					'Pull Year to Date Info only from Core Performance form and only if a new record is being created and only if month > 1
					
				if x(f)="Performance" then					
											
						dim CommunityYTD
						CommunityYTD = 0
						dim SchoolYTD
						SchoolYTD = 0
						dim OtherYTD
						OtherYTD = 0
						dim RevenueYTD
						RevenueYTD = 0
						
					if m > 1 then
	
					 	Set Con = Server.CreateObject("ADODB.Connection")
						Con.Open "BBBSAforms", "sa","12sist12"
						query = "SELECT * FROM tbl_frmPerformance WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & y & " ORDER BY month"
	
						Set GetYTD = Con.Execute(query)	
	
						' Get First Month's Data
						CommunityYTD = GetYTD("ClosedMatchesCommunityBased") 
						SchoolYTD = GetYTD("ClosedMatchesSchoolBased")
						OtherYTD = GetYTD("ClosedMatchesOtherSiteBased") 
						RevenueYTD = GetYTD("Revenue")	
											
						
						count = 0
						for count = 1 to m - 1
								CommunityYTD = CommunityYTD + GetYTD("ClosedMatchesCommunityBased") 
								SchoolYTD = SchoolYTD + GetYTD("ClosedMatchesSchoolBased")
								OtherYTD = OtherYTD + GetYTD("ClosedMatchesOtherSiteBased") 
								RevenueYTD = RevenueYTD + GetYTD("Revenue")
								GetYTD.MoveNext()	
						next		
						
						GetYTD.Close
						Set GetYTD = Nothing		
						
					 	Set Con = Server.CreateObject("ADODB.Connection")
						Con.Open "BBBSAforms", "sa","12sist12"
						query = "SELECT * FROM tbl_frmPerformance WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & y & " AND Month=" & m-1
					
						Set GetYTD = Con.Execute(query)
						CommunityYTD = CommunityYTD + GetYTD("OpenMatchesCommunityBased")		
						SchoolYTD = SchoolYTD + GetYTD("OpenMatchesSchoolBased")	
						OtherYTD = OtherYTD + GetYTD("OpenMatchesOtherSiteBased")	
						
						
						
						GetYTD.Close
						Set GetYTD = Nothing						
							
						
					end if	
					'End Pull Year to Date Info
				end if
			end if
				
			Response.Redirect(x(f) & "_edit.asp?y=" & y & "&m=" & m & "&PrevOpenComm=" & PrevOpenComm & "&PrevOpenSchool=" & PrevOpenSchool & "&PrevOpenOther=" & PrevOpenOther & "&PrevOpenGroup=" & PrevOpenGroup & "&PrevOpenSpecMent=" & PrevOpenSpecMent & "&PrevOpenSpecNonMent=" & PrevOpenSpecNonMent & "&CommunityYTD=" & CommunityYTD & "&SchoolYTD=" & SchoolYTD & "&OtherYTD=" & OtherYTD & "&RevenueYTD=" & RevenueYTD)


			else 
				NoDataMonth=int(m)
				NoDataEntered=1
			end if
		Else		
			'show complete form w/ edit button
			z = x(f) & "ID"
			'z = "PerformanceID"
			id = GetData(z)
			Response.Redirect(x(f) & "_complete.asp?y=" & y & "&m=" & m & "&id=" & id)
		End If
		GetData.Close
		Set GetData = Nothing	
	Con.Close
	Set Con = Nothing
End If



	
	
	
If Request("status") = "bounce" Then
	y = Request("year")
	f = Request("forms")
	Redim x(8)
	x(1) = "BudgetForecast"
'	x(1) = "SDMInformation"
	x(2) = "Income"
	x(3) = "Expenses"
	x(4) = "Benefits"
	x(5) = "BoardMembers"
	if Session("staffFormAccess") then
		x(6) = "Staff"
	end if
	x(7) = "SelfAssessment"
	
	
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
		query = "SELECT " & x(f) & "ID FROM tbl_frm" & x(f) & " WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(y)
		Set GetData = Con.Execute(query)
		If (GetData.EOF OR GetData.BOF) Then
			'show blank form
			Response.Redirect(x(f) & "_edit.asp?y=" & y)
		Else
			'show complete form w/ edit button
			z = x(f) & "ID"
			id = GetData(z)
			Response.Redirect(x(f) & "_complete.asp?y=" & y & "&id=" & id)
		End If
		GetData.Close
		Set GetData = Nothing	
	Con.Close
	Set Con = Nothing
End If

' Yearly Assessment Form Selections

If Request("status") = "BounceAssessment" Then
	y = Request("year")
	f = Request("forms")
	Redim x(8)
	x(1) = "SelfAssessment"
	x(2) = "SelfAssessment"
	dim section
	if f = 1 then
		section = "Operational"
	else
		section = "Program"
	end if
	
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
		query = "SELECT " & x(f) & "ID FROM tbl_frm" & x(f) & " WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(y)
	'	query = "SELECT " & x(f) & "ID FROM tbl_frm" & x(f) & " WHERE AgencyID='9999' AND Year=" & Int(y)	
		Set GetData = Con.Execute(query)
		If (GetData.EOF OR GetData.BOF) Then
'			show blank form
			Response.Redirect(x(f) & "_edit.asp?y=" & y & "&section=" & section)
'			AssessmentExpired="Yes"
'			Response.Redirect("yearly.asp?AssessmentExpired=Yes")
			
		Else
			'show complete form w/ edit button
			z = x(f) & "ID"
			id = GetData(z)
			Response.Redirect(x(f) & "_complete.asp?y=" & y & "&id=" & id & "&section=" & section)
		End If
		GetData.Close
		Set GetData = Nothing	
	Con.Close
	Set Con = Nothing
End If

' Fee Calculation from

If Request("status") = "FeeForm" Then
	y = 2008

'			Response.Redirect("yearly.asp?AssessmentExpired=Yes")
			Response.Redirect("../../../myagency/feeform.asp?year=" & y & "&Agency_ID=" & Session("AgencyIDN"))
			'Response.Redirect("feeform.asp?Agency_ID=" & id & "&AgencyID=" & Session("AgencyIDN"))
End If


 %>
 
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

function updateYear(form) //updates list of years on topic selection
{
	if (form.forms.selectedIndex == 4)
	{
		form.year.options.length = 0;
		var CurrentTime = new Date();
		var startYear = CurrentTime.getFullYear()-1;
		for (i=0;i<1;i++)
		{
			form.year.options[i] = new Option(startYear,startYear);
			startYear = startYear - 1;
		}
	}
	/*else if (form.forms.selectedIndex == 2)
	{
		form.year.options.length = 0;
		var CurrentTime = new Date();
		var startYear = CurrentTime.getFullYear();
		for (i=0;i<4;i++)
		{
			form.year.options[i] = new Option(startYear,startYear);
			startYear = startYear - 1;
		}
	}*/
	else
	{
		form.year.options.length = 0;
		var CurrentTime = new Date();
		//var startYear = CurrentTime.getFullYear(); //uncomment when HR AAI forms are ready for 2008 collection
		var startYear = 2009
		for (i=0;i<1;i++)
		{
			form.year.options[i] = new Option(startYear,startYear);
			startYear = startYear - 1;
		}
	}
}

//  End -->

</SCRIPT> 
 


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Annual Agency Information Forms (AAI)</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">

<% ' <!--#include file="../includes/top_nav_forms_yearly.inc"--><!-- include file has </head> and <body> tags --><br>     %>
<!--#include file="../includes/surveytitle.inc"-->

<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
<td width="220" valign="top">
<img src="../includes/images/photos_baseball.jpg" alt="" width="220" height="477" border="0">
<br><a href="FormAdminLogin.asp">...</a></td>
<td width="100%" valign="top">


<!-- Finance Form -->
		 <br><br>
		 <font class="formIndex">
		 Monthly Performance Forms</font>
		 
		 
		 <table width = 350 cellpadding="3" cellspacing="2" border="1" bordercolor="#800080">
			
		 <%
				' Get all dates for Performance forms that have been previously entered for the respective agency.
				' Add the next year/month to the drop down list
			
				set Con = Server.CreateObject("ADODB.Connection")
				Con.Open "BBBSAforms", "sa", "12sist12"
				'sql = "p_getPerformanceFormDatesTEMP '" & Session("AgencyIDN") & "'"   ' for Pulling ALL Years slp
				sql = "p_getPerformanceFormDates '" & Session("AgencyIDN") & "'"
				Set rs = Con.Execute(sql)
				
				' Check to see if there are any results. If not, just display the current 
				' month and year.
			
				dim isRsEmpty
				dim thisMonth
				dim thisYear
				dim thisMonthName
				dim currentDate
				currentDate = Date
				isRsEmpty = false
				dim setNextMonth
				setNextMonth = true
					
				
				if rs.BOF and rs.EOF then
					
					' Populate thisMonth and thisYear with the last month and year
					' ( BBBSA does not want the current month to show )
					
					isRsEmpty		= true
					if month(currentDate) = 1 then
						thisMonth		= 12
						thisMonthName	= monthname(thisMonth)
						thisYear		= year(currentDate)-1				
					else
						thisMonth		= month(currentDate) - 1
						thisMonthName	= monthname(thisMonth)
						thisYear		= year(currentDate)
					end if
				
				else	
					
					' Add the next month before the result set based on the most recent date 
					' in the resultset. But, do not add the current month if it would be 
					' the next month after the most recent result set. (The current month should
					' not show )
					
					' get the most recent date in the result set
					rs.MoveFirst
					dim newDate
					dim rsMonth
					dim rsYear
					rsMonth = rs("Month")
					rsYear = rs("Year")
					newDate = rsMonth & "-1-" & rsYear
					newDate = DateValue(newDate)
					newDate = DateAdd("m", 1, newDate)
					thisMonth = month(newDate)
					thisMonthName = monthname(thisMonth)
					thisYear = year(newDate)
					
					' determine if the next month/year in the series would be the current month/year
					
					dim currentMonth
					dim currentYear
					currentMonth = month(currentDate)
					currentYear = year(currentDate)
					if (currentMonth = thisMonth) and (currentYear = thisYear) then 
						setNextMonth = false
					end if
					
					' BEGIN APRIL 2003 SPECIAL BRANCH
					' The following branch is a work around just in case someone has already entered 
					' incomplete April 2003 data (the month/year when this logic to not show the current 
					' month/year in the drop down list was added. )
					
					dim showFirstRecordInResultSet
					showFirstRecordInResultSet = true
					
					if (rsMonth = currentMonth) and (rsYear = currentYear) then
						setNextMonth = false
						showFirstRecordInResultSet = false
					end if
					
					' END APRIL 2003 SPECIAL BRANCH
				end if 
				
			%>

			
			<tr>
			<form method="post" action="monthly.asp">
				<input type="hidden" name="CoreStatus" value="Bounce">
				<input type="hidden" name="CoreForms" value="1">
				<td align="left" bgcolor="#cccccc" class="formMainBold">
				&nbsp;Core Business
				</td>
				<% if AimAgency = "n"  or AdminLevel = 1 then %>
				<td align="left" bgcolor="#cccccc" class="formMainBold">
				
					<select name="month" size=1 class="formMain">
					
						<% 	
							if setNextMonth = true then
								Response.Write "<option	value=""" & thisMonth & "-" & thisYear & """ class=""formMain"">" & thisYear & "-" & thisMonthName & "</option>"
							end if
							
							if isRsEmpty = false then
								rs.MoveFirst
								
								' BEGIN APRIL 2003 SPECIAL BRANCH
								' The following branch is a work around just in case someone has already entered 
								' incomplete April 2003 data (the month/year when this logic to not show the current 
								' month/year in the drop down list was added. )
								
								if showFirstRecordInResultSet = false then
									rs.MoveNext
								end if
								
								' END APRIL 2003 SPECIAL BRANCH
								
								
								while not rs.EOF
									thisMonth = rs("Month") 
									thisMonthName = MonthName(thisMonth)
									thisYear = rs("Year")
									Response.Write "<option	value=""" & thisMonth & "-" & thisYear & """ class=""formMain"">" & thisYear & "-" & thisMonthName & "</option>"	
									rs.MoveNext
								wend
							end if 
						%>
					</select>
					<td align="left" bgcolor="#cccccc">&nbsp;<input type="submit" value="Go" class="formMainBold"></td>
				<% else %>
					<td align="left" colspan="2" bgcolor="#cccccc" class="formMainBold">				
					AIM Agency - No Longer Required
				<% end if %>	
				</td>


			</form>
			</tr>

				<%
			
				set Con = Server.CreateObject("ADODB.Connection")
				Con.Open "BBBSAforms", "sa", "12sist12"
				sql = "p_getFinancePerformanceFormDates '" & Session("AgencyIDN") & "'"
				Set rs = Con.Execute(sql)
				
				' Check to see if there are any results. If not, just display the current 
				' month and year.
			

				currentDate = Date
				isRsEmpty = false
				setNextMonth = true
					
				
				if rs.BOF and rs.EOF then
					
					' Populate thisMonth and thisYear with the last month and year
					' ( BBBSA does not want the current month to show )
					
					isRsEmpty		= true
					if month(currentDate) = 1 then
						thisMonth		= 12
						thisMonthName	= monthname(thisMonth)
						thisYear		= year(currentDate)-1				
					else
						thisMonth		= month(currentDate) - 1
						thisMonthName	= monthname(thisMonth)
						thisYear		= year(currentDate)
					end if
				
				else	
					
					' Add the next month before the result set based on the most recent date 
					' in the resultset. But, do not add the current month if it would be 
					' the next month after the most recent result set. (The current month should
					' not show )
					
					' get the most recent date in the result set
					rs.MoveFirst
					rsMonth = rs("Month")
					rsYear = rs("Year")
					newDate = rsMonth & "-1-" & rsYear
					newDate = DateValue(newDate)
					newDate = DateAdd("m", 1, newDate)
					thisMonth = month(newDate)
					thisMonthName = monthname(thisMonth)
					thisYear = year(newDate)
					
					' determine if the next month/year in the series would be the current month/year
					
					currentMonth = month(currentDate)
					currentYear = year(currentDate)
					if (currentMonth = thisMonth) and (currentYear = thisYear) then 
						setNextMonth = false
					end if
				end if 
				
			%>
			
			<tr>
			<form method="post" action="monthly.asp">
				<input type="hidden" name="FinanceStatus" value="Bounce">
				<input type="hidden" name="FinanceForms" value="4">
				<td align="left" bgcolor="#cccccc" class="formMainBold">				
				&nbsp;Revenue / Expense
				</td>
				<td align="left" bgcolor="#cccccc">				
				<select name="month" size=1 class="formMain">
					<% 	
						if setNextMonth = true then
							Response.Write "<option	value=""" & thisMonth & "-" & thisYear & """ class=""formMain"">" & thisYear & "-" & thisMonthName & "</option>"
						end if
						
						if isRsEmpty = false then
							rs.MoveFirst
							
							
							while not rs.EOF
								thisMonth = rs("Month") 
								thisMonthName = MonthName(thisMonth)
								thisYear = rs("Year")
								Response.Write "<option	value=""" & thisMonth & "-" & thisYear & """ class=""formMain"">" & thisYear & "-" & thisMonthName & "</option>"	
								rs.MoveNext
							wend
						end if 
					%>
				</select>
				</td>
				<td align="left" bgcolor="#cccccc">				
				&nbsp;<input type="submit" value="Go" class="formMainBold">
				</td>
				</tr>
			</form>
			</table>
<!-- End Finance Form -->






<br><br>
<font class="formIndex">
HR Annual Agency Information (AAI)</font>
<% if StaffLevel = "Shared" or ReadOnlyLevel = 1 then %>
	<p><span class="formMain" ><em><strong>Please Note: </strong>The Staff Form is not available.  Only users with the "Full Access" password (Agency ED's/CEO's) may access the Staff Form.</em></span></p>
<% End If %>




<table width = 250 cellpadding="3" cellspacing="2" border="1" bordercolor="#800080">
<tr>
<form method="post" action="yearly.asp">
<input type="hidden" name="status" value="bounce">

<td align="left" bgcolor="#c0c0c0">
<select name="forms" size=1 class="formMain" onchange="updateYear(this.form)">
				<% if Session("staffFormAccess") then %>
					 <option value="6" class="formMain">Staff
				<% end if %>
				<option value="1" class="formMain">Staffing Expense Summary
<!-- <option value="1" class="formMain">SDM Information -->
<!-- <option value="2" class="formMain">Revenue -->
		 		<option value="4" class="formMain">Benefits
				<option value="5" class="formMain">Board Members
<!--  <option value="3" class="formMain">Finances -->

<% 'if latestmonth < 10 then %>
<!-- <option value="7" class="formMain">End Of Year Performance -->
<% ' end if %>
</select>&nbsp;
</td>

<td align="left" bgcolor="#c0c0c0">
<select name="year" size=1 class="formMain">
  <% 
  ' y = 2003
  ' ydisplay = 2003
  'y = 2004
  ydisplay = 2009
  'If Year(Now) > (y+1) Then
  	'ydisplay = (Int(Year(Now))+1)
  	'ydisplay = (Int(Year(Now))-1) 'comment of good code - comented prior to starting 2008 HR AAI
  	Do Until ydisplay = (Int(Year(Now))-1)
   %>
   <option value="<%= ydisplay %>" class="formMain"><%= ydisplay %>
  <% 
  		ydisplay = (ydisplay - 1)
  		'y = (y + 1)
  	Loop
  'Else
  %>
  <!--<option value="<%= y %>" class="formMain"><%= y %>-->
  <% 
  'End if
   %>
</select>&nbsp;
</td>

<td align="left" bgcolor="#c0c0c0">
<input type="submit" value="Go" class="formMainBold">
</td>
</form>
</tr>

</table>

<br><br>
<font class="formIndex">
Finance Annual Agency Information (AAI)</font>

<table width = 250 cellpadding="3" cellspacing="2" border="1" bordercolor="#800080">
  <tr>
    <form method="post" action="yearly.asp">
    <input type="hidden" name="status" value="bounce">
  
    <td align="left" bgcolor="#c0c0c0">
      <select name="forms" size=1 class="formMain" onchange="updateYear(this.form)">
      				<option value="3" class="formMain">Finances  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      </select>&nbsp;
    </td>
  
    <td align="left" bgcolor="#c0c0c0">
      <select name="year" size=1 class="formMain">
          <% 
            ydisplay = 2008
          	Do Until ydisplay = (Int(Year(Now))-2)
          %>
             <option value="<%= ydisplay %>" class="formMain"><%= ydisplay %>
          <% 
          		ydisplay = (ydisplay - 1)
          	Loop
          %>
      </select>&nbsp;
    </td>
  
    <td align="left" bgcolor="#c0c0c0">
      <input type="submit" value="Go" class="formMainBold">
    </td>
    </form>
  </tr>
</table>


<br><br>
<font class="formIndex">
Annual Self-Assessment</font>

<table width = 250 cellpadding="3" cellspacing="2" border="1" bordercolor="#800080">
<tr>
<form method="post" action="yearly.asp">
<input type="hidden" name="status" value="BounceAssessment">

<!--<td> <b>Available December 1st</b></td> <!--Remove when self assessment is ready for submission, and uncoment lines below-->

<td align="left" bgcolor="#c0c0c0">
<select name="forms" size=1 class="formMain">
<option value="1" class="formMain">Operational Standards
<option value="2" class="formMain">Program Standards

</select>&nbsp;
</td>
<td align="left" bgcolor="#c0c0c0">
<select name="year" size=1 class="formMain">
<% 
 y = 2008 '2007 to display last to years
 ydisplay = 2006
  If Year(Now) > (y+1) Then
 	ydisplay = (Int(Year(Now))+1) - 2
	Do Until y = (Int(Year(Now))+1) - 1
 %>
<option value="<%= ydisplay %>" class="formMain"><%= ydisplay %>
<% 
		ydisplay = (ydisplay - 1)
		y = (y + 1)
	Loop
Else
 %>
<option value="<%= y %>" class="formMain"><%= y %>
<% 
End if
 %>
</select>&nbsp;
</td>
<td align="left" bgcolor="#c0c0c0">
<!--<input type=button value="Go" class="formMainBold" ID="temp1" NAME="temp1"> <!--Remove when self assessment is ready for submission, and uncoment lines below-->
<input type="submit" value="Go" class="formMainBold">
</td>
</form>
</tr>

</table>

<% if AssessmentExpired = "Yes" then %>
<span class="formMain">
<font color="red">
<strong>
<br>Editing of 2004 Assessment Forms is no longer available.  Please contact us at <a href="mailto:affiliatereview@bbbsa.org" class="cool3">affiliatereview@bbbsa.org</a> with any questions.<br>
</strong>
</font>
</span>

<%end if %>


<br><br>
<form method="post" action="yearly.asp">
<input type="hidden" name="status" value="FeeForm">
<input type="submit" value="Click Here to Submit Fee Calculation Form" class="formMainBold">
</form>

<br>
<span class="formMain">
<!-- Changes have been made to the yearly forms. <a href="../helpfiles/surveyhelp.asp?HelpID=yearly1" onclick="NewWindow(this.href,'name','700','400','yes');return false;">Click Here</a> for an explanation. -->
</span>

<br>
<!--#include file="../includes/contact_info.inc"-->
<br>

<P>

</td>
</tr>
</table>

</body>
</html>
