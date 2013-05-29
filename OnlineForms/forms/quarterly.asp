<!--#include file="../includes/session_stamp.asp"-->

<% Dim ReadOnlyLevel
If Session("ReadOnly") then
	ReadOnlyLevel=1
Else
	ReadOnlyLevel=0
End If
%>

<% Dim AdminLevel
If Session("Admin") then 
	AdminLevel = 1
Else
	AdminLevel = 0
End If
%>


<% Dim NoDataEntered 
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

' Check for DOE Agency
DIM DOEAgency
Set DOECon = Server.CreateObject("ADODB.Connection")
DOECon.Open "BBBSAforms","sa","12sist12"
query = "SELECT * FROM tblGrants_Grantees WHERE GranteeAgencyID = " & int(Session("AgencyIDN")) & " and (GranteesGroupID = 30 or GranteesGroupID = 31) " 

Set DOEQuery = DOECon.Execute(query)
if (DOEquery.eof) then
	DOEAgency = 0
else
	DOEAgency = 1
End if


' Check for MCP Agency
DIM MCPAgency
Set MCPCon = Server.CreateObject("ADODB.Connection")
MCPCon.Open "BBBSAforms","sa","12sist12"
query = "SELECT * FROM tblGrants_Grantees WHERE GranteeAgencyID = " & int(Session("AgencyIDN")) & " and (GranteesGroupID = 33 or GranteesGroupID = 34) " 

Set MCPQuery = MCPCon.Execute(query)
if (MCPquery.eof) then
	MCPAgency = 0
else
	MCPAgency = 1
End if


' Check for SBM Agency
DIM SBMAgency
Set SBMCon = Server.CreateObject("ADODB.Connection")
SBMCon.Open "BBBSAforms","sa","12sist12"
query = "SELECT SBM FROM tbl_AgencyInfo WHERE AgencyID = '" & Session("AgencyIDN") & "' and SBM = -1  " 
Set SBMQuery = SBMCon.Execute(query)
if (SBMquery.eof) then
	SBMAgency = 0
else
	SBMAgency = 1
End if

' Check to see if this form was submitted. If so, redirect to appropriate form.
If Request("QFPStatus") = "Bounce"  Then
    Dim y, q, id, ero
    q = left(Request("quarter"), 1)
    y = mid(Request("quarter"), 2, 4)
    ero = right(Request("quarter"), 1)
    id = Session("AgencyIDN")
    If ero = "e" Then
		Response.Redirect("QuarterlyCashflow_edit.asp?y=" & y & "&q=" & q & "&id=" & id)
	ElseIf ero = "r" Then
		Response.Redirect("QuarterlyCashflow_complete.asp?y=" & y & "&q=" & q & "&id=" & id)
	else
	End If
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

//  End -->

</SCRIPT>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Quarterly Performance Forms</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
	<% ' <!--#include file="../includes/top_nav_forms_monthly.inc"--><!-- include file has </head> and <body> tags --><br>     %>
	<!--#include file="../includes/surveytitle.inc"-->
	
	<table width="100%" cellspacing="0" cellpadding="0" border="0">
		<tr>
			<td width="220" valign="top">
			<img src="../includes/images/photos_football.jpg" alt="" width="220" height="477" border="0">
			<br><a href="FormAdminLogin.asp">...</a>
			</td>
			<td width="100%" valign="top">
			<br><br>
			<span class="formIndex">Quarterly Performance Forms</span><br><br>
			
			<%
				' Get all dates for Performance forms that have been previously entered for the respective agency.
				' Add the next year/month to the drop down list
			
				set Con = Server.CreateObject("ADODB.Connection")
				Con.Open "BBBSAforms", "sa", "12sist12"
				sql = "p_getPerformanceFormDatesTEMP '" & Session("AgencyIDN") & "'"   ' for Pulling All Years Wendy
'				sql = "p_getPerformanceFormDates '" & Session("AgencyIDN") & "'"
				Set rs = Con.Execute(sql)
				
				' Check to see if there are any results. If not, just display the current 
				' month and year.
			
				dim isRsEmpty
				dim thisMonth
				dim thisYear
				dim thisMonthName
				dim currentDate
				dim currentQuarter
				dim inputyear
				dim iQtr
				Dim iYear
				Dim editroswitch
				currentDate = Date
				isRsEmpty = false
				dim setNextMonth
				setNextMonth = true
					
				
				if rs.BOF and rs.EOF then
					
					' Populate thisMonth and thisYear with the last month and year
					' ( BBBSA does not want the current month to show )
					
					isRsEmpty		= true
					thisMonth		= month(currentDate) - 1
					thisMonthName	= monthname(thisMonth)
					thisYear		= year(currentDate)
				
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


			<table width = 400 cellpadding="3" cellspacing="2" border="1" bordercolor="#800080">
			
			<%
			if NoDataEntered = 1 then %>
			
			<span class="formMain">
			
			
				<br><b><em>No Data Has Been Entered for the month of <%=monthname(NoDataMonth)%></em></b>
				<br><br>
				<% if ReadOnlyLevel = 1 then %>
				<hr>
				I was able to enter a new month's data in the past. Why am I restricted from entering data now?<br><a href="..\helpfiles\surveyhelp.asp?HelpID=password1" onclick="NewWindow(this.href,'name','500','250','yes');return false;">Click Here</a> for an explanation.
				<hr>
				<% end if %>
				
			
			</span>	
				
				
			<% end if %>
			
			<% 
				' clean up
				rs.close
				set rs = nothing
				Con.close
				set Con = nothing
			%>
			
			
<!-- Quarterly Finance Performance Input  -->	
<%
'Find starting Quarter
Select case currentMonth
    Case 1, 2, 3 'Q4 Last year
        currentQuarter = 4
		inputyear = currentyear - 1
    Case 4, 5, 6 'Q1 This year
        currentQuarter = 1
        inputyear = currentyear
    Case 7, 8, 9
        currentQuarter = 2
        inputyear = currentyear
    Case 10, 11, 12
        currentQuarter = 3
        inputyear = currentyear
End Select
iYear = inputyear 
iQtr = currentQuarter
editroswitch = "e"
%>
			<tr>

			<form method="post" action="quarterly.asp" id=form1 name=form1>
				<input type="hidden" name="QFPStatus" value="Bounce">
				<input type="hidden" name="QFPForms" value="3">
				<td align="left" bgcolor="#cccccc" class="formMainBold">
				&nbsp;Quarterly Balance Sheet Report
				</td>
				<td align="left" bgcolor="#cccccc">		
						<%if currentMonth-1=0 then
							currentMonth=13
						end if%>
				<select name="quarter" size=1 class="formMain">
					<%
					Do while (iYear & "0" & iQtr) >= 200604 'Change to 200604 @@
					    Response.Write("<option	value=""" & iQtr & iYear & editroswitch & """ class=""formMain"">" & "Q" & iQtr & "-" & iYear & "</option>")
						Select Case iQtr
						    Case 1
								iYear = iYear - 1
								iQtr = 4
						    Case 2,3,4
								iQtr = iQtr - 1
						End Select
						editroswitch = "r"
					Loop					
					%>
				</select>
				</td>
				<td align="left" bgcolor="#cccccc">				
				&nbsp;<input type="submit" value="Go" class="formMainBold" id=submit1 name=submit1>
				</td>
				</tr>
			</form>
<!-- End Quarterly Finance Performance Input  -->	

		</table>

			<br>
			<!--#include file="../includes/contact_info.inc"-->
			<br>
			</td>
		</tr>
	</table>
			
</body>
</html>
