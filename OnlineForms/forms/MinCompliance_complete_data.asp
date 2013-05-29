<% If (GetMinCompliance.EOF) and (GetMinCompliance.BOF) Then%>
	<span class="formMainBold">No Minimum Compliance Data Available for this Agency</span>
<% else %>	
	<%  section = Request("section") %>
	
	<table width="650" border="1" cellspacing="0" cellpadding="3" bordercolordark="#003063" ID="Table1">
			<tr><td align="center"><form name="frmSelfAssessment" action="SelfAssessment_edit.asp?y=<%= Request("y") %>&section=<%=Request("section")%>" method="post" ID="Form1">
			<!-- <tr> 
					<td colspan="3" align="center" valign="top" class="formSubhead">BBBS -  Agency Minimum Compliance Score Card</td>
			</tr> -->
			<tr>
					<td colspan="3" class="formHeader">Agency Compliance Report</td>
			</tr>
			<tr>
					<td align="left" valign="top" class="formMainBold" width=30%>Agency ID: <%= GetMinCompliance("FK_Agency_ID") %></td>
					<td align="left" valign="top" class="formMainBold" width=50%>Name: <%= GetMinCompliance("AgencyName") %></td>
					<td align="center" valign="top" class="formMainBold">State: <%= GetMinCompliance("AgencyState") %></td>
			</tr>
			<tr>
					<td colspan="3" class="formHeaderMedium" align="center">General</td>
			</tr>
			<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" width=30%>Area</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" width=50%>Description</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0">In Compliance<br>(Yes/No)</td>
			</tr>
			<tr>
					<td align="left" valign="top" class="formMain">CEO</td>
					<td align="left" valign="top" class="formMain">CEO Position Filled</td>
					<td align="center" valign="top" class="formMain"><% if GetMinCompliance("CEO_Position_Open") = "True" then%>No<%elseif GetMinCompliance("CEO_Position_Open") = "False" then%>Yes<%else%><font color="#FF0000"><font color="#FF0000"><font color="#FF0000">Not Entered</font></font></font><%end if%></td>
			</tr>
			
			<tr>
					<td align="left" valign="top" class="formMain">CEO Date</td>
					<td align="left" valign="top" class="formMain">CEO Position Open Since: (red if open more than 90 days)</td>
					<td align="center" valign="top" class="formMain"><% if (isNull(GetMinCompliance("CEO_Position_Open_Date")) = "False") and (DateDiff("d",GetMinCompliance("CEO_Position_Open_Date"),Now)>=90) and (GetMinCompliance("CEO_Position_Open") = "True") then%> <font color="red"> <%= GetMinCompliance("CEO_Position_Open_Date")%></font><%elseif (isNull(GetMinCompliance("CEO_Position_Open_Date")) = "False") and (DateDiff("d",GetMinCompliance("CEO_Position_Open_Date"),Now)<90) and (GetMinCompliance("CEO_Position_Open") = "True") then%><font color="green"><%= GetMinCompliance("CEO_Position_Open_Date")%></font><%else%>N/A<%end if%></td>
					<!-- <td align="center" valign="top" class="formMain"><% if (isNull(GetMinCompliance("CEO_Position_Open_Date")) = "False") and ((DateDiff("m",GetMinCompliance("CEO_Position_Open_Date"),Now)>=3)) then%> <font color="red"> <%= MonthName(Month(GetMinCompliance("CEO_Position_Open_Date")))%>&nbsp<%= Year(GetMinCompliance("CEO_Position_Open_Date"))%></font><%elseif (isNull(GetMinCompliance("CEO_Position_Open_Date")) = "False") and ((DateDiff("m",GetMinCompliance("CEO_Position_Open_Date"),Now)<3)) then%><font color="green"><%= MonthName(Month(GetMinCompliance("CEO_Position_Open_Date")))%></font><%else%>N/A<%end if%></td> -->
			</tr>
			<tr>
					<td align="left" valign="top" class="formMain"><%= (GetMinCompliance("Compliance_Year")) %> Core Data</td>
					<td align="left" valign="top" class="formMain">Core Data Submitted</td>
					<td align="center" valign="top" class="formMain"><% if GetMinCompliance("Core_Matches_LastYear_Submited") = "False" then%>No<%elseif GetMinCompliance("Core_Matches_LastYear_Submited") = "True" then%>Yes<%else%><font color="#FF0000"><font color="#FF0000"><font color="#FF0000">Not Entered</font></font></font><%end if%></td>
			</tr>
			<tr>
					<td colspan="3" class="formHeaderMedium" align="center">Fees and Finances</td>
			</tr>
			<tr>
					<td align="left" valign="top" class="formMain">Fees Form</td>
					<td align="left" valign="top" class="formMain">Fee Calculation Form Submitted</td>
					<td align="center" valign="top" class="formMain"><% if GetMinCompliance("Fee_Calculation_Form_Submited") = "False" then%>No<%elseif GetMinCompliance("Fee_Calculation_Form_Submited") = "True" then%>Yes<%else%><font color="#FF0000"><font color="#FF0000"><font color="#FF0000">Not Entered</font></font></font><%end if%></td>
			</tr>
			<tr>
					<td align="left" valign="top" class="formMain">Arrears < 6 mo</td>
					<td align="left" valign="top" class="formMain">All fees paid or delinquency less than 6 months</td>
					<td align="center" valign="top" class="formMain"><% if GetMinCompliance("Fee_Payments_Current") = "False" then%>No<%elseif GetMinCompliance("Fee_Payments_Current") = "True" then%>Yes<%else%><font color="#FF0000"><font color="#FF0000"><font color="#FF0000">Not Entered</font></font></font><%end if%></td>
			</tr>
			<tr>
					<td align="left" valign="top" class="formMain">Audit Report</td>
					<td align="left" valign="top" class="formMain">Audit Report Submitted</td>
					<td align="center" valign="top" class="formMain"><% if GetMinCompliance("Audit_Report_Submited") = "False" then%>No<%elseif GetMinCompliance("Audit_Report_Submited") = "True" then%>Yes<%else%><font color="#FF0000"><font color="#FF0000"><font color="#FF0000">Not Entered</font></font></font><%end if%></td>
			</tr>
			<tr>
					<td colspan="3" class="formHeaderMedium" align="center">Risk Management</td>
			</tr>
			<tr>
					<td align="left" valign="top" class="formMain">Insurance Certificate</td>
					<td align="left" valign="top" class="formMain">Insurance Certificate Submitted</td>
					<td align="center" valign="top" class="formMain"><% if GetMinCompliance("InsuranceCert_Submited") = "False" then%>No<%elseif GetMinCompliance("InsuranceCert_Submited") = "True" then%>Yes<%else%><font color="#FF0000"><font color="#FF0000"><font color="#FF0000">Not Entered</font></font></font><%end if%></td>
			</tr>
			<!--<tr>
					<td align="left" valign="top" class="formMain">Insurance Certificate</td>
					<td align="left" valign="top" class="formMain">Insurance Certificate Effective Date</td>
					<td align="center" valign="top" class="formMain"><% if isNull(GetMinCompliance("InsuranceCert_Effective_Date")) = "False" then%> <%= GetMinCompliance("InsuranceCert_Effective_Date") %> <%elseif isNull(GetMinCompliance("InsuranceCert_Effective_Date")) = "True" then%>Not Entered<%else%><font color="#FF0000"><font color="#FF0000"><font color="#FF0000">Not Entered</font></font></font><%end if%></td>
			</tr>
			<tr>
					<td align="left" valign="top" class="formMain">Insurance Certificate</td>
					<td align="left" valign="top" class="formMain">Insurance Certificate Expiration Date</td>
					<td align="center" valign="top" class="formMain"><% if isNull(GetMinCompliance("InsuranceCert_Expiration_Date")) = "False" then%> <%= GetMinCompliance("InsuranceCert_Expiration_Date") %> <%elseif isNull(GetMinCompliance("InsuranceCert_Expiration_Date")) = "True" then%>Not Entered<%else%><font color="#FF0000"><font color="#FF0000"><font color="#FF0000">Not Entered</font></font></font><%end if%></td>
			</tr>-->
			<tr>
					<td colspan="3" class="formHeaderMedium" align="center">Surveys</td>
			</tr>
			<tr>
					<td align="left" valign="top" class="formMain">AAI Board Members Survey</td>
					<td align="left" valign="top" class="formMain">AAI Board Members Survey Submitted</td>
					<td align="center" valign="top" class="formMain"><% if GetMinCompliance("Survey_Board_Submited") = "False" then%>No<%elseif GetMinCompliance("Survey_Board_Submited") = "True" then%>Yes<%else%><font color="#FF0000"><font color="#FF0000"><font color="#FF0000">Not Entered</font></font></font><%end if%></td>
			</tr>
			<tr>
					<td align="left" valign="top" class="formMain">AAI Expenses Survey</td>
					<td align="left" valign="top" class="formMain">AAI Expenses Survey Submitted</td>
					<td align="center" valign="top" class="formMain"><% if GetMinCompliance("Survey_Expances_Submited") = "False" then%>No<%elseif GetMinCompliance("Survey_Expances_Submited") = "True" then%>Yes<%else%><font color="#FF0000"><font color="#FF0000"><font color="#FF0000">Not Entered</font></font></font><%end if%></td>
			</tr>
			<tr>
					<td align="left" valign="top" class="formMain">AAI Staff Survey</td>
					<td align="left" valign="top" class="formMain">AAI Staff Survey Submitted</td>
					<td align="center" valign="top" class="formMain"><% if GetMinCompliance("Survey_Staff_Submited") = "False" then%>No<%elseif GetMinCompliance("Survey_Staff_Submited") = "True" then%>Yes<%else%><font color="#FF0000"><font color="#FF0000"><font color="#FF0000">Not Entered</font></font></font><%end if%></td>
			</tr>
			<tr>
					<td align="left" valign="top" class="formMain">AAI Benefits Survey</td>
					<td align="left" valign="top" class="formMain">AAI Benefits Survey Submitted</td>
					<td align="center" valign="top" class="formMain"><% if GetMinCompliance("Survey_Benefits") = "False" then%>No<%elseif GetMinCompliance("Survey_Benefits") = "True" then%>Yes<%else%><font color="#FF0000"><font color="#FF0000"><font color="#FF0000">Not Entered</font></font></font><%end if%></td>
			</tr>
			<tr>
					<td align="left" valign="top" class="formMain">AAI Forecast Survey</td>
					<td align="left" valign="top" class="formMain">AAI Budget Forecast Survey Submitted</td>
					<td align="center" valign="top" class="formMain"><% if GetMinCompliance("Survey_Forecast") = "False" then%>No<%elseif GetMinCompliance("Survey_Forecast") = "True" then%>Yes<%else%><font color="#FF0000"><font color="#FF0000"><font color="#FF0000">Not Entered</font></font></font><%end if%></td>
			</tr>
			<tr>
					<td align="left" valign="top" class="formMain">Self Assessment</td>
					<td align="left" valign="top" class="formMain">Operational Section Submitted</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAss("SelfAssessment_Operational_Completed") = "False" then%>No<%elseif GetSelfAss("SelfAssessment_Operational_Completed") = "True" then%>Yes<%else%><font color="#FF0000"><font color="#FF0000"><font color="#FF0000">Not Entered</font></font></font><%end if%></td>
			</tr>
			<tr>
					<td align="left" valign="top" class="formMain">Self Assessment</td>
					<td align="left" valign="top" class="formMain">Program Section Submitted</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAss("SelfAssessment_Program_Completed") = "False" then%>No<%elseif GetSelfAss("SelfAssessment_Program_Completed") = "True" then%>Yes<%else%><font color="#FF0000"><font color="#FF0000"><font color="#FF0000">Not Entered</font></font></font><%end if%></td>
			</tr>
			

<% Dim ReadOnlyLevel
If Session("ReadOnly") then
	ReadOnlyLevel=1
Else
	ReadOnlyLevel=0
End If
%>		
	</table> 
<% end if %>