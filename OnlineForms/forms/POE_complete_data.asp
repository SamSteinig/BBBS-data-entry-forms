<%
Order=request("Order")%>

<!-- RESULTS TABLE STARTS HERE -->
		<% if printform="No" then %>
			<table border="1" cellpadding="2" cellspacing="0" bordercolordark="003063" width="660">
		<% else %>
			<table border="1" cellpadding="2" cellspacing="0" bordercolordark="003063" width="600">		
		<% end if %>
			<tr>
				<td colspan="12" align="center" class="formHeader">POE Data Entry</td>
			</tr>
			
			<form name="frmPOE" action="POE_edit.asp?row=<%=ct%>&AgencyID=<%=AgencyID%>&Order=<%=Order%>" method="post">
			<!--#include file="../includes/form_stamp.asp"-->
			<input type="hidden" name="status" value="newPOE">
						<tr>
			                <td colspan="12" class="formHeader"><input type="submit" value="Add New POE Records" class="formMainBold"></td>
			       		</tr>
			</form>	
			
			<tr>
				<td align="middle" colspan="3" class="formsubhead">
				<% Select Case Order %>
					<% Case "DateAssessmentDone" %>
							<a class="formmain" href="POE_Complete.asp?Order=MatchID&AgencyID=<%=AgencyID%>">Sort by Match ID</a>
					<% Case "MatchID" %>
							<a class="formmain" href="POE_Complete.asp?Order=DateAssessmentDone&AgencyID=<%=AgencyID%>">Sort by Assessment Date</a>
				<% End Select %>
				</td>			
				<td colspan="9">&nbsp;</td>				
			</tr>
			
			<tr>
				<td colspan="12" align="center" class="formSubhead">
				Code Key for Confidence, Competence, and Caring Questions:<br>
				1 = Much Worse; 2 = A little Worse; 3 = No Change; 4 = A Little Better; 5 = Much Better; 6 = Don't Know; 7 = Not a Problem
				</td>
			</tr>			


					<%
						ct = 1
						Set Con = Server.CreateObject("ADODB.Connection")
						Con.Open "BBBSAforms", "sa","12sist12"
						query = "SELECT * FROM tbl_frmPOE WHERE AgencyID='" & Session("AgencyIDN") & "' ORDER BY " & request("Order")
						Set GetPOE = Con.Execute(query)
					%>
<script language="JavaScript">
<!-- 
				function confirmDelete()
	{
		if (confirm("Are you sure you want to delete this record?"))
		{
			location = "POE_edit.asp?status=deleteRow&row=<%= GetPOE("POEID") %>";
			alert("Record deleted.");
		}
		else
		{
		return false;
		}
	}		
// -->
</script>	










					<%
						If GetPOE.EOF OR GetPOE.BOF Then
					%>
					<tr>
	                <td colspan="11" class="formMainBold">No POE Members To List</td>
    		   		</tr>
					<%
						Else
						GetPOE.MoveFirst
						Do Until GetPOE.EOF
					 %>

					 
					 
					 
			<tr>
				<td class="formMainBold">Match ID</td>
				<td class="formMainBold">Source</td>
				<td class="formMainBold">Program Type</td>
				<td class="formMainBold">Date Assessment Done</td>
				<td class="formMainBold">Match Length (Months)</td>
				<td class="formMainBold">Age</td>
				<td class="formMainBold">Gender</td>
				<td class="formMainBold">Ethnicity</td>
				<td class="formMainBold">Self Confidence</td>
				<td class="formMainBold">Express Feelings</td>				
				<td class="formMainBold">Make Decisions</td>					

				<% if printform="No" then %>
					<td align="right" class="formMain" rowspan="6"><a href="POE_edit.asp?status=editRow&row=<%= GetPOE("POEID") %>&AgencyID=<%=AgencyID%>&Order=<%=Order%>">Edit</a><br><br><a href="#" onclick="confirmDelete();return false;">Delete</a></td>				
				<% end if %>					
				
			</tr>
			
			<tr>

				<td class="formMain" align="center" rowspan="5"><%= GetPOE("MatchID") %></td>
				<% 
				Select Case GetPOE("Source")
					Case 1
						Source = "Volunteer"
					Case 2 
						Source = "Parent"
					Case 3 
						Source = "Teacher"
				End Select
				%>
				<td class="formMain" align="center"><%= Source %></td>

				<%
				Select Case GetPOE("ProgramType")
					Case 1
						ProgramType = "Community"
					Case 2
						ProgramType = "School"
					Case 3 
						ProgramType = "Other Site"
				End Select
				%>
				<td class="formMain" align="center"><%= ProgramType %></td>

				<td class="formMain" align="center"><%= GetPOE("DateAssessmentDone") %></td>
				<td class="formMain" align="center"><%= GetPOE("MatchLength") %></td>
				<td class="formMain" align="center"><%= GetPOE("Age")%></td>
				<td class="formMain" align="center"><%= GetPOE("Gender")%></td>
				<td class="formMain" align="center"><%= GetPOE("Ethnicity") %></td>
				<td class="formMain" align="center"><%= GetPOE("Selfconfidence") %></td>
				<td class="formMain" align="center"><%= GetPOE("ExpressFeelings") %></td>				
				<td class="formMain" align="center"><%= GetPOE("MakeDecisions") %></td>								
				
				
				
				
				
			</tr>	
	
	
	
	
			<tr>
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
				<td class="formMain" align="center"><%= GetPOE("ClassroomBehavior") %></td>	
				<td class="formMain" align="center"><%= GetPOE("InterestsHobbies") %></td>				
				<td class="formMain" align="center"><%= GetPOE("Hygiene") %></td>				
				<td class="formMain" align="center"><%= GetPOE("SenseOfFuture") %></td>				
				<td class="formMain" align="center"><%= GetPOE("CommunityResources") %></td>				
				<td class="formMain" align="center"><%= GetPOE("SchoolResources") %></td>	
				<td class="formMain" align="center"><%= GetPOE("AcademicPerformance") %></td>	
				<td class="formMain" align="center"><%= GetPOE("AttitudeTowardSchool") %></td>	
				<td class="formMain" align="center"><%= GetPOE("SchoolPreparedness") %></td>	
				<td class="formMain" align="center"><%= GetPOE("ClassParticipation") %></td>	
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
				<td class="formMain" align="center"><%= GetPOE("AvoidDelinquency") %></td>			
				<td class="formMain" align="center"><%= GetPOE("AvoidSubstanceAbuse") %></td>				
				<td class="formMain" align="center"><%= GetPOE("AvoidEarlyParenting") %></td>
				<td class="formMain" align="center"><%= GetPOE("ShowsTrust") %></td>
				<td class="formMain" align="center"><%= GetPOE("RespectsOtherCultures") %></td>
				<td class="formMain" align="center"><%= GetPOE("RelationshipWithFamily") %></td>
				<td class="formMain" align="center"><%= GetPOE("RelationshipWithPeers") %></td>				
				<td class="formMain" align="center"><%= GetPOE("RelationshipWithOtherAdults") %></td>				
				<td class="formMain" align="center"><%= GetPOE("SubjectImprovement") %></td>	
				<td class="formMain" align="center"><%= GetPOE("NumberOfSubjects") %></td>							
			
				
					


			</tr>
			<tr>
                <td colspan="12" class="formHeader"><img src="../images/spacer.gif" width="1" height="5" alt="" border="0"></td>
       		</tr>

					<% 
							GetPOE.MoveNext
							ct = ct + 1
						Loop
						GetPOE.Close
						Set GetPOE = Nothing
						Con.Close
						Set Con = Nothing
					End If
					 %>

			<% if printform="No" then %>
								 
			<form name="frmPOE" action="POE_edit.asp?row=<%=ct%>&AgencyID=<%=AgencyID%>&Order=<%=Order%>" method="post">
			<!--#include file="../includes/form_stamp.asp"-->
			<input type="hidden" name="status" value="newPOE">
						<tr>
			                <td colspan="12" class="formHeader"><input type="submit" value="Add New POE Records" class="formMainBold"></td>
			       		</tr>
						
						<tr>
							<td colspan="12"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
						</tr>
			</form>
			
			<% end if %>
					 
		</table>
		