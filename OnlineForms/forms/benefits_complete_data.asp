<% Dim ReadOnlyLevel
	If Session("ReadOnly") then
		ReadOnlyLevel=1
	Else
		ReadOnlyLevel=0
	End If %>


<form name="frmExpenses" action="benefits_edit.asp?y=<%= Request("y") %>" method="post">
<!--#include file="../includes/form_stamp.asp"-->
<input type="hidden" name="status" value="editOld">
<table border="1" cellspacing="0" cellpadding="2" width = "660" bordercolordark="#003063">


 

	<tr>
		<td colspan="6" align="center" class="formSubhead">BBBS - <%= y %> Annual Agency Information (AAI)</td>
	</tr>
	
	<% if printform="No" then %>
	
		<tr>
			<td colspan="6" class="formHeader">BENEFITS</td>
		</tr>
	
	<% else %>
	
		<tr>
			<td colspan="6" class="formIndex">BENEFITS</td>
		</tr>
	
	<% end if %>	
	
	<tr>
		<td colspan="6" class="formMainBold">Created: <%= GetBenefits("CreateDate") %><br>
		<% form = "Benefits" %> 
		<% gid = GetBenefits("BenefitsID") %>
		<!--#include file="../includes/lastmodified_stamp.asp"-->
		</td>
	</tr>


	
	<tr>
	
		<td colspan="6">
			<table width="100%" border="1" bordercolordark="#003063" cellspacing="0" cellpadding="2">

			
		<% if PrintForm = "No" then %>

			<% if ReadOnlyLevel=0 then %>	
				<tr>
					<td colspan="6" class="formHeader"><input type="submit" value="Edit Form" class="formMainBold" ID="Submit1" NAME="Submit1"></td>
				</tr>
			<% else %>
				<tr>
					<td colspan="9" class="formMainCentered">Where did the <strong>Edit Button</strong> go?  <a href="..\helpfiles\surveyhelp.asp?HelpID=password1" onclick="NewWindow(this.href,'name','500','250','yes');return false;">Click Here</a> for an explanation.</td>
				</tr>							
			<% end if %>
			
			<tr>
				<td colspan="5" class="formHeaderSmall">BENEFITS - MEDICAL<BR>% paid by BBBS & total benefit cost for Employee and Employee's Family</td>
			</tr>
			
		<% else %>
		
			<tr>
				<td colspan="5" class="formMainCentered"><strong>BENEFITS - MEDICAL<BR>% paid by BBBS & total benefit cost for Employee and Employee's Family</td></strong>
			</tr>		
			
		<% end if %>	
			
			
				
			<tr>
				<td class="formMain" width="20%">&nbsp;</td>
				<td class="formHeaderSmall" align="center" colspan="2" width="40%">Full Time</td>
				<td class="formHeaderSmall" align="center" colspan="2" width="40%">Part Time</td>				
			</tr>
			<tr>
				<td class="formMain" align="center" cellpadding="1">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" ID="Table9">
					<tr>
						<td class="formMain" valign="top" align="center">MEDICAL<br>offered?</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="Center"><% if GetBenefits("BenMedOffered") then%>Yes<%else%>No<%end if%>
						<!--<select name="forms" size=1 class="formMain" ID="Select1">
							<option value="selected" ><% if GetBenefits("BenMedOffered") then%>Yes<%else%>No<%end if%></option>
						</select>&nbsp;-->
						</td>
					</tr>
					</table>
				</td>
				<td>
					<table width="100%" border="0" cellspacing="1" cellpadding="1" ID="Table1">
					<tr>
						<td class="formMain" valign="top" align="center" colspan="2"><Strong>Per Employee</Strong></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
				    <tr>
						<td class="formMain" valign="top" align="left"><strong><% If isNull(GetBenefits("BenMedFullEmployeeAmount")) = false Then %><%= GetBenefits("BenMedFullEmployeeAmount") %><% Else %>N/A<% End If %><br></td>
						<td class="formMain" valign="top" align="left">Average total monthly premium<br>(agency + employee contribution)</td>
					</tr>
					<tr>
						<td class="formMain" valign="top" align="left"><strong><% If isNull(GetBenefits("BenMedFullEmployee")) = false Then %><%= GetBenefits("BenMedFullEmployee")%><% Else %>N/A<% End If %></td>
						<td class="formMain" valign="top" align="left">% of premium <br> paid by agency</td>
					</tr>
					
					</table>
				</td>
				
				<td>
					<table width="100%" border="0" cellspacing="1" cellpadding="1" ID="Table2">
					<tr>
						<td class="formMain" valign="top" align="center" colspan="2"><Strong>Per Family</Strong></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><strong><% If isNull(GetBenefits("BenMedFullFamilyAmount")) = false Then %><%= GetBenefits("BenMedFullFamilyAmount") %><% Else %>N/A<% End If %><br></td>
						<td class="formMain" valign="top" align="left">Average total monthly premium<br>(agency + employee contribution)</td>
					</tr>
					<tr>
						<td class="formMain" valign="top" align="left"><strong><% If isNull(GetBenefits("BenMedFullFamily")) = false Then %><%= GetBenefits("BenMedFullFamily") %><% Else %>N/A<% End If %></td>
						<td class="formMain" valign="top" align="left">% of premium <br> paid by agency</td>
					</tr>
				
					</table>
				</td>
				
				<td>
					<table width="100%" border="0" cellspacing="1" cellpadding="1" ID="Table3">
					<tr>
						<td class="formMain" valign="top" align="center" colspan="2"><Strong>Per Employee<Strong></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><strong><% If isNull(GetBenefits("BenMedPartEmployeeAmount")) = false Then %><%= GetBenefits("BenMedPartEmployeeAmount") %><% Else %>N/A<% End If %><br></td>
						<td class="formMain" valign="top" align="left">Average total monthly premium<br>(agency + employee contribution)</td>
					</tr>
					<tr>
						<td class="formMain" valign="top" align="left"><strong><% If isNull(GetBenefits("BenMedPartEmployee")) = false Then %><%= GetBenefits("BenMedPartEmployee") %><% Else %>N/A<% End If %></td>
						<td class="formMain" valign="top" align="left">% of premium <br> paid by agency</td>
					</tr>
				
					</table>
				</td>		
				
				<td>
					<table width="100%" border="0" cellspacing="1" cellpadding="1" ID="Table4">
					<tr>
						<td class="formMain" valign="top" align="center" colspan="2"><Strong>Per Family</Strong></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><strong><% If isNull(GetBenefits("BenMedPartFamilyAmount")) = false Then %><%= GetBenefits("BenMedPartFamilyAmount") %><% Else %>N/A<% End If %><br></td>
						<td class="formMain" valign="top" align="left">Average total monthly premium<br>(agency + employee contribution)</td>
					</tr>
					<tr>
						<td class="formMain" valign="top" align="left"><strong><% If isNull(GetBenefits("BenMedPartFamily")) = false Then %><%= GetBenefits("BenMedPartFamily") %><% Else %>N/A<% End If %></td>
						<td class="formMain" valign="top" align="left">% of premium <br> paid by agency</td>
					</tr>
					
					</table>
				</td>						
			
			</tr>
			
			<tr>
				<td class="formMain" align="center" cellpadding="1">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" ID="Table5">
					<tr>
						<td class="formMain" valign="top" align="center">DENTAL<br>offered?</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="Center"><% if GetBenefits("BenDentOffered") then%>Yes<%else%>No<%end if%>
						<!--<select name="forms" size=1 class="formMain" ID="Select2">
							<option value="selected"><% if GetBenefits("BenDentOffered") then%>Yes<%else%>No<%end if%></option>
						</select>&nbsp;-->
						</td>
					</tr>
					</table>
				</td>
				<td>
					<table width="100%" border="0" cellspacing="1" cellpadding="1" ID="Table6">
					<tr>
						<td class="formMain" valign="top" align="center" colspan="2"><strong>Per Employee</strong></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><strong><% If isNull(GetBenefits("BenDentFullEmployeeAmount")) = false Then %><%= GetBenefits("BenDentFullEmployeeAmount") %><% Else %>N/A<% End If %><br></td>
						<td class="formMain" valign="top" align="left">Average total monthly premium<br>(agency + employee contribution)</td>
					</tr>
					<tr>
						<td class="formMain" valign="top" align="left"><strong><% If isNull(GetBenefits("BenDentFullEmployee")) = false Then %><%= GetBenefits("BenDentFullEmployee")%><% Else %>N/A<% End If %></td>
						<td class="formMain" valign="top" align="left">% of premium <br> paid by agency</td>
					</tr>
				
					</table>
				</td>
				
				<td>
					<table width="100%" border="0" cellspacing="1" cellpadding="1" ID="Table7">
					<tr>
						<td class="formMain" valign="top" align="center" colspan="2"><strong>Per Family</strong></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><strong><% If isNull(GetBenefits("BenDentFullFamilyAmount")) = false Then %><%= GetBenefits("BenDentFullFamilyAmount") %><% Else %>N/A<% End If %><br></td>
						<td class="formMain" valign="top" align="left">Average total monthly premium<br>(agency + employee contribution)</td>
					</tr>
					<tr>
						<td class="formMain" valign="top" align="left"><strong><% If isNull(GetBenefits("BenDentFullFamily")) = false Then %><%= GetBenefits("BenDentFullFamily") %><% Else %>N/A<% End If %></td>
						<td class="formMain" valign="top" align="left">% of premium <br> paid by agency</td>
					</tr>
					
					</table>
				</td>
				
				<td>
					<table width="100%" border="0" cellspacing="1" cellpadding="1" ID="Table8">
					<tr>
						<td class="formMain" valign="top" align="center" colspan="2"><strong>Per Employee</strong></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><strong><% If isNull(GetBenefits("BenDentPartEmployeeAmount")) = false Then %><%= GetBenefits("BenDentPartEmployeeAmount") %><% Else %>N/A<% End If %><br></td>
						<td class="formMain" valign="top" align="left">Average total monthly premium<br>(agency + employee contribution)</td>
					</tr>
					<tr>
						<td class="formMain" valign="top" align="left"><strong><% If isNull(GetBenefits("BenDentPartEmployee")) = false Then %><%= GetBenefits("BenDentPartEmployee") %><% Else %>N/A<% End If %></td>
						<td class="formMain" valign="top" align="left">% of premium <br> paid by agency</td>
					</tr>
					
					</table>
				</td>		
				
				<td>
					<table width="100%" border="0" cellspacing="1" cellpadding="1" ID="Table10">
					<tr>
						<td class="formMain" valign="top" align="center" colspan="2"><strong>Per Family</strong></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><strong><% If isNull(GetBenefits("BenDentPartFamilyAmount")) = false Then %><%= GetBenefits("BenDentPartFamilyAmount") %><% Else %>N/A<% End If %><br></td>
						<td class="formMain" valign="top" align="left">Average total monthly premium<br>(agency + employee contribution)</td>
					</tr>
					<tr>
						<td class="formMain" valign="top" align="left"><strong><% If isNull(GetBenefits("BenDentPartFamily")) = false Then %><%= GetBenefits("BenDentPartFamily") %><% Else %>N/A<% End If %></td>
						<td class="formMain" valign="top" align="left">% of premium <br> paid by agency</td>
					</tr>
					
					</table>
				</td>						
			
			</tr>
			
			<tr>
				<td class="formMain" align="center" cellpadding="1">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" ID="Table11">
					<tr>
						<td class="formMain" valign="top" align="center">VISION<br>offered?</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="Center"><% if GetBenefits("BenVisOffered") then%>Yes<%else%>No<%end if%>
						<!--<select name="forms" size=1 class="formMain" ID="Select3">
							<option value="selected"><% if GetBenefits("BenVisOffered") then%>Yes<%else%>No<%end if%></option>
						</select>&nbsp;-->
						</td>
					</tr>
					</table>
				</td>
				<td>
					<table width="100%" border="0" cellspacing="1" cellpadding="1" ID="Table12">
					<tr>
						<td class="formMain" valign="top" align="center" colspan="2"><strong>Per Employee</strong></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><strong><% If isNull(GetBenefits("BenVisFullEmployeeAmount")) = false Then %><%= GetBenefits("BenVisFullEmployeeAmount") %><% Else %>N/A<% End If %><br></td>
						<td class="formMain" valign="top" align="left">Average total monthly premium<br>(agency + employee contribution)</td>
					</tr>
					<tr>
						<td class="formMain" valign="top" align="left"><strong><% If isNull(GetBenefits("BenVisFullEmployee")) = false Then %><%= GetBenefits("BenVisFullEmployee")%><% Else %>N/A<% End If %></td>
						<td class="formMain" valign="top" align="left">% of premium <br> paid by agency</td>
					</tr>
					
					</table>
				</td>
				
				<td>
					<table width="100%" border="0" cellspacing="1" cellpadding="1" ID="Table13">
					<tr>
						<td class="formMain" valign="top" align="center" colspan="2"><strong>Per Family</strong></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><strong><% If isNull(GetBenefits("BenVisFullFamilyAmount")) = false Then %><%= GetBenefits("BenVisFullFamilyAmount") %><% Else %>N/A<% End If %><br></td>
						<td class="formMain" valign="top" align="left">Average total monthly premium<br>(agency + employee contribution)</td>
					</tr>
					<tr>
						<td class="formMain" valign="top" align="left"><strong><% If isNull(GetBenefits("BenVisFullFamily")) = false Then %><%= GetBenefits("BenVisFullFamily") %><% Else %>N/A<% End If %></td>
						<td class="formMain" valign="top" align="left">% of premium <br> paid by agency</td>
					</tr>
				
					</table>
				</td>
				
				<td>
					<table width="100%" border="0" cellspacing="1" cellpadding="1" ID="Table14">
					<tr>
						<td class="formMain" valign="top" align="center" colspan="2"><strong>Per Employee</strong></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><strong><% If isNull(GetBenefits("BenVisPartEmployeeAmount")) = false Then %><%= GetBenefits("BenVisPartEmployeeAmount") %><% Else %>N/A<% End If %><br></td>
						<td class="formMain" valign="top" align="left">Average Total monthly premium<br>(agency + employee contribution)</td>
					</tr>
					<tr>
						<td class="formMain" valign="top" align="left"><strong><% If isNull(GetBenefits("BenVisPartEmployee")) = false Then %><%= GetBenefits("BenVisPartEmployee") %><% Else %>N/A<% End If %></td>
						<td class="formMain" valign="top" align="left">% of premium <br> paid by agency</td>
					</tr>
				
					</table>
				</td>		
				
				<td>
					<table width="100%" border="0" cellspacing="1" cellpadding="1" ID="Table15">
					<tr>
						<td class="formMain" valign="top" align="center" colspan="2"><strong>Per Family</strong></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><strong><% If isNull(GetBenefits("BenVisPartFamilyAmount")) = false Then %><%= GetBenefits("BenVisPartFamilyAmount") %><% Else %>N/A<% End If %><br></td>
						<td class="formMain" valign="top" align="left">Average total monthly premium<br>(agency + employee contribution)</td>
					</tr>
					<tr>
						<td class="formMain" valign="top" align="left"><strong><% If isNull(GetBenefits("BenVisPartFamily")) = false Then %><%= GetBenefits("BenVisPartFamily") %><% Else %>N/A<% End If %></td>
						<td class="formMain" valign="top" align="left">% of premium <br> paid by agency</td>
					</tr>
				
					</table>
				</td>						
			
			</tr>
			
			
			<% if PrintForm="No" then %>
				
				<tr>
					<td colspan="6" class="formHeaderSmall">BENEFITS - NON-MEDICAL<br>(check all that apply)</td>
				</tr>	
				
			<% else %>
			
				<tr>
					<td colspan="6" class="formMainCentered"><strong>BENEFITS - NON-MEDICAL<br>(check all that apply)</strong></td>
				</tr>				
				
			<% end if %>				
			<tr>
			<td colspan=6>
				<table width="100%" border="1" bordercolordark="#003063" cellspacing="0" cellpadding="1" ID="Table17">
				<tr>
					<!--<td class="formMain" colspan="3">&nbsp;</td>-->
					<td class="formMain" width="40%">&nbsp;</td>
					<td class="formMain" width="30%">
						<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table16">
						<tr>
							
							<td class="formMain" valign="top" align="center" colspan="4"><strong>Full Time</strong></td>
						</tr>
						<tr><td>&nbsp;</td></tr>
						<tr>
							<td class="formMain" valign="bottom" align="center">Offered</td>
							<!--<td class="formMain" valign="bottom" align="left">Paid</td>-->
							<td class="formMain" valign="bottom" align="left">&nbsp;</td>
							<td class="formMain" valign="bottom" align="left">Total<br>Mthly<br>Prem</td>
							<td class="formMain" valign="bottom" align="left">% of<br>Prem<br>Paid by<br>Ag</td>
							
						</tr>
						</table>
					</td>
					
					<!--<td class="formMain" colspan="3">&nbsp;</td>
					<td class="formMain" align="center">Full Time</td>-->
					<td class="formMain" align="center" width="30%">
						<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table18">
						<tr>
							
							<td class="formMain" valign="top" align="center" colspan="4"><strong>Part Time</strong></td>
						</tr>
						<tr><td>&nbsp;</td></tr>
						<tr>
							<td class="formMain" valign="bottom" align="center">Offered</td>
							<!--<td class="formMain" valign="bottom" align="left">Paid</td>-->
							<td class="formMain" valign="bottom" align="left">&nbsp;</td>
							<td class="formMain" valign="bottom" align="left">Total<br>Mthly<br>Prem</td>
							<td class="formMain" valign="bottom" align="left">% of<br>Prem<br>Paid by<br>Ag</td>
							
						</tr>
						</table>
					</td>

				</tr>
				<tr>
					<td class="formMain" width="40%">Disability Insurance SHORT Term per employee</td>
					<td class="formMain" align="center">
						<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table20">
						<tr>
							<!--<td class=formMain width="25%" align=center><input type="checkbox" readonly class="formMain" name="frmExpensesDisInsShortTermFull" value="1" <% if GetBenefits("DisInsShortTermFull")=true then%>checked="true"<% end if %> onclick="return false"</td>-->
							<td class=formMain width="25%" align=center><% if GetBenefits("DisInsShortTermFull")=true then%>Yes<%else%>No<% end if %></td>
							<!--<td class=formMain width="25%" align=center><input type="checkbox" readonly class="formMain" name="frmExpensesDisInsShortTermFullPaid" value="1" <% if GetBenefits("DisInsShortTermFullPaid")=true then%>checked="true"<% end if %> onclick="return false"</td>-->
							<td class=formMain width="25%" align=center>$<input type="text" readonly size="1" maxlength="3" value="<% If isNull(GetBenefits("DisInsShortTermFullAmount")) = false Then %><%= GetBenefits("DisInsShortTermFullAmount")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesDisInsShortTermFullAmount"></td>
							<td class=formMain width="25%" align=center><input type="text" readonly size="1" maxlength="3" value="<% If isNull(GetBenefits("DisInsShortTermFullPrcnt")) = false Then %><%= GetBenefits("DisInsShortTermFullPrcnt")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesDisInsShortTermFullPrcnt">%</td>
							
						</tr>
						</table>
					</td>
					<td class="formMain" align="center">
						<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table21">
						<tr>
							<!--<td class=formMain width="25%" align="center"><input type="checkbox" readonly class="formMain" name="frmExpensesDisInsShortTermPart" value="1"  <% if GetBenefits("DisInsShortTermPart")=true then%>checked="true"<% end if %> onclick="return false"</td>-->
							<td class=formMain width="25%" align=center><% if GetBenefits("DisInsShortTermPart")=true then%>Yes<%else%>No<% end if %></td>
							<!--<td class=formMain width="25%" align="center"><input type="checkbox" readonly class="formMain" name="frmExpensesDisInsShortTermPartPaid" value="1"  <% if GetBenefits("DisInsShortTermPartPaid")=true then%>checked="true"<% end if %> onclick="return false"</td>-->
							<td class=formMain width="25%" align="center">$<input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("DisInsShortTermPartAmount")) = false Then %><%= GetBenefits("DisInsShortTermPartAmount")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesDisInsShortTermPartAmount"></td>
							<td class=formMain width="25%" align="center"><input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("DisInsShortTermPartPrcnt")) = false Then %><%= GetBenefits("DisInsShortTermPartPrcnt")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesDisInsShortTermPartPrct">%</td>
							
						</tr>
						</table>
					</td>
				</tr>
				
				<tr>
				<td class="formMain" width="40%">Disability Insurance LONG Term per employee</td>				
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table22">
					<tr>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" readonly class="formMain" name="frmExpensesDisInsLongTermFull" value="1" <% if GetBenefits("DisInsLongTermFull")=true then%>checked="true"<% end if %> onclick="return false"</td>
						<td class=formMain width="25%" align=center><input type="checkbox" readonly class="formMain" name="frmExpensesDisInsLongTermFullPaid" value="1" <% if GetBenefits("DisInsLongTermFullPaid")=true then%>checked="true"<% end if %> onclick="return false"</td>-->
						<td class=formMain width="25%" align=center><% if GetBenefits("DisInsLongTermFull")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain width="25%" align=center>$<input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("DisInsLongTermFullAmount")) = false Then%><%= GetBenefits("DisInsLongTermFullAmount")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesDisInsLongTermFullAmount"></td>
						<td class=formMain width="25%" align=center><input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("DisInsLongTermFullPrcnt")) = false Then%><%= GetBenefits("DisInsLongTermFullPrcnt")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesDisInsLongTermFullPrcnt">%</td>
						
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table23">
					<tr>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" readonly class="formMain" name="frmExpensesDisInsLongTermPart" value="1" <% if GetBenefits("DisInsLongTermPart")=true then%>checked="true"<% end if %> onclick="return false"</td>
						<td class=formMain width="25%" align=center><input type="checkbox" readonly class="formMain" name="frmExpensesDisInsLongTermPartPaid" value="1" <% if GetBenefits("DisInsLongTermPartPaid")=true then%>checked="true"<% end if %> onclick="return false"</td>-->
						<td class=formMain width="25%" align=center><% if GetBenefits("DisInsLongTermPart")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain width="25%" align=center>$<input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("DisInsLongTermPartAmount")) = false Then %><%= GetBenefits("DisInsLongTermPartAmount")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesDisInsLongTermPartAmount"></td>
						<td class=formMain width="25%" align=center><input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("DisInsLongTermPartPrcnt")) = false Then %><%= GetBenefits("DisInsLongTermPartPrcnt")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesDisInsLongTermPartPrcnt">%</td>
						
					</tr>
					</table>
				</td>
				</tr>
				
				<tr>
				<td class="formMain" width="40%">EAP: Employee Assistance Programs per employee</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table24">
					<tr>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" readonly class="formMain" name="frmExpensesEAPFull" value="1" <% if GetBenefits("EAPFull")=true then%>checked="true"<% end if %> onclick="return false"</td>		
						<td class=formMain width="25%" align=center><input type="checkbox" readonly class="formMain" name="frmExpensesEAPFullPaid" value="1" <% if GetBenefits("EAPFullPaid")=true then%>checked="true"<% end if %> onclick="return false"</td>-->
						<td class=formMain width="25%" align=center><% if GetBenefits("EAPFull")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain width="25%" align=center>$<input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("EAPFullAmount")) = false Then %><%= GetBenefits("EAPFullAmount")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesEAPFullAmount"></td>
						<td class=formMain width="25%" align=center><input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("EAPFullPrcnt")) = false Then %><%= GetBenefits("EAPFullPrcnt")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesEAPFullPrcnt">%</td>
						
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table25">
					<tr>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" readonly class="formMain" name="frmExpensesEAPPart" value="1" <% if GetBenefits("EAPPart")=true then%>checked="true"<% end if %> onclick="return false"</td>
						<td class=formMain width="25%" align=center><input type="checkbox" readonly class="formMain" name="frmExpensesEAPPartPaid" value="1" <% if GetBenefits("EAPPart")=true then%>checked="true"<% end if %> onclick="return false"</td>-->
						<td class=formMain width="25%" align=center><% if GetBenefits("EAPPart")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain width="25%" align=center>$<input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("EAPPartAmount")) = false Then %><%= GetBenefits("EAPPartAmount")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesEAPPartAmount"></td>
						<td class=formMain width="25%" align=center><input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("EAPPartPrcnt")) = false Then %><%= GetBenefits("EAPPartPrcnt")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesEAPPartPrcnt">%</td>
						
					</tr>
					</table>
				</td>
				</tr>
				
				<!--Commented due to change in requirements (not need to collect this)
				<tr>
				<td class="formMain" width="40%">"Flex" Pre-Tax Plan (medical, dependent)</td>			
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table26">
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesFlexFull" value="1" <% if GetBenefits("FlexFull")=true then%>checked="true"<% end if %></td>				
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesFlexFullPaid" value="1" <% if GetBenefits("FlexFullPaid")=true then%>checked="true"<% end if %></td>
						<td class=formMain width="25%" align=center><input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("FlexFullPrcnt")) = false Then %><%= GetBenefits("FlexFullPrcnt")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesFlexFullPrcnt">%</td>
						<td class=formMain width="25%" align=center>$<input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("FlexFullAmount")) = false Then %><%= GetBenefits("FlexFullAmount")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesFlexFullAmount"></td>
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table27">
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesFlexPart" value="1" <% if GetBenefits("FlexPart")=true then%>checked="true"<% end if %></td>				
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesFlexPartPaid" value="1" <% if GetBenefits("FlexPartPaid")=true then%>checked="true"<% end if %></td>
						<td class=formMain width="25%" align=center><input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("FlexPartPrcnt")) = false Then %><%= GetBenefits("FlexPartPrcnt")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesFlexPartPrcnt">%</td>
						<td class=formMain width="25%" align=center>$<input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("FlexPartAmount")) = false Then %><%= GetBenefits("FlexPartAmount")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesFlexPartAmount"></td>
					</tr>
					</table>
				</td>
				</tr>-->
				
				<tr>
				<td class="formMain" width="40%">Health Club per employee</td>			
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table28">
					<tr>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesHealthClubFull" value="1" <% if GetBenefits("HealthClubFull")=true then%>checked="true"<% end if %> onclick="return false"</td>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesHealthClubFullPaid" value="1" <% if GetBenefits("HealthClubFullPaid")=true then%>checked="true"<% end if %> onclick="return false"</td>-->
						<td class=formMain width="25%" align=center><% if GetBenefits("HealthClubFull")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain width="25%" align=center>$<input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("HealthClubFullAmount")) = false Then %><%= GetBenefits("HealthClubFullAmount")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesHealthClubFullAmount"></td>
						<td class=formMain width="25%" align=center><input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("HealthClubFullPrcnt")) = false Then %><%= GetBenefits("HealthClubFullPrcnt")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesHealthClubFullPrcnt">%</td>
						
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table29">
					<tr>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesHealthClubPart" value="1" <% if GetBenefits("HealthClubPart")=true then%>checked="true"<% end if %> onclick="return false"</td>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesHealthClubPartPaid" value="1" <% if GetBenefits("HealthClubPartPaid")=true then%>checked="true"<% end if %> onclick="return false"</td>-->
						<td class=formMain width="25%" align=center><% if GetBenefits("HealthClubPart")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain width="25%" align=center>$<input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("HealthClubPartAmount")) = false Then %><%= GetBenefits("HealthClubPartAmount")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesHealthClubPartAmount"></td>
						<td class=formMain width="25%" align=center><input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("HealthClubPartPrcnt")) = false Then %><%= GetBenefits("HealthClubPartPrcnt")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesHealthClubPartPrcnt">%</td>
						
					</tr>
					</table>
				</td>
				</tr>
				
				<tr>
				<td class="formMain" width="40%">Life Insurance per employee</td>			
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table30">
					<tr>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesLifeInsuranceFull" value="1" <% if GetBenefits("LifeInsuranceFull")=true then%>checked="true"<% end if %> onclick="return false"</td>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesLifeInsuranceFullPaid" value="1" <% if GetBenefits("LifeInsuranceFullPaid")=true then%>checked="true"<% end if %> onclick="return false"</td>-->
						<td class=formMain width="25%" align=center><% if GetBenefits("LifeInsuranceFull")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain width="25%" align=center>$<input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("LifeInsuranceFullAmount")) = false Then %><%= GetBenefits("LifeInsuranceFullAmount")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesLifeInsuranceFullAmount"></td>
						<td class=formMain width="25%" align=center><input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("LifeInsuranceFullPrcnt")) = false Then %><%= GetBenefits("LifeInsuranceFullPrcnt")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesLifeInsuranceFullPrcnt">%</td>
						
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table31">
					<tr>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesLifeInsurancePart" value="1" <% if GetBenefits("LifeInsurancePart")=true then%>checked="true"<% end if %> onclick="return false"</td>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesLifeInsurancePartPaid" value="1" <% if GetBenefits("LifeInsurancePartPaid")=true then%>checked="true"<% end if %> onclick="return false"</td>-->
						<td class=formMain width="25%" align=center><% if GetBenefits("LifeInsurancePart")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain width="25%" align=center>$<input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("LifeInsurancePartAmount")) = false Then %><%= GetBenefits("LifeInsurancePartAmount")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesLifeInsurancePartAmount"></td>
						<td class=formMain width="25%" align=center><input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("LifeInsurancePartPrcnt")) = false Then %><%= GetBenefits("LifeInsurancePartPrcnt")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesLifeInsurancePartPrcnt">%</td>
						
					</tr>
					</table>
				</td>
				</tr>
			</table>
				
			<table width="100%" border="1" bordercolordark="#003063" cellspacing="0" cellpadding="1" ID="Table19">
			<!--<tr>
				<td class="formMain" width="40%">Paid Time Off (Vacation, <br>Floating Holidays, Personal Days)</td>			
				<td class="formMain" align="center" width="30%">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table32">
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesTimeOffFull" value="1" <% if GetBenefits("TimeOffFull")=true then%>checked="true"<% end if %> onclick="return false"</td>
						<td class=formMain width="75%" align=center colspan="3"><input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("TimeOffFullDays")) = false Then %><%= GetBenefits("TimeOffFullDays")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesTimeOffFullDays"> days</td>
					</tr>
					</table>
				</td>			
				<td class="formMain" align="center" width="30%">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table33">
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesTimeOffPart" value="1" <% if GetBenefits("TimeOffPart")=true then%>checked="true"<% end if %> onclick="return false"</td>
						<td class=formMain width="75%" align=center colspan="3"><input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("TimeOffPartDays")) = false Then %><%= GetBenefits("TimeOffPartDays")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesTimeOffPartDays"> days</td>
					</tr>
					</table>
				</td>
			</tr>-->
			<tr>
				<td class="formMain" width="40%">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table26">
					<tr>
						<td class=formMain width="50%" rowspan=2" valign="top">Paid Time Off (Vacation, Floating Holidays, Personal Days)</td>
						<td class=formMain width="50%" align=left valign="bottom" height="80">Exempt (Salaried)</td>
					</tr>
					<tr>
						<td class=formMain width="50%" align=left valign="bottom" height="20">Non-Exempt (Hourly)</td>
					</tr>
					</table>
				<td class="formMain" align="center" width="30%">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table32">
					<tr>
						<!--<td class=formMain width="25%" align=center valign=bottom><input type="checkbox" class="formMain" name="frmBenefitsTimeOffFull" value="1" <% if GetBenefits("TimeOffFull")=true then%>checked="true"<% end if %> onclick="return false"</td>-->
						<td class=formMain width="25%" align=center valign=bottom>Offered<br><br><% if GetBenefits("TimeOffFull")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain width="25%" align=center>&nbsp;&nbsp;# of days for new employee<br><input type="text" readonly size="1" maxlength="3" value="<% If isNull(GetBenefits("TimeOffFullDays")) = false Then %><%= GetBenefits("TimeOffFullDays")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffFullDays"></td>
						<td class=formMain width="25%" align=center># of years before increase<br><input type="text" readonly size="1" maxlength="3" value="<% If isNull(GetBenefits("TimeOffFullYears")) = false Then %><%= GetBenefits("TimeOffFullYears")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffFullYears"></td>
						<td class=formMain width="25%" align=center>total days after increase<br><input type="text" readonly size="1" maxlength="3" value="<% If isNull(GetBenefits("TimeOffFullDaysIncreased")) = false Then %><%= GetBenefits("TimeOffFullDaysIncreased")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffFullDaysIncreased"></td>
					</tr>
					<tr>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsTimeOffFullNExempt" value="1" <% if GetBenefits("TimeOffFullNExempt")=true then%>checked="true"<% end if %> onclick="return false"</td>-->
						<td class=formMain width="25%" align=center valign=bottom><% if GetBenefits("TimeOffFullNExempt")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain width="25%" align=center><input type="text" readonly size="1" maxlength="3" value="<% If isNull(GetBenefits("TimeOffFullDaysNExempt")) = false Then %><%= GetBenefits("TimeOffFullDaysNExempt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffFullDaysNExempt"></td>
						<td class=formMain width="25%" align=center><input type="text" readonly size="1" maxlength="3" value="<% If isNull(GetBenefits("TimeOffFullYearsNExempt")) = false Then %><%= GetBenefits("TimeOffFullYearsNExempt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffFullYearsNExempt"></td>
						<td class=formMain width="25%" align=center><input type="text" readonly size="1" maxlength="3" value="<% If isNull(GetBenefits("TimeOffFullDaysIncreasedNExempt")) = false Then %><%= GetBenefits("TimeOffFullDaysIncreasedNExempt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffFullDaysIncreasedNExempt"></td>
					</tr>
					</table>
				</td>
				<td class="formMain" align="center" width="30%">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table27">
					<tr>
						<!--<td class=formMain width="25%" align=center valign=bottom><input type="checkbox" class="formMain" name="frmBenefitsTimeOffPart" value="1" <% if GetBenefits("TimeOffPart")=true then%>checked="true"<% end if %> onclick="return false"</td>-->
						<td class=formMain width="25%" align=center valign=bottom>Offered<br><br><% if GetBenefits("TimeOffPart")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain width="25%" align=center>&nbsp;&nbsp;# of days for new employee<br><input type="text" readonly size="1" maxlength="3" value="<% If isNull(GetBenefits("TimeOffPartDays")) = false Then %><%= GetBenefits("TimeOffPartDays")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffPartDays"></td>
						<td class=formMain width="25%" align=center># of years before increase<br><input type="text" size="1" readonly maxlength="3" value="<% If isNull(GetBenefits("TimeOffPartYears")) = false Then %><%= GetBenefits("TimeOffPartYears")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffPartYears"></td>
						<td class=formMain width="25%" align=center>total days after increase<br><input type="text" size="1" readonly maxlength="3" value="<% If isNull(GetBenefits("TimeOffPartDaysIncreased")) = false Then %><%= GetBenefits("TimeOffPartDaysIncreased")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffPartDaysIncreased"></td>
					</tr>
					<tr>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsTimeOffPartNExempt" value="1" <% if GetBenefits("TimeOffPartNExempt")=true then%>checked="true"<% end if %> onclick="return false"</td>-->
						<td class=formMain width="25%" align=center valign=bottom><% if GetBenefits("TimeOffPartNExempt")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain width="25%" align=center><input type="text" readonly size="1" maxlength="3" value="<% If isNull(GetBenefits("TimeOffPartDaysNExempt")) = false Then %><%= GetBenefits("TimeOffPartDaysNExempt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffPartDaysNExempt"></td>
						<td class=formMain width="25%" align=center><input type="text" readonly size="1" maxlength="3" value="<% If isNull(GetBenefits("TimeOffPartYearsNExempt")) = false Then %><%= GetBenefits("TimeOffPartYearsNExempt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffPartYearsNExempt"></td>
						<td class=formMain width="25%" align=center><input type="text" readonly size="1" maxlength="3" value="<% If isNull(GetBenefits("TimeOffPartDaysIncreasedNExempt")) = false Then %><%= GetBenefits("TimeOffPartDaysIncreasedNExempt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffPartDaysIncreasedNExempt"></td>
					</tr>
					</table>
				</td>
			</tr>
				
			<tr>
				<td class="formMain" width="40%">Paid Time Off (Sick Time) per employee</td>			
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table34">
					<tr>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesTimeSickOffFull" value="1" <% if GetBenefits("TimeOffSickFull")=true then%>checked="true"<% end if %> onclick="return false"</td>-->
						<td class=formMain width="25%" align=center><% if GetBenefits("TimeOffSickFull")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain width="75%" align=center colspan="3"><input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("TimeOffSickFullDays")) = false Then %><%= GetBenefits("TimeOffSickFullDays")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesTimeOffSickFullDays"> days</td>
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table35">
					<tr>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesTimeSickOffPart" value="1" <% if GetBenefits("TimeOffSickPart")=true then%>checked="true"<% end if %> onclick="return false"</td>-->
						<td class=formMain width="25%" align=center><% if GetBenefits("TimeOffSickPart")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain width="75%" align=center colspan="3"><input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("TimeOffSickPartDays")) = false Then %><%= GetBenefits("TimeOffSickPartDays")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesTimeOffSickPartDays"> days</td>
					</tr>
					</table>
				</td>
			</tr>													

<!-- Commented due changes from edit by Cindy and Jeff on Sept 6
			<tr>
				<td class="formMain" width="40%">Paid Time Off (Vacation)</td>			
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table36">
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesTimeVacOffFull" value="1" <% if GetBenefits("TimeOffVacFull")=true then%>checked="true"<% end if %></td>
						<td class=formMain width="75%" align=center colspan="3"><input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("TimeOffVacFullDays")) = false Then %><%= GetBenefits("TimeOffVacFullDays")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesTimeOffVacFullDays"> days</td>
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table37">
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesTimeVacOffPart" value="1" <% if GetBenefits("TimeOffVacPart")=true then%>checked="true"<% end if %></td>
						<td class=formMain width="75%" align=center colspan="3"><input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("TimeOffVacPartDays")) = false Then %><%= GetBenefits("TimeOffVacPartDays")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesTimeOffVacPartDays"> days</td>
					</tr>
					</table>
				</td>
			</tr>
-->

			<tr>
				<td class="formMain" width="40%">Professional Dues, Conferences, etc. per employee</td>			
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table38">
					<tr>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesProfDuesFull" value="1" <% if GetBenefits("ProfDuesFull")=true then%>checked="true"<% end if %> onclick="return false"</td>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesProfDuesFullPaid" value="1" <% if GetBenefits("ProfDuesFullPaid")=true then%>checked="true"<% end if %> onclick="return false"</td>-->
						<td class=formMain width="25%" align=center><% if GetBenefits("ProfDuesFull")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain width="25%" align=center>&nbsp;</td>
						<td class=formMain width="50%" align=center colspan="2">avg $ amount per employee<br><input type="text" readonly size="2" maxlength="5" value=" <% If isNull(GetBenefits("ProfDuesFullAmount")) = false Then %><%= GetBenefits("ProfDuesFullAmount")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesProfDuesFullAmount"></td>
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table39">
					<tr>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesProfDuesPart" value="1" <% if GetBenefits("ProfDuesPart")=true then%>checked="true"<% end if %> onclick="return false"</td>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesProfDuesPartPaid" value="1" <% if GetBenefits("ProfDuesPartPaid")=true then%>checked="true"<% end if %> onclick="return false"</td>-->
						<td class=formMain width="25%" align=center><% if GetBenefits("ProfDuesPart")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain width="25%" align=center>&nbsp;</td>
						<td class=formMain width="50%" align=center colspan="2">avg $ amount per employee<br><input type="text" readonly size="2" maxlength="5" value=" <% If isNull(GetBenefits("ProfDuesPartAmount")) = false Then %><%= GetBenefits("ProfDuesPartAmount")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesProfDuesPartAmount"></td>
					</tr>
					</table>
				</td>
			</tr>
				
			<tr>
				<td class="formMain" width="40%">Agency paid Pension Plan per employee</td>			
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table40">
					<tr>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesProfRetirementFull" value="1" <% if GetBenefits("RetirementFull")=true then%>checked="true"<% end if %> onclick="return false"</td>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesRetirementFullPaid" value="1" <% if GetBenefits("RetirementFullPaid")=true then%>checked="true"<% end if %></td>-->
						<td class=formMain width="25%" align=center><% if GetBenefits("RetirementFull")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain width="75%" align=center colspan="3">% of contribution by agency&nbsp;<input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("RetirementFullPrcnt")) = false Then %><%= GetBenefits("RetirementFullPrcnt")%><% Else %>0<% End If %>" class="formMain" name="frmExpensesRetirementFullPrcnt"></td>
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table41">
					<tr>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesProfRetirementPart" value="1" <% if GetBenefits("RetirementPart")=true then%>checked="true"<% end if %> onclick="return false"</td>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesRetirementPartPaid" value="1" <% if GetBenefits("RetirementpartPaid")=true then%>checked="true"<% end if %></td>-->
						<td class=formMain width="25%" align=center><% if GetBenefits("RetirementPart")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain width="75%" align=center colspan="3">% of contribution by agency&nbsp;<input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("RetirementPartPrcnt")) = false Then %><%= GetBenefits("RetirementPartPrcnt")%><% Else %>0<% End If %>" class="formMain" name="frmExpensesRetirementPartPrcnt"></td>
					</tr>
					</table>
				</td>
			</tr>
				
			<tr>
				<td class="formMain" width="40%">403 B per employee</td>			
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table42">
					<tr>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpenses403BFull" value="1" <% if GetBenefits("403BFull")=true then%>checked="true"<% end if %> onclick="return false"</td>-->
						<td class=formMain width="25%" align=center><% if GetBenefits("403BFull")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain align=center valign=bottom>Employer Contribution<br><!--<input type="text" size="1" maxlength="3" value=" <% If say = "edit" Then %><%= GetBenefits("403BFullContrib")%><% Else %>0<% End If %>" class="formMain" name="frmExpenses403BFullContrib">-->
							<select name="403BFullContrib" size=1 class="formMain" ID="Select4">
								<option value="selected"><% if GetBenefits("403BFullContrib") then%>Yes<%else%>No<%end if%></option>
							</select>&nbsp;
						</td>
						<td class=formMain align=center>% of matching<br>contribution<br><input type="text" readonly size="1" maxlength="3" value="<% If isNull(GetBenefits("403BFullContribPrcnt")) = false Then %><%= GetBenefits("403BFullContribPrcnt")%><% Else %>0<% End If %>" class="formMain" name="frmExpenses403BFullContribPrcnt"></td>
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table43">
					<tr>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpenses403BPart" value="1" <% if GetBenefits("403BPart")=true then%>checked="true"<% end if %> onclick="return false"</td>-->
						<td class=formMain width="25%" align=center><% if GetBenefits("403BPart")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain align=center valign=bottom>Employer Contribution<br><!--<input type="text" size="1" maxlength="3" value=" <% If say = "edit" Then %><%= GetBenefits("403BPartContrib")%><% Else %>0<% End If %>" class="formMain" name="frmExpenses403BPartContrib">-->
							<select name="403BPartContrib" size=1 class="formMain" ID="Select5">
								<option value="selected"><% if GetBenefits("403BPartContrib") then%>Yes<%else%>No<%end if%></option>
							</select>&nbsp;
						</td>
						<td class=formMain align=center>% of matching<br>contribution<br><input type="text" readonly size="1" maxlength="3" value="<% If isNull(GetBenefits("403BPartContribPrcnt")) = false Then %><%= GetBenefits("403BPartContribPrcnt")%><% Else %>0<% End If %>" class="formMain" name="frmExpenses403BPartContribPrcnt"></td>
					</tr>
					</table>
				</td>
			</tr>
				
			<tr>
				<td class="formMain" width="40%">Telecommuting per employee</td>			
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table44">
					<tr>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesTelecommFull" value="1" <% if GetBenefits("TelecommFull")=true then%>checked="true"<% end if %> onclick="return false"</td>
						<td class=formMain width="20%" align=center><input type="checkbox" class="formMain" name="frmExpensesTelecommFullPaid" value="1" <% if GetBenefits("TelecommFullPaid")=true then%>checked="true"<% end if %></td>-->
						<td class=formMain width="25%" align=center><% if GetBenefits("TelecommFull")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain width="37%" align=center valign=bottom># of EEs<br><input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("TelecommFullCount")) = false Then %><%= GetBenefits("TelecommFullCount")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesTelecommFullCount"></td>
						<td class=formMain width="37%" align=center>% of EE<br>population<br><input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("TelecommFullPrcnt")) = false Then %><%= GetBenefits("TelecommFullPrcnt")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesTelecommFullPrcnt"></td>
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table45">
					<tr>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesTelecommPart" value="1" <% if GetBenefits("TelecommPart")=true then%>checked="true"<% end if %> onclick="return false"</td>
						<td class=formMain width="20%" align=center><input type="checkbox" class="formMain" name="frmExpensesTelecommPartPaid" value="1" <% if GetBenefits("TelecommPartPaid")=true then%>checked="true"<% end if %></td>-->
						<td class=formMain width="25%" align=center><% if GetBenefits("TelecommPart")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain width="37%" align=center valign=bottom># of EEs<br><input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("TelecommPartCount")) = false Then %><%= GetBenefits("TelecommPartCount")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesTelecommPartCount"></td>
						<td class=formMain width="37%" align=center>% of EE<br>population<br><input type="text" readonly size="1" maxlength="3" value=" <% If isNull(GetBenefits("TelecommPartPrcnt")) = false Then %><%= GetBenefits("TelecommPartPrcnt")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesTelecommPartPrcnt"></td>
					</tr>
					</table>
				</td>
			</tr>
				
			<tr>
				<td class="formMain" width="40%">Tuition Reimbersement per employee</td>			
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table46">
					<tr>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesTuitionFull" value="1" <% if GetBenefits("TuitionFull")=true then%>checked="true"<% end if %> onclick="return false"</td>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesTuitionFullPaid" value="1" <% if GetBenefits("TuitionFullPaid")=true then%>checked="true"<% end if %></td>-->
						<td class=formMain width="25%" align=center><% if GetBenefits("TuitionFull")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain width="75%" align=center colspan="3">Maximum $ paid<br><input type="text" readonly size="2" maxlength="3" value=" <% If isNull(GetBenefits("TuitionFullAmount")) = false Then %><%= GetBenefits("TuitionFullAmount")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesTuitionFullAmount"></td>
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table47">
					<tr>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesTuitionPart" value="1" <% if GetBenefits("TuitionPart")=true then%>checked="true"<% end if %> onclick="return false"</td>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesTuitionPartPaid" value="1" <% if GetBenefits("TuitionPartPaid")=true then%>checked="true"<% end if %></td>-->
						<td class=formMain width="25%" align=center><% if GetBenefits("TuitionPart")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain width="75%" align=center colspan="3">Maximum $ paid<br><input type="text" readonly size="2" maxlength="3" value=" <% If isNull(GetBenefits("TuitionPartAmount")) = false Then %><%= GetBenefits("TuitionPartAmount")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesTuitionPartAmount"></td>
					</tr>
					</table>
				</td>
			</tr>
				
		<tr>
				<td class="formMain" width="40%">Professional Development Budget per employee</td>			
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table48">
					<tr>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesTuitionFull" value="1" <% if GetBenefits("TuitionFull")=true then%>checked="true"<% end if %> onclick="return false"</td>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesTuitionFullPaid" value="1" <% if GetBenefits("TuitionFullPaid")=true then%>checked="true"<% end if %></td>-->
						<td class=formMain width="25%" align=center><% if GetBenefits("ProfDevBudgetFull")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain width="75%" align=center colspan="3">Maximum $ paid<br><input type="text" readonly size="2" maxlength="3" value=" <% If isNull(GetBenefits("ProfDevBudgetFullAmount")) = false Then %><%= GetBenefits("ProfDevBudgetFullAmount")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesProfDevBudgetFullAmount"></td>
					</tr>
					</table>
				</td>
				
		<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table49">
					<tr>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesTuitionPart" value="1" <% if GetBenefits("TuitionPart")=true then%>checked="true"<% end if %> onclick="return false"</td>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmExpensesTuitionPartPaid" value="1" <% if GetBenefits("TuitionPartPaid")=true then%>checked="true"<% end if %></td>-->
						<td class=formMain width="25%" align=center><% if GetBenefits("ProfDevBudgetPart")=true then%>Yes<%else%>No<% end if %></td>
						<td class=formMain width="75%" align=center colspan="3">Maximum $ paid<br><input type="text" readonly size="2" maxlength="3" value=" <% If isNull(GetBenefits("ProfDevBudgetPartAmount")) = false Then %><%= GetBenefits("ProfDevBudgetPartAmount")%><% Else %>N/A<% End If %>" class="formMain" name="frmExpensesProfDevBudgetPartAmount"></td>
					</tr>
					</table>
				</td>
			</tr>
						
				</table>
			</td>
			</tr>
						
			</table>
		</td>
	
	</tr>
	
	
		

	<% if printform="No" then %>	
	
		<% if ReadOnlyLevel=0 then %>	
			<tr>
				<td colspan="6" class="formHeader"><input type="submit" value="Edit Form" class="formMainBold"></td>
			</tr>
		<% else %>
			<tr>
				<td colspan="9" class="formMainCentered">Where did the <strong>Edit Button</strong> go?  <a href="..\helpfiles\surveyhelp.asp?HelpID=password1" onclick="NewWindow(this.href,'name','500','250','yes');return false;">Click Here</a> for an explanation.</td>
			</tr>							
		<% end if %>
		<tr>
			<td colspan="6"><!--#include file="../includes/contact_info.inc"--></td>
		</tr>
		
	<% end if %>
</table>

</form>