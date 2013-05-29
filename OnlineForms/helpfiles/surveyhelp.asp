<%
Dim HelpID
if (Request.QueryString("HelpID") <> "") then HelpID = Request.QueryString("HelpID")

Dim SixMonthsAgo
if (Request.QueryString("SixMonthsAgo")<>"") then SixMonthsAgo = Request.QueryString("SixMonthsAgo")

Dim Now
if (Request.QueryString("Now")<>"") then Now = Request.QueryString("Now")

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>Survey Help</title>
</head>
<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">

<body>


<% if HelpID="rev1" then %>

	<p>
	<span class = "formIndex">Revenue Booked</span>
	</p>
	
	<span class = "formMain">
	<p>This is where you report revenue received or otherwise obligated <b>during </b>the month.  This would include cash and other contributions/assets receivable and booked during the month according to FASB (Financial Accounting Standards Board) guidelines.  </p>
	<p>Please provide your most accurate revenue figures available at this time. At a minimum, cash basis revenue should be reported. If accrual basis is available, please report it. Revenue reported should be based upon NET proceeds from fundraising activities. You may edit prior monthly revenue reports as more accurate figures become available.</p>
	<p>Enter only a whole number (no decimals, no commas, no dollar signs)</p>
	
	</span>
	
<% end if %>	

<% if HelpID="transfer" then %>

	<p>
	<span class = "formIndex">Transfer Matches</span>
	</p>
	
	<span class = "formMain">
	<p>
	If you need to move, or "transfer", matches from one program to another, use this field.
	</p>
	<p>For example, if you wish to transfer six matches from your School-Based program to your Community-Based program, enter a "-6" (negative six) into the School-Based transfer field, and a "+6" (positive six) into the Community-Based transfer field.
	</p>
	<p>
	The sum of all transfer fields must add up to zero (balance out).	
	</p>

	
	</span>
	
<% end if %>	



<% if HelpID = "sdm_vol_yield_inq" then %>
	<p>
	<span class = "formIndex">Yield Rate: Volunteer Inquiry</span>
	</p>
	
	<span class="formMain">
	<p>
	Volunteer Inquiry:  A volunteer is counted as an inquiry if he or she contacts the agency, expresses interest in becoming a Big, meets minimum eligibility criteria, and provides basic contact information (e.g. phone number, email, etc.)
	</p>
	
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 3 of the CB, SB and OSB Monthly Report.</b></font>
	</p>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
	
	</span>
<% end if %>

<% if HelpID = "sdm_child_yield_inq" then %>
	<p>
	<span class = "formIndex">Yield Rate: Child Inquiry </span>
	</p>
	
	<span class="formMain">
	<p>
	Counted if parent contacts agency requesting a Big for their child and child meets minimum eligibility criteria.  If a teacher makes a referral as part of a school-based program, count this as an inquiry.
	</p>
	
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 15 of the CB, SB and OSB Monthly Report.</b></font>
	</p>	
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>



<% if HelpID = "sdm_aml" then %>
	<p>
	<span class = "formIndex">Average Match Length</span>
	</p>
	
	<span class="formMain">
	<p>
	Enter the Average Length (in Months) of all matches <strong>closed during the current month</strong>
	</p>
	
	<p>
	<font color="red"><b>METRICS WORKBOOK:  Row 7 of the 'ALL' Sheet.</b></font>
	</p>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>


<% if HelpID = "sdm_proc_vol_inq_intNUM_CB" then %>
	<p>
	<span class = "formIndex">Processing Time: Inquiry to Interview </span>
	</p>
	
	<span class="formMain">
	<p>
	For individuals who initially contacted the agency to express an interest in becoming a volunteer, Enter the number of volunteers who were interviewed during the month.
	</p>
	
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 4 of the CB Monthly Report.</b></font>
	</p>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>


<% if HelpID = "sdm_proc_vol_inq_intAVG_CB" then %>
	<p>
	<span class = "formIndex">Processing Time: Inquiry to Interview </span>
	</p>
	
	<span class="formMain">
	<p>
	For Volunteers who attended in-person interview, enter the average number of days that elapsed for these individuals between the inquiry date and the interview date.
	</p>
	
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 11 of the CB Monthly Report.</b></font>
	</p>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>


<% if HelpID = "sdm_proc_vol_inq_intNUM_SB" then %>
	<p>
	<span class = "formIndex">Processing Time: Inquiry to Interview </span>
	</p>
	
	<span class="formMain">
	<p>
	For individuals who initially contacted the agency to express an interest in becoming a volunteer, Enter the number of volunteers who were interviewed during the month.
	</p>
	
	<p>
	<font color="red"><b>METRICS WORKBOOK: Rows 4 of the SB Monthly Report.</b></font>
	</p>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>


<% if HelpID = "sdm_proc_vol_inq_intAVG_SB" then %>
	<p>
	<span class = "formIndex">Processing Time: Inquiry to Interview </span>
	</p>
	
	<span class="formMain">
	<p>
	For Volunteers who attended in-person interview, enter the average number of days that elapsed for these individuals between the inquiry date and the interview date.
	</p>
	
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 11 of the SB Monthly Report.</b></font>
	</p>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>


<% if HelpID = "sdm_proc_vol_inq_intNUM_OSB" then %>
	<p>
	<span class = "formIndex">Processing Time: Inquiry to Interview </span>
	</p>
	
	<span class="formMain">
	<p>
	For individuals who initially contacted the agency to express an interest in becoming a volunteer, enter the number of volunteers who were interviewed during the month.
	</p>
	
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 4 of the OSB Monthly Report.</b></font>
	</p>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>


<% if HelpID = "sdm_proc_vol_inq_intAVG_OSB" then %>
	<p>
	<span class = "formIndex">Processing Time: Inquiry to Interview </span>
	</p>
	
	<span class="formMain">
	<p>
	For Volunteers who attended in-person interview, enter the average number of days that elapsed for these individuals between the inquiry date and the interview date.
	</p>
	
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 11 of the OSB Monthly Report.</b></font>
	</p>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>


<% if HelpID = "sdm_proc_vol_int_matchNUM_CB" then %>
	<p>
	<span class = "formIndex">Processing Time: Interview to Matched</span>
	</p>
	
	<span class="formMain">
	<p>
	For individuals who were Interviewed, Enter the number of volunteers who were matched during the month.
	</p>
	
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 5 of the CB Monthly Report.</b></font>
	</p>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>

<% if HelpID = "sdm_proc_vol_int_matchAVG_CB" then %>
	<p>
	<span class = "formIndex">Processing Time: Interview to Matched</span>
	</p>
	
	<span class="formMain">
	<p>
	For Volunteers who are <strong>actually </strong>matched, the average number of days that elapsed between the in-person interview and being matched.
	</p>
	
	<font color="red"><b>METRICS WORKBOOK: Row 12 of the CB Monthly Report.</b></font>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>


<% if HelpID = "sdm_proc_vol_int_matchNUM_SB" then %>
	<p>
	<span class = "formIndex">Processing Time: Interview to Matched</span>
	</p>
	
	<span class="formMain">
	<p>
	For individuals who were Interviewed, Enter the number of volunteers who were matched during the month.
	</p>
	
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 5 of the SB Monthly Report.</b></font>
	</p>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>

<% if HelpID = "sdm_proc_vol_int_matchAVG_SB" then %>
	<p>
	<span class = "formIndex">Processing Time: Interview to Matched</span>
	</p>
	
	<span class="formMain">
	<p>
	For Volunteers who are <strong>actually </strong>matched, the average number of days that elapsed between the in-person interview and being matched.
	</p>
	
	<font color="red"><b>METRICS WORKBOOK: Row 12 of the SB Monthly Report.</b></font>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>


<% if HelpID = "sdm_proc_vol_int_matchNUM_OSB" then %>
	<p>
	<span class = "formIndex">Processing Time: Interview to Matched</span>
	</p>
	
	<span class="formMain">
	<p>
	For individuals who were Interviewed, Enter the number of volunteers who were matched during the month.
	</p>
	
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 5 of the OSB Monthly Report.</b></font>
	</p>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>

<% if HelpID = "sdm_proc_vol_int_matchAVG_OSB" then %>
	<p>
	<span class = "formIndex">Processing Time: Interview to Matched</span>
	</p>
	
	<span class="formMain">
	<p>
	For Volunteers who are <strong>actually </strong>matched, the average number of days that elapsed between the in-person interview and being matched.
	</p>
	
	<font color="red"><b>METRICS WORKBOOK: Row 12 of the OSB Monthly Report.</b></font>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>



<% if HelpID = "sdm_proc_child_inq_intNUM_CB" then %>
	<p>
	<span class = "formIndex">Processing Time: Inquiry to Interview </span>
	</p>
	
	<span class="formMain">
	<p>
	For individuals who contacted the agency requesting a Big for their child,  enter the number of children who were interviewed during the month.
	</p>
	
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 16 of the CB Monthly Report.</b></font>
	</p>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>


<% if HelpID = "sdm_proc_child_inq_intAVG_CB" then %>
	<p>
	<span class = "formIndex">Processing Time: Inquiry to Interview </span>
	</p>
	
	<span class="formMain">
	<p>
	For Children who attended in-person interview, enter the average number of days that elapsed for these individuals between the inquiry date and the interview date.
	</p>
	
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 23 of the CB Monthly Report.</b></font>
	</p>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>



<% if HelpID = "sdm_proc_child_inq_intNUM_SB" then %>
	<p>
	<span class = "formIndex">Processing Time: Inquiry to Interview </span>
	</p>
	
	<span class="formMain">
	<p>
	For individuals who contacted the agency requesting a Big for their child,  enter the number of children who were interviewed during the month.
	</p>
	
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 16 of the SB Monthly Report.</b></font>
	</p>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>


<% if HelpID = "sdm_proc_child_inq_intAVG_SB" then %>
	<p>
	<span class = "formIndex">Processing Time: Inquiry to Interview </span>
	</p>
	
	<span class="formMain">
	<p>
	For Children who attended in-person interview, enter the average number of days that elapsed for these individuals between the inquiry date and the interview date.
	</p>
	
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 23 of the SB Monthly Report.</b></font>
	</p>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>



<% if HelpID = "sdm_proc_child_inq_intNUM_OSB" then %>
	<p>
	<span class = "formIndex">Processing Time: Inquiry to Interview </span>
	</p>
	
	<span class="formMain">
	<p>
	For individuals who contacted the agency requesting a Big for their child,  enter the number of children who were interviewed during the month.
	</p>
	
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 16 of the OSB Monthly Report.</b></font>
	</p>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>


<% if HelpID = "sdm_proc_child_inq_intAVG_OSB" then %>
	<p>
	<span class = "formIndex">Processing Time: Inquiry to Interview </span>
	</p>
	
	<span class="formMain">
	<p>
	For Children who attended in-person interview, enter the average number of days that elapsed for these individuals between the inquiry date and the interview date.
	</p>
	
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 23 of the OSB Monthly Report.</b></font>
	</p>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>



<% if HelpID = "sdm_proc_child_int_matchNUM_CB" then %>
	<p>
	<span class = "formIndex">Processing Time: Interview to Matched</span>
	</p>
	
	<span class="formMain">
	<p>
	For Children who went through the Interview process, enter the number of children who were matched during the month.
	</p>
	
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 17 of the CB Monthly Report.</b></font>
	</p>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>

<% if HelpID = "sdm_proc_child_int_matchAVG_CB" then %>
	<p>
	<span class = "formIndex">Processing Time: Interview to Matched</span>
	</p>
	
	<span class="formMain">
	<p>
	For Children who are <strong>actually </strong>matched, the average number of days that elapsed between the in-person interview and being matched.
	</p>
	
	<font color="red"><b>METRICS WORKBOOK: Row 24 of the CB Monthly Report.</b></font>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>


<% if HelpID = "sdm_proc_child_int_matchNUM_SB" then %>
	<p>
	<span class = "formIndex">Processing Time: Interview to Matched</span>
	</p>
	
	<span class="formMain">
	<p>
	For Children who went through the Interview process, enter the number of children who were matched during the month.
	</p>
	
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 17 of the SB Monthly Report.</b></font>
	</p>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>

<% if HelpID = "sdm_proc_child_int_matchAVG_SB" then %>
	<p>
	<span class = "formIndex">Processing Time: Interview to Matched</span>
	</p>
	
	<span class="formMain">
	<p>
	For Children who are <strong>actually </strong>matched, the average number of days that elapsed between the in-person interview and being matched.
	</p>
	
	<font color="red"><b>METRICS WORKBOOK: Row 24 of the SB Monthly Report.</b></font>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>


<% if HelpID = "sdm_proc_child_int_matchNUM_OSB" then %>
	<p>
	<span class = "formIndex">Processing Time: Interview to Matched</span>
	</p>
	
	<span class="formMain">
	<p>
	For Children who went through the Interview process, enter the number of children who were matched during the month.
	</p>
	
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 17 of the OSB Monthly Report.</b></font>
	</p>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>

<% if HelpID = "sdm_proc_child_int_matchAVG_OSB" then %>
	<p>
	<span class = "formIndex">Processing Time: Interview to Matched</span>
	</p>
	
	<span class="formMain">
	<p>
	For Children who are <strong>actually </strong>matched, the average number of days that elapsed between the in-person interview and being matched.
	</p>
	
	<font color="red"><b>METRICS WORKBOOK: Row 24 of the OSB Monthly Report.</b></font>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>




<% if HelpID = "sdm_freq_match_closures" then %>
	<p>
	<span class = "formIndex">Frequency of Match Closures</span>
	</p>
	
	<span class="formMain">
	<p>
	For community-based and site-based matches, as appropriate, the number that closed that had a monthly length corresponding with the frequency categories listed below.  For example, if three of 10 matches that closed in the month were less than 3 months long at time of closure, enter 3.
	</p>
	
	<p>
	<font color="red"><b>METRICS WORKBOOK: Rows 31 to 36 of the CB, SB & OSB Monthly Reports.</b></font>
	</p>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>

<% if HelpID = "sdm_vol_rematched" then %>
	<p>
	<span class = "formIndex">Volunteers Re-Matched</span>
	</p>
	
	<span class="formMain">
	<p>
	Based on the number of volunteers matched during the month, the number of these that are volunteers who are being re-matched after a previous match experience.
	</p>
	
	<font color="red"><b>METRICS WORKBOOK: Row 6 of the CB, SB & OSB Monthly Reports.</b></font>	
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>


<% if HelpID = "sdm_premature_closure" then %>
	<p>
	<span class = "formIndex">Premature Closure</span>
	</p>
	
	<span class="formMain">
	<p>
	Enter the number of matches that closed prematurely during the month.<p>A premature closure is one that closes before its expected commitment is completed. For example, a community-based match closing before 12 months would be counted as a premature closure.</p>
	</p>
	
	<font color="red"><b>METRICS WORKBOOK: Row 28 of the CB,SB & OSB Monthly Reports.</b></font>	
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>	




<% if HelpID = "sdm_closure_cpstatuschange" then %>
	<p>
	<span class = "formIndex">CHILD/PARENT STATUS CHANGE</span>
	</p>
	
	<span class="formMain">
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 38 of the CB,SB & OSB Monthly Reports.</b></font>	
	</p>
	
	<p>
	If you are not using the SDM Workbook, refer to the "SDM Closure Reasons" to assign closure categories. You can download this information as a Word document by clicking <a href="http://agencies.bbbsa.org/DocumentRepository/SDS/Part3/SDM%20Closure%20Codes.doc">here</a>.
	</p>
		
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>


<% if HelpID = "sdm_closure_volstatuschange" then %>
	<p>
	<span class = "formIndex">VOLUNTEER STATUS CHANGE</span>
	</p>
	
	<span class="formMain">
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 39 of the CB,SB & OSB Monthly Reports.</b></font>	
	</p>
	
	<p>
	If you are not using the SDM Workbook, refer to the "SDM Closure Reasons" to assign closure categories. You can download this information as a Word document by clicking <a href="http://agencies.bbbsa.org/DocumentRepository/SDS/Part3/SDM%20Closure%20Codes.doc">here</a>.
	</p>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>

<% if HelpID = "sdm_closure_cpdissatisfaction" then %>
	<p>
	<span class = "formIndex">CHILD/PARENT DISSATISFACTION</span>
	</p>
	
	<span class="formMain">
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 40 of the CB,SB & OSB Monthly Reports.</b></font>	
	</p>
	
	<p>
	If you are not using the SDM Workbook, refer to the "SDM Closure Reasons" to assign closure categories. You can download this information as a Word document by clicking <a href="http://agencies.bbbsa.org/DocumentRepository/SDS/Part3/SDM%20Closure%20Codes.doc">here</a>.
	</p>	  
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>

<% if HelpID = "sdm_closure_voldissatisfaction" then %>
	<p>
	<span class = "formIndex">VOLUNTEER DISSATISFACTION</span>
	</p>
	
	<span class="formMain">
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 41 of the CB,SB & OSB Monthly Reports.</b></font>	
	</p>
	
	<p>
	If you are not using the SDM Workbook, refer to the "SDM Closure Reasons" to assign closure categories. You can download this information as a Word document by clicking <a href="http://agencies.bbbsa.org/DocumentRepository/SDS/Part3/SDM%20Closure%20Codes.doc">here</a>.
	</p>	
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>

<% if HelpID = "sdm_successfulmatches" then %>
	<p>
	<span class = "formIndex">SUCCESSFUL MATCHES</span>
	</p>
	
	<span class="formMain">
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 34 of the 'ALL' Sheet, OR if row 34 says 'Six Month Retention' in blue,  open the '05 Monthly' sheet and see row 37.</b></font>	
	</p>
	
	<p>
	If you are not using the SDM Workbook, refer to the "SDM Closure Reasons" to assign closure categories. You can download this information as a Word document by clicking <a href="http://agencies.bbbsa.org/DocumentRepository/SDS/Part3/SDM%20Closure%20Codes.doc">here</a>.
	</p>	
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>


<% if HelpID = "sdm_ret_new_matches_6months_ago" then %>
	<p>
	<span class = "formIndex">Total Opened 6 Months Ago</span>
	</p>
	
	<span class="formMain">
	<p>
	Please enter the number of new matches that were opened in <b><%=MonthName(Request("SixMonthsAgo"), False)%></b> (six months ago).
	</p>  
	
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 44 of the CB, SB & OSB Monthly Reports.</b></font>	
	</p>	
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>

<% if HelpID = "sdm16" then %>
	<p>
	<span class = "formIndex">Number Still Open Now</span>
	</p>
	
	<span class="formMain">
	<p>
	Out of the new matches that were opened back in <b><%=MonthName(Request("SixMonthsAgo"), False)%></b> (six months ago), please enter the number of those matches that CLOSED before the end of <b><%=MonthName(Request("Now"), False)%></b>.
	</p>  
	
	<p>
	<font color="red"><b>METRICS WORKBOOK: Row 45 of the CB, SB & OSB Monthly Reports.</b></font>	
	</p>		
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>

<% if HelpID = "cust_sat_1" then %>
	<p>
	<span class = "formIndex">Customer Satisfaction</span>
	</p>
	
	<span class="formMain">
	<p>
	<strong>Enrollment Satisfaction Average Score</strong>
 	</p>	
	<p>
	Enter the average score (1-5) of the ratings given by volunteers on their Customer Satisfaction Surveys (Enrollment) for the most recent  quarter (January - March; April - June; July - September; October - December).
	</p>
	
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>
		
	</span>
<% end if %>

<% if HelpID = "cust_sat_2" then %>
	<p>
	<span class = "formIndex">Customer Satisfaction</span>
	</p>
	
	<span class="formMain">
	<p>
	<strong>Enrollment Satisfaction Count</strong>
 	</p>
	<p>
	Enter the number of volunteers who reported their satisfaction on a Customer Satisfaction Survey (Enrollment) during the most recent quarter (January - March; April - June; July - September; October - December).
	</p>
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>		
	</span>
<% end if %>

<% if HelpID = "cust_sat_3" then %>
	<p>
	<span class = "formIndex">Customer Satisfaction</span>
	</p>
	
	<span class="formMain">
	<p>
	<strong>Match Satisfaction Average Score</strong>
 	</p>
	<p>
	Enter the average score (1-5) of the ratings given by volunteers on their Customer Satisfaction Survey (Match) for the most recent quarter (January - March; April - June; July - September; October - December).
	</p>
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>		
	</span>
<% end if %>

<% if HelpID = "cust_sat_4" then %>

	<p>
	<span class = "formIndex">Customer Satisfaction</span>
	</p>
	
	<span class="formMain">
	<p>
	<strong>Match Satisfaction Count</strong>
 	</p>
	
	<p>
	Enter the number of volunteers who reported their satisfaction on a Customer Satisfaction Survey (Match) during the most recent quarter (January - March; April - June; July - September; October - December).
	</p>
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>		
	</span>
	
<% end if %>

<% if HelpID = "poe_aggregate" then %>

	<p>
	<span class = "formIndex">POE</span>
	</p>
	
	<span class="formMain">
	<p>
	<strong>POE Aggregate Score</strong>
 	</p>
	
	<p>
	Enter the aggregate score (1 - 5) from the POE reports of volunteers as calculated in the POE Workbook at the end of the most recent quarter (March, June, September, or December).
	</p>
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>		
	</span>
	
<% end if %>

<% if HelpID = "poe_count" then %>

	<p>
	<span class = "formIndex">POE</span>
	</p>
	
	<span class="formMain">
	<p>
	<strong>POE Count</strong>
 	</p>
	
	<p>
	Enter the number of volunteers who submitted a POE report for the most recent quarter (January - March; April - June; July - September; or October - December).
	</p>
	<p>
	If you have any additional questions regarding <strong>SDM Metrics only,</strong> please contact us at <a href="mailto:sdm@bbbsa.org">sdm@bbbsa.org</a>
	</p>		
	</span>
	
<% end if %>



<% if HelpID = "pq1" then %>	
	<p>
	<span class = "formIndex">Partnership Questionnaire</span>
	</p>
	
	<span class = "formMain">
	<p>As of May 1, 2003, included in the monthly performance report is a questionnaire related to partnerships.  The purpose of the questionnaire is to gain an understanding of the current status of national volunteer-rich partnerships.  Your feedback will enable us to shape the strategic direction for national volunteer-rich partnerships, in addition to pursue funding opportunities. 
	The questionnaire will appear in the monthly reporting forms bi-annually, in May and December (for the initial round, the questionnaire appears for the April 2003 reporting).
	</p>
	<p>
	For issues related to the <strong>Partnership Questionnaire Only</strong>, please contact Dionne Vernon, Director of Volunteer Development, at <a href="mailto:dvernon@bbbsa.org">dvernon@bbbsa.org</a>
	</p>
	
	</span>
<% end if %>	


<% if HelpID = "password1" then %>

	<p>
	<span class = "formIndex">Data Entry Restrictions</span>
	</p>

	<span class="formMain">
	<p>
	The password security has changed as a result of requests from numerous agencies. An explanation regarding the change was sent out to all agency ED/CEOs in the November 14th, 2003 edition of "The Latest On..." e-bulletin.  
	<br><br>There are now 3 levels of security: 
	<ul>
		<li>Read-Only (which your password now is) </li>
		<li>Limited Access (which is what you need to continue with data entry)</li>
		<li>Full Access (for Agency CEOs/EDs) </li>	
	</ul>
	You'll need to contact your ED/CEO for the Limited Access password. Once you get the Limited Access password, you'll be able to edit all of your data as before. 
	
	</p>
	
	</span>
	
	
<% end if %>


<% if HelpID = "rtbm1" then %>	
	<p>
	<span class = "formIndex">Ready to be Matched</span>
	</p>
	
	<span class = "formMain">
	<p>Enter the number of <b><i>unmatched children</i></b> ("Ready to be Matched") as of the last day of December.  
This figure is requested annually.  It has been moved to this monthly survey (but only asked each December) in order to simplify and streamline the Annual Agency Information survey (previously called the ADS-Agency Demographic Survey).
	</p>

	
	</span>
<% end if %>	


<% if HelpID = "rtbm2" then %>	
	<p>
	<span class = "formIndex">Ready to be Matched</span>
	</p>
	
	<span class = "formMain">
	<p>Enter the number of <b><i>unmatched volunteers</i></b> ("Ready to be Matched") as of the last day of December.  
This figure is requested annually.  It has been moved to this monthly survey (but only asked each December) in order to simplify and streamline the Annual Agency Information survey (previously called the ADS-Agency Demographic Survey).
	</p>
	
	</span>
<% end if %>	


<% if HelpID = "yearly1" then %>
	<p>
	<span class = "formIndex">2003 Annual Agency Information - Overview</span>
	</p>
	<span class="formMain">	
	
	<p>
	The annual portion of the old ADS/ADR has been redesigned and renamed “AAI - Annual Agency Information”.  This document describes the changes you will see when filling out the 2003 AAI survey form (compared to the 2002 ADS survey forms).  As you will see:
	</p>
	
	<p>
	<ul>
	<li>There are fewer forms and fewer questions</li>
	<li>Where there are “new” questions, they are subjective questions that do not require any retroactive record-keeping</li>
	<li>Many of the questions have been updated to reflect Service Delivery Model concepts, functions and definitions</li>
	</ul>
	</p>
	
	<p><em><strong><font color="#ff0000">PLEASE ENTER YOUR 2003 DATA BY FRIDAY, FEBRUARY 15, 2004</font></strong></em></p>
	
	<p>
	<hr>
	<span class = "formIndex">Changes (Releative to 2002 ADS)</span>
	<br><br>	
	</p>
	
	<span class = "formMain">
	<p>
	<b>GENERAL FORM</b>
	<ul>
	<li>This form has been eliminated.  </li>
	<li>A new “SDM Information” form will capture the YIELD and the VOLUNTEER REMATCH RATE information.  This new form is a “transition” form while agencies migrate to using the more complete SDM metrics workbook and/or upgrade to AIM.</li>
	<li>The unmatched-list (RTBM children) will be collected each year in the December monthly match performance survey.</li>
	</ul>
	</p>
	
	<p>
	<b>REVENUE</b>
	<ul>
	<li>BBBSA grants are reported separately.</li>
	<li>RMM revenue is reported separately.</li>
	<li>Online Donations (through BBBSA) are reported separately.</li>
	<li>The six most common Special Events are reported separately.</li>
	<li>A question has been added to identify revenue from grants or restricted funds that are specifically targeted to non-mentoring activities.</li>
	</ul>
	</p>
	
	<p>
	<b>EXPENSES</b>
	<ul>
	<li>Insurance, Marketing and Capital/Furniture/Equipment expenses are reported separately.</li>
	<li>“Functional Allocation” section renamed to “Breakdown by Expense Category”</li>
	<li>Two new “breakdown” sections added.  Agencies are asked, to the best of their ability, to estimate the allocation of expenses in each of these 2 sections so that they total 100%.
	<ul>
		<li>Breakdown by Function (Customer Relations, Enrollment & Matching, Match Support)</li>
		<li>Breakdown by Program Type (Community, School, Site: Non-School)</li>
	</ul>	
	<li>A new “benefits” section has been added.  For both full-time and part-time employees, agencies are asked to check off which benefits they provide and, in the case of medical/dental, the % of costs paid by the agency for the employee and their family.</li>
	</ul>
	</p>
	
	<p>
	<b>BOARD MEMBERS</b>
	</p>
	
	<p>
	<ul>
	<li>Questions added regarding policy on board financial contributions</li>
	</ul>
	</p>
	
	<p>
	<b>STAFF</b>
	</p>
	
	<p>
	<ul>
	<li>No changes to the form</li>
	<li>Choices for “Education” changed:  
	<ul>
		<li>“Professional Degree” is removed and incorporated into “Masters (and beyond) Degree”</li>
	</ul>	
	</li>
	</ul>
	</p>
	
	<p>
	<b>SPECIAL PROGRAMS AND SPECIAL POPULATIONS</b>
	</p>
	
	<p>
	<ul>
	<li>These forms have been eliminated.</li>
	<li>Selected information from these forms will be incorporated into a new “MY PROFILE” of the “MY AGENCY” section of the agencies website.</li>
	</ul>	
	</p>
	
	<p>
	<hr>
	<span class="formIndex">Summary</span>	
	</p>
	
	<p>
	BBBS agencies will benefit from the reports and analysis of this annual data.  The questions have been updated to reflect the Service Delivery Model and the increasing emphasis on Performance Management.  There is increased focus on the Revenue side of our business as well as how benefits are distributed.  Reports on this 2003 data will be released during the 1st quarter of 2004.
	</p>
	
	
	</span>
<% end if %>


<% if HelpID = "total_expenditures" then %>	
	<p>
	<span class = "formIndex">(A) TOTAL EXPENDITURES</span>
	</p>
	
	<span class = "formMain">
	<p>This amount must agree with your last completed financial audit.  If your audit is not yet completed, use your unaudited figures.  Adjustments, if material, will be made when the audited figures are sent to us.  Send your most current audited financial statement and note on the front of the form when your fiscal year ends.
	</p>
	
	</span>
<% end if %>

<% if HelpID = "prior_year_fees_paid_to_bbbsa" then %>	
	<p>
	<span class = "formIndex">(B) Less: Prior Year Fees Paid to BBBSA</span>
	</p>
	
	<span class = "formMain">
	<p>Includes affiliation fees only.
	</p>
	<p>All deductions should be substantiated and will be verified against a copy of your last audit, which must be submitted to BBBSA after the audit has been finalized.
	</p>
	
	</span>
<% end if %>

<% if HelpID = "prior_year_capital_purchases" then %>	
	<p>
	<span class = "formIndex">(C) Less: Prior Year Capital Purchases</span>
	</p>
	
	<span class = "formMain">
	<p>Land, buildings, and equipment purchased <strong>(only if included in your Total Expenditure amount)</strong>.
	</p>
	<p>All deductions should be substantiated and will be verified against a copy of your last audit, which must be submitted to BBBSA after the audit has been finalized.
	</p>
	
	</span>
<% end if %>

<% if HelpID = "prior_year_depreciation" then %>	
	<p>
	<span class = "formIndex">(D) Less: Prior Year Depreciation</span>
	</p>
	
	<span class = "formMain">
	<p>Do not include depreciation expenses for the capital purchases included above.
	</p>
	<p>All deductions should be substantiated and will be verified against a copy of your last audit, which must be submitted to BBBSA after the audit has been finalized.
	</p>
	
	</span>
<% end if %>

<% if HelpID = "prior_year_fundraising_expenses" then %>	
	<p>
	<span class = "formIndex">(E) Less: Prior Year Fundraising Expenses</span>
	</p>
	
	<span class = "formMain">
	<p>Include in this line all expenses related to fundraising activities.  This can include direct salaries and fringes paid to fundraising personnel.  Since this may have a significant impact on your fees, these expenses must be direct fundraising expenses only, not allocated.
	</p>
	<p>All deductions should be substantiated and will be verified against a copy of your last audit, which must be submitted to BBBSA after the audit has been finalized.
	</p>
	
	</span>
<% end if %>

<!-- Monthly Revenue / Expense Help -->
<% if HelpID = "finance_performance_united_way" then %>	
	<p>
	<span class = "formIndex">United Way</span>
	</p>
	
	<span class = "formMain">
	<p>Enter the total of monthly revenue received from United Way(s), including donor designations.
	</p>	
	</span>
<% end if %>


<% if HelpID = "finance_performance_gov_federal_funding" then %>	
	<p>
	<span class = "formIndex">Government - Federal Funding</span>
	</p>
	
	<span class = "formMain">
	<p>Enter the revenue from federal grants either to the agency as a direct recipient or as a sub-recipient.
	</p>	
	</span>
<% end if %>

<% if HelpID = "finance_performance_gov_state_funding" then %>	
	<p>
	<span class = "formIndex">Government - State Funding</span>
	</p>
	
	<span class = "formMain">
	<p>Enter the revenue from a state government funding source, either as a direct recipient or as a sub-recipient.
	</p>	
	</span>
<% end if %>

<% if HelpID = "finance_performance_gov_local_funding" then %>	
	<p>
	<span class = "formIndex">Government - Local Funding</span>
	</p>
	
	<span class = "formMain">
	<p>Enter the revenue from a local government funding source, either as a direct recipient or as a sub-recipient.
	</p>	
	</span>
<% end if %>

<% if HelpID = "finance_performance_foundations_grants" then %>	
	<p>
	<span class = "formIndex">Foundations - Grants</span>
	</p>
	
	<span class = "formMain">
	<p>Enter the revenue from grants given by private or corporate foundations.
	</p>	
	</span>
<% end if %>

<% if HelpID = "finance_performance_corporations" then %>	
	<p>
	<span class = "formIndex">Corporations - Non-event Donations</span>
	</p>
	
	<span class = "formMain">
	<p>Enter revenue of any amount provided from a for-profit business. This may be in the form of grants or direct contributions. Regardless of the purpose of the donation, report it here if the SOURCE of revenue is from a corporation or business entity.
	</p>	
	</span>
<% end if %>

<% if HelpID = "finance_performance_bbbsa_grants" then %>	
	<p>
	<span class = "formIndex">BBBSA (Pass-Through) Grants</span>
	</p>
	
	<span class = "formMain">
	<p>Enter revenue that is passed on to you as a grant from BBBSA (such as SBM grants).
	</p>	
	</span>
<% end if %>

<% if HelpID = "finance_performance_individual_giving" then %>	
	<p>
	<span class = "formIndex">Individual Giving (Non-Event)</span>
	</p>
	
	<span class = "formMain">
	<p>Enter revenue received from individuals, such as pledges for RMM or general contributions. Regardless of the purpose of the donation, report it here if the SOURCE is from an individual.
	</p>	
	</span>
<% end if %>

<% if HelpID = "finance_performance_events" then %>	
	<p>
	<span class = "formIndex">Events</span>
	</p>
	
	<span class = "formMain">
	<p>Enter the <strong>gross</strong> revenue received through events, both individual (BFKS, RMM, etc.) or Corporate (Sponsorships, etc.)
	</p>	
	</span>
<% end if %>


<% if HelpID = "finance_performance_events_individual" then %>	
	<p>
	<span class = "formIndex">Events - Portion From Individuals</span>
	</p>
	
	<span class = "formMain">
	<p>Enter the portion of <b>gross</b> events revenue received from individuals, such as pledges for BFKS or RMM.
	</p>	
	</span>
<% end if %>

<% if HelpID = "finance_performance_events_corporations" then %>	
	<p>
	<span class = "formIndex">Events - Portion From Corporations</span>
	</p>
	
	<span class = "formMain">
	<p>Enter the portion of <b>gross</b> events revenue received from a for-profit business in the form of event sponsorships, pledges, or direct contributions. 
	</p>	
	</span>
<% end if %>

<% if HelpID = "finance_performance_events_dividends_interest" then %>	
	<p>
	<span class = "formIndex">Dividends and Interest</span>
	</p>
	
	<span class = "formMain">
	<p>Enter revenue from dividends or interest bearing accounts. 
	</p>	
	</span>
<% end if %>

<% if HelpID = "finance_performance_other" then %>	
	<p>
	<span class = "formIndex">Other</span>
	</p>
	
	<span class = "formMain">
	<p>Enter evenue from sources not included in any of the above categories. 
	</p>	
	</span>
<% end if %>

<% if HelpID = "finance_performance_event_expenses" then %>	
	<p>
	<span class = "formIndex">Total Direct Expenses from Special Event Fundraising</span>
	</p>
	
	<span class = "formMain">
	<p>Include all expenses directly related to fundraising special events this month.
	</p>	
	</span>
<% end if %>

<% if HelpID = "finance_performance_event_BFKS" then %>	
	<p>
	<span class = "formIndex">Of Total, Amount Raised Through BFKS</span>
	</p>
	
	<span class = "formMain">
	<p>Report here your total NET BFKS revenue for this month. To calculate NET you should deduct direct fundraising expenses from the event, but not including salaries of fundraising personnel. You should report BFKS <b>gross</b> revenue by SOURCES in the categories above (e.g. individual giving or corporate gifts), while reporting the total net for the event here.
	</p>	
	</span>
<% end if %>

<% if HelpID = "finance_performance_event_RMM" then %>	
	<p>
	<span class = "formIndex">Of Total, Amount Raised Through RMM</span>
	</p>
	
	<span class = "formMain">
	<p>Report here your total RMM revenue, which has been reported broken out by SOURCES in the categories above.
	</p>	
	</span>
<% end if %>

<% if HelpID = "BenMedOffered" then %>	
	<p>
	<span class = "formIndex">Medical Insurance</span>
	</p>
<span class = "formMain">
	<p>This is a part of Medical benefits section where you need to provide details about Medical Insurance benefit in your agency.</p>
	<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 width=550>
			<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
				<td width=150 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Field Name<o:p></o:p></b></p>
				</td>
				<td width=400 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Definition<o:p></o:p></b></p>
				</td>
			</tr>

        <tr style='mso-yfti-irow:1'>
				<td width=150 valign=top>
				<p class=MsoNormal>Medical Offered Y/N</p>
				</td>
     <td width=400 valign=top>
					<p class=MsoNormal>Select "YES" if your agency offers Medical Insurance to your employee.<o:p></o:p></p>
				</td>
			</tr>	

    <tr style='mso-yfti-irow:2'>
		<td width=150 valign=top>
	<p class=MsoNormal>Average % Premium paid by agency Per Employee</p>
		</td>
   
     <td width=400 valign=top>
			<p class=MsoNormal>Enter Average Percentage of Total Monthly Premium paid by your
		agency for individual coverage per full time employee. Ex:50%<o:p></o:p></p>
				</td>
			</tr>

     <tr style='mso-yfti-irow:3'>
		<td width=150 valign=top> 
		<p  class=MsoNormal> Average Total Monthly Premium Per Employee</p>
		</td>

    <td width=400 valign=top>
			<p class=MsoNormal>Enter Average Total Monthly Dollar Amount of Premium cost for
		individual coverage per full time employee. This amount includes both Agency
		Contribution and Employee Contribution(Enter the premium even if the agency contribution is 0). Ex:$550<o:p></o:p></p>
				</td>
			</tr>
	
	<tr style='mso-yfti-irow:3'>
		<td width=150 valign=top> 
	<p class=MsoNormal>Average % Premium paid by agency Per Employee Family</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Enter Average Percentage of Total Monthly Premium paid by your
		agency per full time employee family. Ex:75%<o:p></o:p></p>
				</td>
			</tr>
	<tr style='mso-yfti-irow:3'>
		<td width=150 valign=top> 
	<p  class=MsoNormal>Average Total Monthly Premium Per Employee Family</p>
		</td>	
		
		<td width=400 valign=top>
				<p class=MsoNormal>Enter Total Monthly Dollar Amount of Premium cost per full
			time employee family. This amount includes both Agency Contribution and
			Employee Contribution.(Enter the premium even if the agency contribution is 0)<o:p></o:p></p>
		</td>
	</tr>
</table>


<p class=MsoNormal>*&nbsp;The same terms and definitions apply for PART TIME employees section.</p>

	</span>
<% end if %>

<% if HelpID = "BenDentOffered" then %>	
	<p>
	<span class = "formIndex">Dental Insurance</span>
	</p>
	
	<span class = "formMain">
	<p>This is a part of Medical benefits section where you need to provide details about Dental Insurance benefit in your agency.</p>
    <P>If Dental Insurance is part of your medical benefits package and dental premium cost cannot be determined separetely, please select "YES" for Dental offered and leave percentages and premium amounts as 0`s.</p>
      <table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 width=550>
			<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
				<td width=150 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Field Name<o:p></o:p></b></p>
				</td>
				<td width=400 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Definition<o:p></o:p></b></p>
				</td>
			</tr>
			
		        <tr style='mso-yfti-irow:1'>
				<td width=150 valign=top>
				<p class=MsoNormal>Dental Offered Y/N</p>
				</td>
     <td width=400 valign=top>
					<p class=MsoNormal>Select "YES" if your agency offers Dental Insurance to your
		employees.<u><o:p></o:p></u></p>
				</td>
			</tr>	

    <tr style='mso-yfti-irow:2'>
		<td width=150 valign=top>
	<p class=MsoNormal>Average % Premium paid by agency Per Employee</p>
		</td>
   
     <td width=400 valign=top>
			<p class=MsoNormal>Enter Average Percentage of Total Monthly Premium paid by your
		agency for individual coverage per full time employee.<o:p></o:p></p>
				</td>
			</tr>

     <tr style='mso-yfti-irow:3'>
		<td width=150 valign=top> 
		<p  class=MsoNormal>Average Total Monthly Premium Per Employee</p>
		</td>

    <td width=400 valign=top>
			<p class=MsoNormal>Enter Average Total Monthly Dollar Amount of Premium cost for
		individual coverage per full time employee. This amount includes both Agency
		Contribution and Employee Contribution.(Enter the premium even if the agency contribution is 0)<o:p></o:p></p>
				</td>
			</tr>
	
	<tr style='mso-yfti-irow:3'>
		<td width=150 valign=top> 
	<p  class=MsoNormal>Average % Premium paid by agency Per Employee Family</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Enter Average Percentage of Total Monthly Premium paid by your
		agency per full time employee family.<o:p></o:p></p>
				</td>
			</tr>
			
      <tr style='mso-yfti-irow:3'>
		<td width=150 valign=top> 
	<p  class=MsoNormal>Average Total Monthly Premium Per Employee Family</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Enter Average Total Monthly Dollar Amount of Premium cost per full
		time employee family. This amount includes both Agency Contribution and
		Employee Contribution.(Enter the premium even if the agency contribution is 0)<o:p></o:p></p>
				</td>
			</tr>
			</table>

	
		<p class=MsoNormal>*&nbsp;The same terms and
		definitions apply for PART TIME employees section.</p>
		
	</span>
<% end if %>

<% if HelpID = "BenVisOffered" then %>	
	<p>
	<span class = "formIndex">Vision Insurance</span>
	</p>
	
	<span class = "formMain">
	<p>This is a part of Medical benefits section where you need to provide details about Vision Insurance benefit in your agency.</p>
	<P>If Vision Insurance is part of your medical benefits package and vision premium cost cannot be determined separetely, please select "YES" for vision offered and leave percentages and premium amounts as 0`s.</p>

		 <table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 width=550>
			<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
				<td width=150 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Field Name<o:p></o:p></b></p>
				</td>
				<td width=400 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Definition<o:p></o:p></b></p>
				</td>
			</tr>
        <tr style='mso-yfti-irow:1'>
				<td width=150 valign=top>
				<p class=MsoNormal>Vision Offered Y/N</p>
				</td>
     <td width=400 valign=top>
					<p class=MsoNormal>Select "YES" if your agency offers Vision Insurance to your employees.<u><o:p></o:p></u></p>
				</td>
			</tr>	

    <tr style='mso-yfti-irow:2'>
		<td width=150 valign=top>
	<p class=MsoNormal>Average % Premium paid by agency Per Employee</p>
		</td>
   
     <td width=400 valign=top>
			<p class=MsoNormal>Enter Average Percentage of Total Monthly Premium paid by your
		agency per full time employee.<o:p></o:p></p>
				</td>
			</tr>

     <tr style='mso-yfti-irow:3'>
		<td width=150 valign=top> 
		<p  class=MsoNormal>Average Total Monthly Premium per Employee</p>
		</td>

    <td width=400 valign=top>
			<p class=MsoNormal>Enter Average Total Monthly Dollar Amount of Premium cost per full
		time employee. This amount includes both Agency Contribution and Employee
		Contribution.(Enter the premium even if the agency contribution is 0)<o:p></o:p></p>
				</td>
			</tr>
	
	<tr style='mso-yfti-irow:3'>
		<td width=150 valign=top> 
	<p  class=MsoNormal>Average Total Monthly Premium Per Employee Family</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Enter Percentage of Total Monthly Premium paid by your
		agency per full time employee family.<o:p></o:p></p>
				</td>
			</tr>
			
      <tr style='mso-yfti-irow:3'>
		<td width=150 valign=top> 
	<p  class=MsoNormal>Total Monthly Premium per Employee Family</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Enter Total Monthly Dollar Amount of Premium cost per full
		time employee family. This amount includes both Agency Contribution and
		Employee Contribution.(Enter the premium even if the agency contribution is 0)<o:p></o:p></p>
				</td>
			</tr>
			</table>


	<p class=MsoNormal>*&nbsp;The same terms and
	definitions apply for PART TIME employees section.</p>
		
	</span>
<% end if %>

<% if HelpID = "DisInsShortTermFull" then %>	
	<p>
	<span class = "formIndex">Disability Insurance SHORT Term</span>
	</p>
	
	<span class = "formMain">
	<p>This is a part of NON-Medical benefits section where you need to provide details about SHORT Term Disability Insurance benefit in your agency.</p>
	

		<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 width=550>
			<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
				<td width=150 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Field Name<o:p></o:p></b></p>
				</td>
				<td width=400 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Definition<o:p></o:p></b></p>
				</td>
			</tr>
	        <tr style='mso-yfti-irow:1'>
				<td width=150 valign=top>
				<p class=MsoNormal>Offered</p>
				</td>
     <td width=400 valign=top>
					<p class=MsoNormal>Select "YES" if your agency offers SHORT TERM Disability
		Insurance to full time employees.<u><o:p></o:p></u></p>
				</td>
			</tr>	

   <!--<tr style='mso-yfti-irow:2'> // Commented PAID was taked out in 2008 forms 
		<td width=150 valign=top>
	<p class=MsoNormal>PAID</p>
		</td>
   
     <td width=400 valign=top>
			<p class=MsoNormal>Check if your agency pays <b><i>any part</i></b> of SHORT TERM
		Disability Insurance for individual coverage for full time employees.<o:p></o:p></p>
				</td>
			</tr> --->

     <tr style='mso-yfti-irow:3'>
		<td width=150 valign=top> 
		<p  class=MsoNormal>% Premium paid by agency</p>
		</td>

    <td width=400 valign=top>
			<p class=MsoNormal>Enter percentage of total monthly premium your agency pays
		for SHORT TERM disability Insurance for one full time employee.<o:p></o:p></p>
				</td>
			</tr>
	
	<tr style='mso-yfti-irow:3'>
		<td width=150 valign=top> 
	<p  class=MsoNormal>Total Monthly Premium</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Enter dollar amount of total monthly premium for SHORT
		TERM Disability Insurance for individual coverage for full time employee. This amount includes
		both the agency and employee contributions.<o:p></o:p></p>
				</td>
			</tr>
			</table>

		<p class=MsoNormal>*&nbsp;The same terms and
	definitions apply for PART TIME employees section.</p>
		
	</span>
<% end if %>


<% if HelpID = "DisInsLongTermFull" then %>	
	<p>
	<span class = "formIndex">Disability Insurance LONG Term</span>
	</p>
	
	<span class = "formMain">
	<p>This is a part of NON-Medical benefits section where you need to provide details about LONG Term Disability Insurance benefit in your agency.</p>

		<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 width=550>
			<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
				<td width=150 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Field Name<o:p></o:p></b></p>
				</td>
				<td width=400 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Definition<o:p></o:p></b></p>
				</td>
			</tr>
	        <tr style='mso-yfti-irow:1'>
				<td width=150 valign=top>
				<p class=MsoNormal>Offered</p>
				</td>
     <td width=400 valign=top>
					<p class=MsoNormal>Select "YES" if your agency offers LONG TERM Disability Insurance
		to full time employees.<u><o:p></o:p></u></p>
				</td>
			</tr>	

   <!--- <tr style='mso-yfti-irow:2'>
		<td width=150 valign=top>
	<p class=MsoNormal>PAID</p>
		</td>
   
     <td width=400 valign=top>
			<p class=MsoNormal>Check if your agency pays <b><i>any part</i></b> of LONG TERM Disability
		Insurance for full time employees.<o:p></o:p></p>
				</td>
			</tr>--->

     <tr style='mso-yfti-irow:3'>
		<td width=150 valign=top> 
		<p  class=MsoNormal> Premium paid by agency</p>
		</td>

    <td width=400 valign=top>
			<p class=MsoNormal>Enter percentage of total monthly premium your agency pays
		for LONG TERM disability Insurance for one full time employee.<o:p></o:p></p>
				</td>
			</tr>
	
	<tr style='mso-yfti-irow:3'>
		<td width=150 valign=top> 
	<p  class=MsoNormal>Total Monthly Premium</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Enter dollar amount of total monthly premium for LONG TERM
		disability Insurance for one full time employee. This amount includes both
		the agency and employee contributions.
				<p>If monthly premium varies by employee, use the average premium.
		(Calculation of average premium = Total monthly premiums for the agency <u><b>for this benefit</B></u> divided by # of employees receiving <U><B>this benefit</B></U>.</p><o:p></o:p></p>


			
			
			
				</td>
			</tr>

		</table>

		<p class=MsoNormal>*&nbsp;The same terms and
	definitions apply for PART TIME employees section.</p>
		
	</span>
<% end if %>



<% if HelpID = "EAPFull" then %>	
	<p>
	<span class = "formIndex">EAP: Employee Assistance Program</span>
	</p>
	
	<span class = "formMain">
	<p>This is a part of NON-Medical benefits section where you need to provide details about EAP benefit in your agency.</p>
	<!--<p>If EAP benefit is part of your medical benefits package and the EAP premium cost can not be determined separetely, please select that this benefit is offered, and leave percentages and premium amounts as 0`s. </p> commented out for validation--> 
    
     <table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 width=550>
			<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
				<td width=150 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Field Name<o:p></o:p></b></p>
				</td>
				<td width=400 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Definition<o:p></o:p></b></p>
				</td>
			</tr>

		
	  <tr style='mso-yfti-irow:2'>
		<td width=150 valign=top>
		<p class=MsoNormal>Offered</p>
		</td>
		<td width=400 valign=top>
		<p class=MsoNormal>Select "YES" if your agency offers Employee Assistance Programs (EAP) to full time employees.<o:p></o:p></p>
				</td>
	</tr>	
	
		
   <!--- <tr style='mso-yfti-irow:2'>
		<td width=150 valign=top>
		<p class=MsoNormal>PAID</p>
		</td>
		<td width=400 valign=top>
		<p class=MsoNormal>Check if your agency pays <b><i>any part</i></b> of EAP for full time employees<u><o:p></o:p></u></p>
		</td>
	</tr>--->
		
  <tr style='mso-yfti-irow:3'>
		<td width=150 valign=top> 
		<p  class=MsoNormal>% Premium paid by agency</p>
		</td>
		<td width=400 valign=top> 
		<p class=MsoNormal>Enter percentage of total monthly premium your agency pays
		for EAP for one full time employee.</p>
		</td>
  </tr>
		
   <tr style='mso-yfti-irow:4'>
		<td width=150 valign=top> 
		<p class=MsoNormal>Total Monthly Premium:</p>
		</td>
		<td width=400 valign=top> 
		<p class=MsoNormal>Enter dollar amount of total monthly premium for EAP for
		one full time employee. This amount includes both the agency and employee
		contribution.</p>
		</td>
  </tr>
		</table>

		<p class=MsoNormal>*&nbsp;The same terms and
						definitions apply for PART TIME employees section.</p>

	</span>
<% end if %>





<% if HelpID = "FlexFull" then %>	
	<p>
	<span class = "formIndex">"Flex" Pre-Tax Plan (medical, dependent)</span>
	</p>
	
	<span class = "formMain">
	<p>This is a part of NON-Medical benefits section where you need to provide details about "Flex" Pre-Tax Plan benefit in your agency.</p>
	
		<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0
		style='border-collapse:collapse;border:none;mso-border-alt:solid windowtext .5pt;
		mso-yfti-tbllook:480;mso-padding-alt:0in 5.4pt 0in 5.4pt;mso-border-insideh:
		.5pt solid windowtext;mso-border-insidev:.5pt solid windowtext' ID="Table5">
		<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
		<td width=163 valign=top style='width:1.7in;border:solid windowtext 1.0pt;
		mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt'>
		<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Field Name<o:p></o:p></b></p>
		</td>
		<td width=427 valign=top style='width:4.45in;border:solid windowtext 1.0pt;
		border-left:none;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
		solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt'>
		<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Definition<o:p></o:p></b></p>
		</td>
		</tr>
		<tr style='mso-yfti-irow:1'>
		<td width=163 valign=top style='width:1.7in;border:solid windowtext 1.0pt;
		border-top:none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
		padding:0in 5.4pt 0in 5.4pt'>
		<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><i
		style='mso-bidi-font-style:normal'>Offered</i></b><u><o:p></o:p></u></p>
		</td>
		<td width=427 valign=top style='width:4.45in;border-top:none;border-left:
		none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
		mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
		mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt'>
		<p class=MsoNormal>Check if your agency offers “Flex” Pre-Tax Plan to full
		time employees.<o:p></o:p></p>
		<p class=MsoNormal><u><o:p><span style='text-decoration:none'>&nbsp;</span></o:p></u></p>
		</td>
		</tr>
		<tr style='mso-yfti-irow:2'>
		<td width=163 valign=top style='width:1.7in;border:solid windowtext 1.0pt;
		border-top:none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
		padding:0in 5.4pt 0in 5.4pt'>
		<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><i
		style='mso-bidi-font-style:normal'>PAID</i></b><u><o:p></o:p></u></p>
		</td>
		<td width=427 valign=top style='width:4.45in;border-top:none;border-left:
		none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
		mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
		mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt'>
		<p class=MsoNormal>Check if your agency pays <b><i>any part</i></b> of “Flex” Pre-Tax Plan
		to full time employees.<o:p></o:p></p>
		<p class=MsoNormal><u><o:p><span style='text-decoration:none'>&nbsp;</span></o:p></u></p>
		</td>
		</tr>
		<tr style='mso-yfti-irow:3'>
		<td width=163 valign=top style='width:1.7in;border:solid windowtext 1.0pt;
		border-top:none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
		padding:0in 5.4pt 0in 5.4pt'>
		<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><i
		style='mso-bidi-font-style:normal'>% Premium paid by agency:</i></b><span
		style='mso-spacerun:yes'>   </span><u><o:p></o:p></u></p>
		</td>
		<td width=427 valign=top style='width:4.45in;border-top:none;border-left:
		none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
		mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
		mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt'>
		<p class=MsoNormal>Enter percentage of total monthly premium your agency pays
		for “Flex” Pre-Tax Plan for one full time employee.<o:p></o:p></p>
		<p class=MsoNormal><u><o:p><span style='text-decoration:none'>&nbsp;</span></o:p></u></p>
		</td>
		</tr>
		<tr style='mso-yfti-irow:4;mso-yfti-lastrow:yes'>
		<td width=163 valign=top style='width:1.7in;border:solid windowtext 1.0pt;
		border-top:none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
		padding:0in 5.4pt 0in 5.4pt'>
		<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><i
		style='mso-bidi-font-style:normal'>Total Monthly Premium:</i></b><span
		style='mso-spacerun:yes'>    </span><u><o:p></o:p></u></p>
		</td>
		<td width=427 valign=top style='width:4.45in;border-top:none;border-left:
		none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
		mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
		mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt'>
		<p class=MsoNormal>Enter dollar amount of total monthly premium for “Flex”
		Pre-Tax Plan for one full time employee. This amount includes both the agency
		and employee contribution.<o:p></o:p></p>
		<p class=MsoNormal><u><o:p><span style='text-decoration:none'>&nbsp;</span></o:p></u></p>
		</td>
		</tr>
		</table>

		<p class=MsoNormal>*<span style='mso-spacerun:yes'>   </span>The same terms and
						definitions apply for PART TIME employees section.</p>

	</span>
<% end if %>

<% if HelpID = "HealthClubFull" then %>	
	<p>
	<span class = "formIndex">Health Club</span>
	</p>

<span class = "formMain">
	<p>This is a part of NON-Medical benefits section where you need to provide details about Health Club membership assistance benefit in your agency.</p>
    <!--<p>If Health Club benefit is part of your medical benefits package and Health club premium cost can not be determined separetely, please select that this benefit is offered and paid, and leave percentages and premium amounts as 0`s.</p>--->
<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 width=550>
			<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
				<td width=150 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Field Name<o:p></o:p></b></p>
				</td>
				<td width=400 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Definition<o:p></o:p></b></p>
				</td>
			</tr>


        <tr style='mso-yfti-irow:1'>
			<td width=150 valign=top>
				<p class=MsoNormal>Offered</p>
			</td>
     <td width=400 valign=top>
		<p class=MsoNormal>Select "YES" if your agency offers Health Club membership to full time employees.<o:p></o:p></p>
		</td>
		</tr>	

         <!---<tr style='mso-yfti-irow:2'>
			<td width=150 valign=top>
				<p class=MsoNormal>PAID</p>
			</td>
   <td width=400 valign=top>
	  <p class=MsoNormal>Check if your agency pays <b><i>any part</i></b> of Health Club membership
		to full time employees.<o:p></o:p></p>
	  </td>
	  </tr>--->

     <tr style='mso-yfti-irow:3'>
		<td width=150 valign=top> 
		<p  class=MsoNormal>% Premium paid by agency</p>
		</td>
 <td width=400 valign=top>
	<p class=MsoNormal>Enter percentage of total monthly premium your agency pays
		for Health Club membership for one full time employee.<o:p></o:p></p>
	</td>
	</tr>
	
	<tr style='mso-yfti-irow:4'>
		<td width=150 valign=top> 
		<p  class=MsoNormal>Total Monthly Premium</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Enter dollar amount of total monthly premium for Health
		Club membership of one full time employee. This amount includes both the
		agency and employee contribution.<o:p></o:p></p>
				</td>
			</tr>
			</table>
	<span class = "formMain">
	
	<p class=MsoNormal>*&nbsp;The same terms and
						definitions apply for PART TIME employees section.</p>

	</span>
<% end if %>







<% if HelpID = "LifeInsuranceFull" then %>	
	<p>
	<span class = "formIndex">Life Insurance</span>
	</p>
	<span class = "formMain">
	<p>This is a part of NON-Medical benefits section where you need to provide details about Life Insurance benefit in your agency.</p>	

	<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 width=550>
	<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
				<td width=150 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Field Name<o:p></o:p></b></p>
				</td>
				<td width=400 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Definition<o:p></o:p></b></p>
				</td>
	</tr>

 <tr style='mso-yfti-irow:1'>
	  <td width=150 valign=top>
		  <p class=MsoNormal>Offered</p>
	  </td>
      <td width=400 valign=top>
		<p class=MsoNormal>Select "YES" if your agency offers Life Insurance to full time employees.<o:p></o:p></p>
		</td>
</tr>	

	
  <!---<tr style='mso-yfti-irow:2'>
	   <td width=150 valign=top>
		<p class=MsoNormal>PAID</p>
	   </td>
       <td width=400 valign=top>
    	  <p class=MsoNormal>Check if your agency pays <b><i>any part</i></b> of Life Insurance premium
		for full time employees.<o:p></o:p></p>
	  </td>
  </tr>	---->
	
	
		
 <tr style='mso-yfti-irow:3'>
		<td width=150 valign=top> 
		<p  class=MsoNormal>% Premium paid by agency</p>
		</td>
        <td width=400 valign=top>
	    <p class=MsoNormal>Enter percentage of total monthly premium your agency pays
		for Life Insurance for one full time employee. <o:p></o:p></p>
    	</td>
</tr>	
		
		
<tr style='mso-yfti-irow:4'>
		<td width=150 valign=top> 
		<p  class=MsoNormal>Total Monthly Premium</p>
		</td>	
		<td width=400 valign=top>
			<p class=MsoNormal>Enter dollar amount of total monthly premium for Life
		Insurance of one full time employee. This amount includes both the agency and
		employee contribution.
		<p>If monthly premium varies by employee, use the average premium.
		(Calculation of average premium = Total monthly premiums for the agency <u><b>for this benefit</B></u> divided by # of employees receiving <U><B>this benefit</B></U>.</p><o:p></o:p></p>
				</td>
</tr>		
</table>
	
		<p class=MsoNormal>*&nbsp;The same terms and
						definitions apply for PART TIME employees section.</p>

	</span>
<% end if %>

<% if HelpID = "TimeOffFull" then %>	
	<p>
	<span class = "formIndex">Annual Paid Time Off (Vacation, Floating Holidays, Personal Days)</span>
	</p>
	
	<span class = "formMain">
	<p>This is a part of NON-Medical benefits section where you need to provide details about Paid Personal Time Off benefit in your agency.</p>
  <p> For part time employee section, enter information based on <b>half time </b>employee.</P>
<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 width=550>
	<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
				<td width=150 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Field Name<o:p></o:p></b></p>
				</td>
				<td width=400 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Definition<o:p></o:p></b></p>
				</td>
	</tr>
		<tr style='mso-yfti-irow:1'>
	  <td width=150 valign=top>
		  <p class=MsoNormal>Offered</p>
	  </td>
      <td width=400 valign=top>
		<p class=MsoNormal>Select "YES" if your agency offers Annual Paid Time off for
		Vacation, Floating Holidays or Personal Days to full time employees.
		 <font color="ff000000">Please do it for both Exempt (Salaried) and Non-Exempt (Hourly) employees.</font><o:p></o:p></p>
		</td>
</tr>	
	
	<tr style='mso-yfti-irow:2'>
	   <td width=150 valign=top>
		<p class=MsoNormal># of Days for new employee</p>
	   </td>
       <td width=400 valign=top>
    	  <p class=MsoNormal>Enter the number of days per year that your agency offers as
		Paid Time Off to <font color="#ff000000">New</font> full time employees (sum of Vacation, Floating Holidays and Personal Days).<o:p></o:p></p>
	  </td>
	</tr>
	
	<tr style='mso-yfti-irow:3'>
	   <td width=150 valign=top>
		<p class=MsoNormal># of years before increase in # of days</p>
	   </td>
       <td width=400 valign=top>
    	  <p class=MsoNormal>Enter the number of years that it takes for new employee to work before your agency increases
		# of Paid Days Off for full time employees.<o:p></o:p></p>
	  </td>
	</tr>
	
	<tr style='mso-yfti-irow:4'>
	   <td width=150 valign=top>
		<p class=MsoNormal># of Days after increase</p>
	   </td>
       <td width=400 valign=top>
    	  <p class=MsoNormal>Enter the number of days per year that your agency offers as
		Paid Time Off to full time employees <font color="#ff000000">after increase</font> (sum of Vacation, Floating Holidays and Personal Days).<o:p></o:p></p>
	  </td>
	</tr>	
		
	</table>	
	
	<!--	<p class=MsoNormal>*&nbsp;The same terms and definitions apply for PART TIME employees section.</p>-->

	</span>
<% end if %>	
		
		
		
		
		

<% if HelpID = "TimeOffSickFull" then %>	
	<p>
	<span class = "formIndex">Paid Time Off (Sick Time)</span>
	</p>
	
	<span class = "formMain">
	<p>This is a part of NON-Medical benefits section where you need to provide details about Annual Paid Sick Time Off benefit in your agency.</p>
	<p> For part time employee section, enter information based on <b>half time </b>employee.</p>
  <table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 width=550>
	<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
				<td width=150 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Field Name<o:p></o:p></b></p>
				</td>
				<td width=400 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Definition<o:p></o:p></b></p>
				</td>
	</tr>
		
		
		<tr style='mso-yfti-irow:1'>
	  <td width=150 valign=top>
		  <p class=MsoNormal>Offered</p>
	  </td>
      <td width=400 valign=top>
		<p class=MsoNormal>Select "YES" if your agency offers Annual Paid Sick Time off to
		full time employees.<o:p></o:p></p>
		</td>
</tr>	
		
	<tr style='mso-yfti-irow:2'>
	   <td width=150 valign=top>
		<p class=MsoNormal>Days</p>
	   </td>
       <td width=400 valign=top>
    	  <p class=MsoNormal>Enter the number of days per year that your agency offers as
		annual paid sick time off to full time employees.For part time employee section, fill out how you handle paid sick time off for part time employee<o:p></o:p></p>
	  </td>
  </tr>	
</table>		
	<!--<p class=MsoNormal>*&nbsp;The same terms and definitions apply for PART TIME employees section.</p>-->
	
	</span>
<% end if %>

<% if HelpID = "TimeOffVacFull" then %>	
	<p>
	<span class = "formIndex">Paid Time Off (Vacation)</span>
	</p>
	
	<span class = "formMain">
	<p>This is a part of NON-Medical benefits section where you need to provide details about Paid Vacation Time Off benefit in your agency.</p>

	<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 width=550>
	<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
				<td width=150 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Field Name<o:p></o:p></b></p>
				</td>
				<td width=400 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Definition<o:p></o:p></b></p>
				</td>
	</tr>
		
		<tr style='mso-yfti-irow:1'>
	  <td width=150 valign=top>
		  <p class=MsoNormal>Offered</p>
	  </td>
      <td width=400 valign=top>
		<p class=MsoNormal>Select "YES" if your agency offers Annual Paid Vacation Time off
		to full time employees.<o:p></o:p></p>
		</td>
</tr>	
		
	<tr style='mso-yfti-irow:2'>
	   <td width=150 valign=top>
		<p class=MsoNormal>Days</p>
	   </td>
       <td width=400 valign=top>
    	  <p class=MsoNormal>Enter the number of days per year that your agency offers as
		annual paid vacation time off to full time employees.<o:p></o:p></p>
	  </td>
  </tr>	
</table>		
		
		<p class=MsoNormal>*&nbsp;The same terms and
						definitions apply for PART TIME employees section.</p>
	</span>
<% end if %>



<% 
if HelpID = "ProfDuesFull" then %>	
	<p>
	<span class = "formIndex">Professional Dues, Conferences, etc. (Average $ amount per month paid by agency per employee)</span>
	</p>
	
	<span class = "formMain">
	<p>This is a part of NON-Medical benefits section where you need to provide details about benefit that covers employee spendings on Conferences, professional dues and other realted expenses in your agency.</p>
	

		<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 width=550>
	<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
				<td width=150 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Field Name<o:p></o:p></b></p>
				</td>
				<td width=400 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Definition<o:p></o:p></b></p>
				</td>
	</tr>
		
		
        <tr style='mso-yfti-irow:1'>
				<td width=150 valign=top>
					<p class=MsoNormal>Offered</p>
				</td>
     <td width=400 valign=top>
					<p class=MsoNormal>Select "YES" if your agency offers attendance at Conferences
		to full time employees.<o:p></o:p></p>
				</td>
			</tr>	

		
	  <!--- <tr style='mso-yfti-irow:2'>
		<td width=150 valign=top>
		<p class=MsoNormal>PAID</p>
		</td>
   
     <td width=400 valign=top>
			<p class=MsoNormal>Check if your agency pays for attendance at
		Conferences to full time employees.<o:p></o:p></p>
				</td>
			</tr> --->


	   <tr style='mso-yfti-irow:3'>
		<td width=150 valign=top> 
		<p class=MsoNormal> DollarAmount</p>
		</td>

    <td width=400 valign=top>
			<p class=MsoNormal>Enter the average monthly dollar amount per employee that your
		agency pays for Professional Dues, Conferences to full time employees.<o:p></o:p></p>
				</td>
			</tr>

		</table>

		<p class=MsoNormal>*&nbsp;The same terms and
						definitions apply for PART TIME employees section.</p>
	</span>
<% end if %>

<% if HelpID = "RetirementFull" then %>	
	<p>
	<span class = "formIndex">Pension Plan</span>
	</p>
	
	<span class = "formMain">
	<p>This is a part of NON-Medical benefits section where you need to provide details about Pension Plan benefit in your agency.</p>

		<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 width=550>
			<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
				<td width=150 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Field Name<o:p></o:p></b></p>
				</td>
				<td width=400 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Definition<o:p></o:p></b></p>
				</td>
			</tr>

		        <tr style='mso-yfti-irow:1'>
				<td width=150 valign=top>
					<p class=MsoNormal>Offered</p>
				</td>
     <td width=400 valign=top>
					<p class=MsoNormal>Select "YES" if your agency offers a Pension Plan to full time
		employees.<o:p></o:p></p>
				</td>
			</tr>	

     <tr style='mso-yfti-irow:2'>
		<td width=150 valign=top> 
		<p  class=MsoNormal> % of matching contribution</p>
		</td>

    <td width=400 valign=top>
			<p class=MsoNormal>Enter the percentage of employee salary that your agency contributes to
		pension plan for full time employee.<o:p></o:p></p>
				</td>
			</tr>
     </table>

		<p class=MsoNormal>*&nbsp;The same terms and
						definitions apply for PART TIME employees section.</p>
	</span>
<% end if %>

<% if HelpID = "403BFull" then %>	
	<p>
	<span class = "formIndex">403 B</span>
	</p>
	
	<span class = "formMain">
	<p>This is a part of NON-Medical benefits section where you need to provide details about 403B benefit in your agency.</p>

		<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 width=550>
			<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
				<td width=150 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Field Name<o:p></o:p></b></p>
				</td>
				<td width=400 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Definition<o:p></o:p></b></p>
				</td>
			</tr>
			<tr style='mso-yfti-irow:1'>
				<td width=150 valign=top>
					<p class=MsoNormal>Offered</p>
				</td>
				<td width=400 valign=top>
					<p class=MsoNormal>Select "YES" if your agency offers 403(B) to full time employees.<o:p></o:p></p>
				</td>
			</tr>
			<tr style='mso-yfti-irow:2'>
				<td width=150 valign=top>
					<p class=MsoNormal>Employer Contribution (Y/N)</p>
				</td>
				<td width=400 valign=top>
					<p class=MsoNormal>Select Yes if your agency (employer) matches employee contribution to 403(B).<o:p></o:p></p>
				</td>
			</tr>
			<tr style='mso-yfti-irow:3;mso-yfti-lastrow:yes'>
				<td width=150 valign=top>
					<p class=MsoNormal>% of Matching Contribution</p>
				</td>
				<td width=400 valign=top>
					<p class=MsoNormal>Enter the % of matching contribution your agency makes to the 403B plan.<o:p></o:p></p>
				</td>
			</tr>
		</table>

		<p class=MsoNormal>*&nbsp;The same terms and
						definitions apply for PART TIME employees section.</p>
	</span>
<% end if %>

<% if HelpID = "TelecommFull" then %>	
	<p>
	<span class = "formIndex">Telecommuting (employees who work from home on an ongoing basis)</span>
	</p>
	
	<span class = "formMain">
	<p>This is a part of NON-Medical benefits section where you need to provide details about Telecommuting benefit in your agency.</p>

		<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 width=550 ID="Table2">
			<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
				<td width=150 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Field Name<o:p></o:p></b></p>
				</td>
				<td width=400 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Definition<o:p></o:p></b></p>
				</td>
			</tr>
			<tr style='mso-yfti-irow:1'>
				<td width=150 valign=top>
					<p class=MsoNormal>Offered</p>
				</td>
				<td width=400 valign=top>
					<p class=MsoNormal>Select "YES" if your agency offers Telecommuting to full time employees.<o:p></o:p></p>
				</td>
			</tr>
			<tr style='mso-yfti-irow:2'>
				<td width=150 valign=top>
					<p class=MsoNormal>Number of Employees</p>
				</td>
				<td width=400 valign=top>
					<p class=MsoNormal>Enter the number of employees who telecommute in your agency.<o:p></o:p></p>
				</td>
			</tr>
			<tr style='mso-yfti-irow:3;mso-yfti-lastrow:yes'>
				<td width=150 valign=top>
					<p class=MsoNormal>% of Employee Population</p>
				</td>
				<td width=400 valign=top>
					<p class=MsoNormal>Enter the percentage of total employee population who telecommute.<o:p></o:p></p>
				</td>
			</tr>
		</table>
		<p class=MsoNormal>*&nbsp;The same terms and
						definitions apply for PART TIME employees section.</p>
	</span>
<% end if %>

<% if HelpID = "TuitionFull" then %>	
	<p>
	<span class = "formIndex">Tuition Reimbursement</span>
	</p>
	
	<span class = "formMain">
	<p>This is a part of NON-Medical benefits section where you need to provide details about Tuition Reimbursement benefit in your agency.</p>

		<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 width=550>
			<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
				<td width=150 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Field Name<o:p></o:p></b></p>
				</td>
				<td width=400 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Definition<o:p></o:p></b></p>
				</td>
			</tr>

		        <tr style='mso-yfti-irow:1'>
				<td width=150 valign=top>
					<p class=MsoNormal>Offered</p>
				</td>
     <td width=400 valign=top>
					<p class=MsoNormal>Select "YES" if your agency offers tuition reimbursement to full
		time employees.<o:p></o:p></p>
				</td>
			</tr>	

    <tr style='mso-yfti-irow:2'>
		<td width=150 valign=top>
		<p class=MsoNormal>Maximum $ paid</p>
		</td>
   
     <td width=400 valign=top>
			<p class=MsoNormal>Enter the maximum dollar amount per year that your agency can
		pay for tuition reimbursement per full time employee.<o:p></o:p></p>
				</td>
			</tr>

		</table>

		<p class=MsoNormal>*&nbsp;The same terms and
						definitions apply for PART TIME employees section.</p>
	</span>
<% end if %>


<% if HelpID = "Professional Development Budget" then %>	
	<p>
	<span class = "formIndex">Professional Development Budget per employee</span>
	</p>
	
	<span class = "formMain">
	<p>This is a part of NON-Medical benefits section where you need to provide details about Professional Development Budget benefit in your agency.</p>

		<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 width=550>
			<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
				<td width=150 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Field Name<o:p></o:p></b></p>
				</td>
				<td width=400 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Definition<o:p></o:p></b></p>
				</td>
			</tr>

		        <tr style='mso-yfti-irow:1'>
				<td width=150 valign=top>
					<p class=MsoNormal>Offered</p>
				</td>
     <td width=400 valign=top>
					<p class=MsoNormal>Select "YES" if your agency has a  Professional Development Budget for full
		time employees.<o:p></o:p></p>
				</td>
			</tr>	

    <tr style='mso-yfti-irow:2'>
		<td width=150 valign=top>
		<p class=MsoNormal>Maximum $ paid</p>
		</td>
   
     <td width=400 valign=top>
			<p class=MsoNormal>Enter the maximum dollar amount per year that your agency can
		pay for Professional Development Budget per full time employee.<o:p></o:p></p>
				</td>
			</tr>

		</table>

		<p class=MsoNormal>*&nbsp;The same terms and
						definitions apply for PART TIME employees section.</p>
	</span>
<% end if %>














<% if HelpID = "BirthYear" then %>	
	<p>
	<span class = "formIndex"></span>
	</p>
	
	<span class = "formMain">
	<p>Enter the revenue from federal grants either to the agency as a direct recipient or as a sub-recipient.
	</p>	
	</span>
<% end if %>


<% if HelpID = "GoverningBoardMembers" then %>	
	<p>
	<span class = "formIndex"></span>
	</p>
	
	<span class = "formMain">
	<p>Agencies should report only governing board members.
    </p>	
	</span>
<% end if %>


<% if HelpID = "Board100donate" then %>	
	<p>
	<span class = "formIndex"></span>
	</p>
	
	<span class = "formMain">
	<p>If you have a policy of 100% board donating, but do NOT have minimum donation amount - enter zero (0) as minimum donation amount.
	</p>	
	</span>
<% end if %>

<p>
<div align="center"><A HREF="javascript:window.close()"><img src="close.gif" alt="" width="50" height="17" border="0"></a></div>
</p>
</body>
</html>
