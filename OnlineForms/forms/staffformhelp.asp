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


<% if HelpID="BirthYear" then %>
    <p>
	<span class = "formIndex">Birth Year</span>
	</p>
	<span class = "formMain">
	<p>
	Enter the year when the employee was born.
	</p>
	</span>
<% end if %>	


<% if HelpID="EverABig" then %>
    <p>
	<span class = "formIndex">Ever A BIG</span>
	</p>
	<span class = "formMain">
	<p>
	Select "YES" if the employee was ever a big.
	</p>
	</span>
<% end if %>	


<% if HelpID="Position" then %>
    <p>
	<span class = "formIndex">Position </span>
	</p>
	<span class = "formMain">
	<p>Select employee position that He/She holds in the organization.</p>

		 <table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 width=550>
			<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
				<td width=150 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Position<o:p></o:p></b></p>
				</td>
				<td width=400 valign=top>
					<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>Job Capsules<o:p></o:p></b></p>
				</td>
			</tr>
        <tr style='mso-yfti-irow:1'>
				<td width=150 valign=top>
				<p class=MsoNormal>1) Executive Director/CEO/President</p>
				</td>
     <td width=400 valign=top>
					<p class=MsoNormal>Leads agency,board,staff and volunteers in ensuring positive outcomes for children.<u><o:p></o:p></u></p>
				</td>
			</tr>	

    <tr style='mso-yfti-irow:2'>
		<td width=150 valign=top>
	<p class=MsoNormal>2) COO </p>
		</td>
   
     <td width=400 valign=top>
			<p class=MsoNormal>Leads multiple departments in developing strategy, and ensuring operational excellence. <o:p></o:p></p>
				</td>
			</tr>
			
    <tr style='mso-yfti-irow:3'>
		<td width=150 valign=top>
	<p class=MsoNormal>3) Vice President/Associate Director </p>
		</td>
   
     <td width=400 valign=top>
			<p class=MsoNormal>Leads activities of one or more departmetns of the agency. Ensure department achieves goals and budget. <o:p></o:p></p>
				</td>
			</tr>

 <tr style='mso-yfti-irow:3'>
		<td width=150 valign=top>
	<p class=MsoNormal>4) Director, Remote Point of Service</p>
		</td>
   
     <td width=400 valign=top>
			<p class=MsoNormal>Directs and coordinates activities of remote based agency site. Assists in formulating and administering organization policies, developing long range goals and objectives in conjunction with organization`s vision, mission & strategic plan. <o:p></o:p></p>
				</td>
			</tr>



      <tr style='mso-yfti-irow:4'>
		<td width=150 valign=top> 
	<p  class=MsoNormal>7) Director: Sponsored Organization</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Directs and coordinates activities of BBBS programs. 
			                   Develops and administers organization policies, develops long 
			                   range goals and objectives in conjunction with BBBS vision, 
			                   mission & strategic plan.<o:p></o:p></p>
				</td>
			</tr>

<tr style='mso-yfti-irow:5'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 20) Fund Development:  Vice President/Director</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Directs and manages the implementation of contributed income programs on behalf 
			                   of the organization`s goals and objectives.  Maximizes funding potential and increases
			                   base of donor support.<o:p></o:p></p>
				</td>
			</tr>


<tr style='mso-yfti-irow:6'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 22) Events:  Director/Coordinator</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Develops, executes and evaluates strategic fund raising events designed 
			                   to build and communicate the essence of the BBBSA brand.
			                   Provides leadership and direction to BBBS local agencies and 
			                   monitors event satisfaction through feedback, data collection and analysis.<o:p></o:p></p>
				</td>
			</tr>

		
 <tr style='mso-yfti-irow:7'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 23) Individual Giving:  Director/Coordinator</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Directs and manages individual major gifts and planned gifts,
			                   as well as annual giving and membership programs. Responsible for developing 
			                   strategic fund-raising plans, to acquire and sustain major gifts.<o:p></o:p></p>
				</td>
			</tr>
			
	<tr style='mso-yfti-irow:8'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 24) Grants:  Director/Coordinator</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Directs and coordinates the preparation of reports and grant applications and 
			                researches outside agencies and foundations for grant opportunities. 
                            Identifies federal, state, and private funding sources.<o:p></o:p></p>
				</td>
			</tr>

	<tr style='mso-yfti-irow:9'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 25) Major Gifts:  Director/Coordinator</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Develops and directs a programmatic approach to the identification of potential individual major prospects, the development and implementation of appropriate donor cultivation and solicitation strategies.<o:p></o:p></p>
				</td>
			</tr>

<tr style='mso-yfti-irow:10'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 10) Program:  Vice President/Director</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Directs and coordinates activities of the Program department
			                   of the agency. Leads program department staff in creating strong 
			                   results based culture that uses key metrics to measure success. <o:p></o:p></p>
				</td>
			</tr>


<tr style='mso-yfti-irow:11'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 11) Program (all functions): Supervisor/Manager Formerly Casework Supervisor</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Directs and manages the activities of one or more of the Program Department
			                   Service Delivery Model functions (Customer Relations, Enrollment & Matching & 
			                   Match Support), Including Remote locations. <o:p></o:p></p>
				</td>
			</tr>
						

<tr style='mso-yfti-irow:12'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 12) Program (all functions): Staff Formerly Caseworker </p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Provides direct service to volunteers, children and families 
			                   through the execution of the Service Delivery model in any or all of 
			                   the three functions.Utilizes BBBS standards and practices.<o:p></o:p></p>
				</td>
			</tr>
			
			
<tr style='mso-yfti-irow:13'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 14) Customer Relations: Staff </p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Provides high-level customer service in response to all customer and stakeholder 
			                   inquiries and leads.  Additionally responsible for marketing BBBS programs
			                    through telemarketing recruitment and outreach.<o:p></o:p></p>
				</td>
			</tr>			


<tr style='mso-yfti-irow:14'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 15) Enrollment & Matching: Staff </p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Ensures that volunteers and children are appropriately enrolled and matched
			                   while executing a high degree of independent judgment when utilizing BBBS 
			                   standards and practices.   Responsible for focusing on volunteer options and 
			                   child safety throughout enrollment and matching process.<o:p></o:p></p>
				</td>
			</tr>	


<tr style='mso-yfti-irow:15'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 16) Match Support: Staff </p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Provides match support to ensure child safety, positive impacts for youth, 
			                   constructive and satisfying relationships between children and volunteers, 
			                   and a strong sense of affiliation with BBBS on the part of volunteers.<o:p></o:p></p>
				</td>
			</tr>

<tr style='mso-yfti-irow:16'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 17) Customer Relations: Supervisor </p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Directs and manages the SDM Customer Relations program staff.  Leads program staff using key Customer Relations metrics to measure success.<o:p></o:p></p>
				</td>
			</tr>

<tr style='mso-yfti-irow:17'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 18) Enrollment & Matching: Supervisor </p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Directs and manages the SDM Enrollment & Matching program staff.  Leads program staff using key Enrollment metrics to measure success.<o:p></o:p></p>
				</td>
			</tr>

<tr style='mso-yfti-irow:18'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 19) Match Support: Supervisor </p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Directs and manages the SDM Match Support program staff.  Leads program staff using key Match Support metrics to measure success.<o:p></o:p></p>
				</td>
			</tr>

<tr style='mso-yfti-irow:19'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 80) Child Safety/QA Staff </p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Ensures that children are safe and our programs are of the highest quality
			                   possible through a standardized model of one-to-one service delivery. 
			                   Works toward constant improvement in the area of child safety and quality 
			                   assurance through file audits, performance management practices and data management.<o:p></o:p></p>
				</td>
			</tr>

<tr style='mso-yfti-irow:20'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 81) Partnerships:  Director/Coordinator</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Develops, manages and enhances agency`s corporate partnerships building on strong 
			                   brand awareness. Leads corporate identification, cultivation and solicitation; 
			                   plays a major role in the development of agency`s fundraising and awareness building.<o:p></o:p></p>
				</td>
			</tr>
			
     <tr style='mso-yfti-irow:21'>
		<td width=150 valign=top> 
		<p  class=MsoNormal>5) Office: Director/Manager</p>
		</td>

    <td width=400 valign=top>
			<p class=MsoNormal>Directs and manages the organization and supervision all of the administrative activities that facilitate 
			                   the running of the office. Primarily resposible for ensuring that the agency office runs efficiently <o:p></o:p></p>
				</td>
			</tr>
	
	<tr style='mso-yfti-irow:22'>
		<td width=150 valign=top> 
	<p  class=MsoNormal>6) Office: Staff Example: Administrative Assistant, Receptionist </p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Performs administrative and office support activities.Resposible for fielding
			                   telephone calls,receiving and directing visitors, word processing, filing and faxing.<o:p></o:p></p>
				</td>
			</tr>
	
<tr style='mso-yfti-irow:23'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 7) Intern (Paid)</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>N/A<o:p></o:p></p>
				</td>
			</tr>			
			
	<tr style='mso-yfti-irow:24'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 30) Recruitment:  Supervisor/Manager</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Directs and manages the process for identifying the most likely sources of 
			                   suitable candidates for volunteer positions, how to approach those sources,
			                    and then approaching each source. Includes responsibility for identifying
			                     children for participation in the program.<o:p></o:p></p>
				</td>
			</tr>

	<tr style='mso-yfti-irow:25'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 32) Recruitment:  Staff</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Identifies sources for volunteers and children for participation in the program. Coordinates with program staff to identify areas of need.<o:p></o:p></p>
				</td>
			</tr>

<tr style='mso-yfti-irow:26'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 40) Marketing & Communications: Vice President/Director</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Directs and coordinates a broad-based marketing effort aimed at significantly 
			                  increasing agency`s donor and volunteer support. Serves as a high-level strategic
			                  partner responsible for all aspects of brand development and marketing at the agency 
			                  level and through local market affiliates.<o:p></o:p></p>
				</td>
			</tr>

<tr style='mso-yfti-irow:27'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 41) Marketing & Communications: Supervisor/Manager</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Directs and coordinates the articulation of agency messages and value propositions across all media content.  Acts as writer/editor of agency`s M & C written communications, including: magazines; newsletter; brochures; information sheets, position statements, and annual reports.<o:p></o:p></p>
				</td>
			</tr>	
			
	<tr style='mso-yfti-irow:28'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 42) Marketing & Communications: Staff</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Responsible for editorial and design production and distribution of all agency publications.
                               Coordinates media interest in the agency and ensures regular contact with target media 
                               and appropriate response to media requests; as well as the appearance of all agency 
                               print and electronic materials such as letterhead, use of logo, brochures, etc.<o:p></o:p></p>
				</td>
			</tr>	
		
		<tr style='mso-yfti-irow:29'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 50) Information Technology: Vice President/Director</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Directs and coordinates the development and implementation of responsive,
			                   cost-effective and efficient computer systems and operations to meet the needs
			                    of the agency and support the achievement of strategic business goals.  <o:p></o:p></p>
				</td>
			</tr>		
			
		<tr style='mso-yfti-irow:30'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 51) Information Technology: Supervisor/Manager</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Directs and manages all aspects of the procurement, maintenance, development, enhancement and expansion of the agency computer, network and Internet capabilities; to perform project management and supervision of computer staff.<o:p></o:p></p>
				</td>
			</tr>
			
		<tr style='mso-yfti-irow:31'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 52) Information Technology: Staff</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Provides technical support for agency network, web sites, web applications, 
			                   digital media, etc. Develops processes and systems that support agency`s IT systems
			                    and other applicable applications.<o:p></o:p></p>
				</td>
			</tr>
		
		<tr style='mso-yfti-irow:32'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 60) Finance:  Vice President/Director</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Directs and leads the agency`s financial planning and accounting practices,
			                   relationships with lending institutions and the financial community.
			                   Ensures legal and regulatory compliance, oversees cost and general accounting functions,
			                   insurance activities and building operations.<o:p></o:p></p>
				</td>
			</tr>		

		<tr style='mso-yfti-irow:33'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 61) Finance:  Supervisor/Manager</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Manages the agency`s fiscal functions in accordance with generally accepted accounting principles and with a high degree of attention to detail and ethical standards.   Provides proactive checks and balances pertaining to all expenditures and expense requests, resolving all matters in conflict with agency policy or practice.<o:p></o:p></p>
				</td>
			</tr>
		
			<tr style='mso-yfti-irow:34'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 62) Finance:  Staff</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Provides application of agency accounting principles, analyzes and summarizes agency 
			                   financial information.  Utilizes and understands the financial internal control 
			                   environment.<o:p></o:p></p>
				</td>
			</tr>		
		
		
		
				<tr style='mso-yfti-irow:35'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 70) Human Resources: Vice President/Director</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Directs and manages the overall strategic HR leadership to the agency. 
			                   Responsible for leading the development and implementation of human resources policies,
			                   programs, and practices for the agency: including employment, labor relations, 
			                   compensation, learning and development, organizational development, 
			                   and employee services.<o:p></o:p></p>
				</td>
			</tr>

	<tr style='mso-yfti-irow:36'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 71) Human Resources: Supervisor/Manager</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Responsible for the basic process of recruitment, hiring, training, promotion, and retention of the agency`s employees.  Manages the day to day HR issues such as; ensuring compliance of hiring regulations; advising managers on counseling employees regarding performance issues; liaising with other support functions such as Payroll and advising staff on HR policy.<o:p></o:p></p>
				</td>
			</tr>

	<tr style='mso-yfti-irow:37'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 72) Human Resources: Staff</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Responsible for the recruitment and orientation of diverse, talented and skilled staff for agency positions.  Promotes human resource function throughout the agency.<o:p></o:p></p>
				</td>
			</tr>

	<tr style='mso-yfti-irow:37'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 73) Trainer</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Responsible for the Training with in the agency.<o:p></o:p></p>
				</td>
			</tr>
<tr style='mso-yfti-irow:38'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 90) Americorps /Vista</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>N/A<o:p></o:p></p>
				</td>
			</tr>	
		
	<tr style='mso-yfti-irow:38'>
		<td width=150 valign=top> 
	<p  class=MsoNormal> 100) Volunteer</p>
		</td>	
		
	<td width=400 valign=top>
			<p class=MsoNormal>Person who Volunteers at the agency<o:p></o:p></p>
				</td>
			</tr>	
		</table>
		
	</span>
<% end if %>




<% if HelpID="Race" then %>

	<p>
	<span class = "formIndex">Ethnicity</span>
	</p>
	
	<span class = "formMain">
	<p>
	Select ethnicity of the employee (White, Hispanic, Asian, etc.).
	</p>
	
	</span>
	
<% end if %>



<% if HelpID="Sex" then %>

	<p>
	<span class = "formIndex">Gender</span>
	</p>
	
	<span class = "formMain">
	<p>
	Select employee gender (Male/Female).
	</p>
	
	</span>
	
<% end if %>

	

<% if HelpID="Education" then %>

	<p>
	<span class = "formIndex">Education</span>
	</p>
	
	<span class = "formMain">
	<p>
	Select a level of education of the employee. Ex: Bachelors Degree, Graduate Degree.
	</p>
	
	</span>
	
<% end if %>



<% if HelpID="MonthStart" then %>

	<p>
	<span class = "formIndex">Month Hired @ Agency</span>
	</p>
	
	<span class = "formMain">
	<p>
	"MONTH" when employee was hired. e.g. August
	</p>
	
	</span>
	
<% end if %>

	

<% if HelpID="YearStart" then %>

	<p>
	<span class = "formIndex">Year Hired @ Agency</span>
	</p>
	
	<span class = "formMain">
	<p>
	"YEAR" when employee was hired. e.g. 2007
	</p>
	
	</span>
	
<% end if %>


<% if HelpID="YearsInNetwork" then %>

	<p>
	<span class = "formIndex">Years of Service in the BBBS Network</span>
	</p>
	
	<span class = "formMain">
	<p>
		The total number of years that the employee worked at BBBS network.<br><br>
		(e.g. If a staff member has worked in one agency for 2 years and in a 
		different agency for 1 year, Total years of service would be 2+1 = 3.
		&nbsp;So the employee has worked in BBBS network for 3 Years).
	</p>
	
	 <p> If an employee has worked in BBBS network for less than a year please select 0</p>
	</span>
	
<% end if %>

<% if HelpID="Time" then %>

	<p>
	<span class = "formIndex">Employment Status </span>
	</p>
	
	<span class = "formMain">
	<p> If employee is no longer employeed with the agency. Select the "Month" Employee was terminated.<br>
	    (Ex: Employee left the agency in March. Select "March" month in the drop down menu)</p>
	<P> If employee is still employeed with the agency,select "Still Employeed" in the drop down menu</p>
	</span>
	
<% end if %>

<% if HelpID="PositionStartDate" then %>

	<p>
	<span class = "formIndex">Date Started Current Position</span>
	</p>
	
	<span class = "formMain">
	<p>
		Date the employee started his or her current position at this agency.<br><br>
		Ex: 02/25/2007 if employee started on February 25th, 2007.
	</p>
	
	</span>
	
<% end if %>


<% if HelpID="Terminated" then %>

	<p>
	<span class = "formIndex">Terminated</span>
	</p>
	
	<span class = "formMain">
	<p>
	Month when employment ended (e.g. March)<br><br>
    Enter  month only if the employee ended employment during current AAI year being reported.<br>
    If the employee  did not leave the agency, select "Still Employed" option.
   </p>
	
	</span>
	
<% end if %>

<% if HelpID="HoursWeek" then %>

	<p>
	<span class = "formIndex">Hours Per Week</span>
	</p>
	
	<span class = "formMain">
	<p>
	Number of hours per week the employee works. e.g. 40, 30, 20
	</p>
	
	</span>
	
<% end if %>

<% if HelpID="EmployeeName" then %>

	<p>
	<span class = "formIndex">Employee Name</span>
	</p>
	
	<span class = "formMain">
	<p>
	Enter employee Full Name. First, space, Last name. Ex: John Smith.
	</p>
	
	</span>
	
<% end if %>

<% if HelpID="YearlySalary" then %>

	<p>
	<span class = "formIndex">Compensation(Salary + Bonus/Incentives)</span>
	</p>
	
	<span class = "formMain">
	<p>
		<b>Base Salary:</b>  Base salary of the employee. Yearly gross salary.<br>
		Note: Do not enter commas for thousand separation.
	</p>
	<p>
		<b>Bonus/Incentives:</b>  Bonus or incentives on top of base salary.<br>
		Note: Do not enter commas for thousand separation.
    </p>
    <p>
		<b>Total:</b>  This field is calculated automatically by addition of Bonus to the Base Salary.</p>
</span>
	
<% end if %>


<% if HelpID="Percent of your Board Members are Donationg to your Agency" then %>
    <p>
	<span class = "formIndex">Percent of your Board Members are Donationg to your Agency</span>
	</p>
	<span class = "formMain">
	<p>
	Please enter the percentage of board members that are donationg to your agency. (Ex. 3% of board members are donating to your agency)
	</p>
	</span>
<% end if %>	



<% if HelpID="Percent of your Board has connected the agency" then %>
    <p>
	<span class = "formIndex">Percent of your Board has connected the agency </span>
	</p>
	<span class = "formMain">
	<p>
	Please enter the percentage of board members that have connected the  
 agency to potential Corporate and Individual Donors.
 (Ex. 5% of board members that have connected the agency to potential 
   Corporate and Individual Donors.)

	</p>
	</span>
<% end if %>	


<% if HelpID="Average Donation by Board Member" then %>
    <p>
	<span class = "formIndex">Average Donation by Board Member</span>
	</p>
	<span class = "formMain">
	<p>
	Please enter the average $ amount donated by Board member of your agency. (Ex. $300. The donation Amount must be individual giving).

	</p>
	</span>
<% end if %>	








<p>
<div align="center"><A HREF="javascript:window.close()"><img src="close.gif" alt="" width="50" height="17" border="0"></a></div>
</p>
</body>
</html>