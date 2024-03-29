/* Microsoft SQL Server - Scripting			*/
/* Server: IMC2SQL65					*/
/* Database: BBBSA					*/
/* Creation Date 2/6/01 8:01:32 PM 			*/

if exists (select * from sysobjects where id = object_id(N'[dbo].[tbl_frmBoardMembers]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_frmBoardMembers]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[tbl_frmExpenses]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_frmExpenses]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[tbl_frmGeneralInformation]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_frmGeneralInformation]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[tbl_frmIncome]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_frmIncome]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[tbl_frmPerformance]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_frmPerformance]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[tbl_frmSpecialPopulations]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_frmSpecialPopulations]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[tbl_frmSpecialPrograms]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_frmSpecialPrograms]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[tbl_frmStaff]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_frmStaff]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[tbl_ModifyLog]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_ModifyLog]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[tbl_StaffEducation]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_StaffEducation]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[tbl_StaffPosition]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_StaffPosition]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[tbl_StaffRace]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_StaffRace]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[tbl_UserLogins]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_UserLogins]
GO

CREATE TABLE [dbo].[tbl_frmBoardMembers] (
	[BoardMembersID] [int] IDENTITY (1, 1) NOT NULL ,
	[AgencyID] [varchar] (50) NULL ,
	[Year] [int] NULL ,
	[NumberBoardMembers] [int] NULL ,
	[TermLimits] [bit] NOT NULL ,
	[TermLimitsYears] [int] NULL ,
	[AverageTenureYears] [int] NULL ,
	[AverageTenureMonths] [int] NULL ,
	[StandingCommitteesPersonnel] [bit] NOT NULL ,
	[StandingCommitteesProgram] [bit] NOT NULL ,
	[StandingCommitteesExecutive] [bit] NOT NULL ,
	[StandingCommitteesFundDevelopment] [bit] NOT NULL ,
	[StandingCommitteesFinance] [bit] NOT NULL ,
	[StandingCommitteesPublicRelations] [bit] NOT NULL ,
	[StandingCommitteesStrategicPlanning] [bit] NOT NULL ,
	[StandingCommitteesBoardDevelopment] [bit] NOT NULL ,
	[StandingCommitteesVolunteerRecruitment] [bit] NOT NULL ,
	[StandingCommitteesOther] [bit] NOT NULL ,
	[StandingCommitteesOtherText] [varchar] (100) NULL ,
	[FemaleWhite] [int] NULL ,
	[FemaleBlack] [int] NULL ,
	[FemaleHispanic] [int] NULL ,
	[FemaleAsian] [int] NULL ,
	[FemaleIslander] [int] NULL ,
	[FemaleNative] [int] NULL ,
	[FemaleMulti] [int] NULL ,
	[FemaleUnknown] [int] NULL ,
	[MaleWhite] [int] NULL ,
	[MaleBlack] [int] NULL ,
	[MaleHispanic] [int] NULL ,
	[MaleAsian] [int] NULL ,
	[MaleIslander] [int] NULL ,
	[MaleNative] [int] NULL ,
	[MaleMulti] [int] NULL ,
	[MaleUnknown] [int] NULL ,
	[FrequencyMonthly] [bit] NOT NULL ,
	[FrequencyTwoMonths] [bit] NOT NULL ,
	[FrequencyQuarterly] [bit] NOT NULL ,
	[FrequencyOther] [bit] NOT NULL ,
	[FrequencyOtherText] [varchar] (100) NULL ,
	[MoneyMinimum] [bit] NOT NULL ,
	[MoneyInKind] [bit] NOT NULL ,
	[MoneyNotExpected] [bit] NOT NULL ,
	[MoneyNoPolicy] [bit] NOT NULL ,
	[MoneyMinimumAmount] [money] NULL ,
	[YearlyContribution] [money] NULL ,
	[SkillsFinance] [int] NULL ,
	[SkillsLegal] [int] NULL ,
	[SkillsPublicRelations] [int] NULL ,
	[SkillsHumanServicesPractitioner] [int] NULL ,
	[SkillsHumanServicesAdministrator] [int] NULL ,
	[SkillsFullTimeStudent] [int] NULL ,
	[SkillsHumanResources] [int] NULL ,
	[SkillsCorporateCEO] [int] NULL ,
	[SkillsOtherCorporateOfficer] [int] NULL ,
	[SkillsInsurance] [int] NULL ,
	[SkillsSmallBusiness] [int] NULL ,
	[SkillsBig] [int] NULL ,
	[SkillsParentLittle] [int] NULL ,
	[SkillsLittle] [int] NULL ,
	[SkillsLocalGovernment] [int] NULL ,
	[SkillsOther] [int] NULL ,
	[SkillsUnknown] [int] NULL ,
	[CreateDate] [datetime] NULL ,
	[ModifyLogID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_frmExpenses] (
	[ExpensesID] [int] IDENTITY (1, 1) NOT NULL ,
	[AgencyID] [varchar] (50) NULL ,
	[Year] [int] NULL ,
	[SalariesWages] [money] NULL ,
	[EmployeeBenefits] [money] NULL ,
	[ProfessionalFees] [money] NULL ,
	[RentOccupancy] [money] NULL ,
	[Supplies] [money] NULL ,
	[Travel] [money] NULL ,
	[Other] [money] NULL ,
	[Total] [money] NULL ,
	[Administration] [varchar] (10) NULL ,
	[Program] [varchar] (10) NULL ,
	[FundRaising] [varchar] (10) NULL ,
	[CreateDate] [datetime] NULL ,
	[ModifyLogID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_frmGeneralInformation] (
	[GeneralInformationID] [int] IDENTITY (1, 1) NOT NULL ,
	[AgencyID] [varchar] (50) NULL ,
	[Year] [int] NULL ,
	[PopulationSCA] [int] NULL ,
	[SchoolAgeSCA] [int] NULL ,
	[VolunteerInquiries] [int] NULL ,
	[VolunteerApplications] [int] NULL ,
	[VolunteersAccepted] [int] NULL ,
	[StrategicGrowthPlan] [bit] NOT NULL ,
	[ChildrenBy2004] [int] NULL ,
	[SexualPreventionCurriculum] [bit] NOT NULL ,
	[TrainingMentoringOrganizations] [bit] NOT NULL ,
	[TrainingPostMatch] [bit] NOT NULL ,
	[AfterSchoolMentoringProgram] [bit] NOT NULL ,
	[ASMPHowManyChildren] [int] NULL ,
	[CreateDate] [datetime] NULL ,
	[ModifyLogID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_frmIncome] (
	[IncomeID] [int] IDENTITY (1, 1) NOT NULL ,
	[AgencyID] [varchar] (50) NULL ,
	[Year] [int] NULL ,
	[UnitedWay] [money] NULL ,
	[FederalGovernmentFunding] [money] NULL ,
	[StateGovernmentFunding] [money] NULL ,
	[LocalGovernmentFunding] [money] NULL ,
	[FoundationGrants] [money] NULL ,
	[CorporateGifts] [money] NULL ,
	[BowlForKidsSake] [money] NULL ,
	[BFKSBowlingCenters] [int] NULL ,
	[CarsForKidsSake] [money] NULL ,
	[SpecialEvent1] [money] NULL ,
	[SE1Name] [varchar] (50) NULL ,
	[SpecialEvent2] [money] NULL ,
	[SE2Name] [varchar] (50) NULL ,
	[SpecialEvent3] [money] NULL ,
	[SE3Name] [varchar] (50) NULL ,
	[SpecialEvent4] [money] NULL ,
	[SE4Name] [varchar] (50) NULL ,
	[SpecialEvent5] [money] NULL ,
	[SE5Name] [varchar] (50) NULL ,
	[IndividualGiving] [money] NULL ,
	[DividendsInterest] [money] NULL ,
	[OtherFunding] [money] NULL ,
	[OtherFundingType] [varchar] (50) NULL ,
	[Total] [money] NULL ,
	[CreateDate] [datetime] NULL ,
	[ModifyLogID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_frmPerformance] (
	[PerformanceID] [int] IDENTITY (1, 1) NOT NULL ,
	[AgencyID] [varchar] (50) NULL ,
	[Year] [int] NULL ,
	[Month] [int] NULL ,
	[OpenMatchesCommunityBased] [int] NULL ,
	[OpenMatchesSchoolBased] [int] NULL ,
	[OpenMatchesOtherSiteBased] [int] NULL ,
	[OpenMatchesGroupMentoring] [int] NULL ,
	[OpenMatchesSpecialProgramsMentoring] [int] NULL ,
	[OpenMatchesSpecialProgramsNonMentoring] [int] NULL ,
	[ClosedMatchesCommunityBased] [int] NULL ,
	[ClosedMatchesSchoolBased] [int] NULL ,
	[ClosedMatchesOtherSiteBased] [int] NULL ,
	[ClosedMatchesGroupMentoring] [int] NULL ,
	[ClosedMatchesSpecialProgramsMentoring] [int] NULL ,
	[ClosedMatchesSpecialProgramsNonMentoring] [int] NULL ,
	[CreateDate] [datetime] NULL ,
	[ModifyLogID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_frmSpecialPopulations] (
	[SpecialPopulationsID] [int] IDENTITY (1, 1) NOT NULL ,
	[AgencyID] [varchar] (50) NULL ,
	[Year] [int] NULL ,
	[AbusedNeglected] [int] NULL ,
	[AdjudicatedDelinquents] [int] NULL ,
	[AfterSchool] [int] NULL ,
	[AIDSAffected] [int] NULL ,
	[DeafHearingImpaired] [int] NULL ,
	[DevelopmentallyDisabled] [int] NULL ,
	[FosterChildren] [int] NULL ,
	[Homeless] [int] NULL ,
	[IncarceratedParents] [int] NULL ,
	[Institutionalized] [int] NULL ,
	[LearningDisabled] [int] NULL ,
	[PhysicallyDisabled] [int] NULL ,
	[PregnantTeen] [int] NULL ,
	[SchoolDropouts] [int] NULL ,
	[TeenParentsFemale] [int] NULL ,
	[TeenParentsMale] [int] NULL ,
	[VisuallyImpaired] [int] NULL ,
	[OtherType] [varchar] (255) NULL ,
	[Other] [int] NULL ,
	[CreateDate] [datetime] NULL ,
	[ModifyLogID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_frmSpecialPrograms] (
	[SpecialProgramsID] [int] IDENTITY (1, 1) NOT NULL ,
	[AgencyID] [varchar] (50) NULL ,
	[Year] [int] NULL ,
	[FiftyFiveOrOlderBigs] [int] NULL ,
	[AcademicAchievement] [int] NULL ,
	[AfterSchoolMentoring] [int] NULL ,
	[AIDSPreventionIntervention] [int] NULL ,
	[AlcoholAbusePreventionIntervention] [int] NULL ,
	[Camping] [int] NULL ,
	[CharacterCounts] [int] NULL ,
	[ChildrenWithDisabilities] [int] NULL ,
	[CollegeStudentsAsBigs] [int] NULL ,
	[CommunityServiceProjects] [int] NULL ,
	[DropOutPrevention] [int] NULL ,
	[DrugAbusePreventionIntervention] [int] NULL ,
	[EmergencyFinancialAssistance] [int] NULL ,
	[EmployabilityJobReadiness] [int] NULL ,
	[FamilyCounseling] [int] NULL ,
	[GroupActivities] [int] NULL ,
	[HighSchoolStudentsAsBigs] [int] NULL ,
	[LifeSkillsLifeChoices] [int] NULL ,
	[ParentSupportGroups] [int] NULL ,
	[PartnershipsCivicOrganizations] [int] NULL ,
	[PartnershipsCollegesUniversities] [int] NULL ,
	[PartnershipsCorporationsBusinesses] [int] NULL ,
	[PartnershipsOtherYouthServingOrganizations] [int] NULL ,
	[PartnershipsReligiousOrganizations] [int] NULL ,
	[PregnancyPrevention] [int] NULL ,
	[Scholarships] [int] NULL ,
	[SexualAbusePreventionInterventionEmpower] [int] NULL ,
	[SexualAbusePreventionInterventionNOTEmpower] [int] NULL ,
	[SiteBasedMentoringSchoolBased] [int] NULL ,
	[SiteBasedMentoringNOTSchoolBased] [int] NULL ,
	[TeenParenting] [int] NULL ,
	[OtherText] [varchar] (255) NULL ,
	[Other] [int] NULL ,
	[CreateDate] [datetime] NULL ,
	[ModifyLogID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_frmStaff] (
	[StaffID] [int] IDENTITY (1, 1) NOT NULL ,
	[AgencyID] [varchar] (50) NULL ,
	[Year] [int] NULL ,
	[BirthYear] [int] NULL ,
	[Position] [int] NULL ,
	[Race] [int] NULL ,
	[Sex] [char] (2) NULL ,
	[Time] [char] (2) NULL ,
	[Education] [int] NULL ,
	[MonthStart] [int] NULL ,
	[YearStart] [int] NULL ,
	[MonthEnd] [int] NULL ,
	[HoursWeek] [int] NULL ,
	[YearlySalary] [money] NULL ,
	[CreateDate] [datetime] NULL ,
	[ModifyLogID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_ModifyLog] (
	[ModifyId] [int] IDENTITY (1, 1) NOT NULL ,
	[Form] [varchar] (100) NULL ,
	[FormModified] [int] NULL ,
	[Year] [int] NULL ,
	[Month] [int] NULL ,
	[ModifyType] [varchar] (50) NULL ,
	[UserName] [varchar] (50) NULL ,
	[AgencyID] [varchar] (50) NULL ,
	[ModifyDate] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_StaffEducation] (
	[CodeID] [int] IDENTITY (1, 1) NOT NULL ,
	[Code] [int] NULL ,
	[Education] [varchar] (100) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_StaffPosition] (
	[CodeID] [int] IDENTITY (1, 1) NOT NULL ,
	[Code] [int] NULL ,
	[Position] [varchar] (100) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_StaffRace] (
	[CodeID] [int] IDENTITY (1, 1) NOT NULL ,
	[Code] [int] NULL ,
	[Race] [varchar] (100) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_UserLogins] (
	[UserLoginID] [int] IDENTITY (1, 1) NOT NULL ,
	[Username] [varchar] (50) NULL ,
	[Password] [varchar] (50) NULL ,
	[AgencyID] [varchar] (50) NULL ,
	[CreateDate] [datetime] NULL 
) ON [PRIMARY]
GO

