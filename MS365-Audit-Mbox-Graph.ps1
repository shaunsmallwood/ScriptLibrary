# This script has been tested on both Windows PowerShell (v5.x) and macOS PowerShell Core (v7.3) 
# The MgGraph and ExchangeOnline modules are required, if this is your first time running MgGraph/ExchangeOnline 
# modules on your computer please run:
# 
# Install-Module Microsoft.Graph
# Install-Module ExchangeOnlineManagement
#
# Before running this script connect to environment using the following commands:
#
# Connect-MgGraph -Scopes Directory.ReadWrite.All,User.ReadWrite.All,Group.ReadWrite.All
# Connect-ExchangeOnline -UserPrincipalName <username@domain-used-for-mggraph-above>
#
# If running on macOS or Linux you'll need to install PowerShell Core:
# https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell
#
# YOU MAY ALSO NEED TO CHANGE EXECUTION POLICY, you can check current policy by running Get-ExecutionPolicy -List
# then run Set-ExecutionPolicy Unrestricted -Scope <desired_scope>
#

# Define output file name
$OutFile = ".\MailboxReport" + (Get-Date -UFormat "%y%m%d%H%M%S") + ".csv"

# Define array to decipher MS365 licensing "Friendly Name"
$MSProdArray = ConvertFrom-Csv @'
"StringID","ProductName"
"AAD_BASIC","AZURE ACTIVE DIRECTORY BASIC"
"AAD_PREMIUM","AZURE ACTIVE DIRECTORY PREMIUM P1"
"AAD_PREMIUM_P2","AZURE ACTIVE DIRECTORY PREMIUM P2"
"ADALLOM_O365","Office 365 Cloud App Security"
"ADALLOM_STANDALONE","Microsoft Cloud App Security"
"ADV_COMMS","Advanced Communications"
"ATA","Microsoft Defender for Identity"
"ATP_ENTERPRISE","Microsoft Defender for Office 365 (Plan 1)"
"ATP_ENTERPRISE_FACULTY","Microsoft Defender for Office 365 (Plan 1) Faculty"
"ATP_ENTERPRISE_GOV","Microsoft Defender for Office 365 (Plan 1) GCC"
"AX7_USER_TRIAL","Microsoft Dynamics AX7 User Trial"
"BUSINESS_VOICE_DIRECTROUTING","Microsoft 365 Business Voice (without calling plan)"
"BUSINESS_VOICE_DIRECTROUTING_MED","Microsoft 365 Business Voice (without Calling Plan) for US"
"BUSINESS_VOICE_MED2","Microsoft 365 Business Voice"
"BUSINESS_VOICE_MED2_TELCO","Microsoft 365 Business Voice (US)"
"CCIBOTS_PRIVPREV_VIRAL","Power Virtual Agents Viral Trial"
"CDS_DB_CAPACITY","Common Data Service Database Capacity"
"CDS_DB_CAPACITY_GOV","Common Data Service Database Capacity for Government"
"CDS_LOG_CAPACITY","Common Data Service Log Capacity"
"CDSAICAPACITY","AI Builder Capacity add-on"
"CMPA_addon_GCC","Compliance Manager Premium Assessment Add-On for GCC"
"CPC_B_2C_4RAM_64GB","Windows 365 Business 2 vCPU, 4 GB, 64 GB"
"CPC_B_4C_16RAM_128GB_WHB","Windows 365 Business 4 vCPU, 16 GB, 128 GB (with Windows Hybrid Benefit)"
"CPC_E_2C_4GB_64GB","Windows 365 Enterprise 2 vCPU, 4 GB, 64 GB"
"CPC_E_2C_8GB_128GB","Windows 365 Enterprise 2 vCPU, 8 GB, 128 GB"
"CPC_LVL_2","Windows 365 Enterprise 2 vCPU, 8 GB, 128 GB (Preview)"
"CPC_LVL_3","Windows 365 Enterprise 4 vCPU, 16 GB, 256 GB (Preview)"
"CRM_ONLINE_PORTAL","Dynamics 365 Enterprise Edition - Additional Portal (Qualified Offer)"
"CRMINSTANCE","Dynamics 365 - Additional Production Instance (Qualified Offer)"
"CRMPLAN2","MICROSOFT DYNAMICS CRM ONLINE BASIC"
"CRMSTANDARD","MICROSOFT DYNAMICS CRM ONLINE"
"CRMSTORAGE","Dynamics 365 - Additional Database Storage (Qualified Offer)"
"CRMTESTINSTANCE","Dynamics 365 - Additional Non-Production Instance (Qualified Offer)"
"D365_FIELD_SERVICE_ATTACH","Dynamics 365 for Field Service Attach to Qualifying Dynamics 365 Base Offer"
"D365_MARKETING_USER","Dynamics 365 for Marketing USL"
"D365_SALES_ENT_ATTACH","Dynamics 365 Sales Enterprise Attach to Qualifying Dynamics 365 Base Offer"
"D365_SALES_PRO","Dynamics 365 For Sales Professional"
"D365_SALES_PRO_ATTACH","Dynamics 365 Sales Professional Attach to Qualifying Dynamics 365 Base Offer"
"D365_SALES_PRO_IW","Dynamics 365 For Sales Professional Trial"
"DEFENDER_ENDPOINT_P1","Microsoft Defender for Endpoint P1"
"DESKLESSPACK","OFFICE 365 F3"
"DEVELOPERPACK","OFFICE 365 E3 DEVELOPER"
"DEVELOPERPACK_E5","Microsoft 365 E5 Developer (without Windows and Audio Conferencing)"
"DYN365_AI_SERVICE_INSIGHTS","Dynamics 365 Customer Service Insights Trial"
"DYN365_ASSETMANAGEMENT","Dynamics 365 Asset Management Addl Assets"
"DYN365_BUSCENTRAL_ADD_ENV_ADDON","Dynamics 365 Business Central Additional Environment Addon"
"DYN365_BUSCENTRAL_DB_CAPACITY","Dynamics 365 Business Central Database Capacity"
"DYN365_BUSCENTRAL_ESSENTIAL","Dynamics 365 Business Central Essentials"
"DYN365_BUSCENTRAL_PREMIUM","Dynamics 365 Business Central Premium"
"DYN365_BUSINESS_MARKETING","Dynamics 365 for Marketing Business Edition"
"DYN365_CUSTOMER_INSIGHTS_VIRAL","Dynamics 365 Customer Insights vTrial"
"DYN365_CUSTOMER_SERVICE_PRO","Dynamics 365 Customer Service Professional"
"DYN365_CUSTOMER_VOICE_ADDON","Dynamics 365 Customer Voice Additional Responses"
"DYN365_CUSTOMER_VOICE_BASE","Dynamics 365 Customer Voice"
"DYN365_ENTERPRISE_CASE_MANAGEMENT","Dynamics 365 for Case Management Enterprise Edition"
"DYN365_ENTERPRISE_CUSTOMER_SERVICE","Dynamics 365 for Customer Service Enterprise Edition"
"DYN365_ENTERPRISE_FIELD_SERVICE","Dynamics 365 for Field Service Enterprise Edition"
"DYN365_ENTERPRISE_P1_IW","DYNAMICS 365 P1 TRIAL FOR INFORMATION WORKERS"
"DYN365_ENTERPRISE_PLAN1","Dynamics 365 Customer Engagement Plan"
"DYN365_ENTERPRISE_SALES","DYNAMICS 365 FOR SALES ENTERPRISE EDITION"
"DYN365_ENTERPRISE_SALES_CUSTOMERSERVICE","DYNAMICS 365 FOR SALES AND CUSTOMER SERVICE ENTERPRISE EDITION"
"DYN365_ENTERPRISE_TEAM_MEMBERS","DYNAMICS 365 FOR TEAM MEMBERS ENTERPRISE EDITION"
"DYN365_FINANCE","Dynamics 365 Finance"
"DYN365_FINANCIALS_ACCOUNTANT_SKU","Dynamics 365 Business Central External Accountant"
"DYN365_FINANCIALS_BUSINESS_SKU","DYNAMICS 365 FOR FINANCIALS BUSINESS EDITION"
"DYN365_IOT_INTELLIGENCE_ADDL_MACHINES","Sensor Data Intelligence Additional Machines Add-in for Dynamics 365 Supply Chain Management"
"DYN365_IOT_INTELLIGENCE_SCENARIO","Sensor Data Intelligence Scenario Add-in for Dynamics 365 Supply Chain Management"
"DYN365_REGULATORY_SERVICE","Dynamics 365 Regulatory Service - Enterprise Edition Trial"
"DYN365_SCM","DYNAMICS 365 FOR SUPPLY CHAIN MANAGEMENT"
"DYN365_TEAM_MEMBERS","DYNAMICS 365 TEAM MEMBERS"
"Dynamics_365_Customer_Service_Enterprise_viral_trial","Dynamics 365 Customer Service Enterprise Viral Trial"
"Dynamics_365_Field_Service_Enterprise_viral_trial","Dynamics 365 Field Service Viral Trial"
"Dynamics_365_for_Operations","DYNAMICS 365 UNF OPS PLAN ENT EDITION"
"Dynamics_365_for_Operations_Devices","Dynamics 365 Operations - Device"
"Dynamics_365_for_Operations_Sandbox_Tier2_SKU","Dynamics 365 Operations - Sandbox Tier 2:Standard Acceptance Testing"
"Dynamics_365_for_Operations_Sandbox_Tier4_SKU","Dynamics 365 Operations - Sandbox Tier 4:Standard Performance Testing"
"Dynamics_365_Hiring_SKU","Dynamics 365 Talent: Attract"
"DYNAMICS_365_ONBOARDING_SKU","DYNAMICS 365 TALENT: ONBOARD"
"Dynamics_365_Sales_Premium_Viral_Trial","Dynamics 365 Sales Premium Viral Trial"
"E3_VDA_only","Windows 10/11 Enterprise VDA"
"EMS","ENTERPRISE MOBILITY + SECURITY E3"
"EMS_EDU_FACULTY","Enterprise Mobility + Security A3 for Faculty"
"EMS_GOV","Enterprise Mobility + Security G3 GCC"
"EMSPREMIUM","ENTERPRISE MOBILITY + SECURITY E5"
"EMSPREMIUM_GOV","Enterprise Mobility + Security G5 GCC"
"ENTERPRISEPACK","Office 365 E3"
"ENTERPRISEPACK_GOV","OFFICE 365 G3 GCC"
"ENTERPRISEPACK_USGOV_DOD","Office 365 E3_USGOV_DOD"
"ENTERPRISEPACK_USGOV_GCCHIGH","Office 365 E3_USGOV_GCCHIGH"
"ENTERPRISEPACKPLUS_FACULTY","Office 365 A3 for faculty"
"ENTERPRISEPACKPLUS_STUDENT","Office 365 A3 for students"
"ENTERPRISEPREMIUM","Office 365 E5"
"ENTERPRISEPREMIUM_FACULTY","Office 365 A5 for faculty"
"ENTERPRISEPREMIUM_GOV","Office 365 G5 GCC"
"ENTERPRISEPREMIUM_NOPSTNCONF","OFFICE 365 E5 WITHOUT AUDIO CONFERENCING"
"ENTERPRISEPREMIUM_STUDENT","Office 365 A5 for students"
"ENTERPRISEWITHSCAL","OFFICE 365 E4"
"EOP_ENTERPRISE","Exchange Online Protection"
"EOP_ENTERPRISE_PREMIUM","Exchange Enterprise CAL Services (EOP, DLP)"
"EQUIVIO_ANALYTICS","Office 365 Advanced Compliance"
"EQUIVIO_ANALYTICS_GOV","Office 365 Advanced Compliance for GCC"
"EXCHANGE_S_ESSENTIALS","EXCHANGE ONLINE ESSENTIALS"
"EXCHANGEARCHIVE","EXCHANGE ONLINE ARCHIVING FOR EXCHANGE SERVER"
"EXCHANGEARCHIVE_ADDON","EXCHANGE ONLINE ARCHIVING FOR EXCHANGE ONLINE"
"EXCHANGEDESKLESS","EXCHANGE ONLINE KIOSK"
"EXCHANGEENTERPRISE","EXCHANGE ONLINE (PLAN 2)"
"EXCHANGEESSENTIALS","EXCHANGE ONLINE ESSENTIALS (ExO P1 BASED)"
"EXCHANGESTANDARD","Exchange Online (Plan 1)"
"EXCHANGESTANDARD_GOV","Exchange Online (Plan 1) for GCC"
"EXCHANGETELCO","EXCHANGE ONLINE POP"
"EXPERTS_ON_DEMAND","Microsoft Threat Experts - Experts on Demand"
"FLOW_BUSINESS_PROCESS","Power Automate per flow plan"
"FLOW_FREE","MICROSOFT FLOW FREE"
"FLOW_P1_GOV","Power Automate Plan 1 for Government (Qualified Offer)"
"FLOW_P2","MICROSOFT POWER AUTOMATE PLAN 2"
"FLOW_PER_USER","Power Automate per user plan"
"FLOW_PER_USER_DEPT","Power Automate per user plan dept"
"FLOW_PER_USER_GCC","Power Automate per user plan for Government"
"FORMS_PRO","Dynamics 365 Customer Voice Trial"
"Forms_Pro_AddOn","Dynamics 365 Customer Voice Additional Responses"
"Forms_Pro_USL","Dynamics 365 Customer Voice USL"
"GUIDES_USER","Dynamics 365 Guides"
"IDENTITY_THREAT_PROTECTION","Microsoft 365 E5 Security"
"IDENTITY_THREAT_PROTECTION_FOR_EMS_E5","Microsoft 365 E5 Security for EMS E5"
"INFORMATION_PROTECTION_COMPLIANCE","Microsoft 365 E5 Compliance"
"Intelligent_Content_Services","SharePoint Syntex"
"INTUNE_A","INTUNE"
"INTUNE_A_D","Microsoft Intune Device"
"INTUNE_A_D_GOV","MICROSOFT INTUNE DEVICE FOR GOVERNMENT"
"INTUNE_SMB","MICROSOFT INTUNE SMB"
"IT_ACADEMY_AD","MS IMAGINE ACADEMY"
"LITEPACK","OFFICE 365 SMALL BUSINESS"
"LITEPACK_P2","OFFICE 365 SMALL BUSINESS PREMIUM"
"M365_E5_SUITE_COMPONENTS","Microsoft 365 E5 Suite Features"
"M365_F1","Microsoft 365 F1"
"M365_F1_COMM","Microsoft 365 F1"
"M365_F1_GOV","Microsoft 365 F3 GCC"
"M365_G3_GOV","MICROSOFT 365 G3 GCC"
"M365_G5_GCC","Microsoft 365 GCC G5"
"M365_SECURITY_COMPLIANCE_FOR_FLW","Microsoft 365 Security and Compliance for Firstline Workers"
"M365EDU_A1","Microsoft 365 A1"
"M365EDU_A3_FACULTY","Microsoft 365 A3 for Faculty"
"M365EDU_A3_STUDENT","MICROSOFT 365 A3 FOR STUDENTS"
"M365EDU_A3_STUUSEBNFT","Microsoft 365 A3 for students use benefit"
"M365EDU_A3_STUUSEBNFT_RPA1","Microsoft 365 A3 - Unattended License for students use benefit"
"M365EDU_A5_FACULTY","Microsoft 365 A5 for Faculty"
"M365EDU_A5_NOPSTNCONF_STUUSEBNFT","Microsoft 365 A5 without Audio Conferencing for students use benefit"
"M365EDU_A5_STUDENT","MICROSOFT 365 A5 FOR STUDENTS"
"M365EDU_A5_STUUSEBNFT","Microsoft 365 A5 for students use benefit"
"MCOCAP","COMMON AREA PHONE"
"MCOCAP_GOV","Common Area Phone for GCC"
"MCOEV","MICROSOFT 365 PHONE SYSTEM"
"MCOEV_DOD","MICROSOFT 365 PHONE SYSTEM FOR DOD"
"MCOEV_FACULTY","MICROSOFT 365 PHONE SYSTEM FOR FACULTY"
"MCOEV_GCCHIGH","MICROSOFT 365 PHONE SYSTEM FOR GCCHIGH"
"MCOEV_GOV","MICROSOFT 365 PHONE SYSTEM FOR GCC"
"MCOEV_STUDENT","MICROSOFT 365 PHONE SYSTEM FOR STUDENTS"
"MCOEV_TELSTRA","MICROSOFT 365 PHONE SYSTEM FOR TELSTRA"
"MCOEV_USGOV_DOD","MICROSOFT 365 PHONE SYSTEM_USGOV_DOD"
"MCOEV_USGOV_GCCHIGH","MICROSOFT 365 PHONE SYSTEM_USGOV_GCCHIGH"
"MCOEVSMB_1","MICROSOFT 365 PHONE SYSTEM FOR SMALL AND MEDIUM BUSINESS"
"MCOIMP","SKYPE FOR BUSINESS ONLINE (PLAN 1)"
"MCOMEETADV","Microsoft 365 Audio Conferencing"
"MCOMEETADV_GOC","MICROSOFT 365 AUDIO CONFERENCING FOR GCC"
"MCOPSTN_1_GOV","Microsoft 365 Domestic Calling Plan for GCC"
"MCOPSTN_5","MICROSOFT 365 DOMESTIC CALLING PLAN (120 Minutes)"
"MCOPSTN1","SKYPE FOR BUSINESS PSTN DOMESTIC CALLING"
"MCOPSTN2","SKYPE FOR BUSINESS PSTN DOMESTIC AND INTERNATIONAL CALLING"
"MCOPSTN5","SKYPE FOR BUSINESS PSTN DOMESTIC CALLING (120 Minutes)"
"MCOPSTNC","COMMUNICATIONS CREDITS"
"MCOPSTNEAU2","TELSTRA CALLING FOR O365"
"MCOPSTNPP","Skype for Business PSTN Usage Calling Plan"
"MCOSTANDARD","SKYPE FOR BUSINESS ONLINE (PLAN 2)"
"MCOTEAMS_ESSENTIALS","Teams Phone with Calling Plan"
"MDATP_Server","Microsoft Defender for Endpoint Server"
"MDATP_XPLAT","Microsoft Defender for Endpoint P2_XPLAT"
"MEETING_ROOM","Microsoft Teams Rooms Standard"
"MEETING_ROOM_NOAUDIOCONF","Microsoft Teams Rooms Standard without Audio Conferencing"
"MFA_STANDALONE","Microsoft Azure Multi-Factor Authentication"
"MICROSOFT_BUSINESS_CENTER","MICROSOFT BUSINESS CENTER"
"MICROSOFT_REMOTE_ASSIST","Dynamics 365 Remote Assist"
"MICROSOFT_REMOTE_ASSIST_HOLOLENS","Dynamics 365 Remote Assist HoloLens"
"Microsoft_Teams_Audio_Conferencing_select_dial_out","Microsoft Teams Audio Conferencing select dial-out"
"MIDSIZEPACK","OFFICE 365 MIDSIZE BUSINESS"
"MS_TEAMS_IW","Microsoft Teams Trial"
"MTR_PREM","Teams Rooms Premium"
"NONPROFIT_PORTAL","Nonprofit Portal"
"O365_BUSINESS","MICROSOFT 365 APPS FOR BUSINESS"
"O365_BUSINESS_ESSENTIALS","MICROSOFT 365 BUSINESS BASIC"
"O365_BUSINESS_PREMIUM","Microsoft 365 Business Standard"
"OFFICE_PROPLUS_DEVICE1","Microsoft 365 Apps for enterprise (device)"
"OFFICE365_MULTIGEO","Multi-Geo Capabilities in Office 365"
"OFFICESUBSCRIPTION","Microsoft 365 Apps for enterprise"
"OFFICESUBSCRIPTION_FACULTY","Microsoft 365 Apps for Faculty"
"OFFICESUBSCRIPTION_STUDENT","Microsoft 365 Apps for Students"
"PBI_PREMIUM_P1_ADDON","Power BI Premium P1"
"PBI_PREMIUM_PER_USER","Power BI Premium Per User"
"PBI_PREMIUM_PER_USER_ADDON","Power BI Premium Per User Add-On"
"PBI_PREMIUM_PER_USER_DEPT","Power BI Premium Per User Dept"
"PHONESYSTEM_VIRTUALUSER","MICROSOFT 365 PHONE SYSTEM - VIRTUAL USER"
"PHONESYSTEM_VIRTUALUSER_GOV","Microsoft 365 Phone System - Virtual User for GCC"
"POWER_BI_ADDON","POWER BI FOR OFFICE 365 ADD-ON"
"POWER_BI_INDIVIDUAL_USER","Power BI"
"POWER_BI_PRO","Power BI Pro"
"POWER_BI_PRO_CE","Power BI Pro CE"
"POWER_BI_PRO_DEPT","Power BI Pro Dept"
"POWER_BI_STANDARD","Power BI (free)"
"POWERAPPS_DEV","Microsoft PowerApps for Developer"
"POWERAPPS_INDIVIDUAL_USER","POWERAPPS AND LOGIC FLOWS"
"POWERAPPS_P1_GOV","Power Apps Plan 1 for Government"
"POWERAPPS_PER_APP","Power Apps per app plan"
"POWERAPPS_PER_APP_IW","PowerApps per app baseline access"
"POWERAPPS_PER_APP_NEW","Power Apps per app plan (1 app or portal)"
"POWERAPPS_PER_USER","Power Apps per user plan"
"POWERAPPS_PER_USER_GCC","Power Apps per user plan for Government"
"POWERAPPS_PORTALS_LOGIN_T2","Power Apps Portals login capacity add-on Tier 2 (10 unit min)"
"POWERAPPS_PORTALS_LOGIN_T2_GCC","Power Apps Portals login capacity add-on Tier 2 (10 unit min) for Government"
"POWERAPPS_PORTALS_PAGEVIEW_GCC","Power Apps Portals page view capacity add-on for Government"
"POWERAPPS_VIRAL","Microsoft Power Apps Plan 2 Trial"
"POWERAUTOMATE_ATTENDED_RPA","Power Automate per user with attended RPA plan"
"POWERAUTOMATE_UNATTENDED_RPA","Power Automate unattended RPA add-on"
"POWERBI_PRO_GOV","Power BI Pro for GCC"
"POWERFLOW_P2","Microsoft Power Apps Plan 2 (Qualified Offer)"
"PROJECT_MADEIRA_PREVIEW_IW_SKU","Dynamics 365 Business Central for IWs"
"PROJECT_P1","PROJECT PLAN 1"
"PROJECT_PLAN1_DEPT","Project Plan 1 (for Department)"
"PROJECT_PLAN3_DEPT","Project Plan 3 (for Department)"
"PROJECTCLIENT","PROJECT FOR OFFICE 365"
"PROJECTESSENTIALS","Project Online Essentials"
"PROJECTESSENTIALS_GOV","Project Online Essentials for GCC"
"PROJECTONLINE_PLAN_1","PROJECT ONLINE PREMIUM WITHOUT PROJECT CLIENT"
"PROJECTONLINE_PLAN_2","PROJECT ONLINE WITH PROJECT FOR OFFICE 365"
"PROJECTPREMIUM","PROJECT ONLINE PREMIUM"
"PROJECTPREMIUM_GOV","Project Plan 5 for GCC"
"PROJECTPROFESSIONAL","Project Plan 3"
"PROJECTPROFESSIONAL_GOV","Project Plan 3 for GCC"
"RIGHTSMANAGEMENT","AZURE INFORMATION PROTECTION PLAN 1"
"RIGHTSMANAGEMENT_ADHOC","Rights Management Adhoc"
"RMSBASIC","Rights Management Service Basic Content Protection"
"SHAREPOINTENTERPRISE","SHAREPOINT ONLINE (PLAN 2)"
"SHAREPOINTSTANDARD","SHAREPOINT ONLINE (PLAN 1)"
"SHAREPOINTSTORAGE","Office 365 Extra File Storage"
"SHAREPOINTSTORAGE_GOV","Office 365 Extra File Storage for GCC"
"SKU_Dynamics_365_for_HCM_Trial","Dynamics 365 for Talent"
"SMB_APPS","Business Apps (free)"
"SMB_BUSINESS","MICROSOFT 365 APPS FOR BUSINESS"
"SMB_BUSINESS_ESSENTIALS","MICROSOFT 365 BUSINESS BASIC"
"SMB_BUSINESS_PREMIUM","MICROSOFT 365 BUSINESS STANDARD - PREPAID LEGACY"
"SOCIAL_ENGAGEMENT_APP_USER","Dynamics 365 AI for Market Insights (Preview)"
"SPB","Microsoft 365 Business Premium"
"SPE_E3","Microsoft 365 E3"
"SPE_E3_RPA1","Microsoft 365 E3 - Unattended License"
"SPE_E3_USGOV_DOD","Microsoft 365 E3_USGOV_DOD"
"SPE_E3_USGOV_GCCHIGH","Microsoft 365 E3_USGOV_GCCHIGH"
"SPE_E5","Microsoft 365 E5"
"SPE_E5_NOPSTNCONF","Microsoft 365 E5 without Audio Conferencing"
"SPE_F1","Microsoft 365 F3"
"SPE_F5_SEC","Microsoft 365 F5 Security Add-on"
"SPE_F5_SECCOMP","Microsoft 365 F5 Security + Compliance Add-on"
"SPZA_IW","APP CONNECT IW"
"STANDARDPACK","Office 365 E1"
"STANDARDPACK_GOV","Office 365 G1 GCC"
"STANDARDWOFFPACK","OFFICE 365 E2"
"STANDARDWOFFPACK_FACULTY","Office 365 A1 for faculty"
"STANDARDWOFFPACK_IW_FACULTY","Office 365 A1 Plus for faculty"
"STANDARDWOFFPACK_IW_STUDENT","Office 365 A1 Plus for students"
"STANDARDWOFFPACK_STUDENT","Office 365 A1 for students"
"STREAM","MICROSOFT STREAM"
"STREAM_P2","Microsoft Stream Plan 2"
"STREAM_STORAGE","Microsoft Stream Storage Add-On (500 GB)"
"TEAMS_COMMERCIAL_TRIAL","Microsoft Teams Commercial Cloud"
"TEAMS_EXPLORATORY","MICROSOFT TEAMS EXPLORATORY"
"TEAMS_FREE","MICROSOFT TEAMS (FREE)"
"THREAT_INTELLIGENCE","Microsoft Defender for Office 365 (Plan 2)"
"THREAT_INTELLIGENCE_GOV","Microsoft Defender for Office 365 (Plan 2) GCC"
"TOPIC_EXPERIENCES","Viva Topics"
"UNIVERSAL_PRINT","Universal Print"
"VIRTUAL_AGENT_BASE","Power Virtual Agent"
"VISIO_PLAN1_DEPT","Visio Plan 1"
"VISIO_PLAN2_DEPT","Visio Plan 2"
"VISIOCLIENT","VISIO ONLINE PLAN 2"
"VISIOCLIENT_GOV","VISIO PLAN 2 FOR GCC"
"VISIOONLINE_PLAN1","VISIO ONLINE PLAN 1"
"WACONEDRIVEENTERPRISE","ONEDRIVE FOR BUSINESS (PLAN 2)"
"WACONEDRIVESTANDARD","ONEDRIVE FOR BUSINESS (PLAN 1)"
"WIN_DEF_ATP","MICROSOFT DEFENDER FOR ENDPOINT"
"WIN_ENT_E5","Windows 10/11 Enterprise E5 (Original)"
"WIN10_ENT_A3_FAC","Windows 10 Enterprise A3 for faculty"
"WIN10_ENT_A3_STU","Windows 10 Enterprise A3 for students"
"WIN10_PRO_ENT_SUB","WINDOWS 10 ENTERPRISE E3"
"WIN10_VDA_E3","WINDOWS 10 ENTERPRISE E3"
"WIN10_VDA_E5","Windows 10 Enterprise E5"
"WINDOWS_STORE","WINDOWS STORE FOR BUSINESS"
"WINE5_GCC_COMPAT","Windows 10 Enterprise E5 Commercial (GCC Compatible)"
"WORKPLACE_ANALYTICS","Microsoft Workplace Analytics"
"WSFB_EDU_FACULTY","Windows Store for Business EDU Faculty"
'@

#  Pull a list of all mailboxes and defining any non-default properties to query
$mailboxes = Get-EXOMailbox -ResultSize unlimited -Properties ArchiveStatus,ArchiveState,RetentionPolicy,AutoExpandingArchiveEnabled

# Cycle through each mailbox and pull metrics for output to CSV
foreach($mailbox in $mailboxes) {

    $currentuser = $mailbox.UserPrincipalName
    # Skip built-in "DiscoverySearchMailbox"
    if ($currentuser -notlike "*DiscoverySearchMailbox*") {Write-Host "Processing $currentuser" -NoNewLine} else { Write-Host -ForegroundColor Yellow "Discovery Mailbox Detected - Skipping" ; continue}

    $stats = Get-EXOMailboxStatistics $mailbox.UserPrincipalName -ErrorAction SilentlyContinue -Properties LastLogonTime
    $MgUserAccount = Get-MgUser -UserId $mailbox.UserPrincipalName -Property accountEnabled, assignedLicenses, DisplayName, UserPrincipalName -ErrorAction SilentlyContinue
    $MgUserAccountLicenses = (Get-MgUserLicenseDetail -UserId $mailbox.UserPrincipalName -ErrorAction SilentlyContinue).SkuPartNumber

    # Pull archive mailbox statistics if exists, even if not licensed and does not return Enabled attribute, otherwise skip
    if ($mailbox.ArchiveState -eq "Local") {
        $archivemailbox = Get-EXOMailboxStatistics $mailbox.UserPrincipalName -Archive
        if ($null -ne $archivemailbox) {
            $archivemailboxsize = [math]::Round((((($archivemailbox).TotalItemSize).ToString()).Split("(")[1].Split(" ")[0].Replace(",", "")/1GB), 2)
        }
        else {
            $archivemailboxsize = "NOT CREATED"
        }
    }
    else {
         $archivemailboxsize = "N/A"
     }
    
    # Get list of MS365 licensing and match against "Friendly Name" in product array
    if (($MgUserAccountLicenses).Count -gt 0) {
        $MSLicenses = (($MgUserAccountLicenses | ForEach-Object { ($MSProdArray -match "\b$_\b").ProductName }) -Join " | ").ToUpper()
    }
    else {
         $MSLicenses = ""
    }

    # Create/append to array with for each user/mailbox
    $ExportObject = [PSCustomObject]@{
        "Display Name" = $MgUserAccount.DisplayName
        "Username (MG)" = $MgUserAccount.UserPrincipalName
        "Username (EXO)" = $mailbox.UserPrincipalName
        "Primary Email" = $mailbox.PrimarySmtpAddress
        "Email Aliases" = (($mailbox.EmailAddresses | Where-Object {$_ -clike "smtp*" -and $_ -NotLike "*onmicrosoft.com*"}) -replace "smtp:","") -join " | "
        "Mailbox Type" = $mailbox.RecipientTypeDetails
        "Mailbox Size (GB)" = [math]::Round((($stats.Totalitemsize.Value.ToString()).Split("(")[1].Split(" ")[0].Replace(",", "")/1GB), 2)
        "Archive Enabled" = $mailbox.ArchiveStatus
        "Archive Size (GB)" = $archivemailboxsize
        "Auto Expand Archive" = $mailbox.AutoExpandingArchiveEnabled
        "Retention Policy" = $mailbox.RetentionPolicy
        "Licenses Assigned" = "$MSLicenses"
        "Login Allowed" = $MgUserAccount.AccountEnabled
        "Last Logon Time" = $stats.LastLogonTime
    }

# Export final CSV and wrap up
$ExportObject | Export-Csv $OutFile -NoClobber -NoTypeInformation -Append
Write-Host " - Complete!"
}

Write-Host -ForegroundColor Yellow "Script Complete - Output file name: $OutFile"
