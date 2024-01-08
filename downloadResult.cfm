<cfset local.userObject=createObject("component", "Components.userdetails")>
<cfset local.seqno=local.userObject.getMaxSeqNo()>
<cfset local.mySpreadsheet = spreadsheetNew()>
<cfset spreadsheetAddRow(local.mySpreadsheet, 'First Name,Last Name,Address,Email,Phone,DOB,Role,Result,Reason')>
<cfset local.headerFormat = {}>
<cfset local.headerFormat.bold = "true">
<cfset spreadsheetFormatRow(local.mySpreadsheet, local.headerFormat, 1)> 
<cfset local.userObject = createObject("component", "Components.userdetails")>
<cfset local.resultUserDetailsError = local.userObject.getUserDetailsWithError(local.seqNo)>
<cfloop query="local.resultUserDetailsError">
    <cfset local.rolenames ="'#local.resultUserDetailsError.roles#'"> 			
    <cfset local.address = replace(local.resultUserDetailsError.address, ",", " ", "ALL")>      
    <cfset local.combinedValues = '#local.resultUserDetailsError.firstname#,#local.resultUserDetailsError.lastname#,
    #local.address#,#local.resultUserDetailsError.email#,#local.resultUserDetailsError.phone#,
    #local.resultUserDetailsError.dob#,#local.rolenames#,Failed,#local.resultUserDetailsError.result#'>
    <cfset spreadsheetAddRow(local.mySpreadsheet,local.combinedValues)>
</cfloop>		
<cfset local.resultUserDetails = local.userObject.getUserDetails(local.seqNo)>
<cfloop query="local.resultUserDetails">
    <cfset local.rolenames ="'#local.resultUserDetails.rolenames#'"> 
    <cfset local.address = replace(local.resultUserDetails.address, ",", " ", "ALL")>  
    <cfset local.combinedValues = '#local.resultUserDetails.firstname#,#local.resultUserDetails.lastname#,
    #local.address#,#local.resultUserDetails.email#,#local.resultUserDetails.phone#,
    #DateFormat(local.resultUserDetails.dob, "MM/DD/YYYY")#,#local.rolenames#,#local.resultUserDetails.result#'>
    <cfset spreadsheetAddRow(local.mySpreadsheet,local.combinedValues)>
</cfloop>
<cfset local.timestamp = DateFormat(now(), "yyyymmdd_HHmmss")>
<cfset local.uniqueFilename = "ExcelResult_#local.timestamp#.xlsx">
<cfset Spreadsheetwrite(local.mySpreadsheet,'#expandPath('ExcelUploads/Result/')##local.uniqueFilename#',true)>
<cfheader name="Content-Disposition" value="attachment;filename=#local.uniqueFilename#">
<cfcontent file="#expandPath('ExcelUploads/Result/')##local.uniqueFilename#" type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"> 										
