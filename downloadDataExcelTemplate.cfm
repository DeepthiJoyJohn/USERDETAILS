<cfset local.mySpreadsheet = spreadsheetNew()>
<cfset spreadsheetAddRow(local.mySpreadsheet, 'First Name,Last Name,Address,Email,Phone,DOB,Role')>
<cfset local.headerFormat = {}>
<cfset local.headerFormat.bold = "true">
<cfset spreadsheetFormatRow(local.mySpreadsheet, local.headerFormat, 1)> 
<cfset local.userObject = createObject("component", "Components.userdetails")>
<cfset local.resultUserDetails = local.userObject.getUserDetails()>
<cfloop query="local.resultUserDetails">
    <cfset local.rolenames ="'#local.resultUserDetails.rolenames#'">        
    <cfset local.combinedValues = '#local.resultUserDetails.firstname#,#local.resultUserDetails.lastname#,
    #local.resultUserDetails.address#,#local.resultUserDetails.email#,#local.resultUserDetails.phone#,
    #DateFormat(local.resultUserDetails.dob, "MM/DD/YYYY")#,#local.rolenames#'>
    <cfset spreadsheetAddRow(local.mySpreadsheet,local.combinedValues)>
</cfloop>
<cfset local.timestamp = DateFormat(now(), "yyyymmdd_HHmmss")>
<cfset local.uniqueFilename = "ExcelTemplateData_#local.timestamp#.xlsx">
<cfset Spreadsheetwrite(local.mySpreadsheet,'#expandPath('ExcelTemplate/DataExcel/')##local.uniqueFilename#',true)>
<cfheader name="Content-Disposition" value="attachment;filename=#uniqueFilename#">
<cfcontent file="#expandPath('ExcelTemplate/DataExcel/')##uniqueFilename#" type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"> 


