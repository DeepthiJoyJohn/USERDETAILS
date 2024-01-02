<cfset local.userObject=createObject("component", "Components.userdetails")>
<cfset local.resultUserDetails=local.userObject.getUserDetails()>

<cfset local.timestamp = DateFormat(now(), "yyyymmdd_HHmmss")>
<cfset local.uniqueFilename = "ExcelTemplateData_#local.timestamp#.xlsx">



<cfspreadsheet action="write" filename="#expandPath('ExcelTemplate/DataExcel/')##uniqueFilename#" query="local.resultUserDetails">
<cfheader name="Content-Disposition" value="attachment;filename=#uniqueFilename#">
<cfcontent file="#expandPath('ExcelTemplate/DataExcel/')##uniqueFilename#" type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet">