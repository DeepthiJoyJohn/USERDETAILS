<cfset local.mySpreadsheet = spreadsheetNew("Sheet1",true)>
<cfset spreadsheetAddRow(local.mySpreadsheet, 'First Name,Last Name,Address,Email,Phone,DOB,Role')>
<cfset local.headerFormat = {}>
<cfset local.headerFormat.bold = "true">
<cfset spreadsheetFormatRow(local.mySpreadsheet, local.headerFormat, 1)> 
<cfset local.userObject=createObject("component", "Components.userdetails")>
<cfset local.resultUserDetails=local.userObject.getUserDetails()>
<cfset local.rowNum = 2>
<cfloop query="local.resultUserDetails">
    <cfset spreadsheetSetCellValue(local.mySpreadsheet, local.resultUserDetails.firstname, local.rowNum, 1)>
    <cfset spreadsheetSetCellValue(local.mySpreadsheet, local.resultUserDetails.lastname, local.rowNum, 2)>
    <cfset spreadsheetSetCellValue(local.mySpreadsheet, local.resultUserDetails.address, local.rowNum, 3)>
    <cfset spreadsheetSetCellValue(local.mySpreadsheet, local.resultUserDetails.email, local.rowNum, 4)>
    <cfset spreadsheetSetCellValue(local.mySpreadsheet, local.resultUserDetails.phone, local.rowNum, 5)>
    <cfset spreadsheetSetCellValue(local.mySpreadsheet, local.resultUserDetails.dobdisplay, local.rowNum, 6)>
    <cfset spreadsheetSetCellValue(local.mySpreadsheet, local.resultUserDetails.rolenames, local.rowNum, 7)>
    <cfset local.rowNum = local.rowNum+1>
</cfloop>
<cfheader name="Content-Disposition" value="inline;filename=Data.xlsx">
<cfcontent  variable="#spreadsheetReadBinary(local.mySpreadsheet)#" type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"> 



