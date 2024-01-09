<cfset local.mySpreadsheet = spreadsheetNew("Sheet1",true)>
<cfset spreadsheetAddRow(local.mySpreadsheet, 'First Name,Last Name,Address,Email,Phone,DOB,Role')>
<cfset local.headerFormat = {}>
<cfset local.headerFormat.bold = "true">
<cfset spreadsheetFormatRow(local.mySpreadsheet, local.headerFormat, 1)>
<cfheader name="Content-Disposition" value="inline;filename=Plain_Template.xlsx">
<cfcontent  variable="#spreadsheetReadBinary(local.mySpreadsheet)#" type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"> 
