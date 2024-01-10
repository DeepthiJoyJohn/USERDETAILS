<cfcomponent> 
	<cffunction name="getUserDetails" access="public" returntype="query">		
		<cfquery name="local.qgetUserDetails" datasource="#application.datasoursename#">
			SELECT
				u.userid,
				u.firstname,u.lastname,u.address,u.email,u.phone,u.dob,
				DATE_FORMAT(u.dob,'%d-%m-%Y') as dobdisplay,
				GROUP_CONCAT(r.rolename, '') AS rolenames
			FROM
				USER u
			INNER JOIN
				userroles ur ON u.userid = ur.userid
			INNER JOIN
				ROLE r ON ur.roleid = r.roleid			
			GROUP BY
				u.userid								
		</cfquery>
		<cfreturn local.qgetUserDetails> 		
	</cffunction>
	
	<cffunction name="checkEmailExists" access="public" returntype="numeric">
		<cfargument name="email">		
		<cfquery name="local.qcheckEmailExists" datasource="#application.datasoursename#">
			SELECT email
			FROM user 
			WHERE email=<cfqueryparam value="#arguments.email#" cfsqltype="cf_sql_varchar"> 
		</cfquery>
		<cfreturn local.qcheckEmailExists.recordCount>
	</cffunction>

	<cffunction name="uploadExcel111" access="remote" returntype="string">
		<cfargument name="file" type="any" required="true">	
		
			<cffile action="upload" fileField="file" destination="#expandPath('ExcelUploads/Result/tem.xlsx')#" nameConflict="MakeUnique">
			
		
		<cfreturn "done">
	</cffunction>

	<cffunction name="uploadExcel" access="remote">		
		<cfargument name="fileUpload" type="any" required="true">
		<!--- Setting Unique Name for file --->
		<cfset local.timestamp = DateFormat(now(), "yyyy-mm-dd HH-MM-ss")>
		<cfset local.uniqueFilename = "Excel_#local.timestamp#.xlsx"> 
		<!--- Upload the file --->		
		<cffile action="upload" fileField="fileUpload" destination="#expandPath('ExcelUploads/')##local.uniqueFilename#" nameConflict="MakeUnique">
		<cfset filePath = "#expandPath('ExcelUploads/')##cffile.serverFile#">
		<!--- Read the uploaded spreadsheet --->
		<cfspreadsheet action="read" src="#filePath#" query="local.excelData" >	
		<!---Creating a excel --->
		<cfset local.mySpreadsheet = spreadsheetNew("Sheet1",true)>
		<cfset spreadsheetAddRow(local.mySpreadsheet, 'First Name,Last Name,Address,Email,Phone,DOB,Role,Result,Reason')>
		<cfset local.headerFormat = {}>
		<cfset local.headerFormat.bold = "true">
		<cfset spreadsheetFormatRow(local.mySpreadsheet, local.headerFormat, 1)> 				
		<cfset local.rowNum = 2>
		<!--End-->
		<cfif local.excelData.recordCount GT 1>	
			
			<cfloop query="local.excelData" startrow="2">
				<!---Validating Data--->	
				<cfset local.errorFlag=0>	
				<cfset local.errorEmail=0>
				<cfset local.errorMssg="">	
				<cfset local.datepattern = '[0-3][0-9]/[0-1][0-9]/[0-2][0-9][0-9][0-9]'>				
				<cfif Len(trim(local.excelData.COL_1)) EQ 0 OR not isValid("regex", local.excelData.COL_1, "^[a-zA-Z]+$")>
					<cfset local.errorFlag=1>
					<cfset local.errorMssg = "First Name Cant be Null And Should Contain Only characters.">
				<cfelseif(Len(trim(local.excelData.COL_1)) GT 50)>
					<cfset local.errorFlag=1>
					<cfset local.errorMssg = "Maximum length of characters permitted for firastname is 50.">					
				</cfif>
				<cfif Len(trim(local.excelData.COL_2)) EQ 0 OR not isValid("regex", local.excelData.COL_2, "^[a-zA-Z]+$")>
					<cfset local.errorFlag=1>
					<cfset local.errorMssg &= "Last Name Cant be Null And Should Contain Only Characters">
				<cfelseif(Len(trim(local.excelData.COL_2)) GT 50)>
					<cfset local.errorFlag=1>
					<cfset local.errorMssg = "Maximum length of characters permitted for last name is 50.">					
				</cfif>
				<cfif Len(trim(local.excelData.COL_3)) EQ 0>
					<cfset local.errorFlag=1>
					<cfset local.errorMssg &= "Address Cant be Null.">
				<cfelseif(Len(trim(local.excelData.COL_3)) GT 200)>
					<cfset local.errorFlag=1>
					<cfset local.errorMssg = "Maximum length of characters permitted for Address is 200.">				   
				</cfif>
				<cfif Len(trim(local.excelData.COL_4)) EQ 0>
					<cfset local.errorFlag=1>
					<cfset local.errorMssg &= "Email Cant be Null.">
				<cfelseif checkEmailExists(local.excelData.COL_4) EQ 1>
					<cfset local.errorEmail=1>
				<cfelseif NOT isValid("email", local.excelData.COL_4)>
					<cfset local.errorFlag=1>
					<cfset local.errorMssg &= "Enter Valid Email.">					
				</cfif>
				<cfif Len(trim(local.excelData.COL_5)) EQ 0>
					<cfset local.errorFlag=1>
					<cfset local.errorMssg &= "Phone Cant be Null.">
				<cfelseif NOT(REFind("^\d{10}$", local.excelData.COL_5)) OR NOT(Len(local.excelData.COL_5) eq 10)>
					<cfset local.errorFlag=1>
					<cfset local.errorMssg &= "Enter Valid Phone No with 10 digits.">					
				</cfif>
				<cfif Len(trim(local.excelData.COL_6)) EQ 0>
					<cfset local.errorFlag=1>
					<cfset local.errorMssg &= "Date Of Birth Cant be Null.">
				<cfelseif NOT isValid("date", local.excelData.COL_6)><!---OR (REFind(local.datepattern, local.slicedData.COL_6) EQ 0)--->
					<cfset local.errorFlag=1>
					<cfset local.errorMssg &= "Invallid Date.Enter Date in(DD-MM-YYYY).">					
				</cfif>
				<cfif Len(trim(local.excelData.COL_7)) EQ 0>
					<cfset local.errorFlag=1>
					<cfset local.errorMssg &= "Role Cant be Null.">
				</cfif>
				<cfset local.roleArray = ListToArray(local.excelData.COL_7, ",")>
				<!---checking if roles Exists--->	
				<cfset local.roleIDArray = []> 		
				<cfloop array="#local.roleArray#" index="item">
					<cfquery name="local.qGetRoleID" datasource="#application.datasoursename#">
						SELECT roleid
						FROM role
						WHERE rolename = <cfqueryparam value="#item#" cfsqltype="cf_sql_varchar">
					</cfquery>
					<cfset local.roleID = local.qGetRoleID.roleid>
					<cfif Len(trim(local.roleID)) EQ 0>
						<cfset local.errorFlag=1>
						<cfset local.errorMssg &= "Select Predefined roles">
						<cfbreak>
					<cfelse>						
        				<cfset arrayAppend(local.roleIDArray, local.roleID)> 
					</cfif>
				</cfloop>				
				<!---Inserting to user table--->
				<cfif local.errorFlag EQ 0 AND local.errorEmail EQ 0>						
					<cfquery name="local.qInsertUserDetails" datasource="#application.datasoursename#" result="local.rInsertUserDetails">
						INSERT
						INTO 
						user (firstname,lastname,address,email,phone,dob)
						VALUES (<cfqueryparam value="#local.excelData.COL_1#" cfsqltype="cf_sql_varchar">,
								<cfqueryparam value="#local.excelData.COL_2#" cfsqltype="cf_sql_varchar">,
								<cfqueryparam value="#local.excelData.COL_3#" cfsqltype="cf_sql_varchar">,
								<cfqueryparam value="#local.excelData.COL_4#" cfsqltype="cf_sql_varchar">,
								<cfqueryparam value="#local.excelData.COL_5#" cfsqltype="cf_sql_varchar">,
								<cfqueryparam value="#DateFormat(local.excelData.COL_6, "yyyy-mm-dd")#" cfsqltype="cf_sql_date">				
							) 							
					</cfquery>
					<!---Inserting to Excel--->
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, local.excelData.COL_1, local.rowNum, 1)>
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, local.excelData.COL_2, local.rowNum, 2)>
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, local.excelData.COL_3, local.rowNum, 3)>
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, local.excelData.COL_4, local.rowNum, 4)>
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, local.excelData.COL_5, local.rowNum, 5)>
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, local.excelData.COL_6, local.rowNum, 6)>
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, local.excelData.COL_7, local.rowNum, 7)>
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, "Added", local.rowNum, 8)>
					<cfset local.rowNum = local.rowNum+1>
					<!---End--->
					<cfset local.lastInsertedID= local.rInsertUserDetails.generatedkey>													
					<!---Inserting to user role table--->
					<cfloop array="#local.roleIDArray#" index="local.item">											
						<cfquery name="local.qInsertUserRoles" datasource="#application.datasoursename#">
							INSERT
							INTO 
							userroles (userid,roleid)
							VALUES (<cfqueryparam value="#local.lastInsertedID#" cfsqltype="cf_sql_integer">,
									<cfqueryparam value="#local.item#" cfsqltype="cf_sql_integer"> 
								) 
						</cfquery>						
					</cfloop>
				<cfelseif local.errorFlag EQ 0 AND local.errorEmail EQ 1>
					<!---Updating User Table if email Exists--->
					<cfquery name="local.qGetUserId" datasource="#application.datasoursename#">
						SELECT userid FROM user WHERE email=<cfqueryparam value="#local.excelData.COL_4#" cfsqltype="cf_sql_varchar">
					</cfquery>
					<cfquery name="local.qUpdateUserTable" datasource="#application.datasoursename#">
						UPDATE user 
						SET firstname=<cfqueryparam value="#local.excelData.COL_1#" cfsqltype="cf_sql_varchar">,
							lastname=<cfqueryparam value="#local.excelData.COL_2#" cfsqltype="cf_sql_varchar">,
							address=<cfqueryparam value="#local.excelData.COL_3#" cfsqltype="cf_sql_varchar">,
							phone=<cfqueryparam value="#local.excelData.COL_5#" cfsqltype="cf_sql_varchar">,
							dob=<cfqueryparam value="#DateFormat(local.excelData.COL_6, "yyyy-mm-dd")#" cfsqltype="cf_sql_date">							
							WHERE userid=<cfqueryparam value="#qGetUserId.userid#" cfsqltype="cf_sql_integer">
					</cfquery>
					<!---Updating userRole Table--->
					<!---Selecting from userrole table--->
					<cfquery name="local.qSelectUserRoles" datasource="#application.datasoursename#">						
						SELECT roleid
						FROM userroles
						WHERE userid = <cfqueryparam value="#local.qGetUserId.userId#" cfsqltype="cf_sql_integer">
						AND roleid NOT IN (							
							<cfqueryparam value="#ArrayToList(local.roleIDArray)#" cfsqltype="cf_sql_varchar">
						)
					</cfquery>
					<cfloop query="#local.qSelectUserRoles#">	
							<cfquery name="local.qDeleteUserRoles" datasource="#application.datasoursename#">
								DELETE 
								FROM userroles
								WHERE userid=<cfqueryparam value="#qGetUserId.userid#" cfsqltype="cf_sql_integer">
								AND roleid=<cfqueryparam value="#local.qSelectUserRoles.roleid#" cfsqltype="cf_sql_integer">
							</cfquery>
					</cfloop>				
                   <cfloop array="#local.roleIDArray#" index="local.item">
					    <cfquery name="local.qcheckrole" datasource="#application.datasoursename#">
							SELECT roleid 
							FROM userroles
							WHERE roleid=<cfqueryparam value="#local.item#" cfsqltype="cf_sql_integer">
						</cfquery>
						<cfif local.qcheckrole.recordCount EQ 0>
							<cfquery name="local.qInsertUserRoles" datasource="#application.datasoursename#">
								INSERT
								INTO 
								userroles (userid,roleid)
								VALUES (<cfqueryparam value="#local.qGetUserId.userid#" cfsqltype="cf_sql_integer">,
										<cfqueryparam value="#local.item#" cfsqltype="cf_sql_integer"> 
									) 
							</cfquery>
						</cfif>							
					</cfloop>
					<!---Inserting to Excel--->
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, local.excelData.COL_1, local.rowNum, 1)>
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, local.excelData.COL_2, local.rowNum, 2)>
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, local.excelData.COL_3, local.rowNum, 3)>
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, local.excelData.COL_4, local.rowNum, 4)>
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, local.excelData.COL_5, local.rowNum, 5)>
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, local.excelData.COL_6, local.rowNum, 6)>
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, local.excelData.COL_7, local.rowNum, 7)>
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, "Updated", local.rowNum, 8)>
					<cfset local.rowNum = local.rowNum+1>
					<!---End--->					
				<cfelse>
					<!---Inserting to Excel--->
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, local.excelData.COL_1, local.rowNum, 1)>
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, local.excelData.COL_2, local.rowNum, 2)>
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, local.excelData.COL_3, local.rowNum, 3)>
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, local.excelData.COL_4, local.rowNum, 4)>
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, local.excelData.COL_5, local.rowNum, 5)>
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, local.excelData.COL_6, local.rowNum, 6)>
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, local.excelData.COL_7, local.rowNum, 7)>
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, "Failed", local.rowNum, 8)>
					<cfset spreadsheetSetCellValue(local.mySpreadsheet, local.errorMssg, local.rowNum, 9)>
					<cfset local.rowNum = local.rowNum+1>
					<!---End--->					
				</cfif>
			</cfloop>									
			<cfset local.resultMsg=" Data to upload">
		<cfelse>
			<cfset local.resultMsg="No Data to upload">
		</cfif>
		<cfheader name="Content-Disposition" value="inline;filename=Data.xlsx">
		<cfcontent  variable="#spreadsheetReadBinary(local.mySpreadsheet)#" type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"> 
	</cffunction>
</cfcomponent>
