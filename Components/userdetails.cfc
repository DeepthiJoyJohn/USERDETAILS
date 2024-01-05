<cfcomponent> 	
	<cffunction name="getUserDetails" access="public" returntype="query">
		<cfargument name="seqNo">
		<cfquery name="qgetUserDetails" datasource="#application.datasoursename#">
			SELECT
				u.userid,
				u.firstname,u.lastname,u.address,u.email,u.phone,u.dob,u.result,
				GROUP_CONCAT(r.rolename, '') AS rolenames
			FROM
				USER u
			INNER JOIN
				userroles ur ON u.userid = ur.userid
			INNER JOIN
				ROLE r ON ur.roleid = r.roleid
			<cfif arguments.seqNo NEQ 0>
				WHERE u.seq=<cfqueryparam value="#arguments.seqNo#" cfsqltype="cf_sql_integer">
			</cfif>
			GROUP BY
				u.userid,
				u.firstname									
		</cfquery>
		<cfreturn qgetUserDetails> 		
	</cffunction>
	<cffunction name="getUserDetailsWithError" access="public" returntype="query">
		<cfargument name="seqNo">
		<cfquery name="qgetUserDetailsWithError" datasource="#application.datasoursename#">
			SELECT *
			FROM exceluploaderror
			WHERE firstname <> "" AND lastname <> "" AND address <> "" AND email <> "" 
			AND phone <> "" AND dob <> "" AND 
			seq=<cfqueryparam value="#arguments.seqNo#" cfsqltype="cf_sql_integer">												
		</cfquery>
		<cfreturn qgetUserDetailsWithError> 		
	</cffunction>
	<cffunction name="checkEmailExists" access="public" returntype="numeric">
		<cfargument name="email">		
		<cfquery name="qcheckEmailExists" datasource="#application.datasoursename#">
			SELECT email
			FROM user 
			WHERE email=<cfqueryparam value="#arguments.email#" cfsqltype="cf_sql_varchar"> 
		</cfquery>
		<cfreturn qcheckEmailExists.recordCount>
	</cffunction>
	
	<cffunction name="uploadExcel" access="remote" returntype="string">
		<cfargument name="fileUpload" type="any" required="true">
		<!--- Setting Unique Name for file --->
		<cfset local.timestamp = DateFormat(now(), "yyyymmdd_HHmmss")>
		<cfset local.uniqueFilename = "Excel_#local.timestamp#.xlsx"> 
		<!--- Upload the file --->		
		<cffile action="upload" fileField="fileUpload" destination="#expandPath('ExcelUploads/')##uniqueFilename#" nameConflict="MakeUnique">
		<cfset filePath = "#expandPath('ExcelUploads/')##uniqueFilename#">
		<!--- Read the uploaded spreadsheet --->
		<cfspreadsheet action="read" src="#filePath#" query="excelData">
		<!---Getting Seq No for Excel Upload--->
		<cfquery name="qGetSeqNo" datasource="#application.datasoursename#">
			SELECT COALESCE(MAX(seq)+1, 1) AS seqno FROM user
		</cfquery>
		<cfset local.seqNo = qGetSeqNo.seqno>
		<!--- Loop through data from row 2 onwards --->
		<cfset local.startRow = 2>
		<cfset local.numRows = excelData.recordCount - local.startRow + 1>
		<cfif excelData.recordCount GT 1>		
			<cfset local.slicedData = QuerySlice(excelData, local.startRow, local.numRows)>			
			<cfloop query="local.slicedData">
				<!---Validating Data--->	
				<cfset local.errorFlag=0>	
				<cfset local.errorEmail=0>
				<cfset local.errorMssg="">	
				<cfset local.datepattern = '[0-3][0-9]/[0-1][0-9]/[0-2][0-9][0-9][0-9]'>		
				
				<cfif Len(trim(local.slicedData.COL_1)) EQ 0>
					<cfset local.errorFlag=1>
					<cfset local.errorMssg = "First Name Cant be Null.">
				</cfif>
				<cfif Len(trim(local.slicedData.COL_2)) EQ 0>
					<cfset local.errorFlag=1>
					<cfset local.errorMssg &= "Last Name Cant be Null.">
				</cfif>
				<cfif Len(trim(local.slicedData.COL_3)) EQ 0>
					<cfset local.errorFlag=1>
					<cfset local.errorMssg &= "Address Cant be Null.">
				</cfif>
				<cfif Len(trim(local.slicedData.COL_4)) EQ 0>
					<cfset local.errorFlag=1>
					<cfset local.errorMssg &= "Email Cant be Null.">
				</cfif>
				<cfif Len(trim(local.slicedData.COL_5)) EQ 0>
					<cfset local.errorFlag=1>
					<cfset local.errorMssg &= "Phone Cant be Null.">
				</cfif>
				<cfif Len(trim(local.slicedData.COL_6)) EQ 0>
					<cfset local.errorFlag=1>
					<cfset local.errorMssg &= "Date Of Birth Cant be Null.">
				</cfif>
				<cfif Len(trim(local.slicedData.COL_7)) EQ 0>
					<cfset local.errorFlag=1>
					<cfset local.errorMssg &= "Role Cant be Null.">
				</cfif>
				<cfif checkEmailExists(local.slicedData.COL_4) EQ 1>
					<cfset local.errorEmail=1>					
				</cfif>
				<cfif NOT isValid("email", local.slicedData.COL_4)>
					<cfset local.errorFlag=1>
					<cfset local.errorMssg &= "Enter Valid Email.">
				</cfif>
				<cfif NOT(REFind("^\d{10}$", local.slicedData.COL_5)) OR NOT(Len(local.slicedData.COL_5) eq 10)>
					<cfset local.errorFlag=1>
					<cfset local.errorMssg &= "Enter Valid Phone No with 10 digits.">
				</cfif>			
				<cfif NOT(isValid("date", local.slicedData.COL_6)) OR Len(local.slicedData.COL_6) neq 10 OR (REFind(local.datepattern, local.slicedData.COL_6) EQ 0)>
					<cfset local.errorFlag=1>
					<cfset local.errorMssg &= "Invallid Date.Enter Date in(DD-MM-YYYY).">
				</cfif>
				<cfset local.roleArray = ListToArray(local.slicedData.COL_7, ",")>
				<!---checking if roles Exists--->			
				<cfloop array="#local.roleArray#" index="item">
					<cfquery name="qGetRoleID" datasource="#application.datasoursename#">
						SELECT roleid
						FROM role
						WHERE rolename = <cfqueryparam value="#item#" cfsqltype="cf_sql_varchar">
					</cfquery>
					<cfset local.roleID = qGetRoleID.roleid>
					<cfif Len(trim(local.roleID)) EQ 0>
						<cfset local.errorFlag=1>
						<cfset local.errorMssg &= "Select Predefined roles">
						<cfbreak>
					</cfif>
				</cfloop>	
				<!---Inserting to user table--->
				<cfif local.errorFlag EQ 0 AND local.errorEmail EQ 0>						
					<cfquery name="qInsertUserDetails" datasource="#application.datasoursename#">
						INSERT
						INTO 
						user (firstname,lastname,address,email,phone,dob,seq,result)
						VALUES (<cfqueryparam value="#local.slicedData.COL_1#" cfsqltype="cf_sql_varchar">,
								<cfqueryparam value="#local.slicedData.COL_2#" cfsqltype="cf_sql_varchar">,
								<cfqueryparam value="#local.slicedData.COL_3#" cfsqltype="cf_sql_varchar">,
								<cfqueryparam value="#local.slicedData.COL_4#" cfsqltype="cf_sql_varchar">,
								<cfqueryparam value="#local.slicedData.COL_5#" cfsqltype="cf_sql_varchar">,
								<cfqueryparam value="#DateFormat(local.slicedData.COL_6, "yyyy-mm-dd")#" cfsqltype="cf_sql_date">,
								<cfqueryparam value="#local.seqNo#" cfsqltype="cf_sql_integer">,
								<cfqueryparam value="Added" cfsqltype="cf_sql_varchar">															
							) 
					</cfquery>
					<!---Getting the last inserted userid--->
					<cfset local.insertedUserID = "">
					<cfquery name="qGetLastID" datasource="#application.datasoursename#">
						SELECT LAST_INSERT_ID() AS last_id
					</cfquery>			
					<cfset local.lastInsertedID = qGetLastID.last_id>			
					<!---Inserting to user role table--->
					<cfloop array="#local.roleArray#" index="item">
						<!---Getting roleid from role table--->
						<cfquery name="qGetRoleID" datasource="#application.datasoursename#">
							SELECT roleid
							FROM role
							WHERE rolename = <cfqueryparam value="#item#" cfsqltype="cf_sql_varchar">
						</cfquery>
						<cfset local.roleID = qGetRoleID.roleid>
						<cfquery name="qInsertUserRoles" datasource="#application.datasoursename#">
							INSERT
							INTO 
							userroles (userid,roleid)
							VALUES (<cfqueryparam value="#local.lastInsertedID#" cfsqltype="cf_sql_varchar">,
									<cfqueryparam value="#local.roleID#" cfsqltype="cf_sql_varchar"> 
								) 
						</cfquery>
					</cfloop>
				<cfelseif local.errorFlag EQ 0 AND local.errorEmail EQ 1>
					<!---Updating User Table if email Exists--->
					<cfquery name="qGetUserId" datasource="#application.datasoursename#">
						SELECT userid FROM user WHERE email=<cfqueryparam value="#local.slicedData.COL_4#" cfsqltype="cf_sql_varchar">
					</cfquery>
					<cfquery name="qUpdateUserTable" datasource="#application.datasoursename#">
						UPDATE user 
						SET firstname=<cfqueryparam value="#local.slicedData.COL_1#" cfsqltype="cf_sql_varchar">,
							lastname=<cfqueryparam value="#local.slicedData.COL_2#" cfsqltype="cf_sql_varchar">,
							address=<cfqueryparam value="#local.slicedData.COL_3#" cfsqltype="cf_sql_varchar">,
							phone=<cfqueryparam value="#local.slicedData.COL_5#" cfsqltype="cf_sql_varchar">,
							dob=<cfqueryparam value="#DateFormat(local.slicedData.COL_6, "yyyy-mm-dd")#" cfsqltype="cf_sql_date">,
							seq=<cfqueryparam value="#local.seqNo#" cfsqltype="cf_sql_integer">,
							result=<cfqueryparam value="Updated" cfsqltype="cf_sql_varchar">
							WHERE userid=<cfqueryparam value="#qGetUserId.userid#" cfsqltype="cf_sql_integer">
					</cfquery>
					<!---Updating userRole Table--->
					<!---Deleting from userrole table--->
					<cfquery name="qDeletingUserRoles" datasource="#application.datasoursename#">
						DELETE
						FROM userroles
						WHERE userid=<cfqueryparam value="#qGetUserId.userid#" cfsqltype="cf_sql_integer">
					</cfquery>
					<cfloop array="#local.roleArray#" index="item">
						<!---Getting roleid from role table--->
						<cfquery name="qGetRoleID" datasource="#application.datasoursename#">
							SELECT roleid
							FROM role
							WHERE rolename = <cfqueryparam value="#item#" cfsqltype="cf_sql_varchar">
						</cfquery>
						<cfset local.roleID = qGetRoleID.roleid>						
						<!---Inserting to User Roles--->
						<cfquery name="qInsertUserRoles" datasource="#application.datasoursename#">
							INSERT
							INTO 
							userroles (userid,roleid)
							VALUES (<cfqueryparam value="#qGetUserId.userid#" cfsqltype="cf_sql_integer">,
									<cfqueryparam value="#local.roleID#" cfsqltype="cf_sql_integer"> 
								) 
						</cfquery>					
					</cfloop>
				<cfelse>
					<!---Inserting to error table--->					
					<cfquery name="qInsertErrorTable" datasource="#application.datasoursename#">
						INSERT
						INTO 
						exceluploaderror(firstname,lastname,address,email,phone,dob,seq,result,roles)
						VALUES (<cfqueryparam value="#local.slicedData.COL_1#" cfsqltype="cf_sql_varchar">,
								<cfqueryparam value="#local.slicedData.COL_2#" cfsqltype="cf_sql_varchar">,
								<cfqueryparam value="#local.slicedData.COL_3#" cfsqltype="cf_sql_varchar">,
								<cfqueryparam value="#local.slicedData.COL_4#" cfsqltype="cf_sql_varchar">,
								<cfqueryparam value="#local.slicedData.COL_5#" cfsqltype="cf_sql_varchar">,
								<cfqueryparam value="#local.slicedData.COL_6#" cfsqltype="cf_sql_varchar">,
								<cfqueryparam value="#local.seqNo#" cfsqltype="cf_sql_integer">,	
								<cfqueryparam value="#local.errorMssg#" cfsqltype="cf_sql_varchar">,
								<cfqueryparam value="#local.slicedData.COL_7#" cfsqltype="cf_sql_varchar">																													
							) 
					</cfquery>
				</cfif>
			</cfloop>
			<cfset generateResultExcel(local.seqNo)>			
			<cfset local.resultMsg="File uploaded successfully!">
		<cfelse>
			<cfset local.resultMsg="No Data to upload">
		</cfif>
		<cfreturn local.resultMsg>
	</cffunction>
	<cffunction name="generateResultExcel" access="public">
		<cfargument name="seqNo">
		<cfset local.mySpreadsheet = spreadsheetNew()>
		<cfset spreadsheetAddRow(local.mySpreadsheet, 'First Name,Last Name,Address,Email,Phone,DOB,Role,Result,Reason')>
		<cfset local.headerFormat = {}>
		<cfset local.headerFormat.bold = "true">
		<cfset spreadsheetFormatRow(local.mySpreadsheet, local.headerFormat, 1)> 
		<cfset local.userObject = createObject("component", "Components.userdetails")>
		<cfset local.resultUserDetailsError = local.userObject.getUserDetailsWithError(arguments.seqNo)>
		<cfloop query="local.resultUserDetailsError">
			<cfset local.rolenames ="'#local.resultUserDetailsError.roles#'">        
			<cfset local.combinedValues = '#local.resultUserDetailsError.firstname#,#local.resultUserDetailsError.lastname#,
			#local.resultUserDetailsError.address#,#local.resultUserDetailsError.email#,#local.resultUserDetailsError.phone#,
			#local.resultUserDetailsError.dob#,#local.rolenames#,Failed,#local.resultUserDetailsError.result#'>
			<cfset spreadsheetAddRow(local.mySpreadsheet,local.combinedValues)>
		</cfloop>		
		<cfset local.resultUserDetails = local.userObject.getUserDetails(arguments.seqNo)>
		<cfloop query="local.resultUserDetails">
			<cfset local.rolenames ="'#local.resultUserDetails.rolenames#'">        
			<cfset local.combinedValues = '#local.resultUserDetails.firstname#,#local.resultUserDetails.lastname#,
			#local.resultUserDetails.address#,#local.resultUserDetails.email#,#local.resultUserDetails.phone#,
			#DateFormat(local.resultUserDetails.dob, "MM/DD/YYYY")#,#local.rolenames#,#local.resultUserDetails.result#'>
			<cfset spreadsheetAddRow(local.mySpreadsheet,local.combinedValues)>
		</cfloop>		
		<cfset local.timestamp = DateFormat(now(), "yyyymmdd_HHmmss")>
		<cfset local.uniqueFilename = "ExcelTemplateData_#local.timestamp#.xlsx">
		<cfset Spreadsheetwrite(local.mySpreadsheet,'#expandPath('ExcelTemplate/DataExcel/')##local.uniqueFilename#',true)>
		<cfheader name="Content-Disposition" value="attachment;filename=#uniqueFilename#">
		<cfcontent file="#expandPath('ExcelTemplate/DataExcel/')##uniqueFilename#" type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" reset="true"> 				
	</cffunction>	
</cfcomponent>

