<cfcomponent> 	
	<cffunction name="getUserDetails" access="public" returntype="query">
		<cfquery name="qgetUserDetails" datasource="#application.datasoursename#">
			SELECT
				u.userid,
				u.firstname,u.lastname,u.address,u.email,u.phone,u.dob,
				GROUP_CONCAT(r.rolename, '') AS rolenames
			FROM
				USER u
			INNER JOIN
				userroles ur ON u.userid = ur.userid
			INNER JOIN
				ROLE r ON ur.roleid = r.roleid
			GROUP BY
				u.userid,
				u.firstname									
		</cfquery>
		<cfreturn qgetUserDetails> 		
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
		<cfset local.slicedData = QuerySlice(excelData, local.startRow, local.numRows)>			
		<cfloop query="local.slicedData">
			<!---Validating Data--->	
			<cfset local.errorFlag=0>	
			<cfset local.errorMssg="">		
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
			<!---Inserting to user table--->
			<cfif local.errorFlag EQ 0>
				<cfset local.roleArray = ListToArray(local.slicedData.COL_7, ",")>			
				<cfquery name="qInsertUserDetails" datasource="#application.datasoursename#">
					INSERT
					INTO 
					user (firstname,lastname,address,email,phone,dob,seq)
					VALUES (<cfqueryparam value="#local.slicedData.COL_1#" cfsqltype="cf_sql_varchar">,
							<cfqueryparam value="#local.slicedData.COL_2#" cfsqltype="cf_sql_varchar">,
							<cfqueryparam value="#local.slicedData.COL_3#" cfsqltype="cf_sql_varchar">,
							<cfqueryparam value="#local.slicedData.COL_4#" cfsqltype="cf_sql_varchar">,
							<cfqueryparam value="#local.slicedData.COL_5#" cfsqltype="cf_sql_varchar">,
							<cfqueryparam value="#DateFormat(local.slicedData.COL_6, "yyyy-mm-dd")#" cfsqltype="cf_sql_date">,
							<cfqueryparam value="#local.seqNo#" cfsqltype="cf_sql_integer">															
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
			<cfelse>
				<!---Inserting to error table--->	
				<cfquery name="qInsertErrorTable" datasource="#application.datasoursename#">
					INSERT
					INTO 
					exceluploaderror(firstname,lastname,address,email,phone,dob,seq,result)
					VALUES (<cfqueryparam value="#local.slicedData.COL_1#" cfsqltype="cf_sql_varchar">,
							<cfqueryparam value="#local.slicedData.COL_2#" cfsqltype="cf_sql_varchar">,
							<cfqueryparam value="#local.slicedData.COL_3#" cfsqltype="cf_sql_varchar">,
							<cfqueryparam value="#local.slicedData.COL_4#" cfsqltype="cf_sql_varchar">,
							<cfqueryparam value="#local.slicedData.COL_5#" cfsqltype="cf_sql_varchar">,
							<cfqueryparam value="#DateFormat(local.slicedData.COL_6, "yyyy-mm-dd")#" cfsqltype="cf_sql_date">,
							<cfqueryparam value="#local.seqNo#" cfsqltype="cf_sql_integer">,	
							<cfqueryparam value="#local.errorMssg#" cfsqltype="cf_sql_varchar">																													
						) 
				</cfquery>
			</cfif>
		</cfloop>
		<!---<cfset local.result=generateResultExcel()>--->
		<cfreturn "File uploaded successfully!">
	</cffunction>
	<cffunction name="generateResultExcel" access="public">
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
		
	</cffunction>
</cfcomponent>

