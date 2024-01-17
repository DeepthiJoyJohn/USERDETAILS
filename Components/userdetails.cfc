<cfcomponent> 	
	<cffunction name="getUserDetails" access="public" returntype="query">
		<cfset local.qgetUserDetails = QueryNew("userid, firstname, lastname, address, email, phone, dob, dobdisplay, rolenames")>
		<cftry>		
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
			<cfcatch type="any">        
				<cfoutput>#cfcatch.message#in qgetUserDetails</cfoutput>
			</cfcatch>	
		</cftry>
		<cfreturn local.qgetUserDetails>	
	</cffunction>
	
	<cffunction name="checkEmailExists" access="public" returntype="any">
		<cfargument name="email">		
		<cftry>
			<cfquery name="local.qcheckEmailExists" datasource="#application.datasoursename#">
				SELECT COALESCE(userid, 0) AS userid
				FROM user 
				WHERE email=<cfqueryparam value="#arguments.email#" cfsqltype="cf_sql_varchar"> 
			</cfquery>
			<cfreturn local.qcheckEmailExists.userid>
			<cfcatch type="database">        
				<cfoutput>#cfcatch.message#</cfoutput>     
			</cfcatch>
			<cfcatch type="any">			
				<cfoutput>#cfcatch.message#"in checkEmailExists"</cfoutput>
			</cfcatch>
		</cftry>
	</cffunction>	

	<cffunction name="uploadExcel" access="remote" returntype="any">		
		<cfargument name="fileUpload" type="any" required="true">
		<!--- Setting Unique Name for file --->
		<cfset local.timestamp = DateFormat(now(), "yyyy-mm-dd HH-MM-ss")>
		<cfset local.uniqueFilename = "Excel_#local.timestamp#.xlsx"> 
		<!---End--->
		<!--- Upload the file --->				
		<cftry>				
			<!---checking if file is there--->
			<cfif NOT IsDefined("fileUpload") OR fileUpload EQ "">
        		<cfthrow message="No file selected for upload.">			
			</cfif>			
			<cffile action="upload" fileField="fileUpload" destination="#expandPath('ExcelUploads/')##local.uniqueFilename#" nameConflict="MakeUnique">
			<cfset local.filePath = "#expandPath('ExcelUploads/')##cffile.serverFile#">
			<cfcatch type="any">
				<cfthrow message="An error occurred during file upload.">
			</cfcatch>
		</cftry>
		<!---End--->
		<!--- Read the uploaded spreadsheet --->
		<cftry>		
			<cfset local.fileExt = listLast(cffile.serverfile,".")>						
			<cfif NOT FileExists(local.filePath)>			
				<cfthrow message="File not exists.">
			<cfelseif local.fileExt NEQ "xls" AND local.fileExt NEQ "xlsx">
				<cfthrow message="Not a xls or xlsx file.">			
			</cfif>
			<cfspreadsheet action="read" src="#filePath#" query="local.excelData">
			<cfcatch type="any">
				<cfthrow message="An error occurred during the spreadsheet reading process: #cfcatch.message#">
			</cfcatch>
		</cftry>			
		<!---End--->
		<!---Creating a excel --->
		<cfset local.mySpreadsheet = spreadsheetNew("Sheet1",true)>
		<cfset spreadsheetAddRow(local.mySpreadsheet, 'First Name,Last Name,Address,Email,Phone,DOB,Role,Result,Reason')>
		<cfset local.headerFormat = {}>
		<cfset local.headerFormat.bold = "true">
		<cfset spreadsheetFormatRow(local.mySpreadsheet, local.headerFormat, 1)> 				
		<cfset local.rowNum = 2>
		<!--End-->
		<cftry>
			<!---Creating Array to store Excel data--->
			<cfset local.dataArrayInserted = []>			
			<cfset local.dataArrayError = []>
			<cfset local.roleIdList="">
			<!---End--->
			<cfif local.excelData.recordCount GT 1>				
				<cfloop query="local.excelData" startrow="2">
					<!---Validating Data--->	
					<cfset local.errorFlag=0>
					<cfset local.errorMssg="">						
					<cfif Len(trim(local.excelData.COL_1)) EQ 0 OR not isValid("regex", local.excelData.COL_1, "^[a-zA-Z]+$")>
						<cfset local.errorFlag=1>
						<cfset local.errorMssg = "First Name Cant be Null And Should Contain Only characters.">
					<cfelseif(Len(trim(local.excelData.COL_1)) GT 50)>
						<cfset local.errorFlag=1>
						<cfset local.errorMssg = "Maximum length of characters permitted for firstname is 50.">					
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
					<cfelse>	
						<!---checking if roles Exists--->							
						<cfloop list="#local.excelData.COL_7#" index="local.currentItem">							
							<cfquery name="local.qGetRoleID" datasource="#application.datasoursename#">
								SELECT roleid
								FROM role
								WHERE rolename = <cfqueryparam value="#local.currentItem#" cfsqltype="cf_sql_varchar">
							</cfquery>
							<cfset local.roleID = local.qGetRoleID.roleid>
					
							<cfif Len(trim(local.roleID)) EQ 0>
								<cfset local.errorFlag=1>
								<cfset local.errorMssg &= "Select Predefined roles">
								<cfbreak>	
							<cfelse>
								<cfset local.roleIdList &= local.roleID & "," >						 
							</cfif>							
						</cfloop>	
						<cfset local.roleIdList=left(trim(local.roleIdList),len(trim(local.roleIdList))-1)>
					</cfif>						
					<!---Inserting row data of excel to structure--->
					<cfset local.rowData = {
						"Column1" = local.excelData.COL_1,
						"Column2" = local.excelData.COL_2,
						"Column3" = local.excelData.COL_3,
						"Column4" = local.excelData.COL_4,
						"Column5" = local.excelData.COL_5,
						"Column6" = DateFormat(local.excelData.COL_6, "dd-mm-yyyy"),
						"Column7" = local.excelData.COL_7,
						"Column8" = local.errorMssg,
						"Column9" = ""
					}>						
					<!---Inserting to user table--->		
					<cfset local.userId=checkEmailExists(local.excelData.COL_4)>					
					<cfset local.lastUpdatedID = 0>								
					<cfif local.errorFlag EQ 0 AND Len(trim(local.userId)) EQ 0>
						<cftry>
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
						<cfcatch>
							<cfoutput>Error qInsertUserDetails</cfoutput>
						</cfcatch>	
					    </cftry>							
						<cfif local.rInsertUserDetails.keyExists('GENERATEDKEY')>
							<cfset local.lastInsertedID= local.rInsertUserDetails.generatedkey>	
							<!---Inserting to user role table--->
							<cftry>
								<cfloop list="#local.roleIdList#" index="local.item">											
									<cfquery name="local.qInsertUserRoles" datasource="#application.datasoursename#">
										INSERT
										INTO 
										userroles (userid,roleid)
										VALUES (<cfqueryparam value="#local.lastInsertedID#" cfsqltype="cf_sql_integer">,
												<cfqueryparam value="#local.item#" cfsqltype="cf_sql_integer"> 
											) 
									</cfquery>						
								</cfloop>
							<cfcatch>								
																						
							</cfcatch>
							</cftry>
							<cfset local.rowData["Column9"] = "Added">
							<!---Inserting to Array--->	
							<cfset arrayAppend(local.dataArrayInserted, local.rowData)>
							<!---End--->
						</cfif>						
					<cfelseif local.errorFlag EQ 0 AND Len(trim(local.userId)) NEQ 0>
						<!---Updating User Table if email Exists--->
						<cftry>
							<cfquery name="local.qUpdateUserTable" datasource="#application.datasoursename#" result="local.rUpdateUserTable">
								UPDATE user 
								SET firstname=<cfqueryparam value="#local.excelData.COL_1#" cfsqltype="cf_sql_varchar">,
									lastname=<cfqueryparam value="#local.excelData.COL_2#" cfsqltype="cf_sql_varchar">,
									address=<cfqueryparam value="#local.excelData.COL_3#" cfsqltype="cf_sql_varchar">,
									phone=<cfqueryparam value="#local.excelData.COL_5#" cfsqltype="cf_sql_varchar">,
									dob=<cfqueryparam value="#DateFormat(local.excelData.COL_6, "yyyy-mm-dd")#" cfsqltype="cf_sql_date">							
									WHERE userid=<cfqueryparam value="#local.userId#" cfsqltype="cf_sql_integer">
							</cfquery>
						<cfcatch>
							<cfoutput>Error qUpdateUserTable</cfoutput>
						</cfcatch>	
						</cftry>						
						<!---Updating userRole Table--->
						<!---Selecting from userrole table--->											
						<cftry>
							<cfquery name="local.qSelectUserRoles" datasource="#application.datasoursename#">						
								SELECT roleid
								FROM userroles
								WHERE userid = <cfqueryparam value="#local.userId#" cfsqltype="cf_sql_integer">
								AND roleid NOT IN (			
									<cfqueryparam value=#local.roleIdList# cfsqltype="cf_sql_varchar">)
							</cfquery>							
						<cfcatch>
							<cfoutput>Error qSelectuserroles</cfoutput>
						</cfcatch>
						</cftry>						
						<cftry>
							<cfif local.qSelectUserRoles.recordCount>							
								<cfloop query="#local.qSelectUserRoles#">	
										<cfquery name="local.qDeleteUserRoles" datasource="#application.datasoursename#">
											DELETE 
											FROM userroles
											WHERE userid=<cfqueryparam value="#local.userId#" cfsqltype="cf_sql_integer">
											AND roleid=<cfqueryparam value="#local.qSelectUserRoles.roleid#" cfsqltype="cf_sql_integer">
										</cfquery>
								</cfloop>	
							</cfif>
						<cfcatch><cfoutput>Error Delete qSelectUserRoles</cfoutput></cfcatch>
						</cftry>
						
						<!---inseritng to userroles table where roleid is not there from excel upload--->
						<cftry>
							<cfquery name="local.qcheckrole" datasource="#application.datasoursename#">
								 SELECT roleid FROM ROLE 
									WHERE roleid IN 
										(
											<cfqueryparam value="#local.roleIdList#" cfsqltype="cf_sql_varchar" list="true">
										)
									AND roleid NOT IN 
										(
											SELECT roleid FROM userroles WHERE userid = <cfqueryparam value="#local.userId#" cfsqltype="cf_sql_integer">
										)
							</cfquery>	
							<cfcatch>
								<cfoutput>Error database qcheckrole</cfoutput>
							</cfcatch>
						</cftry>	
						<cfif local.qcheckrole.recordCount GT 0>
							<cftry>								
								<cfloop query="#local.qcheckrole#">
									<cfquery name="local.qInsertUserRoles" datasource="#application.datasoursename#">
										INSERT
										INTO 
										userroles (userid,roleid)
										VALUES (<cfqueryparam value="#local.userId#" cfsqltype="cf_sql_integer">,
												<cfqueryparam value="#local.qcheckrole.roleid#" cfsqltype="cf_sql_integer"> 
											) 
									</cfquery>																	
								</cfloop>								
							<cfcatch>
								<cfoutput>Error database qInsertUserRoles</cfoutput>
							</cfcatch>
							</cftry>
						</cfif>						
						<!---Inserting to Array--->
						<cfset local.rowData["Column9"] = "Updated">
						<cfset arrayAppend(local.dataArrayInserted, local.rowData)>
						<!---End--->	
					<cfelse>
						<!---Inserting to Array--->
						<cfset local.rowData["Column9"] = "Failed">
						<cfset arrayAppend(local.dataArrayError, local.rowData)>
						<!---End--->					
					</cfif>			
									
				</cfloop>	
			</cfif>
			<cfcatch>				
				<cfthrow message="#cfcatch.message#">
				<cfdump var="#local.rowNum#" abort>
			</cfcatch>
		</cftry>
		<cftry>					
			<!---Writing to the result excel--->
			
			<cfloop array="#local.dataArrayError#" index="row">
				<cfset spreadsheetSetCellValue(local.mySpreadsheet, row.Column1, local.rowNum, 1)>
				<cfset spreadsheetSetCellValue(local.mySpreadsheet, row.Column2, local.rowNum, 2)>
				<cfset spreadsheetSetCellValue(local.mySpreadsheet, row.Column3, local.rowNum, 3)>
				<cfset spreadsheetSetCellValue(local.mySpreadsheet, row.Column4, local.rowNum, 4)>
				<cfset spreadsheetSetCellValue(local.mySpreadsheet, row.Column5, local.rowNum, 5)>
				<cfset spreadsheetSetCellValue(local.mySpreadsheet, row.Column6, local.rowNum, 6)>
				<cfset spreadsheetSetCellValue(local.mySpreadsheet, row.Column7, local.rowNum, 7)>
				<cfset spreadsheetSetCellValue(local.mySpreadsheet, "Failed", local.rowNum, 8)>
				<cfset spreadsheetSetCellValue(local.mySpreadsheet, row.Column8, local.rowNum, 9)>
				<cfset local.rowNum = local.rowNum+1>
			</cfloop>
			
			<cfloop array="#local.dataArrayInserted#" index="row">				
				<cfset spreadsheetSetCellValue(local.mySpreadsheet, row.Column1, local.rowNum, 1)>
				<cfset spreadsheetSetCellValue(local.mySpreadsheet, row.Column2, local.rowNum, 2)>
				<cfset spreadsheetSetCellValue(local.mySpreadsheet, row.Column3, local.rowNum, 3)>
				<cfset spreadsheetSetCellValue(local.mySpreadsheet, row.Column4, local.rowNum, 4)>
				<cfset spreadsheetSetCellValue(local.mySpreadsheet, row.Column5, local.rowNum, 5)>
				<cfset spreadsheetSetCellValue(local.mySpreadsheet, row.Column6, local.rowNum, 6)>
				<cfset spreadsheetSetCellValue(local.mySpreadsheet, row.Column7, local.rowNum, 7)>
				<cfset spreadsheetSetCellValue(local.mySpreadsheet, row.Column9, local.rowNum, 8)>
				<cfset spreadsheetSetCellValue(local.mySpreadsheet, row.Column8, local.rowNum, 9)>
				<cfset local.rowNum = local.rowNum+1>
			</cfloop>	
			<cfspreadsheet action="write" filename="#expandPath('ExcelUploads/Result/')#UploadResult.xlsx" name="local.myspreadsheet" overwrite="true">
		    
			<!---End--->
			<!---Auto Downloading the Result Excel--->
			<!---<cfheader name="Content-Disposition" value="inline;filename=Data.xlsx">
			<cfcontent  variable="#spreadsheetReadBinary(local.mySpreadsheet)#" type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet">---> 						
			<cfreturn "done">
			<cfcatch type="any">				
				<cfthrow message="An error occurred while processing the spreadsheet: #cfcatch.message#">
			</cfcatch>
		</cftry>
	</cffunction>
</cfcomponent>
