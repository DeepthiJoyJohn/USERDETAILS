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
		<cfset local.timestamp = DateFormat(now(), "yyyymmdd_HHmmss")>
		<cfset local.uniqueFilename = "Excel_#local.timestamp#.xlsx">		 
        <cffile action="upload" fileField="fileUpload" destination="#expandPath('ExcelUploads/')##uniqueFilename#" nameConflict="MakeUnique">
		<cfset filePath = "#expandPath('ExcelUploads/')##uniqueFilename#">
		<cfspreadsheet action="read" src="#filePath#" query="excelData">
			<cfoutput query="excelData">
        		<p>User ID: #excelData.COL_1#</p>
    		</cfoutput>

        <cfreturn "File uploaded successfully!">
	</cffunction>
</cfcomponent>

