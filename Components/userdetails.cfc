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
</cfcomponent>

