<html>
    <head>
        <title>User Details Excel Upload</title>
        <link rel="stylesheet" href="css/userdetails.css"> 
        <link rel="stylesheet" href="font-awesome-4.7.0/css/font-awesome.css">
        <link rel="stylesheet" href="font-awesome-4.7.0/css/font-awesome.min.css">  
        <script src="js/jquery-3.6.0.min.js"></script>         
        <script src="js/userdetails.js"></script>  
    </head>
    <body> 
        <cfoutput>            
            <form action="" method="post" id="myForm" name="myForm" enctype="multipart/form-data">
                <cfset local.spreadSheetObj="">                
                <cfset local.userObject=createObject("component", "Components.userdetails")>  
                <cfif  StructKeyExists(form,"uploadBtn") && NOT IsNull(form.fileUpload)> 
                    <cfset local.spreadSheetObj=local.userObject.uploadExcel(form.fileUpload)>  
                </cfif>  
                <cfset local.resultUserDetails=local.userObject.getUserDetails()>               
                <div class="heading">USER INFORMATION</div>
                <div class="btnDiv">
                    <div class="btnLeft">
                        <button class="plainTemplate" onclick="generateExcelTemplate('plain')">Plain Template</button>
                        <button class="templateWithData" onclick="generateExcelTemplate('data')">Template With Data</button>
                    </div>
                    <div class="btnRight">
                        <button class="browse" onclick="browse()">Browse<input type="file" name="fileUpload" id="fileUpload" class="fileUpload" required="yes" accept=".xlsx, .xls" /></button>                        
                        <div class="selectedFileInfo" id="selectedFileInfo"></div>  
                        <button class="upload"  type="submit" onclick="checkFileExists()" name="uploadBtn">Upload</button>
                    </div>
                </div>                   
                <div class="tableDiv">   
                    <span class="spanTableHeading">Table</i></span>             
                    <table class="table">
                        <tr>
                            <th>First Name</th>
                            <th>Last Name</th>
                            <th>Address</th>
                            <th>Email</th>
                            <th>Phone</th>
                            <th>DOB</th>
                            <th>Role</th>
                        </tr>                        
                        <cfloop query="#local.resultUserDetails#">
                            <tr>
                                <td>#local.resultUserDetails.firstname#</td>
                                <td>#local.resultUserDetails.lastname#</td>
                                <td>#local.resultUserDetails.address#</td>
                                <td>#local.resultUserDetails.email#</td>
                                <td>#local.resultUserDetails.phone#</td>
                                <td>#local.resultUserDetails.dobdisplay#</td>
                                <td>#local.resultUserDetails.rolenames#</td> 
                            </tr>                           
                        </cfloop>                                      
                    </table>
                </div>
            </form>  
            <cfif local.spreadSheetObj EQ "done">
                <meta http-equiv="refresh" content="0;url=downloadExcelTemplate.cfm">
            </cfif>              
            <div class="result" id="result"></div>         
        </cfoutput>          
    </body>     
</html>