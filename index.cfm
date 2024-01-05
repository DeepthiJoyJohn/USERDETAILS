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
            <form action="" method="post" id="myForm" enctype="multipart/form-data" >
                <cfset local.userObject=createObject("component", "Components.userdetails")>
                <cfset local.resultUserDetails=local.userObject.getUserDetails(0)>
                <div class="heading">USER INFORMATION</div>
                <div class="btnDiv">
                    <div class="btnLeft">
                        <button class="plainTemplate" onclick="generateExcelTemplate()">Plain Template</button>
                        <button class="templateWithData" onclick="generateDataExcelTemplate()">Template With Data</button>
                    </div>
                    <div class="btnRight">
                        <button class="browse" onclick="browse()">Browse<input type="file" name="fileUpload" id="fileUpload" class="fileUpload" required="yes" accept=".xlsx, .xls" /></button>                        
                        <div class="selectedFileInfo" id="selectedFileInfo"></div>  
                        <button class="upload" onclick="upload()" type="submit" name="uploadBtn">Upload</button>
                    </div>
                </div>   
                <div class="tableDiv">   
                    <span class="spanTableHeading">Table</span>             
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
                        <cfloop query="local.resultUserDetails">
                            <tr>
                                <td>#firstname#</td>
                                <td>#lastname#</td>
                                <td>#address#</td>
                                <td>#email#</td>
                                <td>#phone#</td>
                                <td>#dob#</td>
                                <td>#rolenames#</td> 
                            </tr>                           
                        </cfloop>                                      
                    </table>
                </div>   
            </form>  
            <cfset local.resultExcelUpload="">
            <cfif  StructKeyExists(form,"uploadBtn") && NOT IsNull(form.fileUpload)> 
                <cfset local.resultExcelUpload=local.userObject.uploadExcel(#form.fileUpload#)>
                <!---<a href="index.cfm" target="_blank">Download Excel</a>--->
                <div class="result" id="result1">#local.resultExcelUpload#</div>
                <cflocation url="index.cfm">
            </cfif>                   
            <div class="result" id="result"></div>         
        </cfoutput>
    </body>     
</html>