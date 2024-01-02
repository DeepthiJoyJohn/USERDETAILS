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
            <form action="" method="post" enctype="multipart/form-data" >
                <cfset local.userObject=createObject("component", "Components.userdetails")>
                <cfset local.resultUserDetails=local.userObject.getUserDetails()>
                <div class="heading">USER INFORMATION</div>
                <div class="btnDiv">
                    <div class="btnLeft">
                        <button class="plainTemplate" onclick="generateExcelTemplate()">Plain Template</button>
                        <button class="templateWithData">Template With Data</button>
                    </div>
                    <div class="btnRight">
                        <button class="browse" onclick="browse()">Browse<input type="file" name="fileUpload" id="fileUpload" class="fileUpload" required="yes" accept=".xlsx, .xls" /></button>                        
                        <div class="selectedFileInfo" id="selectedFileInfo"><span id="closeIcon"><i class="fa fa-close"></i></span></div>  
                        <button class="upload">Upload</button>
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
                        <tr>
                            <cfloop query="local.resultUserDetails">
                                <td>#firstname#</td>
                                <td>#lastname#</td>
                                <td>#address#</td>
                                <td>#email#</td>
                                <td>#phone#</td>
                                <td>#dob#</td>
                                <td>#rolenames#</td>                            
                            </cfloop>
                        </tr>               
                    </table>
                </div>
            </form>
        </cfoutput>
    </body>     
</html>