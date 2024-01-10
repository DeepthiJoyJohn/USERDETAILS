$(document).ready(function(){
 
  
});

function generateExcelTemplate(){  
  window.location.href="downloadExcelTemplate.cfm"
  
}
function generateDataExcelTemplate(){
  window.location.href="downloadDataExcelTemplate.cfm"
}
function cancelSelection(){
  window.location.reload();
  
}
function upload111() { 
  
  var fileInput = document.getElementById('fileUpload');
  var result = document.getElementById('result');  
  if (fileInput.files.length === 0) {
      result.innerHTML = "Please choose a file to upload.";
      return false;
  }else{
    var formData = new FormData();
    var fileInput = document.getElementById('fileUpload').files[0]; // Get the selected file
    var fileInput = $('#fileUpload')[0].files[0];
    formData.append('file', fileInput);   
    alert(formData);
    $.ajax({
      url: 'Components/userdetails.cfc?method=uploadExcel111',
      type: 'POST',
      data: formData,
      processData: false,
      contentType: false,
      success: function(response) {
        alert(response);
          console.log(response); // Handle the response from the server
      },
      error: function(xhr, status, error) {
          console.error('Error:', error);
      }
  });
    // $.ajax({
    //   type: "POST",      
    //   url: 'Components/userdetails.cfc?method=uploadExcel111',
    //   data: formData,
    //   processData: false, 
    //   contentType: false,        
    //   cache: false,
    //   success: function(data){ 
        // alert(data);
        // var retval = $(data).find("string").text();
        // alert(retval);
        // window.location.href="downloadDataExcelTemplate1.cfm?filename="+retval;
        //   setTimeout(() => {
        //     window.location.href = 'index.cfm'; // Redirect after download completes
        //   }, 3000);
      //  }          
      //});
    }    
}

function upload1() {
  
  var fileInput = document.getElementById('fileUpload');
  var result = document.getElementById('result');
  
  if (fileInput.files.length === 0) {
      result.innerHTML = "Please choose a file to upload.";
      return false;
  }else{   
    setTimeout(() => {
          window.location.href = 'index.cfm'; // Redirect after download completes
        }, 3000);

  } 
}

function browse() {  
 
  var result = document.getElementById('result'); 
  result.innerHTML = "";
  document.getElementById('fileUpload').click();
  
  document.getElementById('fileUpload').addEventListener('change', function() {
    const selectedFile = this.files[0]; 
    var fileExtension = selectedFile.name.slice((selectedFile.name.lastIndexOf(".") - 1 >>> 0) + 2);
    var allowedExtensions = [".xls", ".xlsx"];     
    if (!allowedExtensions.includes("." + fileExtension.toLowerCase())) {
      result.innerHTML = "Please select a file with .xls or .xlsx extension.";
      document.getElementById('fileUpload').value = '';      
      return false;
    }else if (selectedFile) {            
      const fileInfo = document.getElementById('selectedFileInfo');
      fileInfo.innerHTML = selectedFile.name + ' ' + `<i title="Cancel Selection" onclick="cancelSelection()" class="fa fa-close fa-lg cancel"></i>`;
    }else {      
      document.getElementById('selectedFileInfo').innerHTML = 'No file selected';
    }
  });
}
function download(){
  window.location.href="downloadResult.cfm"
}
