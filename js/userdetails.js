function generateExcelTemplate(data){ 
  window.location.href="downloadExcelTemplate.cfm?data="+data;  
}

function cancelSelection(){
  window.location.reload();  
}
function onSubmitFunction(){  
 // window.location.href="downloadExcelTemplate.cfm?data="+data;
 
}
function checkFileExists(){
  var fileInput = document.getElementById('fileUpload');
  var result = document.getElementById('result'); 
  if (fileInput.files.length === 0) {
      result.innerHTML = "Please choose a file to upload.";
      return false;
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
