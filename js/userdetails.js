function generateExcelTemplate(){  
  window.location.href="downloadExcelTemplate.cfm";
  
}
function generateDataExcelTemplate(){
  window.location.href="downloadDataExcelTemplate.cfm"
}
function cancelSelection(){
  window.location.reload();  
}
function onSubmitFunction(){
  setTimeout(function() {
    window.location.href="index.cfm" // Refresh parent page
  }, 3000);
  
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
