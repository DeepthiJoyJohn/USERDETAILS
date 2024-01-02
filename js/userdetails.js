$(document).ready(function () { 
  var closeIcon=document.getElementById("closeIcon").style.display;
 document.getElementById("closeIcon").style.display="none";
});
function generateExcelTemplate(){
  window.location.href="downloadExcelTemplate.cfm"
}
function browse(){
  document.getElementById('fileUpload').click();
  document.getElementById('fileUpload').addEventListener('change', function() {
    const selectedFile = this.files[0]; 
    if (selectedFile) {
        const fileInfo = document.getElementById('selectedFileInfo');
        fileInfo.textContent = selectedFile.name;
       
    } else {
        document.getElementById('selectedFileInfo').textContent = 'No file selected';
    }
  });
}
