function generateExcelTemplate(){  
  window.location.href="downloadExcelTemplate.cfm"
  
}
function generateDataExcelTemplate(){
  window.location.href="downloadDataExcelTemplate.cfm"
}
function cancelSelection(){
  window.location.reload();
  
}
function upload1() {  
  var fileInput = document.getElementById('fileUpload');
  var result = document.getElementById('result');
  
  if (fileInput.files.length === 0) {
      result.innerHTML = "Please choose a file to upload.";
      return false;
  }else{
    var formData = new FormData();
    var fileInput = document.getElementById('fileUpload').files[0]; // Get the selected file
    formData.append('fileUpload', fileInput);
    alert(formData);
    
    $.ajax({
      type: "POST",      
      url: 'Components/userdetails.cfc?method=uploadExcel',
      data: formData,
      processData: false, 
      contentType: false, 
      enctype: 'multipart/form-data', 
      cache: false,
      success: function(data){ 
        var retval = $(data).find("string").text();
        alert(retval);
        window.location.href="downloadDataExcelTemplate1.cfm?filename="+retval;
          setTimeout(() => {
            window.location.href = 'index.cfm'; // Redirect after download completes
          }, 3000);
        }
          
      });
    }    
}

function upload11() {
  alert("df");
  var fileInput = document.getElementById('fileUpload');
  var result = document.getElementById('result');
  
  if (fileInput.files.length === 0) {
      result.innerHTML = "Please choose a file to upload.";
      return false;
  }else{   
    $.ajax({
      type: "POST",
      url: 'Components/userdetails.cfc?method=uploadExcel',
      cache: false,
      success: function(data){ 
      }
        
    });

  } 
}
function upload() {
  $.ajax({
    type: "POST",
    url: 'Components/userdetails.cfc?method=uploadExcel2',
    cache: false,
    success: function(data){ 
      fetch('downloadResult.cfm') // Replace 'downloadFile.cfm' with your file download URL
      .then(response => {
      // Open the file download in a new tab or window
      window.open('downloadResult.cfm', '_blank');
      
      // Redirect to another page after a delay (adjust delay as needed)
      setTimeout(() => {
        window.location.href = 'index.cfm'; // Replace with your redirect URL
      }, 3000); // Redirect after 3 seconds (adjust delay as needed)
    })
    .catch(error => {
      console.error('Error:', error);
    });
    }
      
  });
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
      document.getElementById('selectedFileInfo').textContent = 'No file selected';
    }
  });
}
function download(){
  window.location.href="downloadResult.cfm"
}
