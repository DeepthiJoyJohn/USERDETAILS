<cfheader name="Content-Disposition" value="inline;filename=Result.xlsx">
<cfcontent  file="#expandPath('ExcelUploads/Result/')#UploadResult.xlsx" type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"> 



