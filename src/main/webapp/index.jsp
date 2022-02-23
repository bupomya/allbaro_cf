<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Insert title here</title>
</head>
<body>
	<form action="AllbaroExcelServlet" method  ="post" enctype = "multipart/form-data" class = "allbaroExcel">
		<input type = "file" id = "excel"  accept = ".xls,.xlsx" name = "excel" multiple>
		<input type = "submit" value = "보내기">
	</form>
</body>
</html>