<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>deleteModel</title>
</head>
<body>
<center>
   <h3>deletModel</h3><br>
<form action="/springMVC/mvc/deleteModel" method="post">

  <label>标识：</label><input type="text" name="num"/><br><br>
  <label>姓名：</label><input type="text" name="name"/><br><br>
  <input type="submit" value="put"/>
  
  </form>
</center>
</body>
</html>