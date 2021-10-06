<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>addModel</title>
</head>
<body>
<center>

<h3>addModel</h3><br>
<form action="/springMVC/mvc/addModel" method="post">


  <label>姓名：</label><input type="text" name="name"/><br><br>
  <label>年龄：</label><input type="text" name="age"/><br><br>
  <label>性别：</label><input type="text" name="gender"/><br><br>
  <label>邮箱：</label><input type="text" name="email"/><br><br>
  <input type="submit" value="put"/>
  
  </form>
</center>
</body>
</html>