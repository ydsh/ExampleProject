<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>studentDetail</title>
</head>
<body>
<center>

<h3>studentDetail</h3><br>

  <label style="display: none;">标识：</label><input style="display: none;" type="text" name="num" value="${requestScope.student.num}" readonly="readonly"/><br><br>
  <label>姓名：</label><input type="text" name="name" value="${requestScope.student.name}" readonly="readonly"/><br><br>
  <label>年龄：</label><input type="text" name="age" value="${requestScope.student.age}" readonly="readonly"/><br><br>
  <label>性别：</label><input type="text" name="gender" value="${requestScope.student.gender}" readonly="readonly"/><br><br>
  <label>邮箱：</label><input type="text" name="email" value="${requestScope.student.email}" readonly="readonly"/><br><br>
  
</center>
</body>
</html>