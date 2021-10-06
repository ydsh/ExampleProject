<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">

<title>Insert title here</title>
<!-- 严格按顺序，先加在dwr引擎 -->
<script type="text/javascript" src="/DwrDemo/dwr/engine.js"></script>
<script type="text/javascript"  src="/DwrDemo/dwr/util.js"></script>
<!-- 引入js,这里是在dwr.xml中自定义JavaScript -->
<script type="text/javascript" src="/DwrDemo/dwr/interface/FirstDemo.js"></script>
<script type="text/javascript">
   
   
   function showData1(){
	   
   var name = document.getElementById('dwr').value;
      
   //FirstDemo是在dwr.xml中自定义JavaScript，fun1是java类中的方法，
	   FirstDemo.fun1(name,function(data){//这里的function是回调函数
		   
		   document.getElementById("show1").innerHTML=data;
	   });
   
   }
   
   function showData2(){
	   var result = document.getElementById("test").value;
	   
	   FirstDemo.fun2(result,function(data){
		   document.getElementById("show2").innerHTML=data;
	   });
   }
   
</script>
</head>
<body>
 <div align="center">
     <input id="dwr" type="text" name="name1">
     <input type="button" value="提交" onclick="showData1()"><br><br>
     <h3 id="show1" style="color: red;"></h3><br><br>
      <input id="test" type="text" name="name2">
     <input type="button" value="test" onclick="showData2()"><br><br>
     <h3 id="show2" style="color: blue;"></h3>
 </div>
</body>
</html>