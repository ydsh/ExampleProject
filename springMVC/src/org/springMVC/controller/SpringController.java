package org.springMVC.controller;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springMVC.dao.StudentDao;
import org.springMVC.model.Student;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;

@Controller
@RequestMapping("/mvc")
public class SpringController {

	@Autowired
	private StudentDao studentDao;

	public StudentDao getStudentDao() {
		return studentDao;
	}

	public void setStudentDao(StudentDao studentDao) {
		this.studentDao = studentDao;
	}
	
	@RequestMapping(value="/err")
	public void toError(HttpServletRequest req,HttpServletResponse res) throws Exception{
		
		req.setAttribute("msg", "您好，您已经成功转发！");
		req.getRequestDispatcher("user/error.jsp").forward(req, res);
		//return "redirect:/user/error.jsp";
	}

	
	
	@RequestMapping(value="/add")
	public String toAddModel(){
		
		
		return "addModel";
	}
	
	
	@RequestMapping(value="/addModel")
	public String addModel(Student stu,String msg,Model model){
		
		if(stu!=null){
			System.out.println(stu.getName()+" , "+stu.getAge()+" , "+stu.getGender()+" , "+stu.getEmail());
			if(studentDao == null){
				System.out.println("studentDao 为 null~~~~~~~~~~~~~~~~~~~~~~~~~~~");
				return "addModel";
			}else{
				boolean b = studentDao.addStudent(stu);
				
				if(b){
					msg = "addModel SUCCESS!";
					model.addAttribute("msg", msg);
					
					return "success";
				}
			}
			
		}
		return "addModel";
	}	
	
	@RequestMapping(value="/delete")
	public String toDeleteModel(){
		return "deleteModel";
	}
	
	@RequestMapping(value="deleteModel")
	public String deleteModel(Student stu,String msg,Model model){
		if(stu!=null){
			System.out.println(stu.getNum()+" "+stu.getName());
			if(studentDao == null){
				System.out.println("studentDao 为 null~~~~~~~~~~~~~~~~~~~~~~~~~~~");
				return "addModel";
			}else{
				boolean b = studentDao.deleteStudent(stu);
				
				if(b){
					msg = "deleteModel SUCCESS!";
					model.addAttribute("msg", msg);
					
					return "success";
				}
			}
			
		}
		return "deleteModel";
	}
	
	@RequestMapping(value="/query")
	public String toQueryModel(){
		return "queryModel";
	}
	
	@RequestMapping(value="/queryModel")
	public String queryModel(Student stu,Model model){
		Student student = null;
		
		if(stu != null){
			student = studentDao.queryStudent(stu);
			if(student!=null){
				
				model.addAttribute("student",student);
				
				return "studentDetail";
			}
		}
		return "queryModel";
	}
}
