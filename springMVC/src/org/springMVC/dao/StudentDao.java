package org.springMVC.dao;

import java.util.List;

import org.springMVC.model.Student;

public interface StudentDao {

	public boolean addStudent(Student stu);
	
	public Student queryStudent(Student stu);
	
	public boolean deleteStudent(Student stu);
	
	public boolean updateStudent(Student stu);
	
	public List<Student> queryAllStudent();
	
}
