package org.springMVC.impl;

import java.util.LinkedList;
import java.util.List;

import org.hibernate.Session;
import org.hibernate.SessionFactory;
import org.hibernate.query.Query;
import org.springMVC.dao.StudentDao;
import org.springMVC.model.Student;
import org.springframework.beans.factory.annotation.Autowired;

public class StudentImpl implements StudentDao {
	
	@Autowired
	private SessionFactory sessionFactory;
	
	public SessionFactory getSessionFactory() {
		return sessionFactory;
	}

	public void setSessionFactory(SessionFactory sessionFactory) {
		this.sessionFactory = sessionFactory;
	}
	

	@Override
	public boolean addStudent(Student stu) {
		System.out.println("这里是否执行到？");
		boolean b = false;
		Session session  = sessionFactory.getCurrentSession();
		session.getTransaction().begin();
		int i = Integer.parseInt(session.save(stu).toString());
		
		if(i > 0){
			
			b = true;
		}
		
		session.getTransaction().commit();
		return b;
	}

	@Override
	public Student queryStudent(Student stu) {
		
		int num = stu.getNum();
		String name = stu.getName();
		Student student = null;
		
		Session session = sessionFactory.getCurrentSession();
		session.getTransaction().begin();
		if(num!=0&&name.trim().length()==0){
			Query<Student> query = session.createQuery("from Student where num=?").setParameter(0, num);
			if(query!=null){
				
				student = query.getResultList().get(0);
			}
		}
		if(num==0&&name.trim().length()!=0){
			Query<Student> query = session.createQuery("from Student where name=?").setParameter(0, name);
			
           if(query!=null){
				
				student = query.getResultList().get(0);
			}
		}
		if(num!=0&&name.trim().length()!=0){
			Query<Student> query = session.createQuery("from Student where num=? and name=?").setParameter(0, num).setParameter(1, name);
            
			student = query.getResultList().get(0);
		}
		
		
		return student;
	}

	@Override
	public boolean deleteStudent(Student stu) {
		boolean b = false;
		
		int num = stu.getNum();
		String name = stu.getName();
		Session session = sessionFactory.getCurrentSession();
		session.getTransaction().begin();
		
		if(num!=0&&name.trim().length()==0){
			int i = session.createQuery("delete from Student where num=?").setParameter(0, num).executeUpdate();
			if(i > 0){
				b=true;
			}
		}
		if(num==0&&name.trim().length()>0){
			
			int i = session.createQuery("delete from Student where name=?").setParameter(0, name).executeUpdate();
			if(i > 0){
				b = true;
			}
		}
		
		if(num!=0&&name.trim().length()>0){
			int i =session.createQuery("delete from Student where num=? and name=?").setParameter(0, num).setParameter(1, name).executeUpdate();
			if(i > 0){
				b = true;
			}
		}
		
		if(num==0&&name.trim().length()==0){
			int i = session.createQuery("delete from Student").executeUpdate();
			
			if(i > 0){
				b =true;
			}
		}
		
		session.getTransaction().commit();
		return b;
	}

	@Override
	public boolean updateStudent(Student stu) {
		boolean b= false;
		
		Session session = sessionFactory.getCurrentSession();
		session.getTransaction().begin();
		
		if(stu!=null){
			int i = session.createQuery("update Student set age=?,gender=?,email=? where num=? and name=?")
			.setParameter(0, stu.getAge())
			.setParameter(1, stu.getGender())
			.setParameter(2, stu.getEmail())
			.setParameter(3, stu.getNum())
			.setParameter(4, stu.getName())
			.executeUpdate();
			
			if(i > 0){
				b = true;
			}
		}
		session.getTransaction().commit();
		
		return b;
	}

	@Override
	public List<Student> queryAllStudent() {
		List<Student> studentList = new LinkedList<Student>();
		Session session = sessionFactory.getCurrentSession();
		session.getTransaction().begin();
		studentList = session.createQuery("from Student").getResultList();
		
		return studentList;
	}

}
