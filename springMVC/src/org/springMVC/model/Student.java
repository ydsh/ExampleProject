package org.springMVC.model;

public class Student {

	private int num;
	private String name;
	private String gender;
	private int age;
	private String email;
	
	public Student(){
		
	}

	public Student(int num,String name, String gender, int age, String email) {
		this.num = num;
		this.name = name;
		this.gender = gender;
		this.age = age;
		this.email = email;
	}
	

	public int getNum() {
		return num;
	}

	public void setNum(int num) {
		this.num = num;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getGender() {
		return gender;
	}

	public void setGender(String gender) {
		this.gender = gender;
	}

	public int getAge() {
		return age;
	}

	public void setAge(int age) {
		this.age = age;
	}

	public String getEmail() {
		return email;
	}

	public void setEmail(String email) {
		this.email = email;
	}
	
	
}
