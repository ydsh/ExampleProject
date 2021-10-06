package com.example.comm;

import java.util.Date;

public class User {
	/**
	 * ID
	 */
	@Excel(order = 0, name = "ID")
	private String id;
	/**
	 * 姓名
	 */
	@Excel(order = 1, name = { "基本信息", "姓名" })
	private String name;

	/**
	 * 年龄
	 */
	@Excel(order = 2, name = { "基本信息", "年龄" }, fmt = "#0.000")
	private int age;

	/**
	 * 居住地
	 */
	@Excel(order = 3, name = { "基本信息", "地址" })
	private String location;
	/**
	 * 邮箱
	 */
	@Excel(order = 4, name = { "基本信息", "邮箱" })
	private String email;

	/**
	 * 职业
	 */
	@Excel(order = 5, name = { "基本信息", "职业" })
	private String job;
	/**
	 * 入职时间
	 */
	@Excel(order = 6, name = "入行时间", fmt = "yyyy-MM-dd")
	private Date time;
	/**
	 * 简介
	 */
	@Excel(order = 7, name = "简介")
	private String intro;
	/**
	 * 推荐人
	 */
	@Excel(order = 8, name = "推荐人")
	private String reference;
	/**
	 * 毕业院校
	 */
	@Excel(order = 9, name = { "教育信息", "毕业院校" })
	private String graduateSchool;
	/**
	 * 专业
	 */
	@Excel(order = 10, name = { "教育信息", "专业" })
	private String professional;
	/**
	 * 学历
	 */
	@Excel(order = 11, name = { "教育信息", "学历" })
	private String degree;
	/**
	 * 毕业时间
	 */
	@Excel(order = 12, name = { "教育信息", "毕业时间" }, fmt = "yyyy-MM-dd")
	private Date graduateTime;

	public String getId() {
		return id;
	}

	public void setId(String id) {
		this.id = id;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public int getAge() {
		return age;
	}

	public void setAge(int age) {
		this.age = age;
	}

	public String getLocation() {
		return location;
	}

	public void setLocation(String location) {
		this.location = location;
	}

	public String getEmail() {
		return email;
	}

	public void setEmail(String email) {
		this.email = email;
	}

	public String getJob() {
		return job;
	}

	public void setJob(String job) {
		this.job = job;
	}

	public Date getTime() {
		return time;
	}

	public void setTime(Date time) {
		this.time = time;
	}

	public String getIntro() {
		return intro;
	}

	public void setIntro(String intro) {
		this.intro = intro;
	}

	public String getReference() {
		return reference;
	}

	public void setReference(String reference) {
		this.reference = reference;
	}

	public String getGraduateSchool() {
		return graduateSchool;
	}

	public void setGraduateSchool(String graduateSchool) {
		this.graduateSchool = graduateSchool;
	}

	public String getProfessional() {
		return professional;
	}

	public void setProfessional(String professional) {
		this.professional = professional;
	}

	public String getDegree() {
		return degree;
	}

	public void setDegree(String degree) {
		this.degree = degree;
	}

	public Date getGraduateTime() {
		return graduateTime;
	}

	public void setGraduateTime(Date graduateTime) {
		this.graduateTime = graduateTime;
	}

	@Override
	public String toString() {
		return "{\"id\":" + id + ", \"name\":" + name + ", \"age\":" + age + ", \"location\":" + location
				+ ", \"email\":" + email + ", \"job\":" + job + ", \"time\":" + time + ", \"intro\":" + intro
				+ ", \"reference\":" + reference + ", \"graduateSchool\":" + graduateSchool + ", \"professional\":"
				+ professional + ", \"degree\":" + degree + ", \"graduateTime\":" + graduateTime + "}";
	}
    
}
