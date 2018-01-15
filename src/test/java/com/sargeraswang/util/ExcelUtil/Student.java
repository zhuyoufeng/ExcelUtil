package com.sargeraswang.util.ExcelUtil;

import java.util.Arrays;
import java.util.Date;

/**
 * Student
 *
 * @author zhuyoufeng
 */
public class Student {
    @ExcelCell(index = 0)
    private String name;
    @ExcelCell(index = 1)
    private Integer age;
    @ExcelCell(index = 2)
    private Date birthday;
    @ExcelCell(index = 3)
    private String gender;
    @ExcelCell(index = 4)
    private String[] habbit;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public Date getBirthday() {
        return birthday;
    }

    public void setBirthday(Date birthday) {
        this.birthday = birthday;
    }

    public String getGender() {
        return gender;
    }

    public void setGender(String gender) {
        this.gender = gender;
    }

    public String[] getHabbit() {
        return habbit;
    }

    public void setHabbit(String[] habbit) {
        this.habbit = habbit;
    }

    @Override
    public String toString() {
        return "Student{" +
                "name='" + name + '\'' +
                ", age=" + age +
                ", birthday=" + birthday +
                ", gender='" + gender + '\'' +
                ", habbit=" + Arrays.toString(habbit) +
                '}';
    }
}