package ru.academits.java.kononov.excel.person;

public class Person {
    private final String name;
    private final String surname;
    private final int age;
    private final long phone;

    public Person(String name, String surname, int age, long phone) {
        this.name = name;
        this.surname = surname;
        this.age = age;
        this.phone = phone;
    }

    public String getName() {
        return name;
    }

    public String getSurname() {
        return surname;
    }

    public int getAge() {
        return age;
    }

    public long getPhone() {
        return phone;
    }
}
