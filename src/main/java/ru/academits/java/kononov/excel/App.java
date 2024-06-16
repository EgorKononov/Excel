package ru.academits.java.kononov.excel;

import ru.academits.java.kononov.excel.person.Person;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class App {
    private static final String PERSONS_EXCEL_FILE_PATH = "persons.xlsx";

    public static void main(String[] args) {
        try {
            PersonsExcelFileCreator.create(createPersonsList(), PERSONS_EXCEL_FILE_PATH);
        } catch (IOException e) {
            System.out.println("Exception: " + e.getMessage() + e);
        }
    }

    private static List<Person> createPersonsList() {
        List<Person> personList = new ArrayList<>();
        personList.add(new Person("Ivan", "Ivanov", 26, 71234567890L));
        personList.add(new Person("Peter", "Petrov", 27, 71234567891L));
        personList.add(new Person("Egor", "Egorov", 28, 71234567892L));
        personList.add(new Person("Anton", "Antonov", 29, 71234567893L));

        return personList;
    }
}
