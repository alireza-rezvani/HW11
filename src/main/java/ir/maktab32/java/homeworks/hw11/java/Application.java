package ir.maktab32.java.homeworks.hw11.java;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class Application {
    public static void main(String[] args) {

        List<Person> personList = new ArrayList<>();
        personList.add(new Person(1L, "Ali", "Alavi", "09121111111"));
        personList.add(new Person(2L, "Reza", "Razavi", "09101111111"));
        personList.add(new Person(3L, "Ahmad", "Ahmadi", "09351111111"));

        writePersonToExcel(personList);
        readPersonFromExcel().forEach(System.out::println);
    }

    public static void writePersonToExcel(List<Person> personList){
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("PersonSheet");

        Row headRow = sheet.createRow(0);
        headRow.createCell(0).setCellValue("ID");
        headRow.createCell(1).setCellValue("First Name");
        headRow.createCell(2).setCellValue("Last Name");
        headRow.createCell(3).setCellValue("Phone");

        for (int i = 1; i <= personList.size(); i++){
            Row row = sheet.createRow(i);
            row.createCell(0).setCellValue(personList.get(i-1).getId());
            row.createCell(1).setCellValue(personList.get(i-1).getFirstName());
            row.createCell(2).setCellValue(personList.get(i-1).getLastName());
            row.createCell(3).setCellValue(personList.get(i-1).getPhoneNumber());
        }

        try {
            FileOutputStream output = new FileOutputStream(new File("Person.xlsx"));
            workbook.write(output);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public static List<Person> readPersonFromExcel(){
        List<Person> result = new ArrayList<>();
        try {
            FileInputStream input = new FileInputStream("Person.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(input);
            XSSFSheet sheet = workbook.getSheetAt(0);

            Iterator<Row> rowIterator = sheet.rowIterator();
            rowIterator.next();

            while (rowIterator.hasNext()){
                Person person = new Person();
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();

                person.setId((long) cellIterator.next().getNumericCellValue());
                person.setFirstName(cellIterator.next().getStringCellValue());
                person.setLastName(cellIterator.next().getStringCellValue());
                person.setPhoneNumber(cellIterator.next().getStringCellValue());

                result.add(person);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return result;
    }
}
