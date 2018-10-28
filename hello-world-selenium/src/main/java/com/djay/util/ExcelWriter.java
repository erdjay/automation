package com.djay.util;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;


/**
 * The Class ExcelWriter.
 */
public class ExcelWriter {

    /** The columns. */
    private static String[] columns = {"Name", "Email", "Date Of Birth", "Salary", "Formula"};

    /** The employees. */
    private static List<Employee> employees =  new ArrayList<>();

    static {
        Calendar dateOfBirth = Calendar.getInstance();
        dateOfBirth.set(1992, 7, 21);
        employees.add(new Employee("Rajeev Singh", "rajeev@example.com",
                dateOfBirth.getTime(), 1200000.0));

        dateOfBirth.set(1965, 10, 15);
        employees.add(new Employee("Thomas cook", "thomas@example.com",
                dateOfBirth.getTime(), 1500000.0));

        dateOfBirth.set(1987, 4, 18);
        employees.add(new Employee("Steve Maiden", "steve@example.com",
                dateOfBirth.getTime(), 1800000.0));
    }

    /**
     * The main method.
     *
     * @param args the arguments
     * @throws IOException Signals that an I/O exception has occurred.
     * @throws InvalidFormatException the invalid format exception
     */
    public static void main(String[] args) throws IOException, InvalidFormatException {

        // Create a Workbook
        //Workbook workbook = new XSSFWorkbook();     
        Workbook workbook =  new HSSFWorkbook(); //for generating `.xls` file

        /* CreationHelper helps us create instances for various things like DataFormat,
           Hyperlink, RichTextString etc in a format (HSSF, XSSF) independent way */
        CreationHelper createHelper = workbook.getCreationHelper();

        // Create a Sheet
        Sheet sheet = workbook.createSheet("Employee");

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        // Create a Row
        Row headerRow = sheet.createRow(0);

        // Creating cells
        for(int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        // Cell Style for formatting Date
        CellStyle dateCellStyle = workbook.createCellStyle();
        dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));

        // Create Other rows and cells with employees data
        int rowNum = 1;
        for(Employee employee: employees) {
            Row row = sheet.createRow(rowNum++);

            row.createCell(0)
                    .setCellValue(employee.getName());

            row.createCell(1)
                    .setCellValue(employee.getEmail());

            Cell dateOfBirthCell = row.createCell(2);
            dateOfBirthCell.setCellValue(employee.getDateOfBirth());
            dateOfBirthCell.setCellStyle(dateCellStyle);

            row.createCell(3)
                    .setCellValue(employee.getSalary());
            
            row.createCell(4).setCellType(CellType.FORMULA);
            row.createCell(4).setCellFormula("CONCATENATE(B2,A2)");
        }

        // Resize all columns to fit the content size
        for(int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("poi-generated-file.xls");
        workbook.write(fileOut);
        fileOut.close();

        workbook.close();
    }


    /**
     * Modify existing workbook.
     *
     * @throws InvalidFormatException the invalid format exception
     * @throws IOException Signals that an I/O exception has occurred.
     */
    // Example to modify an existing excel file
    private static void modifyExistingWorkbook() throws InvalidFormatException, IOException {
        // Obtain a workbook from the excel file
        Workbook workbook = WorkbookFactory.create(new File("existing-spreadsheet.xlsx"));

        // Get Sheet at index 0
        Sheet sheet = workbook.getSheetAt(0);

        // Get Row at index 1
        Row row = sheet.getRow(1);

        // Get the Cell at index 2 from the above row
        Cell cell = row.getCell(2);

        // Create the cell if it doesn't exist
        if (cell == null)
            cell = row.createCell(2);

        // Update the cell's value
        cell.setCellType(CellType.STRING);
        cell.setCellValue("Updated Value");

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("existing-spreadsheet.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
        
    }
    
   }


class Employee {
    private String name;

    private String email;

    private Date dateOfBirth;

    private double salary;

    public Employee(String name, String email, Date dateOfBirth, double salary) {
        this.name = name;
        this.email = email;
        this.dateOfBirth = dateOfBirth;
        this.salary = salary;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public Date getDateOfBirth() {
        return dateOfBirth;
    }

    public void setDateOfBirth(Date dateOfBirth) {
        this.dateOfBirth = dateOfBirth;
    }

    public double getSalary() {
        return salary;
    }

    public void setSalary(double salary) {
        this.salary = salary;
    }
}
