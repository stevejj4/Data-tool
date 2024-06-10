import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

public class Name { 
    private String firstName;
    private String lastName;
    private char middleInitial;

    // Default constructor
    public Name() {
    }

    // Parameterized constructor to access the instance variables
    // use this keyword 
    public Name(String firstName, String lastName, char middleInitial) {
        this.firstName = firstName;
        this.lastName = lastName;
        this.middleInitial = middleInitial;
    }

    // Getter for firstName
    public String getFirstName() {
        return firstName;
    }

    // Setter for firstName
    public void setFirstName(String firstName) {
        this.firstName = firstName;
    }

    // Getter for lastName
    public String getLastName() {
        return lastName;
    }

    // Setter for lastName
    public void setLastName(String lastName) {
        this.lastName = lastName;
    }

    // Getter for middleInitial
    public char getMiddleInitial() {
        return middleInitial;
    }

    // Setter for middleInitial
    public void setMiddleInitial(char middleInitial) {
        this.middleInitial = middleInitial;
    }

    // Method to display the full name
    public String getFullName() {
        if (middleInitial == '\0') {
            return firstName + " " + lastName;
        } else {
            return firstName + " " + middleInitial + ". " + lastName;
        }
    }

    // Method to collect name details from the user
    public void collectNameDetails() {
        Scanner scanner = new Scanner(System.in);

        System.out.print("Enter first name: ");
        this.firstName = scanner.nextLine();

        System.out.print("Enter middle initial (press Enter if none): ");
        String middleInitialInput = scanner.nextLine();
        if (!middleInitialInput.isEmpty()) {
            this.middleInitial = middleInitialInput.charAt(0);
        }

        System.out.print("Enter last name: ");
        this.lastName = scanner.nextLine();

        System.out.println("Name details collected successfully.");
    }

    // Method to save the name details into an Excel file
    public void saveToExcel(String fileName) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Name Details");

        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("First Name");
        headerRow.createCell(1).setCellValue("Middle Initial");
        headerRow.createCell(2).setCellValue("Last Name");

        Row dataRow = sheet.createRow(1);
        dataRow.createCell(0).setCellValue(firstName);
        dataRow.createCell(1).setCellValue(middleInitial == '\0' ? "" : String.valueOf(middleInitial));
        dataRow.createCell(2).setCellValue(lastName);

        try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
            workbook.write(fileOut);
            System.out.println("Name details saved to " + fileName);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public static void main(String[] args) {
        // Creating an instance of Name
        Name personName = new Name();

        // Collecting name details from the user
        personName.collectNameDetails();

        // Displaying the collected name details
        System.out.println("Collected Name Details:");
        System.out.println("First Name: " + personName.getFirstName());
        System.out.println("Middle Initial: " + personName.getMiddleInitial());
        System.out.println("Last Name: " + personName.getLastName());
        System.out.println("Full Name: " + personName.getFullName());

        // Saving the name details to an Excel file
        personName.saveToExcel("NameDetails.xlsx");
    }
}
