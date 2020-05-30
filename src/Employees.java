import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;
 
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


class Employee{
    int age;
    String name, address, gender;
    Scanner get = new Scanner(System.in);
    Employee()
    {
    	  System.out.println("=============================="+"\n"+"Employee Details"+"\n"+"=============================="+"\n");
        System.out.println("Enter Name of the Employee:");
        name = get.nextLine();
        System.out.println("Enter Gender of the Employee:");
        gender = get.nextLine();
        System.out.println("Enter Address of the Employee:");
        address = get.nextLine();
        System.out.println("Enter Age:");
        age = get.nextInt();
    }
 
    void display()
    {
        System.out.println("Employee Name: "+name);
        System.out.println("Age: "+age);
        System.out.println("Gender: "+gender);
        System.out.println("Address: "+address);
    }
}
 
class fullTimeEmployees extends Employee{
    float salary;
    String des;
    int workinghrs;
    fullTimeEmployees()
    {
    
        System.out.println("Enter Designation:");
        des = get.next();
        System.out.println("Enter Salary:");
        salary = get.nextFloat();
        System.out.println("Enter Number of Working Hours:");
        workinghrs = get.nextInt();
    }

    void display()
    {
    	super.display();
        System.out.println("Salary: "+salary);
        System.out.println("Designation: "+des);
        System.out.println("Number of hours: "+workinghrs);
        
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("EmployeeDetails");
        
        header(sheet);
        
        try (FileOutputStream outputStream = new FileOutputStream("EmployeeDetails.xlsx")) {
            workbook.write(outputStream);
        } catch (IOException e) {
			e.printStackTrace();
		}
        
        
       writer(sheet,name,age, address, gender,salary, des, workinghrs );
       
       try (FileOutputStream outputStream = new FileOutputStream("EmployeeDetails.xlsx")) {
           workbook.write(outputStream);
       } catch (IOException e) {
			e.printStackTrace();
		}
    }
    
    
    static void header(XSSFSheet sheet) {
  	  CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        cellStyle.setFont(font);
        
  	
  	Row row = sheet.createRow(0);
      Cell cellName = row.createCell(1);   
      cellName.setCellStyle(cellStyle);
      cellName.setCellValue("Name");
      
      Cell cellAge = row.createCell(2);
      cellAge.setCellStyle(cellStyle);
      cellAge.setCellValue("Age");
   
      Cell cellGender = row.createCell(3);
      cellGender.setCellStyle(cellStyle);
      cellGender.setCellValue("Gender");
      
      Cell cellAddress = row.createCell(4);
      cellAddress.setCellStyle(cellStyle);
      cellAddress.setCellValue("Address");
   
      Cell cellSalary = row.createCell(5);
      cellSalary.setCellStyle(cellStyle);
      cellSalary.setCellValue("Salary");
      
      Cell cellDesignation = row.createCell(6);
      cellDesignation.setCellStyle(cellStyle);
      cellDesignation.setCellValue("Designation");
   
      Cell cellHours = row.createCell(7);
      cellHours.setCellStyle(cellStyle);
      cellHours.setCellValue("Hours Worked");
      
      
  }
  static void writer(XSSFSheet sheet,String name, int age,String address, String gender, float salary, String desc, int workinghrs ) {
  	
  	CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
  	
  	Row row = sheet.createRow(1);
      Cell cellName = row.createCell(1);   
      cellName.setCellStyle(cellStyle);
      cellName.setCellValue(name);
      
      Cell cellAge = row.createCell(2);
      cellAge.setCellStyle(cellStyle);
      cellAge.setCellValue(age);
   
      Cell cellGender = row.createCell(3);
      cellGender.setCellStyle(cellStyle);
      cellGender.setCellValue(gender);
      
      Cell cellAddress = row.createCell(4);
      cellAddress.setCellStyle(cellStyle);
      cellAddress.setCellValue(address);
   
      Cell cellSalary = row.createCell(5);
      cellSalary.setCellStyle(cellStyle);
      cellSalary.setCellValue(salary);
      
      Cell cellDesignation = row.createCell(6);
      cellDesignation.setCellStyle(cellStyle);
      cellDesignation.setCellValue(address);
   
      Cell cellHours = row.createCell(7);
      cellHours.setCellStyle(cellStyle);
      cellHours.setCellValue(workinghrs);
  }
    
}

 
class Employees
{
    public static void main(String args[]) 
    {

        fullTimeEmployees ob1 = new fullTimeEmployees();
      
        ob1.display();
        
        
        
    }
    
   
}