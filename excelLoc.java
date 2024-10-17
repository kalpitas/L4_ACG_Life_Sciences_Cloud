package verishield300;
 
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.edge.EdgeOptions;
import org.openqa.selenium.support.ui.Select;

import java.io.FileInputStream;
import java.io.IOException;
 
public class excelLoc {
    public static void main(String[] args) throws InterruptedException, IOException {
        // Set the path to the WebDriver executable
        System.setProperty("webdriver.edge.driver", "msedgedriver.exe");
        EdgeOptions options = new EdgeOptions();
        WebDriver driver = new EdgeDriver();
        driver.manage().window().maximize();
 
        try (FileInputStream fis = new FileInputStream("ACG_Common_Workbook.xlsx");
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {
 
            XSSFSheet sheet = workbook.getSheet("partner");  // Select the correct sheet by name
	    XSSFSheet sheet1 = workbook.getSheet("URL_Login_cred_Tenant");  // Select the correct sheet by name
 
            Thread.sleep(3000);
            // Open URL
            driver.get("https://proud-mud-040601f00.4.azurestaticapps.net/");
            Thread.sleep(3000);
 
            // Perform login
            WebElement usernamelogin = driver.findElement(By.xpath("/html/body/div[1]/div/div/div[4]/div/div[3]/i"));
            usernamelogin.click();
            Thread.sleep(3000);
 
	    Row row1 = sheet1.getRow(2);
 
            String userName = getCellValueAsString(row1.getCell(1));
            String password = getCellValueAsString(row1.getCell(2));
 
            WebElement email = driver.findElement(By.xpath("//input[@placeholder='Enter username']"));
            email.sendKeys(userName);
            Thread.sleep(3000);
 
            WebElement password = driver.findElement(By.xpath("//input[@placeholder='Enter a password']"));
            password.sendKeys(password);
            Thread.sleep(3000);
 
            WebElement login = driver.findElement(By.xpath("/html/body/div[1]/div/div/div[4]/div/div[2]/div/button"));
            login.click();
            Thread.sleep(6000);
 
            // Navigate to the location creation page
            driver.get("https://proud-mud-040601f00.4.azurestaticapps.net/get/locations");
            Thread.sleep(4000);
 
            // Loop through all the rows in the Excel sheet
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
 
                // Click 'New Location' button at the beginning of each loop
                WebElement newloc = driver.findElement(By.xpath("//b[text()='+ New Location']"));
                newloc.click();
                Thread.sleep(4000);
 
                Row row = sheet.getRow(rowIndex);
 
                String locationName = getCellValueAsString(row.getCell(0));
                String locationId = getCellValueAsString(row.getCell(1));
                String state = getCellValueAsString(row.getCell(2));
                String city = getCellValueAsString(row.getCell(3));
                String address = getCellValueAsString(row.getCell(4));
                String postalCode = getCellValueAsString(row.getCell(5));
                String contactPerson = getCellValueAsString(row.getCell(6));
                String emailId = getCellValueAsString(row.getCell(7));
                String phoneNumber = getCellValueAsString(row.getCell(8));
                String website = getCellValueAsString(row.getCell(9));
 
                // Fill in the location details
                WebElement locname = driver.findElement(By.xpath("//input[@placeholder='Enter location name']"));
                locname.sendKeys(locationName);
                Thread.sleep(3000);
                
                Select locType = new Select(driver.findElement(By.xpath("//select[@name='locationType']")));
    	   		locType.selectByValue("Physical Site");
    	   		Thread.sleep(3000);
    	   		
    	   		Select entity = new Select(driver.findElement(By.xpath("//select[@name='entity']")));
    	   		entity.selectByValue("Warehousing Site");
    	   		Thread.sleep(3000);
    	   		
    	   		Select businessEntity = new Select(driver.findElement(By.xpath("//select[@name='locName']"))); 
    	   		businessEntity.selectByValue("HETERO LABS LIMITED");
    	   		Thread.sleep(3000);
    	   		
    	   		Select IDtype = new Select(driver.findElement(By.xpath("//select[@name='locationIdType']"))); 
    	   		IDtype.selectByValue("GLN");
    	   		Thread.sleep(3000);
    	   		
    	   		WebElement identifier = driver.findElement(By.xpath("//input[@name='locationId']"));
                identifier.sendKeys(locationId);
                Thread.sleep(3000);
 
                WebElement next1 = driver.findElement(By.xpath("//button[text()='Next']"));
                next1.click();
                Thread.sleep(3000);
 
                // Fill in the address details
                Select country = new Select(driver.findElement(By.xpath("//select[@name='country']")));
    	   		country.selectByValue("India");
    	   		Thread.sleep(3000);
    	   		
                WebElement stateElement = driver.findElement(By.xpath("//input[@name='state']"));
                stateElement.sendKeys(state);
                Thread.sleep(1000);
 
                WebElement cityElement = driver.findElement(By.xpath("//input[@name='city']"));
                cityElement.sendKeys(city);
                Thread.sleep(1000);
 
                WebElement addressElement = driver.findElement(By.xpath("//input[@name='address']"));
                addressElement.sendKeys(address);
                Thread.sleep(1000);
 
                WebElement pcode = driver.findElement(By.xpath("//input[@name='postalCode']"));
                pcode.sendKeys(postalCode);
                Thread.sleep(1000);
 
                // Scroll into view and fill in contact details
                WebElement element = driver.findElement(By.xpath("//input[@name='contactPersonName']"));
                ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", element);
                Thread.sleep(1000);
 
                WebElement conper = driver.findElement(By.xpath("//input[@placeholder='Enter name']"));
                conper.sendKeys(contactPerson);
                Thread.sleep(1000);
 
                WebElement emailElement = driver.findElement(By.xpath("//input[@placeholder='Enter email']"));
                emailElement.sendKeys(emailId);
                Thread.sleep(1000);
 
                WebElement phno = driver.findElement(By.xpath("//input[@placeholder='Enter phone number']"));
                phno.sendKeys(phoneNumber);
                Thread.sleep(1000);
 
                WebElement websiteElement = driver.findElement(By.xpath("//input[@placeholder='Enter website']"));
                websiteElement.sendKeys(website);
                Thread.sleep(1000);
 
                // Submit the form
                WebElement submit = driver.findElement(By.xpath("//button[text()='Submit']"));
                submit.click();
                Thread.sleep(3000);
 
                // After submitting, wait for the page to refresh before clicking 'New Location' again
                Thread.sleep(4000); // Adjust as necessary
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            // Close the WebDriver session
            // driver.quit();
        }
    }
 
    // Helper method to handle different cell types
    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
 
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf((long) cell.getNumericCellValue());  // Handles numeric values as long
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
}

 