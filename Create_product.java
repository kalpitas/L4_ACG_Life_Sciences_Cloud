package verishield300;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.time.Duration;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.edge.EdgeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

public class add_Product100 {

	public static void main(String[] args) throws InterruptedException, FileNotFoundException, IOException {
		
		 System.setProperty("webdriver.edge.driver",
		 		"msedgedriver.exe");
	        EdgeOptions options = new EdgeOptions();
	        WebDriver driver = new EdgeDriver();
	        driver.manage().window().maximize();
	        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

            try (FileInputStream fis = new FileInputStream("ACG_Common_Workbook.xlsx");
   	             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {
   	            XSSFSheet sheet = workbook.getSheet("Product");  // Select the correct sheet by name
                    XSSFSheet sheet1 = workbook.getSheet("URL_Login_cred_Tenant");
	            Thread.sleep(3000);

	            // Open URL
	            driver.get("https://proud-mud-040601f00.4.azurestaticapps.net/");
	            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(3));
	            Thread.sleep(3000);
	            
	            Row row1 = sheet1.getRow(2);
 
            	    String userName = getCellValueAsString(row1.getCell(1));
                    String password1 = getCellValueAsString(row1.getCell(2));
 
                    WebElement email = driver.findElement(By.xpath("//input[@placeholder='Enter username']"));
                    email.sendKeys(userName);
                    Thread.sleep(3000);
 
                    WebElement password = driver.findElement(By.xpath("//input[@placeholder='Enter a password']"));
                    password.sendKeys(password1);
                    Thread.sleep(3000);

	            WebElement login = driver.findElement(By.xpath("/html/body/div[1]/div/div/div[4]/div/div[2]/div/button"));
	            login.click();
	            
	            Thread.sleep(6000);
	            // Navigate to the Add Product page
	            wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[text()= 'Master Data']")));
	    		driver.findElement(By.xpath("//span[text()= 'Master Data']")).click();
	            
	            // Loop through all the rows in the Excel sheet
	            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
	            	
	            // Click 'Add Product' button at the beginning of each loop
	           // wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//a[@href= '/add/product']")));
	            WebElement addProd = driver.findElement(By.xpath("//a[@href= '/add/product']"));
	            addProd.click();
	           
	               Thread.sleep(4000);
	                XSSFRow row = sheet.getRow(rowIndex);
	                String ProductIdentifier = getCellValueAsString(row.getCell(0));
	                System.out.println(ProductIdentifier);
	                String productName = getCellValueAsString(row.getCell(1));
	                String productDescription = getCellValueAsString(row.getCell(2));
	                String manufacturerName= getCellValueAsString(row.getCell(3));
	                String GS1CompanyPrefix = getCellValueAsString(row.getCell(4));
	                String GLN = getCellValueAsString(row.getCell(5));
	                String ProductIdentifier2 = getCellValueAsString(row.getCell(6));
	                String packagingLevelIndicator = getCellValueAsString(row.getCell(7));
	                String genericName = getCellValueAsString(row.getCell(8));
	                String MinTemp = getCellValueAsString(row.getCell(9));
	                String weight = getCellValueAsString(row.getCell(10));
	                String strength = getCellValueAsString(row.getCell(11));
	                
	                
            WebElement productIdentifierType = driver.findElement(By.xpath("//select[@name= 'productIdentifierType']"));
       		Select s=new Select(productIdentifierType);
       		s.selectByIndex(2);
       		Thread.sleep(3000);
       	// Fill in the Product details
       		//1.product Identifier
            WebElement ProductIdentifier1 = driver.findElement(By.xpath("//input[@name='productIdentifier']"));
            ProductIdentifier1.sendKeys(ProductIdentifier);
            Thread.sleep(3000);
            
            //2.product Name
           WebElement ProductName = driver.findElement(By.xpath("//input[@name='productName']"));
           ProductName.sendKeys(productName);
           Thread.sleep(3000);
           
           //3.product Description
           WebElement ProductDescription = driver.findElement(By.xpath("//input[@name='productDescription']"));
           ProductDescription.sendKeys(productDescription);
           Thread.sleep(3000);
               
            WebElement element = driver.findElement(By.xpath("//select[@name= 'manufacturer']"));
       		JavascriptExecutor js = (JavascriptExecutor) driver;
       		js.executeScript("arguments[0].scrollIntoView()",element);
       		
       		WebElement manufacturer = driver.findElement(By.xpath("//select[@name= 'manufacturer']"));
       		manufacturer.click();
       		
       		Select s1=new Select(manufacturer);
       		s1.selectByVisibleText("Others");
       		
       		WebElement ManufacturerName = driver.findElement(By.xpath("//input[@name='manufacturerName']"));
       		ManufacturerName.sendKeys(manufacturerName);
       		Thread.sleep(3000);
       		       		
	       	WebElement GS1CompanyPrefix2 = driver.findElement(By.xpath("//input[@name='manufacturerOtherGCP']"));
	       	GS1CompanyPrefix2.sendKeys(GS1CompanyPrefix);
	        Thread.sleep(3000);
	        
	        WebElement AddGLNClick = driver.findElement(By.xpath("//b[text()='+ Add GLN ']"));
	        AddGLNClick.click();
       			        
	        WebElement enterGLN = driver.findElement(By.xpath("//input[@placeholder='Enter GLN']"));
	       	enterGLN.sendKeys(GLN);
	       	Thread.sleep(3000);
       		
	         //NextTo_Packaging Details:
    		WebElement Clicknext = driver.findElement(By.xpath("//button[text()= 'Next']"));
    		Clicknext.click();
    		Thread.sleep(3000);
    		
//    		Packaging Details: //button[text()= 'Add']
    		WebElement packagingCodeType = driver.findElement(By.xpath("//select[@name='packagingCodeType']"));
    		packagingCodeType.click();
    		Select s2=new Select(packagingCodeType);
    		s2.selectByIndex(1);
    		Thread.sleep(3000);
    		
    		//GTIN-14:
    		WebElement packagingProductIdentifier = driver.findElement(By.xpath("//input[@name='packagingProductIdentifier']"));
    		packagingProductIdentifier.sendKeys(ProductIdentifier2);
    		Thread.sleep(3000);
    		//Packaging Level:
    		WebElement packagingLevel1 = driver.findElement(By.xpath("//select[@name='packagingLevel']"));
    		packagingLevel1.click();
    		Thread.sleep(3000);
    		Select s3=new Select(packagingLevel1);
    		s3.selectByIndex(1);
    		Thread.sleep(3000);
    		//Packaging Level Indicator:
    		WebElement packagingLevel2 = driver.findElement(By.xpath("//input[@name='packagingLevelIndicator']"));
    		packagingLevel2.sendKeys(packagingLevelIndicator);
    		Thread.sleep(3000);
    		
    		//GTIN-14: Add button
    		WebElement addGTIN = driver.findElement(By.xpath("//button[text()= 'Add']"));
    		addGTIN.click();
    		Thread.sleep(3000);
    		
    		//NextTo_Other Details:
    		WebElement nextOD = driver.findElement(By.xpath("//button[text()= 'Next']"));
    		nextOD.click();
    		Thread.sleep(3000);
    		
    		WebElement genericName1 = driver.findElement(By.xpath("//input[@placeholder='Enter generic name']"));
    		genericName1.sendKeys(genericName);
    		Thread.sleep(3000);
    		
    		WebElement enterMinTemp = driver.findElement(By.xpath("//input[@placeholder='Enter min temperature']"));
    		enterMinTemp.sendKeys(MinTemp);
    		Thread.sleep(3000);    		
    		
       		//NextTo_Regulatory Details:
    		WebElement nextRegDetails = driver.findElement(By.xpath("//button[text()= 'Next']"));
    		nextRegDetails.click();
    		Thread.sleep(3000);
    		
    		WebElement clickNewRegulation = driver.findElement(By.xpath("//b[text()= '+ New Regulation']"));
    		clickNewRegulation.click();
    		Thread.sleep(3000);
    		
    		WebElement countryselect = driver.findElement(By.xpath("//select[@name= 'country']"));
    		countryselect.click();
    		Thread.sleep(3000);
    		Select s4=new Select(countryselect);
    		s4.selectByIndex(1);
    		Thread.sleep(3000);
    		
    		WebElement selectRegulation = driver.findElement(By.xpath("//select[@name= 'regulation']"));
    		selectRegulation.click();
    		Thread.sleep(3000);
    		Select s5=new Select(selectRegulation);
    		s5.selectByIndex(1);
    		Thread.sleep(3000);
    		
    		WebElement enterWgt1 = driver.findElement(By.xpath("//input[@placeholder= 'Enter weight (gm)']"));
    		enterWgt1.sendKeys(weight);
    		Thread.sleep(3000);
    		
    		WebElement strngth = driver.findElement(By.xpath("//input[@name='strength (mg)']"));
    		strngth.sendKeys(strength);
    		Thread.sleep(3000);
    		//button[text()= 'Accept']
    		WebElement acceptclick = driver.findElement(By.xpath("//button[text()= 'Accept']"));
    		acceptclick.click();
    		Thread.sleep(3000);
    		
    		//NextTo_custom Details:
    		WebElement nextCustomDetails = driver.findElement(By.xpath("//button[text()= 'Next']"));
    		nextCustomDetails.click();
    		Thread.sleep(3000);
    		
    		//Submit Details:
    		WebElement submit = driver.findElement(By.xpath("//button[text()= 'Submit']"));
    		submit.click();
    		Thread.sleep(3000);
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
