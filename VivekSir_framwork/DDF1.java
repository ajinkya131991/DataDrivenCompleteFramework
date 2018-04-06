package VivekSir_framwork;


	import java.io.File;
	import java.io.FileInputStream;
	import java.io.FileOutputStream;
	import java.util.concurrent.TimeUnit;

	import org.apache.poi.hssf.usermodel.HSSFCell;
	import org.apache.poi.hssf.usermodel.HSSFRow;
	import org.apache.poi.hssf.usermodel.HSSFSheet;
	import org.apache.poi.hssf.usermodel.HSSFWorkbook;
	import org.junit.After;
	import org.junit.Before;
	import org.junit.Test;
	import org.openqa.selenium.By;
	import org.openqa.selenium.WebDriver;
	import org.openqa.selenium.firefox.FirefoxDriver;

	public class DDF1{
		// This is our 1st WebDriver code
		
		// 1. Define the Selenium WebDriver.
		WebDriver driver;
		String vUrl;
		String xlPath, xlSheet, xlPathRes;
		String[][] xlTestData; // Define a 2D string
		int xRows, xCols;
		String vActualValue;
		
		@Before // Run this before any @Test
		public void myBefore() throws Exception{
			
			driver = new FirefoxDriver();
		    driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
		    
		    vUrl = "http://www.bankrate.com/calculators/auto/auto-loan-calculator.aspx";
		    
		    xlPath = "C:\\Users\\ajinkya.bh\\workspace\\Viveksir_framwork\\DDF Plan.xls";
		    xlPathRes = "C:\\DDF Result1.xls";
		    xlSheet = "TestData";
		    xlTestData = readXL(xlPath, xlSheet);
		}
		
		@Test // Main the test in the code
		public void myTest()throws Exception{
			String vTDSet, vLoanAmt, vTermYears, vRate, vExpectedValue;
			for (int i=1; i<xRows; i++ ){
				vTDSet = xlTestData[i][0];
				vLoanAmt = xlTestData[i][1];
				vTermYears = xlTestData[i][2];
				vRate = xlTestData[i][3];
				vExpectedValue = xlTestData[i][4];

				xlTestData[i][6] = bankRateTest1(vUrl, vLoanAmt, vTermYears, vRate, vExpectedValue);
				xlTestData[i][7] = bankRateTest2(vTermYears);
				xlTestData[i][5] = vActualValue;
			}

			// **** Test 2. See if the months are calculated correctly
			
			// 7. Close the website	-
			driver.quit();
		}
		
		@After //Runs after any @Test
		public void myAfter() throws Exception{
			System.out.println("Hi after");
			writeXL(xlPathRes, "TestDataRes", xlTestData);
		}
		
		public String bankRateTest1(String fURL, String fLA, String fTerm, String fRate, String fEV) throws InterruptedException{
			// Input : URL, LA, Term, Rate, ExpectedValue
			// Output : Pass or Fail
			
			// Driver, go to the main url or the base url
			// 1. Open the website	
			driver.get(fURL);
			
			// 2. Enter Auto Loan Amount	23000
			driver.findElement(By.xpath("//input[@id='loanAmount']")).clear();
			driver.findElement(By.xpath("//input[@id='loanAmount']")).sendKeys(fLA);
					
			// 3. Enter Term in years	5
			driver.findElement(By.xpath("//input[@id='years']")).clear();
			driver.findElement(By.xpath("//input[@id='years']")).sendKeys(fTerm);
					
			// 4. Rate Per Year	2.99
			driver.findElement(By.xpath("//input[@id='interestRate']")).clear();
			driver.findElement(By.xpath("//input[@id='interestRate']")).sendKeys(fRate);
			
			// 5. Click on Calculate	-
			driver.findElement(By.xpath("//button[@id='calcButton']")).click();
			Thread.sleep(4000);// Dear program, at this step wait for 2 seconds.
					
			// 6. Verify the payment	$413.18
			// Where is the payment?
			vActualValue = driver.findElement(By.xpath("//span[@id='mpay']")).getText();
			// What is the value?
			System.out.println("The value on the website is " + vActualValue);
			// Compare with what is expected.
		
			System.out.println("The expected value is " + fEV);
			
			if (vActualValue.equals(fEV)){
				System.out.println("Test is a pass");
				return "Pass";
			} else {
				System.out.println("Test is a fail");
				return "Fail";
			}
		
		}

		public String bankRateTest2(String fTerm){
			// Calculate the term in months
			int fTermMonths;
			String fTermActual;
			
			fTermMonths = Integer.parseInt(fTerm) * 12;
			System.out.println("Expected Term in months is " + fTermMonths);
			
			// Get the actual value from the website
				// Where is the Element
				fTermActual = driver.findElement(By.xpath("//input[@id='terms']")).getAttribute("value");
				// Which Attribute has that value
			
			// Compare if they are the same
			if (Integer.parseInt(fTermActual) == fTermMonths){
				System.out.println("Test2 Pass");
				return "Pass";
			} else {
				System.out.println("Test2 Fail");
				return "Fail";
			}
			
		}
		 // Teach Java to R/W from MS Excel
			// Method to read XL
			public String[][] readXL(String fPath, String fSheet) throws Exception{
				// Inputs : XL Path and XL Sheet name
				// Output : 
				
					String[][] xData;   

					File myxl = new File(fPath);                                
					FileInputStream myStream = new FileInputStream(myxl);                                
					HSSFWorkbook myWB = new HSSFWorkbook(myStream);                                
					HSSFSheet mySheet = myWB.getSheet(fSheet);                                 
					xRows = mySheet.getLastRowNum()+1;                                
					xCols = mySheet.getRow(0).getLastCellNum();   
					System.out.println("Total Rows in Excel are " + xRows);
					System.out.println("Total Cols in Excel are " + xCols);
					xData = new String[xRows][xCols];        
					for (int i = 0; i < xRows; i++) {                           
							HSSFRow row = mySheet.getRow(i);
							for (int j = 0; j < xCols; j++) {                               
								HSSFCell cell = row.getCell(j);
								String value = "-";
								if (cell!=null){
									value = cellToString(cell);
								}
								xData[i][j] = value;      
								System.out.print(value);
								System.out.print("----");
								}        
							System.out.println("");
							}    
					myxl = null; // Memory gets released
					return xData;
			}
			
			//Change cell type
			public static String cellToString(HSSFCell cell) { 
				// This function will convert an object of type excel cell to a string value
				int type = cell.getCellType();                        
				Object result;                        
				switch (type) {                            
					case HSSFCell.CELL_TYPE_NUMERIC: //0                                
						result = cell.getNumericCellValue();                                
						break;                            
					case HSSFCell.CELL_TYPE_STRING: //1                                
						result = cell.getStringCellValue();                                
						break;                            
					case HSSFCell.CELL_TYPE_FORMULA: //2                                
						throw new RuntimeException("We can't evaluate formulas in Java");  
					case HSSFCell.CELL_TYPE_BLANK: //3                                
						result = "%";                                
						break;                            
					case HSSFCell.CELL_TYPE_BOOLEAN: //4     
						result = cell.getBooleanCellValue();       
						break;                            
					case HSSFCell.CELL_TYPE_ERROR: //5       
						throw new RuntimeException ("This cell has an error");    
					default:                  
						throw new RuntimeException("We don't support this cell type: " + type); 
						}                        
				return result.toString();      
				}
			
			// Method to write into an XL
			public void writeXL(String fPath, String fSheet, String[][] xData) throws Exception{

			    	File outFile = new File(fPath);
			        HSSFWorkbook wb = new HSSFWorkbook();
			        HSSFSheet osheet = wb.createSheet(fSheet);
			        int xR_TS = xData.length;
			        int xC_TS = xData[0].length;
			    	for (int myrow = 0; myrow < xR_TS; myrow++) {
				        HSSFRow row = osheet.createRow(myrow);
				        for (int mycol = 0; mycol < xC_TS; mycol++) {
				        	HSSFCell cell = row.createCell(mycol);
				        	cell.setCellType(HSSFCell.CELL_TYPE_STRING);
				        	cell.setCellValue(xData[myrow][mycol]);
				        }
				        FileOutputStream fOut = new FileOutputStream(outFile);
				        wb.write(fOut);
				        fOut.flush();
				        fOut.close();
			    	}
				}
				

	}



