package VivekSir_framwork;


	import java.io.File;
	import java.io.FileInputStream;
	import java.io.FileOutputStream;
	import java.io.IOException;
	import java.util.concurrent.TimeUnit;

	import org.apache.commons.io.FileUtils;
	import org.apache.poi.hssf.usermodel.HSSFCell;
	import org.apache.poi.hssf.usermodel.HSSFRow;
	import org.apache.poi.hssf.usermodel.HSSFSheet;
	import org.apache.poi.hssf.usermodel.HSSFWorkbook;
	import org.junit.After;
	import org.junit.Before;
	import org.junit.Test;
	import org.openqa.selenium.By;
	import org.openqa.selenium.OutputType;
	import org.openqa.selenium.TakesScreenshot;
	import org.openqa.selenium.WebDriver;
	import org.openqa.selenium.firefox.FirefoxDriver;

	public class Kdf1 {
		// Global Variables
		String xlPath, xlRes_TS, xlRes_TC;
		int xRows_TC, xRows_TS, xCols_TC, xCols_TS;
		String[][] xlTC, xlTS;
		String vKW, vIP1, vIP2;
		WebDriver driver;
		String vTS_Res, vTC_Res;
		
		
		@Before // Run this before any @Test
		public void myBefore() throws Exception{
			// driver = new FirefoxDriver();
		    // driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
		    
		    xlPath = "C:\\KDF4.xls";
		    xlRes_TS= "C:\\KDF4_TS_Res.xls";
		    xlRes_TC= "C:\\KDF4_TC_Res.xls";
			xlTC = readXL(xlPath, "Test Cases");
			xlTS = readXL(xlPath, "Test Steps");
			
			xRows_TC = xlTC.length;
			xCols_TC = xlTC[0].length;
			System.out.println("TC Rows are " + xRows_TC);
			System.out.println("TC Cols are " + xCols_TC);
			
			xRows_TS = xlTS.length;
			xCols_TS = xlTS[0].length;
			System.out.println("TS Rows are " + xRows_TS);
			System.out.println("TS Cols are " + xCols_TS);
		}    
		
		@Test
		public void mainTest() throws Exception{
			/*
			 * 1. Read the Excel sheet ... TC / TS
			 * 2. Go to each row in the TC sheet, see if it is ready to execute
			 * 3. Go to each row in TS sheet, and see if it is corresponding to that Test Case
			 * 4. Get the KW, IP1, IP2 for each step
			 * 5. Call the corresponding function
			 */
			
			
			for (int i=1; i<xRows_TC ; i++){
				if (xlTC[i][2].equals("Y")){
					System.out.println("TC ready for execution : " + xlTC[i][0]);
					vTC_Res = "Pass";
					for (int j=1; j<xRows_TS; j++){
						if (xlTC[i][0].equals(xlTS[j][0])){
							vKW = xlTS[j][3];
							vIP1 = xlTS[j][4];
							vIP2 = xlTS[j][5];
							vTS_Res = "Pass";
							System.out.println("KW: " + vKW);
							System.out.println("IP1: " + vIP1);
							System.out.println("IP2: " + vIP2);
							try {
								executeKW(vKW, vIP1, vIP2);
								if (vTS_Res.equals("Pass")){
									vTS_Res = "Pass";
								} else {
									vTS_Res = "Verification Failed";
									vTC_Res = "Fail";
									xlTS[j][7] = "Look at the screenshot.";
									takeScreenshot("C:\\SLT_Oct_2015\\"+xlTC[i][0]+"_"+j+".jpg");
								}
							} catch (Exception myError){
								System.out.println("Error : " + myError);
								vTS_Res = "Fail";
								vTC_Res = "Fail";
								xlTS[j][7] = "Error : " + myError;
								takeScreenshot("C:\\SLT_Oct_2015\\"+xlTC[i][0]+"_"+j+".jpg");
							}
								
							xlTS[j][6] = vTS_Res;
						}
					}	
					xlTC[i][3] = vTC_Res;
				} else {
					System.out.println("TC NOT ready for execution : " + xlTC[i][0]);
				}
			}
		}
		
		@After
		public void myAfterTest() throws Exception{
			writeXL(xlRes_TS, "TestSteps", xlTS);
			writeXL(xlRes_TC, "TestCases", xlTC);
		}
		public void takeScreenshot(String fPath) throws Exception{
			File scrFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
			// Now you can do whatever you need to do with it, for example copy somewhere
			FileUtils.copyFile(scrFile, new File(fPath));
		}
		public void executeKW(String fKW, String fIP1, String fIP2){
			// Purpose: Executes the corr. function
			// I/P: KW, IP1, IP2
			// O/P:
			
			switch (fKW){
				case "goToUrl":
						goToUrl(fIP1);
					break;
				case "clearText":
						clearText(fIP1);
						break;
				case "typeText":
						typeText(fIP1, fIP2);
						break;
				case "clickElement":
						clickElement(fIP1);
						break;
				case "closeBrowser":
						closeBrowser();
						break;
				case "verifyText":
						vTS_Res = verifyText(fIP1, fIP2);
						break;
				case "verifyValue":
					vTS_Res = verifyValue(fIP1, fIP2);
					break;
				case "launchDriver":
						launchDriver();
						break;
				default :
					System.out.println("Keyword is missing.");
			
			}
		}
		
		// Reusable web based actions performed by users
		
			public void launchDriver(){
				// Purpose: Launches a firefoxdriver
				// I/P: -
				// O/P: -
				driver = new FirefoxDriver();
			    driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
			}
			public void clearText(String fXPath){
				// Purpose: Clears any text present in a editable text field
				// I/P: xPath of the element that you want to clear
				// O/P:
				
				driver.findElement(By.xpath(fXPath)).clear();
			}
			
			public void typeText(String fXPath, String fText){
				// Purpose: Types text into an editable text field
				// I/P: xPath of the element, and the text you need to enter
				// O/P:
				
				driver.findElement(By.xpath(fXPath)).sendKeys(fText);
			}
			
			public void clickElement(String fXPath){
				// Purpose: Clicks on any element on webpage
				// I/P: xPath of the element
				// O/P:
				
				driver.findElement(By.xpath(fXPath)).click();
			}
			
			public void goToUrl(String fUrl){
				// Purpose: Takes the browser to a URL
				// I/P: URL
				// O/P:
				
				driver.get(fUrl);
			}
			
			public void waitFor(int fMiliSeconds) throws Exception{
				// Purpose: Make the program wait for certain time
				// I/P: Milli seconds to wait
				// O/P:
				Thread.sleep(fMiliSeconds);
				
			}
			
			public void closeBrowser(){
				// Purpose: Close the browser
				// I/P: 
				// O/P:
				
				driver.quit();
			}
			
			public String verifyText(String fXP, String fText){
				// Purpose: Verifies a text in a specific element
				// I/P: xPath, Text to verify
				// O/P: pass or fail
				
				String fAppText;
				
				fAppText = driver.findElement(By.xpath(fXP)).getText();
				
				if (fAppText.equals(fText)){
					return "Pass";
				} else {
					return "Fail";
				}
			}
			
			public String verifyValue(String fXP, String fText){
				// Purpose: Verifies a value in a specific element
				// I/P: xPath, Text to verify
				// O/P: pass or fail
				
				String fAppText;
				
				fAppText = driver.findElement(By.xpath(fXP)).getAttribute("value");
				
				if (fAppText.equals(fText)){
					return "Pass";
				} else {
					return "Fail";
				}
			}
		
		// Teach Java to R/W from MS Excel
		// Method to read XL
		public String[][] readXL(String fPath, String fSheet) throws Exception{
			// Inputs : XL Path and XL Sheet name
			// Output : 
			
				String[][] xData;  
				int xRows, xCols;

				File myxl = new File(fPath);                                
				FileInputStream myStream = new FileInputStream(myxl);                                
				HSSFWorkbook myWB = new HSSFWorkbook(myStream);                                
				HSSFSheet mySheet = myWB.getSheet(fSheet);                                 
				xRows = mySheet.getLastRowNum()+1;                                
				xCols = mySheet.getRow(0).getLastCellNum();   
				//System.out.println("Total Rows in Excel are " + xRows);
				//System.out.println("Total Cols in Excel are " + xCols);
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
							//System.out.print(value);
							//System.out.print("----");
							}        
						//System.out.println("");
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



