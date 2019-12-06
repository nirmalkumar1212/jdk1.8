package Timesheet;
import org.testng.annotations.Test;
import Excel.Excel;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Parameters;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
//import java.util.Iterator;
import java.util.Properties;
//import java.util.Set;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.ie.InternetExplorerDriver;
//import org.openqa.selenium.interactions.Action;
import org.openqa.selenium.interactions.Actions;
import org.testng.annotations.AfterTest;

public class Timesheet {
	
	FileInputStream ip; 
 	 Properties prop=new Properties();
 	 WebDriver driver;
 	
 	@Parameters("Browser")
 	
  	  @BeforeTest
  	  public void launch(String b) throws IOException, InterruptedException
  	  {
  		 ip=new FileInputStream("C:\\Users\\701288\\eclipse-workspace\\Timesheet\\Properties\\Config.properties");
          prop.load(ip);
  		
  		if(b.equalsIgnoreCase("chrome"))
  		{
  			
  		    String web=prop.getProperty("path");
 	         String loc=prop.getProperty("location");
      	
  		 System.setProperty(web,loc);
  			ChromeOptions options=new ChromeOptions(); 
  	      options.addArguments("--disable-notifications");
           
  	   driver=new ChromeDriver(options);
  	  }
  		else if(b.equalsIgnoreCase("ie"))
  		{
  			String web1=prop.getProperty("path1");
 	         String loc1=prop.getProperty("location1");
  			System.setProperty(web1,loc1);
  		      driver= new InternetExplorerDriver();
  		     // InternetExplorerOptions options1=new InternetExplorerOptions();   	        
  	  }
  		 String url1=prop.getProperty("url");
  		driver.get(url1);
         driver.manage().window().maximize();
         Thread.sleep(9000);
  }
 	
  	@Test
    public void Submit() throws IOException, InterruptedException {
  		
  		Excel clone=new Excel();
		 
 		//Entering time sheet in search box
       String b= clone.getdata(0, 0, 0);
       String Search=prop.getProperty("Searchbox");
       driver.findElement(By.xpath(Search)).sendKeys(b);
       Thread.sleep(3000);
       driver.findElement(By.xpath(Search)).sendKeys(Keys.chord(Keys.ARROW_DOWN,Keys.ENTER));
       Thread.sleep(2000);
       
       
     //  clicking on submit time sheet icon on second page 
            WebElement iframe=driver.findElement(By.xpath("//iframe[@id='appFrame']"));
	        driver.switchTo().frame(iframe);
	        driver.findElement(By.xpath("//div[@class='jspContainer']")).click();
	       Actions actions = new Actions(driver);
            WebElement target = driver.findElement(By.xpath("(//div[@class='inActive'])[1]"));
           actions.moveToElement(target).perform();
           target.click();
           Thread.sleep(12000);
          
            //To open second Tab
          Actions act = new Actions(driver);      
          act.keyDown(Keys.CONTROL).sendKeys("t").keyUp(Keys.CONTROL).build().perform();
          
          //to be in second tab
          ArrayList<String> tabs2 = new ArrayList<String> (driver.getWindowHandles());
          driver.switchTo().window(tabs2.get(1));
          Thread.sleep(15000);
          driver.manage().window().maximize();
          Thread.sleep(8000);
          
         // clicking on time sheet submit date 
          driver.findElement(By.xpath("//a[@id='CTS_TS_LAND_PER_DESCR30$0']")).click();
          Thread.sleep(7000);
          
          
          //To open third Tab
          Actions act1 = new Actions(driver);      
          act1.keyDown(Keys.CONTROL).sendKeys("t").keyUp(Keys.CONTROL).build().perform();
          
          //to be in third tab
          ArrayList<String> tabs3 = new ArrayList<String> (driver.getWindowHandles());
          driver.switchTo().window(tabs3.get(1));
          
         WebElement iframe2=driver.findElement(By.xpath("//iframe[@id='ptifrmtgtframe']"));
          driver.switchTo().frame(iframe2);
           driver.findElement(By.xpath("//form[@id='TE_TIME_ENTRY']")).click();
          driver.findElement(By.xpath("//div[@class='ps_pagecontainer']")).click();
          
          //clicking on img icon to select project 
           driver.findElement(By.xpath("//img[@id='PROJECT_CODE$prompt$img$0']")).click();
          Thread.sleep(5000);
          
        //to go screen for select project 
          driver.switchTo().window(tabs3.get(1));
          WebElement iframe8=driver.findElement(By.xpath("//iframe[@id='ptModFrame_0']"));
          driver.switchTo().frame(iframe8);
          Thread.sleep(3000);
          
        //enter project from excel
          String project= clone.getdata(0, 7, 0);
          String Pj1=prop.getProperty("Pj");
          driver.findElement(By.xpath(Pj1)).sendKeys(project);
          Thread.sleep(3000);
          
          //to search given project
          driver.findElement(By.xpath(Pj1)).sendKeys(Keys.chord(Keys.ENTER));
          Thread.sleep(1000);
          
          //select project
          driver.findElement(By.xpath("//a[@id='SEARCH_RESULT1']")).click();
          Thread.sleep(3000);
          
          //switching to back and select activity 
          driver.switchTo().defaultContent();
          driver.switchTo().window(tabs3.get(1));
          WebElement iframe3=driver.findElement(By.xpath("//iframe[@id='ptifrmtgtframe']"));
          driver.switchTo().frame(iframe3);
          driver.findElement(By.xpath("//form[@id='TE_TIME_ENTRY']")).click();
          driver.findElement(By.xpath("//div[@class='ps_pagecontainer']")).click();
          
          
          
          
          
          
          
          
          //clicking on search img icon to select activity
          driver.findElement(By.xpath("//img[@id='CTS_EX_ACT_VW_DESCR$prompt$img$0']")).click();
         Thread.sleep(5000);
         
         
          
          
       //to go screen for select activity 
         driver.switchTo().window(tabs3.get(1));
          WebElement iframe4=driver.findElement(By.xpath("//iframe[@id='ptModFrame_1']"));
          driver.switchTo().frame(iframe4);
          Thread.sleep(3000);
          
         //enter activity from excel
          String Activity= clone.getdata(0, 1, 0);
          String Act=prop.getProperty("Ac");
          driver.findElement(By.xpath(Act)).sendKeys(Activity);
          Thread.sleep(3000);
          
          //to search given activity
          driver.findElement(By.xpath(Act)).sendKeys(Keys.chord(Keys.ENTER));
          Thread.sleep(1000);
          
          //select activity
          driver.findElement(By.xpath("//a[@id='SEARCH_RESULT1']")).click();
          Thread.sleep(3000);
          
          //switching to back and enter time 
          driver.switchTo().defaultContent();
          driver.switchTo().window(tabs3.get(1));
          WebElement iframe10=driver.findElement(By.xpath("//iframe[@id='ptifrmtgtframe']"));
          driver.switchTo().frame(iframe10);
          driver.findElement(By.xpath("//form[@id='TE_TIME_ENTRY']")).click();
          driver.findElement(By.xpath("//div[@class='ps_pagecontainer']")).click();
          
          //enter time from Monday to Friday
          String Monday= clone.getdata(0, 2, 0);
          String MT=prop.getProperty("MTF");
          driver.findElement(By.xpath(MT)).sendKeys(Monday);
          Thread.sleep(3000);
          
          String Tuesday= clone.getdata(0, 3, 0);
          String TT=prop.getProperty("TTF");
          driver.findElement(By.xpath(TT)).sendKeys(Tuesday);
          Thread.sleep(3000);
          
        String Wednesday= clone.getdata(0, 4, 0);
          String WT=prop.getProperty("WTF");
          driver.findElement(By.xpath(WT)).sendKeys(Wednesday);
          Thread.sleep(3000);
          
          String Thursday= clone.getdata(0, 5, 0);
          String TTs=prop.getProperty("TTE");
          driver.findElement(By.xpath(TTs)).sendKeys(Thursday);
          Thread.sleep(3000);
          
          String Friday= clone.getdata(0, 6, 0);
          String FT=prop.getProperty("FTF");
          driver.findElement(By.xpath(FT)).sendKeys(Friday);
          Thread.sleep(3000);
          
          //take screenshot
          
          TakesScreenshot tk = (TakesScreenshot)driver;
		  File f = tk.getScreenshotAs(OutputType.FILE);
		 File f1=new File(System.getProperty("user.dir")+"/sample.png");
		   FileUtils.copyFile(f,f1 );
          //click on update total button
          driver.findElement(By.xpath("//input[@id='PB_UPDATE_2']")).click();
          Thread.sleep(6000);
          
          //click on submit button 
          driver.findElement(By.xpath("//input[@id='EX_TIME_HDR_WRK_PB_SUBMIT']")).click();
          Thread.sleep(2000);
          driver.switchTo().activeElement();
          Keys.chord(Keys.ENTER);
          Thread.sleep(2000);
          
         // driver.switchTo().defaultContent();
          
        
          //ensure to submit by click Ok button 
          
        //  driver.switchTo().defaultContent();
         // driver.switchTo().window(tabs3.get(1));
        
          
        /*  WebElement iframe9=driver.findElement(By.xpath("(//iframe[@frameborder='0'])[4]"));
          driver.switchTo().frame(iframe9);
          Keys.chord(Keys.TAB);
          Keys.chord(Keys.ESCAPE);*/
          
          
         // Keys.chord(Keys.ESCAPE);
        //  Keys.chord(Keys.ENTER);
          
          
          
      /*    driver.findElement(By.xpath("//img[@class='PSSTATICIMAGE']")).click();
          
          WebElement confirm=driver.findElement(By.xpath("//span[text()='Click OK to submit, or click Cancel to return to the time report without submitting.']"));
          String Display=confirm.getText();
          System.out.print("Text is :"+Display);*/
          
        //span[text()='Click OK to submit, or click Cancel to return to the time report without submitting.']
          
         /* WebElement iframe0=driver.findElement(By.xpath("//iframe[@id='ptifrmtgtframe']"));
	        driver.switchTo().frame(iframe0);
	        driver.findElement(By.xpath("//div[@class='ps_pagecontainer']")).click();
	       Actions actions1 = new Actions(driver);
          WebElement target1 = driver.findElement(By.xpath("//div[@class='PSMODALCLOSE']"));
         actions1.moveToElement(target1).perform();
         target.click();
         Thread.sleep(12000);
        
          Keys.chord(Keys.ESCAPE);*/
          
         //  driver.findElement(By.xpath("//div[@class='PSMODALCLOSE']")).click();
           
          
         /*
          
          driver.findElement(By.xpath("//div[@id='ptModTable_1']")).click();
          
        
          
           driver.findElement(By.xpath("//form[@id='TE_TIME_ENTRY']")).click();
          driver.findElement(By.xpath("//div[@class='ps_pagecontainer']")).click();
          Thread.sleep(6000);
          
          */
          
          
         Thread.sleep(5000);
          
        
              
          }

  @AfterTest
  public void afterTest() {
	  
	 // driver.close();
  }

}
