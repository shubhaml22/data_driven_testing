package pages;


import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import utils.ExcelUtils;

import java.io.IOException;
import java.time.Duration;

public class FD_calculator {
    public static void main(String[] args) throws IOException, InterruptedException {
        WebDriver driver=new ChromeDriver();
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
        driver.get("https://www.moneycontrol.com/fixed-income/calculator/state-bank-of-india/fixed-deposit-calculator-SBI-BSB001.html");
        driver.manage().window().maximize();

        String filepath=System.getProperty("user.dir")+"//test_data//demo_interest_data.xlsx";
        int rownum= ExcelUtils.getRowCount(filepath,"sheet1");

        for(int i=1;i<=rownum;i++){
            //1) Reading the Excel file data
            String pri=ExcelUtils.getCellData(filepath,"Sheet1",i,0);
            String rateOfIn=ExcelUtils.getCellData(filepath,"Sheet1",i,1);
            String pr1=ExcelUtils.getCellData(filepath,"Sheet1",i,2);
            String pr2=ExcelUtils.getCellData(filepath,"Sheet1",i,3);
            String fre=ExcelUtils.getCellData(filepath,"Sheet1",i,4);

            String exp_val=ExcelUtils.getCellData(filepath,"Sheet1",i,5);

//            System.out.println(pri+" "+rateOfIn+" "+pr1+" "+pr2+" "+fre);
            //2)Put above data in website
            driver.findElement(By.xpath("//input[@id='principal']")).sendKeys(pri);
            driver.findElement(By.xpath("//input[@id='interest']")).sendKeys(rateOfIn);
            driver.findElement(By.xpath("//input[@id='tenure']")).sendKeys(pr1);

            Select pr_time=new Select(driver.findElement(By.xpath("//select[@id='tenurePeriod']")));
            pr_time.selectByVisibleText(pr2);

            Select fredwn=new Select(driver.findElement(By.xpath("//select[@id='frequency']")));
            fredwn.selectByIndex(0);
//            driver.findElement(By.xpath("//*[@id=\"fdMatVal\"]//div[2]//a[1]")).click();
            // Wait for overlay to disappear
            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
            wait.until(ExpectedConditions.invisibilityOfElementLocated(By.className("wzrk-overlay")));

// Then click the element
            WebElement element = driver.findElement(By.xpath("//*[@id='fdMatVal']/div[2]/a[1]"));
            element.click();


            String act_val=driver.findElement(By.xpath("//*[@id=\"resp_matval\"]//strong")).getText();
            //3)validation
            if(Double.parseDouble(act_val)==Double.parseDouble(exp_val)){
                System.out.println("Test Pass!");
                ExcelUtils.fillGreenColor(filepath,"Sheet1",i,6);
            }
            else {
                System.out.println("Test Failed!");
                ExcelUtils.fillRedColor(filepath,"Sheet1",i,6);
            }
            Thread.sleep(10000);
            driver.findElement(By.xpath("//img[@class='PL5']")).click();
        }

        driver.quit();


    }
}
