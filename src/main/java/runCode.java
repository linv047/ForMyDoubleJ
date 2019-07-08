import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;


public class runCode {
    WebDriver driver;
    WebDriverWait wait;

    @BeforeMethod
    private void setUp() throws Exception {
        System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir") + "\\src\\main\\chromeDriver\\chromedriver.exe");
        ChromeOptions chromeOptions = new ChromeOptions();
//        chromeOptions.addArguments("--headless");
        chromeOptions.addArguments("user-data-dir=" + System.getProperty("user.home") + "\\AppData\\Local\\Google\\Chrome\\User Data");
        driver = new ChromeDriver(chromeOptions);
        wait = new WebDriverWait(driver, 20);


    }

    @AfterMethod
    public void teardown() {
        driver.close();
    }

    private List<String> searchInfo(String name) throws Exception {


        List<String> list = new ArrayList<String>();
        String url = "https://www.qichacha.com/";

        driver.get(url + "/search?key=" + name);

        Thread.sleep(3500);

        WebElement conpany ;
        try {
            conpany= wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//tbody[@id = 'search-result']/tr[1]//a")));
            Thread.sleep(1500);
        } catch (TimeoutException e) {
            list.add("未找到匹配公司");
            list.add("未找到匹配公司");
            return list;
        }
        if (!name.equalsIgnoreCase(conpany.getText())) {
            System.out.println(name + "---" + "未找到完全匹配公司");
            list.add("未找到完全匹配公司");
            list.add("未找到完全匹配公司");
        } else {
            String conpanyLink = conpany.getAttribute("href");
            try {
                driver.get(conpanyLink);
            } catch (TimeoutException e) {
                driver.get(conpanyLink);
            }
            WebElement code = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//td[text()='统一社会信用代码']/following-sibling::td[1]")));
            WebElement boss = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@class='boss-td']//a/h2")));
            System.out.println(name + "--" + boss.getText() + "--" + code.getText());
            list.add(boss.getText());
            list.add(code.getText());
        }
        return list;
    }

    public void readExcel(String fileName) throws Exception {
        File file = new File(fileName);
        if (!"xlsx".equalsIgnoreCase(fileName.substring(fileName.lastIndexOf(".") + 1))) {
            System.out.println("文件格式不对");
            return;
        }

        if (!file.exists()) {
            System.out.println("文件不存在");
            return;
        }

        FileInputStream inputStream = new FileInputStream(file);
        XSSFWorkbook wordBook = new XSSFWorkbook(inputStream);
        //读取工作表,从0开始
        XSSFSheet sheet = wordBook.getSheetAt(0);
        int rowNum = sheet.getLastRowNum();
        for (int i = 1; i <= rowNum; i++) {
            System.out.println("============" + i + "===========");
            XSSFRow row = sheet.getRow(i);
            //读取单元格
            XSSFCell cell = row.getCell(0);//获取单元格对象
            String value = cell.getStringCellValue();
            XSSFCell fcell = row.getCell(1);
            XSSFCell ccell = row.getCell(2);
            List<String> list = null;
            try {
                list = searchInfo(value);
            } catch (Exception e) {
                FileOutputStream outputStream = new FileOutputStream(file);
                wordBook.write(outputStream);
                inputStream.close();
                outputStream.close();
                wordBook.close();
                break;
            }
            if (list.size() == 2) {
                fcell.setCellValue(list.get(0));
                ccell.setCellValue(list.get(1));
            }

        }
        FileOutputStream outputStream = new FileOutputStream(file);
        wordBook.write(outputStream);
        inputStream.close();
        outputStream.close();
        wordBook.close();
    }


    @Test
    public void testRun() throws Exception {
//
//        String filename = "C:\\Users\\victor.lin\\Documents\\Tencent Files\\407004720\\FileRecv\\test3.xlsx";
//        readExcel(filename);

        String name = "桥西鼎兴实业房屋中介";
        searchInfo(name);

    }

}
