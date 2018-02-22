import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.winium.DesktopOptions;
import org.openqa.selenium.winium.WiniumDriver;
import org.openqa.selenium.winium.WiniumDriverService;

import java.io.File;
import java.io.IOException;

public class WiniumTest {

    public static void main(String[] args) throws IOException, InterruptedException {
        DesktopOptions options = new DesktopOptions();
        options.setApplicationPath("C:\\windows\\system32\\calc.exe");
        File wDriverExe = new File("C:\\Program Files (x86)\\winium\\Winium.Desktop.Driver.exe");
        WiniumDriverService  service = new WiniumDriverService.Builder().usingDriverExecutable(wDriverExe).usingPort(9999).withVerbose(true).withSilent(false).buildDesktopService();
        service.start(); //Build and Start a Winium Driver service
        WiniumDriver driver = new WiniumDriver(service, options); //Start a winium driver
        Thread.sleep(3000);
        WebElement window =  driver.findElementByClassName("CalcFrame");
        WebElement menuItem = window.findElement(By.id("MenuBar")).findElement(By.name("View"));
        menuItem.click();
        driver.findElementByName("Scientific").click();

        window.findElement(By.id("MenuBar")).findElement(By.name("View")).click();
        driver.findElementByName("History").click();
        service.stop();
        driver.close();
    }
}
