package tests;

import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;
import pages.LoginPage;
import utils.ExcelReader;
import java.io.IOException;

public class LoginTest extends BaseTest {
    @DataProvider(name = "loginData")
    public Object[][] loginDataProvider() throws IOException {
        String filePath = "C:\\Users\\GAIF\\eclipse-workspace\\Homework2\\src\\testData\\LoginData.xlsx";
        return ExcelReader.readExcelData(filePath, "Sheet1");
    }

    @Test(dataProvider = "loginData")
    public void testLogin(String email, String password, String expectedMessage) {
        LoginPage loginPage = new LoginPage(driver);
        loginPage.enterEmail(email);
        loginPage.enterPassword(password);
        loginPage.clickLogin();

        String actualMessage = loginPage.getDisplayedMessage();
        Assert.assertEquals(actualMessage, expectedMessage, "The displayed message does not match the expected result.");
    }
}





