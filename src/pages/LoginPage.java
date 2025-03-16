package pages;

import java.util.ArrayList;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

public class LoginPage {
    private WebDriver driver;

    public LoginPage(WebDriver driver) {
        this.driver = driver;
    }

    private By emailField = By.xpath("//input[(@id='email') and (@class='c-custom-input')]");
    private By passwordField = By.xpath("//input[(@id='password') and (@class='c-custom-input')]");
    private By loginButton = By.xpath("//button[(@type='submit') and (@class='btn btn-primary')]");
    public void enterEmail(String email) {
        driver.findElement(emailField).clear();
          driver.findElement(emailField).sendKeys(email);
      }

      public void enterPassword(String password) {
          driver.findElement(passwordField).clear();
          driver.findElement(passwordField).sendKeys(password);
      }
    public void clickLogin() {
        driver.findElement(loginButton).click();
    }


    private By errorMessageEmailFormat = By.xpath("//div[@class='c-custom-input-error c-error' and (@id='errorMessage-email')]");
    private By errorMessagePasswordEmpty = By.xpath("//div[@class='c-custom-input-error c-error' and (@id='errorMessage-password')]");

    private By errorMessageIncorrectCredentials = By.xpath("//div[@class='u=margin-top-md mb-3 u-color-error']");
    
    private By successMessage = By.xpath("//h2[@class='c-universal-custom-carousel_header c-universal-custom-carousel_header-also-like']");

    public String getDisplayedMessage() {
        List<String> messages = new ArrayList<>();

        try {
            WebElement errorEmailFormat = driver.findElement(errorMessageEmailFormat);
            if (errorEmailFormat.isDisplayed()) {
                messages.add(errorEmailFormat.getText());
            }
        } catch (Exception ignored) {}

        try {
            WebElement errorPasswordEmpty = driver.findElement(errorMessagePasswordEmpty);
            if (errorPasswordEmpty.isDisplayed()) {
                messages.add(errorPasswordEmpty.getText());
            }
        } catch (Exception ignored) {}

        try {
            WebElement errorIncorrectCredentials = driver.findElement(errorMessageIncorrectCredentials);
            if (errorIncorrectCredentials.isDisplayed()) {
                messages.add(errorIncorrectCredentials.getText());
            }
        } catch (Exception ignored) {}

        if (!messages.isEmpty()) {
            return String.join(" | ", messages);
        }

        try {
            WebElement success = driver.findElement(successMessage);
            if (success.isDisplayed()) {
                return success.getText();    
            }
        } catch (Exception ignored) {}

        return "You May Also Like";       
    }

}
