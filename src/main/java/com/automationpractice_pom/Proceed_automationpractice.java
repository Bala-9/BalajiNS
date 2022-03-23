package com.automationpractice_pom;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class Proceed_automationpractice {

	public WebDriver driver;

	@FindBy(xpath ="/html/body/div/div[1]/header/div[3]/div/div/div[4]/div[1]/div[2]/div[4]/a")

	private WebElement proceed1;

	public Proceed_automationpractice(WebDriver driver2) {
		this.driver = driver2;
		PageFactory.initElements(driver, this);

	}

	public WebElement getProceed1() {
		return proceed1;
	}

	public WebElement getProceed2() {
		return proceed2;
	}

	public WebElement getProceed3() {
		return proceed3;
	}

	public WebElement getTc() {
		return tc;
	}

	public WebElement getProceed4() {
		return proceed4;
	}

	public WebElement getPayment() {
		return payment;
	}

	public WebElement getConfirm() {
		return confirm;
	}

	@FindBy(xpath = "(//a[@title=\"Proceed to checkout\"])[2]")

	private WebElement proceed2;

	@FindBy(xpath = "(//button[@type=\"submit\"])[2]")

	private WebElement proceed3;

	@FindBy(id = "cgv")

	private WebElement tc;

	@FindBy(xpath = "(//button[@type=\"submit\"])[2]")

	private WebElement proceed4;

	@FindBy(xpath = "//a[@class=\"cheque\"]")

	private WebElement payment;

	@FindBy(xpath = "(//button[@type=\"submit\"])[2]")

	private WebElement confirm;

}
