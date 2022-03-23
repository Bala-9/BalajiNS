package com.automationpractice_pom;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class Signout_automationpractice {

	public WebDriver driver;

	@FindBy(xpath = "//a[@class=\"logout\"]")

	private WebElement Signout;

	public Signout_automationpractice(WebDriver driver2) {
		this.driver = driver2;
		PageFactory.initElements(driver, this);

	}

	public WebElement getSignout() {
		return Signout;

	}

}
