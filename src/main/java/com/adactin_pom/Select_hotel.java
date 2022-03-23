package com.adactin_pom;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class Select_hotel {

	public WebDriver driver;

	@FindBy(id = "radiobutton_0")

	private WebElement radiobutton;

	public Select_hotel(WebDriver driver2) {
		this.driver = driver2;
		
		PageFactory.initElements(driver, this);

	}

	public WebElement getRadiobutton() {
		return radiobutton;
	}

	public WebElement getContinue() {
		return Continue;
	}

	@FindBy(id = "continue")

	private WebElement Continue;

}
