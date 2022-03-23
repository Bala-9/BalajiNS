package com.automationpractice_pom;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class Home_page_automationpractice {
	
	public WebDriver driver;
	
	
	@FindBy(xpath="//a[@class=\"login\"]")
	
	private WebElement signin_btn ;

	public Home_page_automationpractice(WebDriver driver2) {
		this.driver = driver2;
		
		PageFactory.initElements(driver, this);
	
	}

	public WebElement getSignin_btn() {
		return signin_btn;
	}
	
	

}
