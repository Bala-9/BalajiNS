package com.adactin_pom;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class Book_hotels {
	
	public WebDriver driver;
	
	@FindBy(id="first_name")
	
	private WebElement firstname;

	public Book_hotels(WebDriver driver2) {
		this.driver = driver2;
		
		PageFactory.initElements(driver, this);
	
	}

	public WebElement getFirstname() {
		return firstname;
	}

	public WebElement getLastname() {
		return lastname;
	}

	public WebElement getBilladdress() {
		return billaddress;
	}

	public WebElement getCcnum() {
		
		return ccnum;
	}

	public WebElement getCctype() {
		return cctype;
	}

	public WebElement getCcexpmonth() {
		return ccexpmonth;
	}

	public WebElement getCcexpyeaar() {
		return ccexpyeaar;
	}

	public WebElement getCccvv() {
		return cccvv;
	}

	public WebElement getBooknow() {
		return booknow;
	}

	@FindBy(id="last_name")
	
	private WebElement lastname;
	
	@FindBy(id="address")
	
	private WebElement billaddress;
	
	@FindBy(id="cc_num")
	
	private WebElement ccnum;
	
	@FindBy(id="cc_type")
	
	private WebElement cctype;
	
	@FindBy(id="cc_exp_month")
	
	private WebElement ccexpmonth;
	
	@FindBy(id="cc_exp_year")
	
	private WebElement ccexpyeaar;
	
	@FindBy(id="cc_cvv")
	
	private WebElement cccvv;
	
	@FindBy(id="book_now")
	
	private WebElement booknow;

}
