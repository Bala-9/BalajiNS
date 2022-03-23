package com.automationpractice_pom;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class Cart_automationpractice {
	
	public WebDriver driver ;
	
	@FindBy(xpath="(//a[@data-field-qty=\"qty\"])[2]")
	
	private WebElement Qtyplus ; 
	
	public Cart_automationpractice(WebDriver driver2) {
		
		this.driver = driver2;
		
		PageFactory.initElements(driver, this);
	}

	public WebElement getQtyplus() {
		return Qtyplus;
	}

	public WebElement getQtyminus() {
		return Qtyminus;
	}

	public WebElement getSize() {
		return size;
	}

	public WebElement getCart() {
		return cart;
	}

	@FindBy(xpath="(//a[@data-field-qty=\"qty\"])[1]")
	
	private WebElement Qtyminus ;
	
	@FindBy(id="group_1")
	
	private WebElement size ;
	
	@FindBy(xpath="(//button[@type=\"submit\"])[3]")
	
	private WebElement cart;

}
