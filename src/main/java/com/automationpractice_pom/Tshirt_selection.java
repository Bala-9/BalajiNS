package com.automationpractice_pom;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class Tshirt_selection {

	public WebDriver driver;

	@FindBy(id = "layered_id_attribute_group_13")

	private WebElement color;

	public Tshirt_selection(WebDriver driver2) {
		this.driver = driver2;

		PageFactory.initElements(driver, this);

	}

	public WebElement getColor() {
		return color;
	}

	public WebElement getPricesort() {
		return pricesort;
	}

	public WebElement getImage() {
		return image;
	}

	@FindBy(id = "selectProductSort")

	private WebElement pricesort;

	@FindBy(xpath = "(//img[@class=\"replace-2x img-responsive\"])[2]")

	private WebElement image;

}
