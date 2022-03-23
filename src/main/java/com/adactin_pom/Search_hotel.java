package com.adactin_pom;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class Search_hotel {
	
	public static WebDriver driver;
	
	@FindBy(id="location")
	
	private WebElement location  ;
	
	public Search_hotel(WebDriver driver2) {
	
		this.driver = driver2;
		
		PageFactory.initElements(driver, this);
	
	}

	public WebElement getLocation() {
		return location;
	}

	public WebElement getHotels() {
		return hotels;
	}

	public WebElement getRoomtype() {
		return roomtype;
	}

	public WebElement getNoofrooms() {
		return noofrooms;
	}

	public WebElement getIndate() {
		return indate;
	}

	public WebElement getOutdate() {
		return outdate;
	}

	public WebElement getAdults() {
		return adults;
	}

	public WebElement getChild() {
		return child;
	}
	public WebElement getSearch() {
		return search;
	}


	@FindBy(id="hotels")
	
	private WebElement hotels;
	
	@FindBy(id="room_type")
	
	private WebElement roomtype;
	
	@FindBy(id="room_nos")
	
	private WebElement noofrooms;
	
	@FindBy(id="datepick_in")
	
	private WebElement indate;
	
	@FindBy(id="datepick_out")
	
	private WebElement outdate;
	
	@FindBy(id="adult_room")
	
	private WebElement adults;
	
	@FindBy(id="child_room")
	
	private WebElement child;
	
	@FindBy(id="Submit")
	
	private WebElement search;

	
	
}
