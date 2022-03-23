package com.page_object_manager;

import org.openqa.selenium.WebDriver;

import com.adactin_pom.Book_hotels;
import com.adactin_pom.Login_page;
import com.adactin_pom.Logout_page;
import com.adactin_pom.Search_hotel;
import com.adactin_pom.Select_hotel;

public class Adactin_page_object_manager {

	public WebDriver driver;

	private Login_page login;
	private Search_hotel search;
	private Select_hotel select;
	private Book_hotels book;
	private Logout_page logout;

	public Adactin_page_object_manager(WebDriver driver2) {
	
		this.driver = driver2;
	
	}

	public Login_page Get_Instance_login() {

		if (login == null) {

			login = new Login_page(driver);

		}

		return login;

	}

	public Search_hotel Get_Instance_search() {

		if (search == null) {

			search = new Search_hotel(driver);

		}
		return search;

	}
	
	public Select_hotel Get_Instance_select() {
		
		if (select == null) {
			
			select = new Select_hotel(driver);
			
		}	
		
		return select;
		
	}
	
	public Book_hotels Get_Instance_Booking() {
		
		if (book == null) {
			
			book = new Book_hotels(driver);
			
		}
		return book;
		
	}
	
	public Logout_page Get_Instance_Logout() {
		if (logout == null ) {
			
			logout = new Logout_page(driver);
			
		}
		return logout;
		
	}

}
