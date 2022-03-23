package com.page_object_manager;

import org.openqa.selenium.WebDriver;

import com.automationpractice_pom.Account_page_automationpractice;
import com.automationpractice_pom.Cart_automationpractice;
import com.automationpractice_pom.Home_page_automationpractice;
import com.automationpractice_pom.Login_page_automationpractice;
import com.automationpractice_pom.Proceed_automationpractice;
import com.automationpractice_pom.Signout_automationpractice;
import com.automationpractice_pom.Tshirt_selection;

public class Automation_Page_object_manager {

	public WebDriver driver;

	private Home_page_automationpractice hp;
	private Login_page_automationpractice login;
	private Account_page_automationpractice tshirt;
	private Tshirt_selection model;
	private Cart_automationpractice cart;
	private Proceed_automationpractice proceed;
	private Signout_automationpractice signout;

	public Automation_Page_object_manager(WebDriver driver2) {

		this.driver = driver2;
	}  

	public Home_page_automationpractice get_Instance_Hp() {

		if (hp == null) {

			hp = new Home_page_automationpractice(driver);
		}

		return hp;

	}

	public Login_page_automationpractice get_Instance_Login() {

		if (login == null) {

			login = new Login_page_automationpractice(driver);

		}

		return login;
	}

	public Account_page_automationpractice get_Instance_Tshirt() {

		if (tshirt == null) {

			tshirt = new Account_page_automationpractice(driver);

		}

		return tshirt;
	}

	public Tshirt_selection get_Instance_model() {

		if (model == null) {

			model = new Tshirt_selection(driver);

		}

		return model;
	}
	
	public Cart_automationpractice get_Instance_Cart() {
		
		if (cart == null) {
			
			cart = new Cart_automationpractice(driver);
			
		}
		
		return cart;
	}

	public Proceed_automationpractice get_Instance_proceed() {

		if (proceed == null) {

			proceed = new Proceed_automationpractice(driver);

		}

		return proceed;
	}

	public Signout_automationpractice get_Instance_signout() {

		if (signout == null) {

			signout = new Signout_automationpractice(driver);

		}

		return signout;
	}
}
