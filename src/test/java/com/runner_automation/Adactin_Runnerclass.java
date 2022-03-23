 package com.runner_automation;

import java.io.IOException;
import java.util.concurrent.TimeUnit;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import com.adactin_pom.Book_hotels;
import com.adactin_pom.Login_page;
import com.adactin_pom.Logout_page;
import com.adactin_pom.Search_hotel;
import com.adactin_pom.Select_hotel;
import com.baseautomation.Base_class;
import com.page_object_manager.Adactin_page_object_manager;

public class Adactin_Runnerclass extends Base_class {
	public static WebDriver driver = browser_configuration("chrome");

	public static Adactin_page_object_manager pom = new Adactin_page_object_manager(driver);

	public static void main(String[] args) throws IOException {

		geturl("https://adactin.com/HotelApp/index.php");

		implicitwait(10, TimeUnit.SECONDS);

		String username = Particular_Cell_Data(
				"C:\\Users\\ELCOT\\eclipse-workspace\\March22_PB\\ADACTIN_MANUAL_TEST_CASE_DEVELOPMENT.xlsx", 2, 5);
		String password = Particular_Cell_Data(
				"C:\\Users\\ELCOT\\eclipse-workspace\\March22_PB\\ADACTIN_MANUAL_TEST_CASE_DEVELOPMENT.xlsx", 3, 5);
		String datein = Particular_Cell_Data(
				"C:\\Users\\ELCOT\\eclipse-workspace\\March22_PB\\ADACTIN_MANUAL_TEST_CASE_DEVELOPMENT.xlsx", 10, 5);
		String dateout = Particular_Cell_Data(
				"C:\\Users\\ELCOT\\eclipse-workspace\\March22_PB\\ADACTIN_MANUAL_TEST_CASE_DEVELOPMENT.xlsx", 11, 5);
		String firstname = Particular_Cell_Data(
				"C:\\Users\\ELCOT\\eclipse-workspace\\March22_PB\\ADACTIN_MANUAL_TEST_CASE_DEVELOPMENT.xlsx", 19, 5);
		String lastname = Particular_Cell_Data(
				"C:\\Users\\ELCOT\\eclipse-workspace\\March22_PB\\ADACTIN_MANUAL_TEST_CASE_DEVELOPMENT.xlsx", 20, 5);
		String billaddress = Particular_Cell_Data(
				"C:\\Users\\ELCOT\\eclipse-workspace\\March22_PB\\ADACTIN_MANUAL_TEST_CASE_DEVELOPMENT.xlsx", 21, 5);
		String ccnum = Particular_Cell_Data(
				"C:\\Users\\ELCOT\\eclipse-workspace\\March22_PB\\ADACTIN_MANUAL_TEST_CASE_DEVELOPMENT.xlsx", 22, 5);
		String cvv = Particular_Cell_Data(
				"C:\\Users\\ELCOT\\eclipse-workspace\\March22_PB\\ADACTIN_MANUAL_TEST_CASE_DEVELOPMENT.xlsx", 25, 5);

		inputValueelement(pom.Get_Instance_login().getUsername(), username);
		inputValueelement(pom.Get_Instance_login().getPassword(), password);
		clickOnelement(pom.Get_Instance_login().getLogin());

		selectdropdownvalue(pom.Get_Instance_search().getLocation(), "index", "3");
		selectdropdownvalue(pom.Get_Instance_search().getHotels(), "index", "3");
		selectdropdownvalue(pom.Get_Instance_search().getRoomtype(), "value", "Super Deluxe");
		selectdropdownvalue(pom.Get_Instance_search().getNoofrooms(), "index", "4");
//		pom.Get_Instance_search().getIndate().clear();
//		inputValueelement(pom.Get_Instance_search().getIndate(), datein);
//		pom.Get_Instance_search().getOutdate().clear();
//		inputValueelement(pom.Get_Instance_search().getOutdate(), dateout);
		selectdropdownvalue(pom.Get_Instance_search().getAdults(), "value", "3");
		selectdropdownvalue(pom.Get_Instance_search().getChild(), "value", "4");
		clickOnelement(pom.Get_Instance_search().getSearch());

		clickOnelement(pom.Get_Instance_select().getRadiobutton());
		clickOnelement(pom.Get_Instance_select().getContinue());

		inputValueelement(pom.Get_Instance_Booking().getFirstname(), firstname);
		inputValueelement(pom.Get_Instance_Booking().getLastname(), lastname);
		inputValueelement(pom.Get_Instance_Booking().getBilladdress(), billaddress);
		inputValueelement(pom.Get_Instance_Booking().getCcnum(), ccnum);
		selectdropdownvalue(pom.Get_Instance_Booking().getCctype(), "value", "MAST");
		selectdropdownvalue(pom.Get_Instance_Booking().getCcexpmonth(), "index", "10");
		selectdropdownvalue(pom.Get_Instance_Booking().getCcexpyeaar(), "index", "12");
		inputValueelement(pom.Get_Instance_Booking().getCccvv(), cvv);

		clickOnelement(pom.Get_Instance_Booking().getBooknow());

		clickOnelement(pom.Get_Instance_Logout().getLogout());

		close();
	}
}