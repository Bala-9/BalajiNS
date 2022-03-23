package com.runner_automation;
import java.util.concurrent.TimeUnit;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import com.automationpractice_pom.Account_page_automationpractice;
import com.automationpractice_pom.Cart_automationpractice;
import com.automationpractice_pom.Home_page_automationpractice;
import com.automationpractice_pom.Login_page_automationpractice;
import com.automationpractice_pom.Proceed_automationpractice;
import com.automationpractice_pom.Signout_automationpractice;
import com.automationpractice_pom.Tshirt_selection;
import com.baseautomation.Base_class;
import com.page_object_manager.Automation_Page_object_manager;

 public class Automation_practice_Runnerclass extends Base_class {

	public static WebDriver driver = browser_configuration("chrome");
	public static Automation_Page_object_manager pom = new Automation_Page_object_manager(driver);
	public static Logger log = Logger.getLogger(Automation_practice_Runnerclass.class);

	public static void main(String[] args) throws Throwable {
		
		PropertyConfigurator.configure("log4j.properties");
		

		geturl("http://automationpractice.com/index.php");
		
		log.info("after Url launch");
		
		implicitwait(10, TimeUnit.SECONDS);
		
		clickOnelement(pom.get_Instance_Hp().getSignin_btn());
		
		String username = Particular_Cell_Data("C:\\Users\\ELCOT\\eclipse-workspace\\March22_PB\\AUTOMATIONPRACTICE_MANUAL_TC_DEV.xlsx", 2, 5 );
		String password = Particular_Cell_Data("C:\\Users\\ELCOT\\eclipse-workspace\\March22_PB\\AUTOMATIONPRACTICE_MANUAL_TC_DEV.xlsx", 3, 5);
		
		log.info("before signin");
		inputValueelement(pom.get_Instance_Login().getEmail(), username);
		inputValueelement(pom.get_Instance_Login().getPassword(), password); 
		clickOnelement(pom.get_Instance_Login().getSignin_btn());
		log.info("after signin");

		clickOnelement(pom.get_Instance_Tshirt().getTshirt());
		
		log.info("before select tshirt model");
		clickOnelement(pom.get_Instance_model().getColor());
		selectdropdownvalue(pom.get_Instance_model().getPricesort(), "index", "3");
		clickOnelement(pom.get_Instance_model().getImage());
		log.info("after select tshirt model");

		log.info("before add cart");
		clickOnelement(pom.get_Instance_Cart().getQtyplus());
		clickOnelement(pom.get_Instance_Cart().getQtyplus());
		selectdropdownvalue(pom.get_Instance_Cart().getSize(), "value", "3");
		clickOnelement(pom.get_Instance_Cart().getCart());
		log.info("after add cart");
		
		log.info("before payment proceed");
		clickOnelement(pom.get_Instance_proceed().getProceed1());
		clickOnelement(pom.get_Instance_proceed().getProceed2());
		clickOnelement(pom.get_Instance_proceed().getProceed3());
		clickOnelement(pom.get_Instance_proceed().getTc());
		clickOnelement(pom.get_Instance_proceed().getProceed4());
		clickOnelement(pom.get_Instance_proceed().getPayment());
		clickOnelement(pom.get_Instance_proceed().getConfirm());
		log.info("after payment proceed");

		log.info("before signout");
		clickOnelement(pom.get_Instance_signout().getSignout());
		log.info("after signout");

		close();

	}

}
