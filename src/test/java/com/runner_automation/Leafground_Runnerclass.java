package com.runner_automation;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.Select;

import com.baseautomation.Base_class;

public class Leafground_Runnerclass extends Base_class {

	public static WebDriver driver;

	public static void main(String[] args) throws Throwable {

		driver = browser_configuration("chrome");
		
		implicitwait(10, TimeUnit.SECONDS);

		geturl("http://leafground.com/");

//		WebElement windows = driver.findElement(By.xpath("(//a[@class=\"wp-categories-link maxheight\"])[11]"));
//
//		clickOnelement(windows);

		WebElement dragdrop = driver.findElement(By.xpath("(//a[@class=\"wp-categories-link maxheight\"])[14]"));
		clickOnelement(dragdrop);

		WebElement drag = driver.findElement(By.id("draggable"));
		WebElement drop = driver.findElement(By.id("droppable"));
		Action("drag", drag, drop);

		Screenshot("util", "C:\\Users\\ELCOT\\eclipse-workspace\\March22_PB\\Screenshots\\image2.png");

		WebElement dropdown = driver.findElement(By.xpath("(//a[@class=\"wp-categories-link maxheight\"])[5]"));
		clickOnelement(dropdown);

		WebElement dr1 = driver.findElement(By.xpath("//select[@multiple]"));
		selectdropdownvalue(dr1, "index", "2");

		navigate("back", null);

		WebElement alert = driver.findElement(By.xpath("(//a[@class=\"wp-categories-link maxheight\"])[9]"));
		clickOnelement(alert);

		WebElement alertbox = driver.findElement(By.xpath("//button[@onclick=\"normalAlert()\"]"));
		clickOnelement(alertbox);
		Alerts("ok", null);

		WebElement confirmbox = driver.findElement(By.xpath("//button[@onclick=\"confirmAlert()\"]"));
		clickOnelement(confirmbox);
		Alerts("ok", null);
		clickOnelement(confirmbox);
		Alerts("dismiss", null);

		WebElement promptbox = driver.findElement(By.xpath("//button[@onclick=\"confirmPrompt()\"]"));
		clickOnelement(promptbox);

		Alerts("text", "greens technology");
		Alerts("ok", null);

		navigate("back", null);

		WebElement frames = driver.findElement(By.xpath("(//a[@class=\"wp-categories-link maxheight\"])[10]"));
		clickOnelement(frames);

		frames(frames, "index", "0");
		WebElement frame1 = driver.findElement(By.id("Click"));
		clickOnelement(frame1);

		defaultcontent();

		frames(frames, "index", "1");
		frames(frames, "index", "0");
		WebElement frame2 = driver.findElement(By.id("Click1"));
		clickOnelement(frame2);

		close();

	}

}
