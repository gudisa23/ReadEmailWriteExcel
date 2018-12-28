package com.AutomationBPSReport;

import org.springframework.beans.factory.config.PropertiesFactoryBean;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;
import org.springframework.core.io.ClassPathResource;

@SpringBootApplication
public class AutomationBpsReportApplication implements CommandLineRunner{
	
	
	@Bean
	XavientMail xavientMail() {
		return new XavientMail();
	}

	public static void main(String[] args) {
		/*XavientMail xavientMail = new XavientMail();
		xavientMail.getMails("null", "null"); */
		SpringApplication.run(AutomationBpsReportApplication.class, args);
	}
	
	@Override
	public void run(String... args) throws Exception {
		//xavientMail().getMails("", "");
	}
	
	@Bean(name = "keyProperties")
	public static PropertiesFactoryBean mapper() {
	    PropertiesFactoryBean bean = new PropertiesFactoryBean();
	    bean.setLocation(new ClassPathResource("application.properties"));
	    return bean;
	}
}
