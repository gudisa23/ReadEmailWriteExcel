package com.AutomationBPSReport;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.mail.Authenticator;
import javax.mail.BodyPart;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.NoSuchProviderException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.internet.MimeMultipart;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GetPOPMails {

	private static final String MAIL_POP_HOST = "pop.gmail.com";
	private static final String MAIL_STORE_TYPE = "pop3";
	private static final String POP_USER = "bpstestingco@gmail.com";
	private static final String POP_PASSWORD = "1234@bps";
	private static final String POP_PORT = "995";
	private final String fileLocation="src\\main\\resources\\BPS_Report_Template.xlsx";
	private final String propertiesLocation="src\\main\\resources\\application.properties";

	public Properties getProperties() {
		Properties properties = new Properties();
		properties.put("mail.pop3.host", MAIL_POP_HOST);
		properties.put("mail.pop3.port", POP_PORT);
		properties.put("mail.pop3.starttls.enable", "true");
		properties.put("mail.pop3.socketFactory.class", "javax.net.ssl.SSLSocketFactory");
		return properties;
	}
	
	public void getMails(String user, String password) {
		try {
			Session emailSession = Session.getDefaultInstance(getProperties(), new Authenticator() {
				@Override
				protected PasswordAuthentication getPasswordAuthentication() {
					return new PasswordAuthentication(POP_USER, POP_PASSWORD);
				}
			});
			Store store = emailSession.getStore(MAIL_STORE_TYPE);

			store.connect(MAIL_POP_HOST, user, password);

			Folder emailFolder = store.getFolder("INBOX");
			emailFolder.open(Folder.READ_ONLY);

			Message[] messages = emailFolder.getMessages();
			System.out.println("messages.length---" + messages.length);

			//Creating a XSSFWorkbook object for file "BPS_Report_Template.xlsx"
			FileInputStream file = new FileInputStream(new File(fileLocation));
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			Properties propertiesOfKey = new Properties();
			InputStream input = null;
			input = new FileInputStream(propertiesLocation);
			propertiesOfKey.load(input);
			
            HashMap<String,String> subjectMap=new HashMap<String,String>();
            
            //set sysout output to a file name out.txt
            PrintStream fileOut = new PrintStream("./out.txt");
            System.setOut(fileOut);
            
			for (int i = 0, n = messages.length; i < n; i++) {
				Message message = messages[i];
				System.out.println("---------------------------------");
				System.out.println("Email Number " + (i + 1));
				System.out.println("Subject: " + message.getSubject());
				subjectMap.put(message.getSubject(), "");
				System.out.println("From: " + message.getFrom()[0]);
				System.out.println("Body: " + getTextFromMessage(message, message.getSubject(), workbook, propertiesOfKey));
			}
			if (propertiesOfKey != null) {
				Set<String> keys = propertiesOfKey.stringPropertyNames();
				for (String key : keys) {
					if (!subjectMap.containsKey(key)) {
						System.out.println("*********Subject  = " + key + " ******** does not contain any mail");
					}
				}
			}
		
			FileOutputStream outputStream = new FileOutputStream(fileLocation);
			workbook.write(outputStream);
			workbook.close();
			outputStream.close();
			emailFolder.close(false);
			store.close();

		} catch (NoSuchProviderException e) {
			e.printStackTrace();
		} catch (MessagingException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private String getTextFromMessage(Message message, String subjectName, XSSFWorkbook workbook,Properties propertiesOfKey) throws MessagingException, IOException {
		String result = "";
		if (message.isMimeType("text/plain")) {
			result = message.getContent().toString();
		} else if (message.isMimeType("multipart/*")) {
			MimeMultipart mimeMultipart = (MimeMultipart) message.getContent();
			result = getTextFromMimeMultipart(mimeMultipart, subjectName, workbook, propertiesOfKey);
		}
		return result;
	}

	private String getTextFromMimeMultipart(MimeMultipart mimeMultipart, String subjectName, XSSFWorkbook workbook,Properties propertiesOfKey) throws MessagingException, IOException {
		String result = "";
		XSSFSheet sheet=null,sheetResponseCode=null;
		int count = mimeMultipart.getCount();
		if(workbook!=null) {
			sheet = workbook.getSheetAt(1);
			sheetResponseCode = workbook.getSheetAt(2);
		}

		for (int i = 0; i < count; i++) {
			BodyPart bodyPart = mimeMultipart.getBodyPart(i);
			if (bodyPart.isMimeType("text/plain")) {
				result = result + "\n" + bodyPart.getContent();
				Pattern regex = Pattern.compile("(\\d+(?:\\.\\d+)?)");
				Matcher matcher = regex.matcher(result);
				String[] indexArray = null;
				if (subjectName !=null && propertiesOfKey.getProperty(subjectName) != null) {
					indexArray = propertiesOfKey.getProperty(subjectName).split(",");
				}
				if (subjectName !=null && subjectName.startsWith("Install") || subjectName.startsWith("Disconnects")|| subjectName.startsWith("Transfers") || subjectName.startsWith("Swaps")) {
					List<Integer> listField = new ArrayList<Integer>();
					double percentageValue = 0;
					while (matcher.find() && indexArray != null && indexArray.length > 1) {

						String temp = matcher.group();
						if (temp !=null && !temp.contains(".")) {
							listField.add(Integer.parseInt(temp));
						} else {
							double percentage = Double.parseDouble(temp);
							DecimalFormat df = new DecimalFormat("#.##");

							df.setRoundingMode(RoundingMode.FLOOR);

							percentageValue = new Double(df.format(percentage));
						}

					}
					Collections.sort(listField);
					if (listField != null && !listField.isEmpty() && listField.size() == 2) {
						Cell cell2Updatetotal=null,cell2UpdatePercentage=null,cell2Updatesucess=null;
                         if(sheet!=null && sheet.getRow(Integer.parseInt(indexArray[0])) !=null)
                         {
                        	 if(sheet.getRow(Integer.parseInt(indexArray[0])).getCell(Integer.parseInt(indexArray[1]))!=null )
                        	 {
                        			cell2Updatetotal = sheet.getRow(Integer.parseInt(indexArray[0])).getCell(Integer.parseInt(indexArray[1]));
                                	if(listField!=null)
                                	{
                                		cell2Updatetotal.setCellValue(listField.get(1));
            							System.out.println("****Total Value Enter For " + subjectName + "with value" + listField.get(1)+ "Row Index Value" + indexArray[0] + "And Column Index" + indexArray[1]);
                                	}
                        	 }
                        	 if(sheet.getRow(Integer.parseInt(indexArray[0])).getCell(Integer.parseInt(indexArray[3]))!=null)
                        	 {
                        		 cell2UpdatePercentage = sheet.getRow(Integer.parseInt(indexArray[0])).getCell(Integer.parseInt(indexArray[3]));
                        		 cell2UpdatePercentage.setCellValue(percentageValue);
     							 System.out.println("****Percentage Value Enter For " + subjectName + "with value" + percentageValue+ "Row Index Value" + indexArray[0] + "And Column Index" + indexArray[3]);
                        	 }
                        	 if(sheet.getRow(Integer.parseInt(indexArray[0])).getCell(Integer.parseInt(indexArray[2]))!=null)
                        	 {
                        		 cell2Updatesucess = sheet.getRow(Integer.parseInt(indexArray[0])).getCell(Integer.parseInt(indexArray[2]));
                        		 cell2Updatesucess.setCellValue(listField.get(0));
     							 System.out.println("****Sucess Value Value Enter For " + subjectName + "with value" + listField.get(0)+ "Row Index Value" + indexArray[0] + "And Column Index" + indexArray[2]);
                        	 }
                        
                         }

					}
				} else {
					ArrayList<String> list = new ArrayList<>();
					while (matcher.find() && indexArray != null && indexArray.length > 1) {
						String value = matcher.group();
						if (value!=null && !value.contains(".")) {
							list.add(value);
						}
					}
					HashMap<String, String> map = new HashMap<String, String>();
					String[] array = (String[]) list.toArray(new String[list.size()]);
					for (int j = 0; j < array.length; j = j + 2) {
						map.put(array[j + 1], array[j]);
					}
					int countIndex = 2;
					for (Map.Entry<String, String> entry : map.entrySet()) {

						String key = entry.getKey();
						String value = entry.getValue();
						if (key !=null && !key.equals("000") && !key.equals("0000")) {
							try {
								Cell cell2Update = sheetResponseCode.getRow(Integer.parseInt(indexArray[countIndex])).getCell(Integer.parseInt(indexArray[1]));
								if(value!=null && cell2Update!=null) {
									try {
										cell2Update.setCellValue(Integer.parseInt(value));
									}catch(Exception e) {
									System.out.println("********** Parse type Exception for "+value +"For subject ==="+subjectName);
									}
									
								}
							

								Cell cell1Update = sheetResponseCode.getRow(Integer.parseInt(indexArray[countIndex])).getCell(Integer.parseInt(indexArray[0]));
								if (subjectName.startsWith("Voice-Residential") && key !=null && cell1Update!=null) {
									
									cell1Update.setCellValue("E" + key);
								} else if(key !=null){
									try {
										cell1Update.setCellValue(Integer.parseInt(key));
									} catch (Exception e) {
										System.out.println("********** Parse type Exception for " + key+ "For subject ===" + subjectName);
									}
									
								}
                                if(value!=null || key!=null)
                                {
                                	System.out.println("Subject of Mail--" + subjectName + "Cell with Value=="+ Integer.parseInt(value) + " Row Index--" + indexArray[countIndex]+ " Column Index Value--" + indexArray[1]);
    								System.out.println("Key Column Index value==" + Integer.parseInt(indexArray[0])+ "Key Column Value" + key);
                                }
								

								countIndex++;
							} catch (Exception e) {
								e.printStackTrace();
							}

						}

					}
				}
				break;
			} else if (bodyPart.isMimeType("text/html")) {
				String html = (String) bodyPart.getContent();
				result = result + "\n" + org.jsoup.Jsoup.parse(html).text();
			} else if (bodyPart.getContent() instanceof MimeMultipart) {
				result = result + getTextFromMimeMultipart((MimeMultipart) bodyPart.getContent(), subjectName, workbook,
						propertiesOfKey);
			}
		}
		return result;
	}

	public static void main(String[] args) {
		GetPOPMails getPOPMails = new GetPOPMails();
		getPOPMails.getMails(POP_USER, POP_PASSWORD);

	}

}
