package com.AutomationBPSReport;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.math.RoundingMode;
import java.net.URI;
import java.net.URISyntaxException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.annotation.Resource;
import javax.mail.MessagingException;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.output.TeeOutputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.BodyType;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.search.LogicalOperator;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.EmailMessageSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.FolderId;
import microsoft.exchange.webservices.data.property.complex.MessageBody;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

public class XavientMail {
	private final String fileLocation = "C:\\Users\\akumar18\\workspace\\automationbpsreport\\src\\main\\resources\\";
	 
	 @Resource(name = "keyProperties")
	 private Map<String, String> keyProperties;
	 SimpleDateFormat formatter=null;
	 Date date=null;
	 Calendar calendar=null;
	 String updateFileLocation=null;
	 
	public void getMails(String user, String password) {

		System.out.println(keyProperties.get("Swaps-Converge-Commercial-Lincoln"));
		ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
		String username = "akumar18@xavient.com";
		String Password = "denverco@0103";
		
		ExchangeCredentials credentials = new WebCredentials(username, Password);
		service.setCredentials(credentials);
		try {
			service.setUrl(new URI("https://outlook.xavient.com/ews/exchange.asmx"));
		} catch (URISyntaxException e1) {
			System.out.println("Not able to connect with URI"+e1.getMessage());
		}
		try {
			
			formatter = new SimpleDateFormat("MMMM dd, yyyy", Locale.US);
			calendar = Calendar.getInstance();
			calendar.add(Calendar.DATE, -1);
			date=calendar.getTime();
			FileUtils.copyFile(new File(fileLocation+"BPS_Report_Template.xlsx"),new File(fileLocation+"BPS-Report-"+formatter.format(date)+".xlsx"));
			
			updateFileLocation=fileLocation+"BPS-Report-"+formatter.format(date)+".xlsx";
		} catch (IOException e2) {
			// TODO Auto-generated catch block
			e2.printStackTrace();
		}
		FileInputStream file = null;
		try {
			file = new FileInputStream(new File(updateFileLocation));
		} catch (FileNotFoundException e) {
			System.out.println("BPS Report xslx file not found"+e.getMessage());
		}
		XSSFWorkbook workbook = null;
		try {
			workbook = new XSSFWorkbook(file);
		} catch (IOException e) {
			System.out.println(e.getMessage());
		}
        File f=new File("./out.txt");
        try {
            FileOutputStream fos = new FileOutputStream(f);
            Runtime.getRuntime().addShutdownHook(new Thread(() -> {
                try {
                    fos.flush();
                }
                catch (Throwable t) {
                    // Ignore
                }
            }, "Shutdown hook Thread flushing " + f));
            TeeOutputStream myOut=new TeeOutputStream(System.out, fos);
            PrintStream ps = new PrintStream(myOut, true); //true - auto-flush after println
            System.setOut(ps);
        } catch (Exception e) {
            e.printStackTrace();
        }
        
		try {

			WellKnownFolderName sd = WellKnownFolderName.Inbox;
			FolderId folderId = new FolderId(sd);
			Folder folder = Folder.bind(service, folderId);
			SearchFilter sf = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
			HashMap<String, String> subjectMap = new HashMap<String, String>();
			ItemView view = new ItemView(200);
			FindItemsResults<Item> findResults = service.findItems(folder.getId(),sf, view);
			System.out.println("Total Number of Email ="+findResults.getTotalCount());
			XSSFSheet sheet = null, sheetResponseCode = null,sheetGraph=null;
			if (workbook != null) {
				sheet = workbook.getSheetAt(1);
				sheetGraph=workbook.getSheetAt(0);
				Cell cellDate = sheet.getRow(5).getCell(2);
				
				cellDate.setCellValue(formatter.format(date));
				
				Cell cellGraph=sheetGraph.getRow(3).getCell(2);
				cellGraph.setCellValue(formatter.format(date));
				//sheetResponseCode = workbook.getSheetAt(2);
			}
			for (Item item : findResults.getItems()) {
				PropertySet psPropset = new PropertySet();
				psPropset.setRequestedBodyType(BodyType.Text);
				psPropset.setBasePropertySet(BasePropertySet.FirstClassProperties);

				EmailMessage emailMessage = EmailMessage.bind(service, item.getId(), psPropset);
				System.out.println("sub==========" + item.getSubject());
				System.out.println(emailMessage.getBody().toString());
				subjectMap.put(item.getSubject(), "");
				if(item.getSubject().equals("Install-Voice-Residential-Wisconsin-Native"))
				{
					System.out.println("Install-Voice-Residential-Wisconsin-Native");
				}
				if (keyProperties.get(item.getSubject())!=null) {
					getTextFromMimeMultipart(emailMessage.getBody().toString(), item.getSubject(), sheet,sheetResponseCode);
				}

			}
			if (keyProperties != null) {
				Iterator<Map.Entry<String, String>> iterator=keyProperties.entrySet().iterator();
				while (iterator.hasNext()) {
					Map.Entry<String,String> entry = (Map.Entry<String,String>) iterator.next();
					
					if (!subjectMap.containsKey(entry.getKey())) {
						System.out.println("*********Subject  = " + entry.getKey() + " ******** does not contain any mail");
					}
					
				}
				
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		FileOutputStream outputStream = null;
		try {
			outputStream = new FileOutputStream(new File(updateFileLocation));
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		try {
			workbook.write(outputStream);
			workbook.close();
			outputStream.close();
		} catch (IOException e1) {
			System.out.println(e1.getMessage());
		}
		sendEWSMail();
	}

	private String getTextFromMimeMultipart(String body, String subjectName, XSSFSheet sheet,XSSFSheet sheetResponseCode) throws MessagingException, IOException {
		Pattern regex = Pattern.compile("(\\d+(?:\\.\\d+)?)");
		Matcher matcher = regex.matcher(body);
		String[] indexArray = null;
		if (subjectName != null && keyProperties.get(subjectName) != null) {
			indexArray =keyProperties.get(subjectName).split(",");
		}
		if (subjectName != null && subjectName.startsWith("Install") || subjectName.startsWith("Disconnects") || subjectName.startsWith("Transfers") || subjectName.startsWith("Swaps")) {
			List<Integer> listField = new ArrayList<Integer>();
			double percentageValue = 0;
			while (matcher.find() && indexArray != null && indexArray.length > 1) {

				String temp = matcher.group();
				if (temp != null && !temp.contains(".")) {
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
				Cell cell2Updatetotal = null, cell2UpdatePercentage = null, cell2Updatesucess = null;
				if (sheet != null && sheet.getRow(Integer.parseInt(indexArray[0])) != null) {
					if (sheet.getRow(Integer.parseInt(indexArray[0])).getCell(Integer.parseInt(indexArray[1])) != null) {
						cell2Updatetotal = sheet.getRow(Integer.parseInt(indexArray[0])).getCell(Integer.parseInt(indexArray[1]));
						if (listField != null) {
							cell2Updatetotal.setCellValue(listField.get(1));
							System.out.println("****Total Value Enter For " + subjectName + "with value" + listField.get(1)+ "Row Index Value" + indexArray[0] + "And Column Index" + indexArray[1]);
						}
					}
					if (sheet.getRow(Integer.parseInt(indexArray[0])).getCell(Integer.parseInt(indexArray[3])) != null) {
						cell2UpdatePercentage = sheet.getRow(Integer.parseInt(indexArray[0])).getCell(Integer.parseInt(indexArray[3]));
						cell2UpdatePercentage.setCellValue(percentageValue);
						System.out.println("****Percentage Value Enter For " + subjectName + "with value" + percentageValue+ "Row Index Value" + indexArray[0] + "And Column Index" + indexArray[3]);
					}
					if (sheet.getRow(Integer.parseInt(indexArray[0])).getCell(Integer.parseInt(indexArray[2])) != null) {
						cell2Updatesucess = sheet.getRow(Integer.parseInt(indexArray[0])).getCell(Integer.parseInt(indexArray[2]));
						cell2Updatesucess.setCellValue(listField.get(0));
						System.out.println("****Sucess Value Value Enter For " + subjectName + "with value" + listField.get(0)+ "Row Index Value" + indexArray[0] + "And Column Index" + indexArray[2]);
					}

				}

			}
			
			if (listField != null && !listField.isEmpty() && listField.size() == 1) {
				Cell cell2Updatetotal = null, cell2UpdatePercentage = null, cell2Updatesucess = null;
				if (sheet != null && sheet.getRow(Integer.parseInt(indexArray[0])) != null) {
					if (sheet.getRow(Integer.parseInt(indexArray[0])).getCell(Integer.parseInt(indexArray[1])) != null) {
						cell2Updatetotal = sheet.getRow(Integer.parseInt(indexArray[0])).getCell(Integer.parseInt(indexArray[1]));
						if (listField != null) {
							cell2Updatetotal.setCellValue(0);
							//System.out.println("****Total Value Enter For " + subjectName + "with value" + listField.get(1)+ "Row Index Value" + indexArray[0] + "And Column Index" + indexArray[1]);
						}
					}
					if (sheet.getRow(Integer.parseInt(indexArray[0])).getCell(Integer.parseInt(indexArray[3])) != null) {
						cell2UpdatePercentage = sheet.getRow(Integer.parseInt(indexArray[0])).getCell(Integer.parseInt(indexArray[3]));
						cell2UpdatePercentage.setCellValue(0);
						//System.out.println("****Percentage Value Enter For " + subjectName + "with value" + percentageValue+ "Row Index Value" + indexArray[0] + "And Column Index" + indexArray[3]);
					}
					if (sheet.getRow(Integer.parseInt(indexArray[0])).getCell(Integer.parseInt(indexArray[2])) != null) {
						cell2Updatesucess = sheet.getRow(Integer.parseInt(indexArray[0])).getCell(Integer.parseInt(indexArray[2]));
						cell2Updatesucess.setCellValue(0);
						//System.out.println("****Sucess Value Value Enter For " + subjectName + "with value" + listField.get(0)+ "Row Index Value" + indexArray[0] + "And Column Index" + indexArray[2]);
					}

				}

			}
			
			
		} else {
			ArrayList<String> list = new ArrayList<>();
			while (matcher.find() && indexArray != null && indexArray.length > 1) {
				String value = matcher.group();
				if (value != null && !value.contains(".")) {
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
				if (key != null && !key.equals("000") && !key.equals("0000")) {
					try {
						Cell cell2Update = sheetResponseCode.getRow(Integer.parseInt(indexArray[countIndex]))
								.getCell(Integer.parseInt(indexArray[1]));
						if (value != null && cell2Update != null) {
							try {
								cell2Update.setCellValue(Integer.parseInt(value));
							} catch (Exception e) {
								System.out.println("********** Parse type Exception for " + value + "For subject ==="+ subjectName);
							}

						}

						Cell cell1Update = sheetResponseCode.getRow(Integer.parseInt(indexArray[countIndex])).getCell(Integer.parseInt(indexArray[0]));
						if (subjectName.startsWith("Voice-Residential") && key != null && cell1Update != null) {

							cell1Update.setCellValue("E" + key);
						} else if (key != null) {
							try {
								cell1Update.setCellValue(Integer.parseInt(key));
							} catch (Exception e) {
								System.out.println("********** Parse type Exception for " + key + "For subject ===" + subjectName);
							}

						}
						if (value != null || key != null) {
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
		return body;
	}
	public boolean sendEWSMail(){
	    ExchangeService service = new ExchangeService();
	    String username = "akumar18@xavient.com";
		String Password = "denverco@0103";
		String cc="aaluvala@xavient.com;sgurijala@xavient.com;srath@xavient.com";
	    EmailMessage msg = null; 
	    ExchangeCredentials credentials = null;
	    String domain = "";
	    if (domain == null || domain.equals("")) {
	        credentials = new WebCredentials(username, 
	        		Password);
	    } else {
	        credentials = new WebCredentials(username, 
	        		Password, domain);
	    }
	    service.setCredentials(credentials);
	    try {
	    	String body="Hi Team, <br> <br> Here is the BPS Report/Dashboard for " + formatter.format(date) + " <br><br>Thanks<br>Abhishek Kumar";
	    	String to="rgudisa@xavient.com";
	        service.setUrl(new URI("https://outlook.xavient.com/ews/exchange.asmx"));
	        msg = new EmailMessage(service);
	        msg.setSubject("BPS Report/Dashboard"+formatter.format(date)); 
	        msg.setBody(MessageBody.getMessageBodyFromText(body));
	        msg.getAttachments().addFileAttachment(updateFileLocation);
	        if(to == null || to.equals("")){
	        }else{
	            String[] mailTos = to.split(";");
	            for(String mailTo : mailTos){
	                if(mailTo != null && !mailTo.isEmpty())
	                msg.getToRecipients().add(mailTo);
	            }
	            if(cc != null && !cc.isEmpty()){
	                String[] mailCCs = cc.split(";");
	                for(String mailCc : mailCCs){
	                    if(mailCc != null && !mailCc.equals(""))
	                    msg.getCcRecipients().add(mailCc);
	                }
	            }
	            msg.send();
	            System.out.println("Successfully e-mail has been sent.");
	            return true;
	        }
	    } catch (Exception e) {
	    }
	    return false;
	}
}
