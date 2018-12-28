package com.AutomationBPSReport;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.env.Environment;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class TestCon {
	private final String fileLocation="src\\main\\resources\\BPS_Report_Template.xlsx";
	
	@Autowired
	Environment environment;

	@RequestMapping(value="/bps",method=RequestMethod.POST)
	public String getData(@RequestBody SearchData searchData) {
		/*XSSFSheet sheet=null,sheetResponseCode=null;
		String subjectName=searchData.getSid();
		FileInputStream file=null;
		try {
			file = new FileInputStream(new File(fileLocation));
		} catch (FileNotFoundException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		XSSFWorkbook workbook=null;
		try {
			if(file!=null) {
				workbook = new XSSFWorkbook(file);				
			}
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		if(workbook!=null) {
			sheet = workbook.getSheetAt(1);
			sheetResponseCode = workbook.getSheetAt(2);
		}
		System.out.println(environment.getProperty("Install-HSD-Residential-Kansas"));
		
		
		String[] indexArray = null;
		if (subjectName !=null && environment.getProperty(subjectName) != null) {
			indexArray = environment.getProperty(subjectName).split(",");
		}
		if (subjectName !=null && subjectName.startsWith("Install") || subjectName.startsWith("Disconnects")|| subjectName.startsWith("Transfers") || subjectName.startsWith("Swaps")) {
			List<Integer> listField = new ArrayList<Integer>();
			double percentageValue = 0;

				String temp = searchData.getResult().getSuccess();
				listField.add(Integer.parseInt(searchData.getResult().getSuccessfullInstalls()));
				listField.add(Integer.parseInt(searchData.getResult().getTotallInstalls()));

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
			//TODO
			//Need to implement the logic regarding code alert
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
		System.out.println(searchData.getSid());
		System.out.println(searchData.getResult().getSourcetype()+"SuccessfullInstalls"+searchData.getResult().getSuccessfullInstalls());*/
		return "Data";
	}

}
