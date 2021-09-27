import io.restassured.RestAssured;
import io.restassured.builder.ResponseBuilder;
import io.restassured.http.ContentType;
import io.restassured.matcher.ResponseAwareMatcher;
import io.restassured.path.json.JsonPath;
import io.restassured.response.ExtractableResponse;
import io.restassured.response.Response;
import io.restassured.response.ResponseBody;
import io.restassured.response.ValidatableResponse;
import io.restassured.response.ValidatableResponseOptions;
import io.restassured.specification.RequestSpecification;


import org.junit.Before;
//import org.junit.jupiter.api.*;
import org.testng.Assert;
import org.testng.annotations.AfterTest;
import org.testng.annotations.Test;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;

//import Utils.Excelutils;

//import org.testng.annotations.Test1;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONObject;

import static io.restassured.RestAssured.given;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.HashMap;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.testng.annotations.Test;




public class Restapi {
	
	public static String wsid= "NA";
	public static String brandid= "NA";
	public static String macroid= "NA";
	public static String questionnaireId= "NA";
	public static String surveyid= "NA";
	
@Test
	public void Restapi() throws IOException {
		//public static void Btoken(String[] args){
	
	
	String dataPath= System.getProperty("user.dir");
			dataPath=dataPath+"./data/RestAPI.xlsx";

			File excelFile = new File(dataPath);
			FileInputStream file = new FileInputStream (excelFile);
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet= workbook.getSheetAt(0);
			
			int rows=sheet.getLastRowNum();
			//int cols=sheet1.getRow(1).getLastCellNum();
            int sheetnum=1;
			for(int k=1;k<=rows;k++)
			{
			XSSFRow row=sheet.getRow(k);

			String dtURL=sheet.getRow(k).getCell(1).getStringCellValue();
			//String dtrequestbody=sheet.getRow(k).getCell(3).getStringCellValue();
			String dtparameterkey1=sheet.getRow(k).getCell(3).getStringCellValue();
			
			String dtparametervalue1 = null;
			

			if(sheet.getRow(k).getCell(4).getCellType()==CellType.STRING) 
				dtparametervalue1 = sheet.getRow(k).getCell(4).getStringCellValue();
			else if(sheet.getRow(k).getCell(4).getCellType()==CellType.NUMERIC)
				
				dtparametervalue1 = String.valueOf((int)sheet.getRow(k).getCell(4).getNumericCellValue());
			
			//double dtparametervalue1=sheet1.getRow(r).getCell(4).getNumericCellValue();
			
			String dtparameterkey2=sheet.getRow(k).getCell(5).getStringCellValue();
			//String dtparametervalue2=sheet.getRow(k).getCell(6).getStringCellValue();
			String dtparametervalue2 = null;

		if(sheet.getRow(k).getCell(6).getCellType()==CellType.STRING) 
			dtparametervalue2 = sheet.getRow(k).getCell(6).getStringCellValue();
		else if(sheet.getRow(k).getCell(6).getCellType()==CellType.NUMERIC) 
			dtparametervalue2 = String.valueOf((int)sheet.getRow(k).getCell(6).getNumericCellValue());
			//String dtUsername=sheet.getRow(k).getCell(4).getStringCellValue();
			//String dtpassword=sheet.getRow(k).getCell(5).getStringCellValue();
			String dtpath=sheet.getRow(k).getCell(2).getStringCellValue();
			String dtsheetname=sheet.getRow(k).getCell(7).getStringCellValue();

			RestAssured.baseURI=dtURL;
			RequestSpecification req=RestAssured.given();
			req.header("Accept", ContentType.JSON.getAcceptHeader());
			//req.auth().preemptive().basic(dtUsername, dtpassword);
			JSONObject requestParams = new JSONObject();
			requestParams.put(dtparameterkey1, dtparametervalue1);
			requestParams.put(dtparameterkey2, dtparametervalue2);
			//req1.queryParam(dtparameterkey1,dtparametervalue1);
			//req.queryParam(dtparameterkey2,dtparametervalue2);
		    req.body(requestParams.toJSONString());
			//req.body(dtrequestbody);
			req.contentType(ContentType.JSON);
			Response response=req.post(dtpath);
			
			System.out.println("Response code:" +response.asString());
			System.out.println("Response code:" +response.statusCode());
			String body1 = response.body().asString();
			System.out.println(body1);
			
    JsonPath js=new JsonPath(body1); //for parsing Json
    
	String token=js.getString("accessToken");
	String userid=js.getString("userId");
	
	
	
	System.out.println(token);
	System.out.println(userid);

	if(response.statusCode()==200)
	{
	  sheet.getRow(k).createCell(8).setCellType(CellType.STRING);
	  sheet.getRow(k).createCell(9).setCellType(CellType.NUMERIC);
	  sheet.getRow(k).createCell(10).setCellType(CellType.STRING);
	  sheet.getRow(k).getCell(8).setCellValue("PASS");
	  sheet.getRow(k).getCell(9).setCellValue(response.statusCode()); 
	  sheet.getRow(k).getCell(10).setCellValue(body1); 
	}
	else{
	sheet.getRow(k).createCell(8).setCellType(CellType.STRING);
	  sheet.getRow(k).createCell(9).setCellType(CellType.NUMERIC);
	  sheet.getRow(k).createCell(10).setCellType(CellType.STRING);
	  sheet.getRow(k).getCell(8).setCellValue("FAIL");
	  sheet.getRow(k).getCell(9).setCellValue(response.statusCode()); 
	  sheet.getRow(k).getCell(10).setCellValue(body1);

	}
			

	XSSFSheet sheet1= workbook.getSheet(dtsheetname);
	
	int rows1=sheet1.getLastRowNum();
	//int cols=sheet1.getRow(1).getLastCellNum();

	for(int r=1;r<=rows1;r++)
	{
	XSSFRow row1=sheet1.getRow(r);

	String dtapitype=sheet1.getRow(r).getCell(1).getStringCellValue();
	
	if(dtapitype.equalsIgnoreCase("GET")){
	String dtapiURL=sheet1.getRow(r).getCell(2).getStringCellValue();
	//String URL = dtapiURL.replace("{userid}",userid); // Replace 'h' with 's'
	//System.out.println(URL);
	String dtapipath=sheet1.getRow(r).getCell(3).getStringCellValue();
	String path2 = dtapipath.replace("{wsid}",wsid);
	String path3 = path2.replace("{brandid}",brandid);
	String dtrequestbody1=sheet.getRow(k).getCell(4).getStringCellValue();
	String dtgetting=sheet1.getRow(r).getCell(5).getStringCellValue();
	
	
			RestAssured.baseURI=dtapiURL;
			RequestSpecification req1=RestAssured.given();
			req1.header("Authorization", "Bearer " + token);
			req1.header("Accept", "*/*");
			//req1.auth().preemptive().basic(dtUsername, dtpassword);
			req1.contentType(ContentType.JSON);
			//req1.body(dtrequestbody1);
			Response response1=req1.get(path3);
			
			//System.out.println("Response:" +response1);
			System.out.println("Response code:" +response1.statusCode());
			String body = response1.body().asString();
			System.out.println(body);
			//System.out.println("Response code:" +response1.statusCode());
			
			//JsonPath js1=new JsonPath(body); //for parsing Json
			
			//String completeResponse = JsonPath.read(response1, "$");
		    
			//String roleid=js1.getString("\"name\":\"Administrator\"");
			//System.out.println(roleid);
	
			if(response1.statusCode()==200)
			{
			  
			  sheet1.getRow(r).createCell(6).setCellType(CellType.NUMERIC);
			  sheet1.getRow(r).createCell(7).setCellType(CellType.STRING);
			  //sheet1.getRow(r).createCell(17).setCellType(CellType.STRING);
			  
			 
			  sheet1.getRow(r).getCell(6).setCellValue(response1.statusCode()); 
			  sheet1.getRow(r).getCell(7).setCellValue(body); 
			 // sheet1.getRow(r).getCell(17).setCellValue("PASSED");
			}
			else{
			
			  sheet1.getRow(r).createCell(6).setCellType(CellType.NUMERIC);
			  sheet1.getRow(r).createCell(7).setCellType(CellType.STRING);
			  //sheet1.getRow(r).createCell(17).setCellType(CellType.STRING);
			  sheet1.getRow(r).getCell(6).setCellValue(response1.statusCode()); 
			  sheet1.getRow(r).getCell(7).setCellValue(body);
			 // sheet1.getRow(r).getCell(17).setCellValue("FAILED");

			}

			
	}
	
	if(dtapitype.equalsIgnoreCase("POST")){
		//private static final String wsid = null;
		//String wsid = null;
		String dtapiURL=sheet1.getRow(r).getCell(2).getStringCellValue();
		//String URL = dtapiURL.replace("{userid}",userid); // Replace 'h' with 's'
		//System.out.println(URL);
		String dtapipath=sheet1.getRow(r).getCell(3).getStringCellValue();
		String dtsidpassing=dtapipath.replace("{surveyid}",surveyid);
		String dtqidpassing=dtsidpassing.replace("{questionnaireId}",questionnaireId);
		String dtbidpassing=dtqidpassing.replace("{brandid}",brandid);
		String dtrequestbody1=sheet1.getRow(r).getCell(4).getStringCellValue();
		String email = dtrequestbody1.replace("{Email}",dtparametervalue1);
		String brand = dtrequestbody1.replace("{wsid}",wsid);
		String brandid1 = brand.replace("{brandid}",brandid);
		String q1id = brandid1.replace("{questionnaireId}",questionnaireId);
		String m1id = q1id.replace("{macroid}",macroid);
		String dtgetting=sheet1.getRow(r).getCell(5).getStringCellValue();
		
		//CharSequence wsid = null;
		//String wsid2=wsid.replace("null","NA");
		
		//String wsid1 = email.replace("{ID}",wsid);
		
		
//				RestAssured.baseURI=dtapiURL;
//				RequestSpecification req1=RestAssured.given();
//				req1.header("Authorization", "Bearer " + token);
//				//req1.auth().preemptive().basic(dtUsername, dtpassword);
//				req1.contentType(ContentType.JSON);
//				JSONObject requestParams = new JSONObject();
//				requestParams.put(dtparameterkey1, dtparametervalue1);
//				//req1.queryParam(dtparameterkey1,dtparametervalue1);
//				req1.queryParam(dtparameterkey2,dtparametervalue2);
//				req1.queryParam(dtparameterkey3,dtparametervalue3);
//				req1.queryParam(dtparameterkey4,dtparametervalue4);
//				req1.queryParam(dtparameterkey5,dtparametervalue5);
//				req1.body(requestParams.toJSONString());
//				//Response response = request.post("/register");
//				Response response1=req1.post(dtapipath);
				
				RestAssured.baseURI=dtapiURL;
				RequestSpecification req1=RestAssured.given();
				req1.header("Authorization", "Bearer " + token);
				req1.header("Accept", "*/*");
				//req1.header("Content-Type", "text/plain");
				//req1.header("Accept", ContentType.JSON.getAcceptHeader());
				//req.auth().preemptive().basic(dtUsername, dtpassword);
				req1.body(m1id);
				req1.contentType(ContentType.JSON);
				Response response1=req1.post(dtbidpassing);
				
				//System.out.println("Response:" +response1);
				System.out.println("Response code:" +response1.statusCode());
				String body = response1.body().asString();
				System.out.println(body);

				JsonPath js1=new JsonPath(body); //for parsing Json
				
				
				if(dtgetting.equalsIgnoreCase("Workspace ID")){
				wsid=js1.getString("id");
				//String userid=js.getString("userId");
			    System.out.println("Workspaceid:"+ wsid);
				}
				
				if(dtgetting.equalsIgnoreCase("Brand ID")){
					brandid=js1.getString("id");
					//String userid=js.getString("userId");
				    System.out.println("Brandid:"+brandid);
					}
				
				if(dtgetting.equalsIgnoreCase("Macro ID")){
					macroid=js1.getString("id");
					//String userid=js.getString("userId");
				    System.out.println("Macroid:"+macroid);
					}
				
				if(dtgetting.equalsIgnoreCase("questionnaireId")){
					questionnaireId=js1.getString("id");
					//String userid=js.getString("userId");
				    System.out.println("questionnaireId:"+questionnaireId);
					}
				
				if(dtgetting.equalsIgnoreCase("Survey ID")){
					surveyid=js1.getString("id");
					//String userid=js.getString("userId");
				    System.out.println("surveyid:"+surveyid);
					}
				
				
			//System.out.println("Workspace id:" +workspaceid);
				//System.out.println("brand id:" +brandid);
		
				if(response1.statusCode()==200)
				{
				  
				  sheet1.getRow(r).createCell(6).setCellType(CellType.NUMERIC);
				  sheet1.getRow(r).createCell(7).setCellType(CellType.STRING);
				  //sheet1.getRow(r).createCell(17).setCellType(CellType.STRING);
				  
				 
				  sheet1.getRow(r).getCell(6).setCellValue(response1.statusCode()); 
				  sheet1.getRow(r).getCell(7).setCellValue(body); 
				 // sheet1.getRow(r).getCell(17).setCellValue("PASSED");
				}
				else{
				
				  sheet1.getRow(r).createCell(6).setCellType(CellType.NUMERIC);
				  sheet1.getRow(r).createCell(7).setCellType(CellType.STRING);
				  //sheet1.getRow(r).createCell(17).setCellType(CellType.STRING);
				  sheet1.getRow(r).getCell(6).setCellValue(response1.statusCode()); 
				  sheet1.getRow(r).getCell(7).setCellValue(body);
				 // sheet1.getRow(r).getCell(17).setCellValue("FAILED");

				}

				
		}
	
	if(dtapitype.equalsIgnoreCase("PATCH")){
		//private static final String wsid = null;
		//String wsid = null;
		String dtapiURL=sheet1.getRow(r).getCell(2).getStringCellValue();
		//String URL = dtapiURL.replace("{userid}",userid); // Replace 'h' with 's'
		//System.out.println(URL);
		String dtapipath=sheet1.getRow(r).getCell(3).getStringCellValue();
		String bid = dtapipath.replace("{brandid}",brandid);
		String qid = bid.replace("{questionnaireId}",questionnaireId);
		String dtrequestbody1=sheet1.getRow(r).getCell(4).getStringCellValue();
		String wk1 = dtrequestbody1.replace("{wsid}",wsid);
		String wk2 = wk1.replace("{macroid}",macroid);
		String wk3 = wk2.replace("{brandid}",brandid);
		//String email = dtrequestbody1.replace("{Email}",dtparametervalue1);
		String dtgetting=sheet1.getRow(r).getCell(5).getStringCellValue();
		
		//CharSequence wsid = null;
		//String wsid2=wsid.replace("null","NA");
		
		//String wsid1 = email.replace("{ID}",wsid);
		
		
//				RestAssured.baseURI=dtapiURL;
//				RequestSpecification req1=RestAssured.given();
//				req1.header("Authorization", "Bearer " + token);
//				//req1.auth().preemptive().basic(dtUsername, dtpassword);
//				req1.contentType(ContentType.JSON);
//				JSONObject requestParams = new JSONObject();
//				requestParams.put(dtparameterkey1, dtparametervalue1);
//				//req1.queryParam(dtparameterkey1,dtparametervalue1);
//				req1.queryParam(dtparameterkey2,dtparametervalue2);
//				req1.queryParam(dtparameterkey3,dtparametervalue3);
//				req1.queryParam(dtparameterkey4,dtparametervalue4);
//				req1.queryParam(dtparameterkey5,dtparametervalue5);
//				req1.body(requestParams.toJSONString());
//				//Response response = request.post("/register");
//				Response response1=req1.post(dtapipath);
				
				RestAssured.baseURI=dtapiURL;
				RequestSpecification req1=RestAssured.given();
				req1.header("Authorization", "Bearer " + token);
				req1.header("Accept", "*/*");
				//req1.header("Content-Type", "text/plain");
				//req1.header("Accept", ContentType.JSON.getAcceptHeader());
				//req.auth().preemptive().basic(dtUsername, dtpassword);
				req1.body(wk3);
				req1.contentType(ContentType.JSON);
				Response response1=req1.patch(qid);
				
				//System.out.println("Response:" +response1);
				System.out.println("Response code:" +response1.statusCode());
				String body = response1.body().asString();
				System.out.println(body);

				JsonPath js1=new JsonPath(body); //for parsing Json
			    
				//wsid=js1.getString("id");
				//String userid=js.getString("userId");
				//System.out.println(wsid);
		
				if(response1.statusCode()==200)
				{
				  
				  sheet1.getRow(r).createCell(6).setCellType(CellType.NUMERIC);
				  sheet1.getRow(r).createCell(7).setCellType(CellType.STRING);
				  //sheet1.getRow(r).createCell(17).setCellType(CellType.STRING);
				  
				 
				  sheet1.getRow(r).getCell(6).setCellValue(response1.statusCode()); 
				  sheet1.getRow(r).getCell(7).setCellValue(body); 
				 // sheet1.getRow(r).getCell(17).setCellValue("PASSED");
				}
				else{
				
				  sheet1.getRow(r).createCell(6).setCellType(CellType.NUMERIC);
				  sheet1.getRow(r).createCell(7).setCellType(CellType.STRING);
				  //sheet1.getRow(r).createCell(17).setCellType(CellType.STRING);
				  sheet1.getRow(r).getCell(6).setCellValue(response1.statusCode()); 
				  sheet1.getRow(r).getCell(7).setCellValue(body);
				 // sheet1.getRow(r).getCell(17).setCellValue("FAILED");

				}

				
		}
	
	
	}
	sheetnum=sheetnum+1;
			}
			
			FileOutputStream outFile1 = new FileOutputStream(excelFile);
			workbook.write(outFile1);
			workbook.close();
			outFile1.close();
			file.close();

		
	 
}

		//return true;	





}



	

    
	
