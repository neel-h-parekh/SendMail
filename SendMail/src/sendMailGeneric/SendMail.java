package sendMailGeneric;

import java.util.Properties;

import javax.mail.Session;

//import javax.activation.CommandMap;
import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
//import javax.activation.MailcapCommandMap;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

//import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.mail.*;
//import javax.mail.internet.*;
//import javax.activation.*;
//import java.util.*;
//import java.util.Map.Entry;
//import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
//import java.io.FileReader;
//import java.io.IOException;
//import java.io.InputStream;
//import java.io.InputStreamReader;

public class SendMail 
{
	String EmailConfigFIleName;
	File EmailConfigFile;
	FileInputStream file;
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	
	String Host;
	String EmailFlag;
	String EmailFrom;
	String EmailTo;
	String EmailCc;
	String EmailSubject;
	Address[] EmailToId;
	Address[] EmailCcId;
	BodyPart MessageBodyPart;
	Multipart MessageMultiPart;
	
	Properties Property;
	Session session;
	MimeMessage message;
	
	String ConfigParsingError;
	
	public void ParseEmailConfigData(String FileName, String SheetName)
	{
		try
		{
			EmailConfigFIleName=FileName;
			file = new FileInputStream(new File(EmailConfigFIleName));
			workbook = new XSSFWorkbook(file);
			sheet = workbook.getSheet(SheetName);
			
			for (Row row : sheet)
			{
				//Host
				if (row.getRowNum() == 0 && row.getCell(1) != null)
				{
					Host = row.getCell(1).toString();
					Property = System.getProperties();
					Property.setProperty("mail.smtp.host", Host);
					session = Session.getDefaultInstance(Property);
					message = new MimeMessage(session);
//					System.out.println(Host);
				}
				
				//EmailFlag
				if (row.getRowNum() == 1 && row.getCell(1) != null)
				{
					EmailFlag = row.getCell(1).toString(); 
//					System.out.println(EmailFlag);
				}
				
				//EmailFrom
				if (row.getRowNum() == 2 && row.getCell(1) != null)
				{
					EmailFrom = row.getCell(1).toString();
					if (EmailFrom != null && !EmailFrom.isEmpty())
					{
						message.setFrom(new InternetAddress(EmailFrom));
					}
//					System.out.println("Email From "+EmailFrom);
				}
				
				//EmailTo
				if (row.getRowNum() == 3 && row.getCell(1) != null)
				{
					EmailTo = row.getCell(1).toString();
					if (EmailTo != null && !EmailTo.isEmpty())
					{
						message.addRecipients(Message.RecipientType.TO,InternetAddress.parse(EmailTo));
						EmailToId = message.getRecipients(Message.RecipientType.TO);
						message.setRecipients(Message.RecipientType.TO, EmailToId);
					}
					else
					{
						System.out.println("Email To is null");
					}
//					System.out.println("Email To "+EmailTo);
				}
				
				//EmailCc
				if (row.getRowNum() == 4 && row.getCell(1) != null)
				{
					EmailCc = row.getCell(1).toString();
					if (EmailCc != null && !EmailCc.isEmpty())
					{
						message.addRecipients(Message.RecipientType.CC,InternetAddress.parse(EmailCc));
						EmailCcId = message.getRecipients(Message.RecipientType.CC);
						message.setRecipients(Message.RecipientType.CC, EmailCcId);
					}
//					System.out.println("Email Cc "+EmailCc);
				}
				
				//EmailSubject
				if (row.getRowNum() == 5 && row.getCell(1) != null)
				{
					EmailSubject = row.getCell(1).toString();
					if (EmailSubject != null && !EmailSubject.isEmpty())
					{
						message.setSubject(EmailSubject);
					}
//					System.out.println(EmailSubject);
				}
			}
			
		} catch (Exception E) {
			ConfigParsingError="True";
			E.printStackTrace();
//			System.exit(0);
		}
	}
	
	public boolean MandatoryFieldsCheck()
	{
		if (Host == null || EmailFrom == null || EmailTo == null)
		{
			return false;
		}
		else
		{
			return true;
		}
	}
	
	public void ShowEmailConfigDetails()
	{
		//Host
		if (Host != null)
		{
			System.out.println("Host		:"+Host);			
		}
		else
		{
			System.out.println("Host Is Empty");
		}

		//EmailFlag
		if (EmailFlag != null)
		{
			System.out.println("\nEmail Flag	:"+EmailFlag);			
		}
		else
		{
			System.out.println("\nEmail Flag Is Empty");
		}
		
		//EmailFrom
		if (EmailFrom != null)
		{
			System.out.println("\nEmail From	:"+EmailFrom);
		}
		else
		{
			System.out.println("\nEmail From Is Empty");
		}

		//EmailToId
		if (EmailToId != null)
		{
			System.out.println("\nEmail To -");
			for (int i = 0; i < EmailToId.length; i++)
			{
				System.out.println(EmailToId[i].toString());					
			}
		}
		else
		{
			System.out.println("\nEmail To Is Empty");
		}
		
		//EmailCcId
		if (EmailCcId != null)
		{
			System.out.println("\nEmail Cc -");
			for (int i = 0; i < EmailCcId.length; i++)
			{
				System.out.println(EmailCcId[i].toString());					
			}
		}
		else
		{
			System.out.println("\nEmail Cc Is Empty");
		}
		
		//EmailSubject
		if (EmailSubject != null)
		{
			System.out.println("\nEmail Subject	:"+EmailSubject);			
		}
		else
		{
			System.out.println("\nEmail Subject Is Empty");
		}
	}
	
	public void SubmitEmail(String BodyText)
	{
		try
		{
			message.setText(BodyText);
			System.out.println(BodyText);
	        Transport.send(message);
		} catch (MessagingException e) {
			e.printStackTrace();
		}
	}
	
	public void SubmitEmail(String BodyText, String AttachFileName, String AttachFileNamePath)
	{
		if (ConfigParsingError != "True" && EmailFlag != "N")
		{
			try 
			{
				BodyPart messageBodyPart = new MimeBodyPart();
				
				Multipart multipart = new MimeMultipart();
				
				messageBodyPart.setText(BodyText);
				multipart.addBodyPart(messageBodyPart);
				
				messageBodyPart = new MimeBodyPart();
				String filename = AttachFileName;
				DataSource source = new FileDataSource(AttachFileNamePath);
				messageBodyPart.setDataHandler(new DataHandler(source));
		        messageBodyPart.setFileName(filename);
		        multipart.addBodyPart(messageBodyPart);
		        message.setContent(multipart);
		        Transport.send(message);
			} catch (MessagingException e) {
				e.printStackTrace();
			}
		}
	}

//	public static void main(String[] args) 
//	{
//		SendMail SM = new SendMail();
//		SM.ParseEmailConfigData("TEST.xlsx","email");
//		SM.ShowEmailConfigDetails();
//		
//		String Test;
//		Test = "Line 1";
//		Test = Test + "\nLine 2";
//		
//		System.out.println("");
//		
//		if (SM.MandatoryFieldsCheck())
//		{
//			SM.SubmitEmail(Test);
//			SM.SubmitEmail(Test,"TestFile.jpg","C:\\workspace\\SterlingTS\\SendMail\\download.jpg");
//		}
//		else
//		{
//			System.out.println("Send Mail---------"+SM.MandatoryFieldsCheck());
//		}
//	}

}
