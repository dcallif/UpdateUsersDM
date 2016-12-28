package com.digabit.admin;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.rmi.RemoteException;
import java.util.Iterator;

import org.apache.axis2.AxisFault;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.digabit.webservices.client.reference.autogen.impl.AdminService20Stub;
import com.digabit.webservices.client.reference.autogen.impl.AdminService20Stub.UpdateUserRequestDto;
import com.digabit.webservices.client.reference.autogen.impl.AdminService20Stub.UpdateUserResponseDto;
import com.digabit.webservices.client.reference.autogen.impl.AdminService20Stub.UpdateUsers;
import com.digabit.webservices.client.reference.autogen.impl.AdminService20Stub.UpdateUsersResponse;
import com.digabit.webservices.client.reference.autogen.impl.AdminService20Stub.UserDto;

/**
 * 
 * @author dcallif
 */
public class UpdateUsersFromXlsx 
{
	protected static AdminService20Stub service;

	/**
	 * Column 1 must be userName
	 * Column 2 must be organization name
	 * Column 4 must be tek
	 * Column 5 must be password to update
	 * 
	 * @param spreadsheet
	 * @param sheetNum
	 * @param uri
	 */
	private void doIt(String spreadsheet, int sheetNum, String uri)
	{
		File updateUserSpreadsheet = new File( spreadsheet );
		FileInputStream updateUserS;
		try 
		{
			updateUserS = new FileInputStream( updateUserSpreadsheet );
			
			//Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook( updateUserS );
			
			//Get second sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt( sheetNum );
			
			//Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();
			while ( rowIterator.hasNext() ) 
			{
				Row row = rowIterator.next();
				
				//ignore first row. don't care about header
				if( row.getRowNum() >= 1 )
				{
					String userName = "";
					String organizationName = "";
					String encryptedKey = "";
					String password = "";
					
					//For each row, iterate through all the columns
					Iterator<Cell> cellIterator = row.cellIterator();

					//gets only the Page data I need for Media XML
					while( cellIterator.hasNext() ) 
					{
						Cell cell = cellIterator.next();

						//userName
						if( cell != null && cell.getColumnIndex() == 0 )
						{
							userName = cell.getStringCellValue();
						}
						
						//organization name
						if( cell != null && cell.getColumnIndex() == 1 )
						{
							organizationName = cell.getStringCellValue();
						}

						//encryptedKey
						if( cell != null && cell.getColumnIndex() == 3 )
						{
							encryptedKey = cell.getStringCellValue();
						}
						
						//password
						if( cell != null && cell.getColumnIndex() == 4 )
						{
							password = cell.getStringCellValue();
						}
					}
					
					service = new AdminService20Stub( uri );

					System.out.println( "Updating user: " + userName );
					
					//construct the request
					UpdateUserRequestDto requestDto1 = new UpdateUserRequestDto();
					requestDto1.setTenantEncryptedKey( encryptedKey );
					//requestDto1.setOrganizationName( PropertiesUtil.getPropertyAsString( PropertyNames.ORGANIZATION1 ) );
					UserDto userDto = new UserDto();
					userDto.setUserName( userName );
					userDto.setPassword( password );

					requestDto1.setUser( userDto );
					requestDto1.setOrganizationName( organizationName );

					//construct the array of requestDtos. for this example, we have just 1 dto. should add multiple for batching requests
					UpdateUserRequestDto[] requestDtos = new UpdateUserRequestDto[1];
					requestDtos[0] = requestDto1;

					//construct and populate the wrapper around the dto
					UpdateUsers request = new UpdateUsers();
					request.setRequests( requestDtos );

					//invoke the web service
					UpdateUsersResponse response = new UpdateUsersResponse();
					try 
					{
						response = service.updateUsers( request );
					} 
					catch (RemoteException e) 
					{
						e.printStackTrace();
					}

					//do something with the response
					UpdateUserResponseDto[] responseDtos = response.get_return();
					for( UpdateUserResponseDto responseDto : responseDtos )
					{
						System.out.println( "updateUsersSuccess: responseCode = " + responseDto.getResponseCode() );      
						System.out.println( "updateUsersSuccess: responseMessage = " + responseDto.getResponseMessage() );
					}
				}
			}
		} 
		catch (IOException e1) 
		{
			e1.printStackTrace();
		}
	}

	public static void main(String[] args) throws AxisFault 
	{
		String spreadsheet = "/Users/dcallif/Desktop/UpdateUserSpreadsheets/UpdateUsers.xlsx";
		int index = 0;
		String uri = "https://documoto.digabit.com/dws/services/AdminService20.AdminService20HttpSoap12Endpoint/";
		//check how many arguments were passed in
	    if( args.length < 3 )
	    {
	        System.out.println( "Usage: java -jar jarfile.jar 'spreadsheet' sheetNum 'uri'" );
	        System.exit( 0 );
	    }
	    else
	    {
	    	UpdateUsersFromXlsx test = new UpdateUsersFromXlsx();
			test.doIt( args[0], Integer.parseInt( args[1] ), args[2] );
	    }
	}
}
