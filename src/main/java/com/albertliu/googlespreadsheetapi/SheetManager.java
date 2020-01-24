/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.albertliu.googlespreadsheetapi;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetRequest;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetResponse;
import com.google.api.services.sheets.v4.model.CellData;
import com.google.api.services.sheets.v4.model.CellFormat;
import com.google.api.services.sheets.v4.model.ClearValuesRequest;
import com.google.api.services.sheets.v4.model.ClearValuesResponse;
import com.google.api.services.sheets.v4.model.Color;
import com.google.api.services.sheets.v4.model.DeleteDimensionRequest;
import com.google.api.services.sheets.v4.model.DimensionRange;
import com.google.api.services.sheets.v4.model.ExtendedValue;
import com.google.api.services.sheets.v4.model.GridRange;
import com.google.api.services.sheets.v4.model.RepeatCellRequest;
import com.google.api.services.sheets.v4.model.Request;
import com.google.api.services.sheets.v4.model.TextFormat;
import com.google.api.services.sheets.v4.model.ValueRange;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedList;
import java.util.List;
import java.util.Properties;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 *
 * @author albertliu
 */
public class SheetManager extends GooglesheetCredentials{
    
    private final Sheets sheetsService; 
    private final String spreadsheetKey;
    
    public SheetManager(Properties prop, String filepath, String applicationName, String sid) throws IOException, GeneralSecurityException {
        super(prop, filepath);
        
        sheetsService = createSheetsService(applicationName);
        this.spreadsheetKey = sid;
    }
    
    public static Sheets createSheetsService(String applicationName) throws IOException, GeneralSecurityException {
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
        Credential credential = getCredentials(HTTP_TRANSPORT, 19022); 
        
        return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
        .setApplicationName(applicationName).build(); 
    }     
    
    /*******
     * Set one cell in the specified location bold. 
     * If the cell is already bold, leave it bold, otherwise, set it bold. 
     * @param sheetID
     * @param value
     * @param startingrowindex
     * @param startcolumnindex
     * @throws IOException 
     */
    public void setBold(String sheetID, String value, int startingrowindex, int startcolumnindex) throws IOException {
        setbold(sheetID, true, value, startingrowindex, startingrowindex + 1, startcolumnindex, startcolumnindex + 1);         
    }
    
    
    /*******
     * Helper method for set bold. 
     * @param sheetID
     * @param isbold
     * @param val
     * @param sri
     * @param eri
     * @param sci
     * @param eci
     * @throws IOException 
     */
    private void setbold(String sheetID, boolean isbold, String val, int sri, int eri, int sci, int eci) throws IOException
    {
        List<Request> requests = new ArrayList<>();

        requests.add 
            ( new Request()
                .setRepeatCell(new RepeatCellRequest()
                    .setCell(new CellData()
                        .setUserEnteredValue(
                            new ExtendedValue()
                                .setStringValue(val)
                        )
                        .setUserEnteredFormat(
                            new CellFormat().setTextFormat(
                                new TextFormat().setBold(isbold)
                            )
                        )
                    )
                    .setRange(new GridRange()
                        .setSheetId(Integer.parseInt(sheetID))
                        .setStartRowIndex(sci)
                        .setEndRowIndex(eci)
                        .setStartColumnIndex(sri)
                        .setEndColumnIndex(eri)
                    )
                    .setFields("*")
                )
            );
        
        
        BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest()
                .setRequests(requests);
        
        BatchUpdateSpreadsheetResponse batchUpdateResponse = sheetsService.spreadsheets().batchUpdate(spreadsheetKey, batchUpdateRequest)
                .execute();
    }
    
    /***********
     * Set the background color of the index
     * @param blue
     * @param red
     * @param green
     * @param startingrowindex
     * @param startcolumnindex 
     */
    public void setColor(String sheetID, float blue, float red, float green, int startingrowindex, int startcolumnindex)
    {
        List<Request> requests = new ArrayList<Request>();
        
        requests.add 
            ( new Request()
                .setRepeatCell(new RepeatCellRequest()
                    .setCell( new CellData()
                        .setUserEnteredFormat( new CellFormat()
                            .setBackgroundColor( new Color()
                                .setRed(red)
                                .setBlue(blue)
                                .setGreen(green)
                            )
                        )
                    )
                    .setRange(new GridRange()
                        .setSheetId(Integer.parseInt(sheetID))
                        .setStartRowIndex(startcolumnindex)
                        .setEndRowIndex(startcolumnindex + 1)
                        .setStartColumnIndex(startingrowindex)
                        .setEndColumnIndex(startingrowindex + 1)
                    )
                    .setFields("*")
                )
            );
    }
    
    /*******
     * Delete the specified number of rows starting from the beginning of the spreadsheet (row 0). 
     * @param numofrows
     * @throws IOException
     * @throws InterruptedException 
     */
    public void deleterows(int numofrows) throws IOException, InterruptedException
    {
        BatchUpdateSpreadsheetRequest content = new BatchUpdateSpreadsheetRequest();
        Request request = new Request()
                .setDeleteDimension(new DeleteDimensionRequest()
                  .setRange(new DimensionRange()
                    .setSheetId(Integer.parseInt(spreadsheetKey))
                    .setDimension("ROWS")
                    .setStartIndex(0)
                    .setEndIndex(numofrows)
                  )
                );

        List<Request> requests = new ArrayList<Request>();
        requests.add(request);
        content.setRequests(requests);

        sheetsService.spreadsheets().batchUpdate(spreadsheetKey, content).execute();
    } 

    /*******
     * Allows for accessing the full list of data. 
     * @param spreadsheetid
     * @param sheetname
     * @return
     * @throws IOException
     * @throws GeneralSecurityException
     * @throws InterruptedException 
     */
    public List<List<Object>> accessdata(String spreadsheetid, String sheetname) throws IOException, GeneralSecurityException, InterruptedException
    {
        ValueRange response = sheetsService.spreadsheets().values().get(spreadsheetid, sheetname).execute(); 
        
        List<List<Object>> values = response.getValues();
        
        if (values == null || response.isEmpty())
        {
            return null; 
        }
        else
        {
            return values; 
        }
    }
    
    /******
     * Clears all rows in the Googlesheet
     * @throws IOException
     * @throws InterruptedException 
     */
    public void cleartable(String sheetID) throws IOException, InterruptedException 
    {
        ClearValuesRequest requestBody = new ClearValuesRequest();

        Sheets.Spreadsheets.Values.Clear request =
            sheetsService.spreadsheets().values().clear(spreadsheetKey, sheetID, requestBody);
        sleepABit();

        ClearValuesResponse response = request.execute();
    }
        
    private String timetostr(Timestamp t)
    {
        if (t == null)
        {
            return "NULL";
        }
        else
        {
            if (!t.toString().contains("-"))
                System.out.println(t.toString());
            return t.toString().split(" ")[0]; 
        }
    }
    
    // allows the program to pause a bit for the Google server to catch up. 
    private static void sleepABit() throws InterruptedException {    
        Thread.sleep(5000);
    }
}
