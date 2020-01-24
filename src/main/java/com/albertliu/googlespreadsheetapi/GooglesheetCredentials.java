/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.albertliu.googlespreadsheetapi;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.SheetsScopes;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.Arrays;
import java.util.List;
import java.util.Properties;

/**
 *
 * @author Albert Liu
 */
public class GooglesheetCredentials {

    protected static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
    private static final String TOKENS_DIRECTORY_PATH = "tokens1";
    public final Properties prop;
    protected static final List<String> SCOPES = Arrays.asList(SheetsScopes.SPREADSHEETS);
    private static String CREDENTIALS_FULL_PROJECT_PATH;

    protected static Credential getCredentials(final HttpTransport HTTP_TRANSPORT, int portnum) throws IOException {
        
        InputStream in = new FileInputStream(CREDENTIALS_FULL_PROJECT_PATH);
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                .setDataStoreFactory(new FileDataStoreFactory(new File(TOKENS_DIRECTORY_PATH)))
                .setAccessType("offline")
                .build();
        LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(portnum).build();
        return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
    }
 
    public GooglesheetCredentials(Properties prop, final String credentialFilePath) {
        CREDENTIALS_FULL_PROJECT_PATH = credentialFilePath;
        this.prop = prop;
    }

}