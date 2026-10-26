package xyz.cuddlecloud.e2ss;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.Reader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.AddSheetRequest;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetRequest;
import com.google.api.services.sheets.v4.model.ClearValuesRequest;
import com.google.api.services.sheets.v4.model.DeleteSheetRequest;
import com.google.api.services.sheets.v4.model.Request;
import com.google.api.services.sheets.v4.model.SheetProperties;
import com.google.api.services.sheets.v4.model.Spreadsheet;
import com.google.api.services.sheets.v4.model.ValueRange;

public class ExcelToSheetsSync {

	private static final String APP_NAME = "Excel To Sheets Sync";
	private static final String CREDENTIALS_FILE = "credentials.json";
	private static final String TOKEN_DIR = "tokens";

	private static Sheets createSheetsService() throws IOException, GeneralSecurityException {

		NetHttpTransport transport = new NetHttpTransport();

		GoogleClientSecrets clientSecrets;
		try (Reader reader = Files.newBufferedReader(Path.of(CREDENTIALS_FILE))) {
			clientSecrets = GoogleClientSecrets.load(GsonFactory.getDefaultInstance(), reader);
		}

		GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(transport, GsonFactory.getDefaultInstance(), clientSecrets, Collections.singletonList(SheetsScopes.SPREADSHEETS)).setDataStoreFactory(new FileDataStoreFactory(new File(TOKEN_DIR))).setAccessType("offline").build();

		Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver.Builder().setPort(8888).build()).authorize("user");

		return new Sheets.Builder(transport, GsonFactory.getDefaultInstance(), credential).setApplicationName(APP_NAME).build();
	}

	private static Map<String, List<List<Object>>> readExcel(String excelPath) throws IOException {

		Map<String, List<List<Object>>> result = new LinkedHashMap<>();

		DataFormatter formatter = new DataFormatter();

		try (InputStream in = Files.newInputStream(Path.of(excelPath)); Workbook workbook = new XSSFWorkbook(in)) {

			FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

			for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
				Sheet sheet = workbook.getSheetAt(s);
				List<List<Object>> rows = new ArrayList<>();

				int lastRow = sheet.getLastRowNum();
				int maxColumns = 0;

				for (Row row : sheet) {
					maxColumns = Math.max(maxColumns, row.getLastCellNum());
				}

				for (int r = 0; r <= lastRow; r++) {
					Row row = sheet.getRow(r);
					List<Object> cells = new ArrayList<>();

					for (int c = 0; c < maxColumns; c++) {
						Cell cell = row == null ? null : row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

						String value = cell == null ? "" : formatter.formatCellValue(cell, evaluator);

						cells.add(value);
					}

					rows.add(cells);
				}

				result.put(sheet.getSheetName(), rows);
			}
		}

		if (result.isEmpty()) {
			throw new IllegalArgumentException("Excel ไม่มี worksheet");
		}

		return result;
	}

	private static String quoteSheetName(String name) {
		return "'" + name.replace("'", "''") + "'";
	}

	public static void sync(String excelPath, String spreadsheetId) throws IOException, GeneralSecurityException {

		Map<String, List<List<Object>>> excelSheets = readExcel(excelPath);

		Sheets service = createSheetsService();

		Spreadsheet spreadsheet = service.spreadsheets().get(spreadsheetId).setFields("sheets.properties").execute();

		List<com.google.api.services.sheets.v4.model.Sheet> existingSheets = spreadsheet.getSheets();

		Map<String, Integer> existingByName = new LinkedHashMap<>();

		for (com.google.api.services.sheets.v4.model.Sheet sheet : existingSheets) {
			SheetProperties p = sheet.getProperties();
			existingByName.put(p.getTitle(), p.getSheetId());
		}

		List<Request> addRequests = new ArrayList<>();

		for (String name : excelSheets.keySet()) {
			if (!existingByName.containsKey(name)) {
				addRequests.add(new Request().setAddSheet(new AddSheetRequest().setProperties(new SheetProperties().setTitle(name))));
			}
		}

		if (!addRequests.isEmpty()) {
			service.spreadsheets().batchUpdate(spreadsheetId, new BatchUpdateSpreadsheetRequest().setRequests(addRequests)).execute();
		}

		spreadsheet = service.spreadsheets().get(spreadsheetId).setFields("sheets.properties").execute();

		existingSheets = spreadsheet.getSheets();
		existingByName.clear();

		for (com.google.api.services.sheets.v4.model.Sheet sheet : existingSheets) {
			SheetProperties p = sheet.getProperties();
			existingByName.put(p.getTitle(), p.getSheetId());
		}

		List<Request> deleteRequests = new ArrayList<>();

		for (com.google.api.services.sheets.v4.model.Sheet sheet : existingSheets) {
			SheetProperties p = sheet.getProperties();

			if (!excelSheets.containsKey(p.getTitle())) {
				deleteRequests.add(new Request().setDeleteSheet(new DeleteSheetRequest().setSheetId(p.getSheetId())));
			}
		}

		if (!deleteRequests.isEmpty()) {
			service.spreadsheets().batchUpdate(spreadsheetId, new BatchUpdateSpreadsheetRequest().setRequests(deleteRequests)).execute();
		}

		for (Map.Entry<String, List<List<Object>>> entry : excelSheets.entrySet()) {

			String name = entry.getKey();
			List<List<Object>> values = entry.getValue();
			String range = quoteSheetName(name);
			String clearRange = range + "!A:ZZZ";

			service.spreadsheets().values().clear(spreadsheetId, clearRange, new ClearValuesRequest()).execute();

			if (!values.isEmpty() && values.stream().anyMatch(row -> row.stream().anyMatch(value -> value != null && !value.toString().isEmpty()))) {

				ValueRange body = new ValueRange().setMajorDimension("ROWS").setValues(values);

				service.spreadsheets().values().update(spreadsheetId, range + "!A1", body).setValueInputOption("RAW").execute();
			}

			System.out.println("Synced sheet: " + name);
		}

		System.out.println("Sync complete.");
		System.out.println("Spreadsheet: https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/edit");
	}

	public static void main(String[] args) throws Exception {
		// https://docs.google.com/spreadsheets/d/1CMP6faExxQ7x9s7aNOi53fdLe0q4JctkHMCYrDImW8A/edit
		String excelPath = "multiple_table.xlsx";
		if (args.length > 0) {
			excelPath = args[0];
		}

		String spreadsheetId = "1CMP6faExxQ7x9s7aNOi53fdLe0q4JctkHMCYrDImW8A";

		sync(excelPath, spreadsheetId);
	}
}
