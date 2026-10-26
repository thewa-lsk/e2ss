package xyz.cuddlecloud.e2ss;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.Reader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

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
import com.google.gson.Gson;

public class ExcelToSheetsSync {

	private static final Logger LOGGER = LoggerFactory.getLogger(ExcelToSheetsSync.class);
	private static final String DEFAULT_CONFIG_FILE = "config.json";
	public static final String appName = "Excel To Sheets Sync";

	public static class Config {
		public String credentialsFile = "credentials.json";
		public String tokenDir = "tokens";
		public int port = 12751;
		public List<SheetMapping> sheets = new ArrayList<>();
	}

	public static class SheetMapping {
		public int id;
		public String source;
		public String sheetId;
	}

	private static Config loadConfig(String configPath) throws IOException {
		try (Reader reader = Files.newBufferedReader(Path.of(configPath))) {
			return new Gson().fromJson(reader, Config.class);
		}
	}

	private static Sheets createSheetsService(Config config) throws IOException, GeneralSecurityException {

		NetHttpTransport transport = new NetHttpTransport();

		GoogleClientSecrets clientSecrets;
		try (Reader reader = Files.newBufferedReader(Path.of(config.credentialsFile))) {
			clientSecrets = GoogleClientSecrets.load(GsonFactory.getDefaultInstance(), reader);
		}

		GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(transport, GsonFactory.getDefaultInstance(), clientSecrets, Collections.singletonList(SheetsScopes.SPREADSHEETS)).setDataStoreFactory(new FileDataStoreFactory(new File(config.tokenDir))).setAccessType("offline").build();

		Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver.Builder().setPort(config.port).build()).authorize("user");

		return new Sheets.Builder(transport, GsonFactory.getDefaultInstance(), credential).setApplicationName(appName).build();
	}

	private static Map<String, List<List<Object>>> readSource(String sourcePath) throws IOException {
		if (sourcePath.toLowerCase(Locale.ROOT).endsWith(".csv")) {
			return readCsv(sourcePath);
		}
		return readExcel(sourcePath);
	}

	private static Map<String, List<List<Object>>> readCsv(String csvPath) throws IOException {
		Path path = Path.of(csvPath);
		String fileName = path.getFileName().toString();
		String sheetName = fileName.toLowerCase(Locale.ROOT).endsWith(".csv") ? fileName.substring(0, fileName.length() - 4) : fileName;
		if (sheetName.isEmpty()) {
			sheetName = "Sheet1";
		}

		String content = Files.readString(path, StandardCharsets.UTF_8);
		if (content.startsWith("\uFEFF")) {
			content = content.substring(1);
		}

		List<List<Object>> rows = parseCsv(content);
		Map<String, List<List<Object>>> result = new LinkedHashMap<>();
		result.put(sheetName, rows);
		return result;
	}

	private static List<List<Object>> parseCsv(String content) {
		List<List<Object>> rows = new ArrayList<>();
		List<Object> currentRow = new ArrayList<>();
		StringBuilder currentCell = new StringBuilder();
		boolean inQuotes = false;
		int len = content.length();

		for (int i = 0; i < len; i++) {
			char c = content.charAt(i);
			if (inQuotes) {
				if (c == '"') {
					if (i + 1 < len && content.charAt(i + 1) == '"') {
						currentCell.append('"');
						i++;
					} else {
						inQuotes = false;
					}
				} else {
					currentCell.append(c);
				}
			} else {
				if (c == '"' && currentCell.length() == 0) {
					inQuotes = true;
				} else if (c == ',') {
					currentRow.add(currentCell.toString());
					currentCell.setLength(0);
				} else if (c == '\r' || c == '\n') {
					if (c == '\r' && i + 1 < len && content.charAt(i + 1) == '\n') {
						i++;
					}
					currentRow.add(currentCell.toString());
					currentCell.setLength(0);
					rows.add(currentRow);
					currentRow = new ArrayList<>();
				} else {
					currentCell.append(c);
				}
			}
		}

		if (currentCell.length() > 0 || !currentRow.isEmpty()) {
			currentRow.add(currentCell.toString());
			rows.add(currentRow);
		}

		int maxColumns = 0;
		for (List<Object> row : rows) {
			maxColumns = Math.max(maxColumns, row.size());
		}
		for (List<Object> row : rows) {
			while (row.size() < maxColumns) {
				row.add("");
			}
		}

		return rows;
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
			throw new IllegalArgumentException("Excel file contains no worksheets");
		}

		return result;
	}

	private static String quoteSheetName(String name) {
		return "'" + name.replace("'", "''") + "'";
	}

	public static void sync(Sheets service, String sourcePath, String spreadsheetId) throws IOException {

		Map<String, List<List<Object>>> excelSheets = readSource(sourcePath);

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

			LOGGER.info("Synced sheet: " + name);
		}

		LOGGER.info("Sync complete: " + sourcePath);
		LOGGER.info("Spreadsheet: https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/edit");
	}

	public static void main(String[] args) throws Exception {
		String configPath = DEFAULT_CONFIG_FILE;
		if (args.length > 0) {
			configPath = args[0];
		}

		Config config = loadConfig(configPath);
		if (config.sheets == null || config.sheets.isEmpty()) {
			LOGGER.info("No sheets configured in: " + configPath);
			return;
		}

		Sheets service = createSheetsService(config);

		for (SheetMapping mapping : config.sheets) {
			LOGGER.info("Syncing [#" + mapping.id + "] " + mapping.source + " -> " + mapping.sheetId);
			sync(service, mapping.source, mapping.sheetId);
		}
	}
}
