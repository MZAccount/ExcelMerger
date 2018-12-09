package main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.collect.Ordering;

public class ExcelMergerMain {

	private enum Location {
		@Deprecated
		Onefile, Folder, DelimiterSeparatedFiles
	}

	public static void main(String[] args) {

		Map<String, String> x = new HashMap();
		x.put("fileLocations", args[0]);
		String fileLocations = x.get("fileLocations");

		List<String> fileLocationsList = getLocations(fileLocations, Location.DelimiterSeparatedFiles);

		List<Sheet> checkedExcelSheets = getSheets(fileLocationsList);

		if (checkedExcelSheets.isEmpty())
			return;

		List<String> datasetExcelHeaders = new ArrayList<String>();

//file <- system.file("tests", "test_import.xlsx", package = "xlsx")
//res <- read.xlsx(file, 1)  # read first sheet

		for (Sheet sheet : checkedExcelSheets) {

//	checkOnlyText(excelData)

			List<String> excelHeaders = getHeaders(sheet);

//	checkNoDuplicate(excelHeader)

			datasetExcelHeaders.addAll(excelHeaders);

		}

		datasetExcelHeaders.sort(Ordering.usingToString());
		// assert datasetExcelHeaders.isSortedAlpha()

		File file = new File("C:\\Users\\Uber\\git\\ExcelMerger\\ExcelMerger\\testData\\test1\\output\\out.xlsx");
		Workbook workbook;
		// Create a Workbook
		workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

		// Create a Sheet
		Sheet finalDatasetData = workbook.createSheet();

		int numberOfFinalRows = checkedExcelSheets.size() + 1;
		for (int i = 0; i < numberOfFinalRows; i++) {

			finalDatasetData.createRow(i);
			for (int j = 0; j < datasetExcelHeaders.size(); j++) {
				finalDatasetData.getRow(i).createCell(j);
			}
		}

		// Fill headers in final excel
		{
			Row headersRow = finalDatasetData.getRow(0);
			int i = 0;
			for (Cell cell : headersRow) {
				String value = datasetExcelHeaders.get(i);
				cell.setCellValue(value);
				i++;
			}
		}


		// Fill other rows with data from all sheets, insert data into corresponding
		// dataset/global header index
		{
			int i = 0;
			for (Sheet sheet : checkedExcelSheets) {
				List<String> headers = getHeaders(sheet);
				int j = 0;
				for (String header : headers) {

					int column = i;
					int finalColumn = datasetExcelHeaders.indexOf(header);
					finalDatasetData.getRow(i + 1).getCell(finalColumn)
							.setCellValue(sheet.getRow(1).getCell(j).toString());
					j++;
				}
				i++;
			}
		}

		try {
			workbook.write(new FileOutputStream(file));
		} catch (IOException e) {
			System.out.println("Couldn't write to file");
			e.printStackTrace();
		}
	}

	private static List<String> getHeaders(Sheet sheet) {

		Row headers = sheet.getRow(0);
		List l = new ArrayList();
		for (Cell cell : headers) {
			l.add(cell.getStringCellValue());
		}

		return l;
	}

	private static List<Sheet> getSheets(List<String> fileLocationsList) {
		// TODO Auto-generated method stub
		List l = new ArrayList();
		for (String fileLocation : fileLocationsList) {

			FileInputStream file;
			try {
				file = new FileInputStream(new File(fileLocation));
				Workbook workbook;
				try {
					workbook = new XSSFWorkbook(file);
					Sheet sheet = workbook.getSheetAt(0);
					l.add(sheet);
				} catch (IOException e) {
				}
			} catch (FileNotFoundException e1) {
			}
		}
		return l;
	}

	private static List<String> getLocations(String fileLocations, Location locationType) {

		List l = new ArrayList();
		switch (locationType) {
		case Onefile:
			l.add(fileLocations);
			return l;

			
			case DelimiterSeparatedFiles:
			
				l.addAll(Arrays.asList(fileLocations.split(" ")));
				return l;

		default:
			throw new EnumConstantNotPresentException(Location.class, "locationType");

		}
	}

}
