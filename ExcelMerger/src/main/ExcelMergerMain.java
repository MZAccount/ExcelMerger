package main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeSet;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.collect.Ordering;

public class ExcelMergerMain {

	public static void main(String[] args) {

		Map<String, String> x = new HashMap();
		x.put("fileLocations", args[0]);
		String fileLocations = x.get("fileLocations");

		List<String> fileLocationsList = getLocations(fileLocations);

		List<Sheet> checkedExcelSheets = getSheets(fileLocationsList);

		if (checkedExcelSheets.isEmpty())
			return;

		List<String> datasetExcelHeaders = new ArrayList<String>();

//file <- system.file("tests", "test_import.xlsx", package = "xlsx")
//res <- read.xlsx(file, 1)  # read first sheet

		for (Sheet sheet : checkedExcelSheets) {

//	checkOnlyText(excelData)

			Map<Integer, String> excelHeaders = getHeaders(sheet);

//	checkNoDuplicate(excelHeader)

			datasetExcelHeaders.addAll(excelHeaders.values());

		}

		datasetExcelHeaders.sort(Ordering.usingToString());
		// assert datasetExcelHeaders.isSortedAlpha()

		File file = new File("C:\\Users\\Uber\\git\\ExcelMerger\\ExcelMerger\\testData\\test1\\output\\out.xlsx");
		Workbook workbook;
		try {
			workbook = new XSSFWorkbook(file);
		} catch (InvalidFormatException | IOException e) {
			return;
		}

		Sheet finalDatasetData = workbook.getSheetAt(0);

//finalDatasetData=new datasetData[datasetExcelData.size()][datasetExcelHeaders.size()]	
		{
			int i = 0;
			for (Sheet sheet : checkedExcelSheets) {

				int column = i;
				int finalColumn = datasetExcelHeaders.indexOf(sheet.getRow(0).getCell(column).getStringCellValue());
				finalDatasetData.getRow(i).getCell(finalColumn)
						.setCellValue(sheet.getRow(1).getCell(column).getStringCellValue());
				i++;
			}
		}

	}

	private static Map<Integer, String> getHeaders(Sheet sheet) {
		// TODO Auto-generated method stub
		return null;
	}

	private static List<Sheet> getSheets(List<String> fileLocationsList) {
		// TODO Auto-generated method stub
		String fileLocation = fileLocationsList.get(0);
		List l=new ArrayList();

		FileInputStream file;
		try {
			file = new FileInputStream(new File(fileLocation));
		Workbook workbook;
		try {
			workbook = new XSSFWorkbook(file);
		Sheet sheet = workbook.getSheetAt(0);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		} catch (FileNotFoundException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		 
		return l;
	}

	private static List<String> getLocations(String fileLocations) {
		// TODO Auto-generated method stub
		List l=new ArrayList();
		l.add(fileLocations);
		return l;
	}

}
