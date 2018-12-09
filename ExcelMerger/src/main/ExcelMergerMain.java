package main;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeSet;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelMergerMain {

	public static void main(String[] args) {

		Map<String, String> x = new HashMap();
		x.put("fileLocations", "");
		String fileLocations = x.get("fileLocations");

		List<String> fileLocationsList = getLocations(fileLocations);

		List<Sheet> checkedExcelSheets = getSheets(fileLocationsList);

		if (checkedExcelSheets.isEmpty())
			return;

		TreeSet< String> datasetExcelHeaders = new TreeSet< String>();

//file <- system.file("tests", "test_import.xlsx", package = "xlsx")
//res <- read.xlsx(file, 1)  # read first sheet

		for (Sheet sheet : checkedExcelSheets) {

//	checkOnlyText(excelData)
			
			Map<Integer, String> excelHeaders = getHeaders(sheet);

//	checkNoDuplicate(excelHeader)

			datasetExcelHeaders.addAll(excelHeaders.values());

		}

	//assert 	datasetExcelHeaders.isSortedAlpha()
		
//finalDatasetData=new datasetData[datasetExcelData.size()][datasetExcelHeaders.size()]	
//
//foreach excelData//i=0
//	foreach column
//		finalColumn=datasetExcelHeaders.getIndex(excelData[0][column])
//		finalDatasetData[i][finalColumn]=datasetExcelHeaders[1][column]
//		i++
//
//
	}

	private static Map<Integer, String> getHeaders(Sheet sheet) {
		// TODO Auto-generated method stub
		return null;
	}

	private static List<Sheet> getSheets(List<String> fileLocationsList) {
		// TODO Auto-generated method stub
		return null;
	}

	private static List<String> getLocations(String fileLocations) {
		// TODO Auto-generated method stub
		return null;
	}

}
