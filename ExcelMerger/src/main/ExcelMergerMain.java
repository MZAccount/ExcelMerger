package main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.Stream;

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

		Map<String, String> x = new HashMap<String, String>();
//		x.put("fileLocations", args[0]);
		x = getArguments(args);

		String fileLocations = x.get("--input");

		Location locationOption = Location.Folder;
		List<String> fileLocationsList = getLocations(fileLocations, locationOption);

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
//TODO:change output
		String outputPath = x.get("--output");
		File file = new File(outputPath);
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
			workbook.close();//TODO: should open as resource?
		} catch (IOException e) {
			System.out.println("Couldn't write to file");
			e.printStackTrace();
		}
	}

	private static enum Options {
		Input("-i", "--input"), Output("-o", "--output");

		private List<String> mS;

		// Constructor
		Options(String... s) {
			mS = Arrays.asList(s);
		}

		public static Options match(String s) {
			Options[] values = Options.values();
			for (Options option : values) {
				if (option.mS.contains(s)) {
					return option;
				}
			}
			return null;
		}

		@Override
		public String toString() {
			return mS.get(mS.size() - 1);
		}
	}

	//TODO:remake
	private static Map<String, String> getArguments(String[] ar) {
		String[] args=ar[0].split(" ");
		HashMap<String, String> x = new HashMap<String, String>();
		for (int i = 0; i < args.length; i++) {
			if (Options.match(args[i]) != null) {
				Options crt = Options.match(args[i]);
				i++;
				StringBuilder s = new StringBuilder();
				for(;(i < args.length)&&(Options.match(args[i]) == null);i++) {
					s.append(args[i]);
				}
				i--;
				x.put(crt.toString(), s.toString());
			}
		}

		return x;
	}

	@SuppressWarnings("unused")
	final private static String getCurrentDirectory() {
//		final String dir = System.getProperty("user.dir");
//		System.out.println("current dir1 = " + dir);
//		Path currentWorkingDir = Paths.get("").toAbsolutePath();
//		System.out.println("current dir2 (java7)= "+currentWorkingDir.normalize().toString());
//		System.out.println("current dir3= "+(new File("")).getAbsolutePath());

		return Paths.get("").toAbsolutePath().normalize().toString();

	}

	private static List<String> getHeaders(Sheet sheet) {

		Row headers = sheet.getRow(0);
		List<String> l = new ArrayList<String>();
		for (Cell cell : headers) {
			l.add(cell.getStringCellValue());
		}

		return l;
	}

	private static List<Sheet> getSheets(List<String> fileLocationsList) {
		// TODO Auto-generated method stub
		List<Sheet> l = new ArrayList<Sheet>();
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

		List<String> l = new ArrayList<String>();
		switch (locationType) {
		case Onefile:
			l.add(fileLocations);
			return l;

		case DelimiterSeparatedFiles:

			l.addAll(Arrays.asList(fileLocations.split(" ")));
			return l;
		case Folder:
			try (Stream<Path> paths = Files.walk(Paths.get(fileLocations))) {
				l = paths.filter(Files::isRegularFile).filter(p -> p.getFileName().toString().matches(".+\\.xlsx"))
						.filter(p -> !p.getFileName().toString().equals("output.xlsx"))
						.filter(p -> !p.getFileName().toString().matches("~.*")).map(p -> p.toString())
						.collect(Collectors.toList());
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			return l;

		default:
			throw new EnumConstantNotPresentException(Location.class, "locationType");

		}
	}

}
