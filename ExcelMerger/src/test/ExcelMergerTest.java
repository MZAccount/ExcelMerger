package test;

import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TestName;

import main.ExcelMergerMain;

public class ExcelMergerTest {

	@Rule
	public TestName name = new TestName();

	@Test
	public void test1() {
		String testName = name.getMethodName();
		new ExcelMergerMain();

		ExcelMergerMain.main(new String[] {
				"-i testData\\" + testName + "\\input\\input1.xlsx -o testData\\" + testName + "\\output\\out.xlsx" });

		assertOutputFiles(testName);
	}

	@Test
	public void test2() {
		String testName = name.getMethodName();
		new ExcelMergerMain();

		ExcelMergerMain.main(new String[] { "-i testData\\" + testName + "\\input\\input1.xlsx testData\\" + testName
				+ "\\input\\input2.xlsx -o testData\\" + testName + "\\output\\out.xlsx" });

		assertOutputFiles(testName);
	}

	@Test
	public void test3() {
		String testName = name.getMethodName();
		new ExcelMergerMain();

		ExcelMergerMain.main(new String[] {
				"-i testData\\" + testName + "\\input -o testData\\" + testName + "\\output\\out.xlsx" });

		assertOutputFiles(testName);
	}

	private void assertOutputFiles(String testName) {
		File file1 = new File("testData\\" + testName + "\\output\\out.xlsx");
		File file2 = new File("testData\\" + testName + "\\referenceOutput\\out.xlsx");

		Sheet sheet1 = null;
		Sheet sheet2 = null;
		try {
			FileInputStream file;
			file = new FileInputStream(file1);

			try(Workbook workbook = new XSSFWorkbook(file)){
			Sheet sheet = workbook.getSheetAt(0);
			sheet1 = sheet;
			}
		} catch (IOException e) {
		}

		assertNotNull(sheet1);
		try {
			FileInputStream file;
			file = new FileInputStream(file2);
			try {
				try(Workbook workbook = new XSSFWorkbook(file)){
				Sheet sheet = workbook.getSheetAt(0);
				sheet2 = sheet;}
			} catch (IOException e) {
			}
		} catch (FileNotFoundException e1) {
		}
		assertNotNull(sheet2);

		boolean isTwoEqual = checkEqual(sheet1, sheet2);
		assertTrue(isTwoEqual);
	}

	private boolean checkEqual(Sheet sheet1, Sheet sheet2) {
		ArrayList<String> l1 = new ArrayList<String>();
		ArrayList<String> l2 = new ArrayList<String>();
		for (Row row : sheet1) {
			l1.add(row.toString());
		}
		for (Row row : sheet2) {
			l2.add(row.toString());
		}
		return l1.equals(l2);
	}

}
