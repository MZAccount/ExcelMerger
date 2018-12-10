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
import org.junit.jupiter.api.Test;

import main.ExcelMergerMain;

public class ExcelMergerTest {

	


    @Test
    public void test1() {
        ExcelMergerMain tester = new ExcelMergerMain(); // ExcelMergerMain is tested

        ExcelMergerMain.main(new String[]{"C:\\Users\\Uber\\git\\ExcelMerger\\ExcelMerger\\testData\\test1\\input\\input1.xlsx"});
        
    }


    @Test
    public void test2() {
        ExcelMergerMain tester = new ExcelMergerMain(); // ExcelMergerMain is tested

        ExcelMergerMain.main(new String[]{"C:\\Users\\Uber\\git\\ExcelMerger\\ExcelMerger\\testData\\test2\\input\\input1.xlsx C:\\Users\\Uber\\git\\ExcelMerger\\ExcelMerger\\testData\\test1\\input\\input2.xlsx"});
        
    }
	


    @Test
    public void test3() {
        ExcelMergerMain tester = new ExcelMergerMain(); // ExcelMergerMain is tested

        ExcelMergerMain.main(new String[]{"-i testData\\test3\\input -o testData\\test3\\output\\out.xlsx"});
        
        File file1 = new File("testData\\test3\\output\\out.xlsx");
        File file2 = new File("testData\\test3\\referenceOutput\\out.xlsx");
        

        Sheet sheet1 = null;
        Sheet sheet2 = null;
		try {
			FileInputStream file;
			file = new FileInputStream(file1);
			
				Workbook workbook;
				workbook = new XSSFWorkbook(file);
				Sheet sheet = workbook.getSheetAt(0);
				sheet1 = sheet;
			} catch (IOException e) {
				@SuppressWarnings("unused")
				String s="it broke";
			}
		
		
		assertNotNull(sheet1);
		try {
			FileInputStream file;
			file = new FileInputStream(file2);
			try {
				Workbook workbook;
				workbook = new XSSFWorkbook(file);
				Sheet sheet = workbook.getSheetAt(0);
				sheet2 = sheet;
			} catch (IOException e) {
			}
		} catch (FileNotFoundException e1) {
		}
		assertNotNull(sheet2);

		
		
        boolean isTwoEqual = checkEqual(sheet1,sheet2);
		assertTrue(isTwoEqual);
    }


	private boolean checkEqual(Sheet sheet1, Sheet sheet2) {
		ArrayList l1=new ArrayList();
		ArrayList l2=new ArrayList();
		for (Row row : sheet1) {
			l1.add(row.toString());
		}
		for (Row row : sheet2) {
			l2.add(row.toString());
		}
		return l1.equals(l2);
	}
	
	
}
