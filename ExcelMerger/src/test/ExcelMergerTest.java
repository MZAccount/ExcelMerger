package test;

import org.junit.jupiter.api.Test;

import main.ExcelMergerMain;

public class ExcelMergerTest {

	


    @Test
    public void test1() {
        ExcelMergerMain tester = new ExcelMergerMain(); // ExcelMergerMain is tested

        tester.main(new String[]{"C:\\Users\\Uber\\git\\ExcelMerger\\ExcelMerger\\testData\\test1\\input\\input1.xlsx"});
        
    }


    @Test
    public void test2() {
        ExcelMergerMain tester = new ExcelMergerMain(); // ExcelMergerMain is tested

        tester.main(new String[]{"C:\\Users\\Uber\\git\\ExcelMerger\\ExcelMerger\\testData\\test1\\input\\input1.xlsx C:\\Users\\Uber\\git\\ExcelMerger\\ExcelMerger\\testData\\test1\\input\\input2.xlsx"});
        
    }
	
	
}
