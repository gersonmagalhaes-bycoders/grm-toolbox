package sofia.toolbox.msoffice.testcase;

import static org.junit.Assert.*;

import java.io.File;
import java.io.IOException;

import org.junit.Test;

import sofia.toolbox.io.FileSystem;
import sofia.toolbox.msoffice.Excel97;


public class Excel97Test {
	
	String currentDir = new FileSystem().getCurrentDirectory() + File.separator + "src" + File.separator + "test" + File.separator + "resources" + File.separator;	

	@Test
	public void testCreateWorkbook() throws IOException {
		
		File file = new File(currentDir + "Resultado.xls");
		if (file.exists()) file.delete();
		
        Excel97 excel = new Excel97();
        excel.createWorkbook("Resultado.xls", new String[]{"Nome Documento"}, new int[]{8000});

        excel.createRow();
        excel.createColumn("ADP2-VT3-DFH-00002_B.DOC", Excel97.ALIGN_LEFT);

        excel.createRow();
        excel.createColumn("ADP2-VT3-DFH-00002_C.DOC", Excel97.ALIGN_LEFT);
        
        excel.setCellValue(15, 15, "Gerson", Excel97.ALIGN_CENTER);

        excel.save("src/test/resources/resultado.xls") ;
		assertTrue("Value of Cell (0, 0) is wrong", excel.getStringCellValue(0,  0).equals("Nome Documento"));
		assertTrue("Value of Cell (0, 1) is wrong", excel.getStringCellValue(0,  1).equals("ADP2-VT3-DFH-00002_B.DOC"));
		assertTrue("Value of Cell (0, 2) is wrong", excel.getStringCellValue(0,  2).equals("ADP2-VT3-DFH-00002_C.DOC"));	
		
		assertTrue("Value of Cell (15, 15) is wrong", excel.getStringCellValue(15,  15).equals("Gerson"));		
	}

}
