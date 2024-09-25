
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;

import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.*;

public class TestPoi {

	Path inputFilePath = Paths.get("src", "test", "data", "smallBook.xlsx");

	@Test
	public void testReadSpreadSheet() throws IOException {
		FileInputStream fin = new FileInputStream(inputFilePath.toFile());
		Workbook wb = WorkbookFactory.create(fin);
		Sheet sheet = wb.getSheet("Sheet1");
		assertNotNull(sheet);
		Row row = sheet.getRow(0);
		assertNotNull(row);
		Cell c = row.getCell(0);
		assertNotNull(c);
		String value = evaluateCell(c, wb);
		assertThat(value, is("1.0"));

		Cell c2 = row.getCell(1);
		String value2 = evaluateCell(c2, wb);
		assertThat(value2, is("2.0"));

		wb.close();
	}


    @Test
	public void testWriteSpreadSheet() throws IOException {
        Path outputDir = Paths.get("build", "testData");
        Files.createDirectories(outputDir);
        Path outputSS = outputDir.resolve("ss.xlsx");
        Files.copy(inputFilePath, outputSS, StandardCopyOption.REPLACE_EXISTING);

		FileInputStream fin = new FileInputStream(outputSS.toFile());
		Workbook wb = WorkbookFactory.create(fin);
		Sheet sheet = wb.getSheet("Sheet1");
		assertNotNull(sheet);
		Row row = sheet.getRow(0);
		assertNotNull(row);
		Cell c = row.getCell(0);
		assertNotNull(c);

        c.setCellValue(5.0);

		String value = evaluateCell(c, wb);
		assertThat(value, is("5.0"));

		Cell c2 = row.getCell(1);
		String value2 = evaluateCell(c2, wb);
		assertThat(value2, is("10.0"));

		wb.close();
	}



	private String evaluateCell(Cell c, Workbook workbook) {
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        String value = "";
        if (c != null) {
            CellValue cellValue = null;
            try {
                cellValue = evaluator.evaluate(c);
            } catch (Exception ex) {
                return "**err**";
            }
            if (cellValue != null) {
                CellType cellType = cellValue.getCellType();
                switch (cellType) {
                case STRING:
                    value = cellValue.getStringValue();
                    break;
                case NUMERIC:
                    value = "" + cellValue.getNumberValue();
                    if (value.matches("^[+-]?[0-9]*[.][0-9][0-9][0-9][0-9]*")) {
                        Double d = Double.parseDouble(value);
                        value = String.format("%.2f", d);
                    }
                    break;
                case BOOLEAN:
                    value = "" + cellValue.getBooleanValue();
                    break;
                default:
                    return "??";
                }
            }
        }
        return value;
    }



}
