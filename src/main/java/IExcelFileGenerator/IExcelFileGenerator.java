package IExcelFileGenerator;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Method;
import java.util.Date;
import java.util.LinkedHashMap;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;

import IExcelFileGeneratorExceptions.IExcelFileGeneratorException;

public class IExcelFileGenerator {
	
	private HSSFWorkbook workbook;

	private HSSFSheet sheet;

	private Integer rowCount = 0;

	private Integer columnCount = 0;

	private LinkedHashMap<String, Integer> sheetNameSet = new LinkedHashMap<>();

	private final MethodUtil cache = new MethodUtil();

	private File file;

	/* Constructor to create workbook and sheet in excel */
	public IExcelFileGenerator(String file, String sheetName) throws Exception {

		this.workbook = new HSSFWorkbook();

		if (file != null)
			this.file = File.createTempFile(file, ".xlsx");
		else
			this.file = File.createTempFile("dummy", ".xlsx");

		this.sheet = workbook.createSheet(sheetName);

		this.sheetNameSet.put(sheetName, 0);

	}

	/* This method will write headings of columns */
	public void writeHeaderColumns(String[] headers) {

		if (headers == null) {
			throw new IExcelFileGeneratorException("Headers Cannot be null");
		}
		HSSFRow headerRow = sheet.createRow(rowCount++);

		int cellCount = 0;

		for (String header : headers) {
			headerRow.createCell(cellCount++).setCellValue(header);
		}
		rowCount++;
		columnCount = headers.length;
	}

	/* This method will write data row wise 
	 * Pass object of cellstyle for styling the columns*/
	public void write(Object source, String[] headers, CellStyle style) {
		if (source == null) {
			throw new IExcelFileGeneratorException("Object Should not be null");
		}
		if (headers == null) {
			throw new IExcelFileGeneratorException("Headers Should not be null");
		}
		if (style == null) {
			style = getWorkbook().createCellStyle();
		}
		HSSFRow aRow = sheet.createRow(rowCount++);

		int cellCount = 0;

		for (String header : headers) {
			try {
				Method getMethod = cache.getGetMethod(source, header);

				if (getMethod.getReturnType().toString().equalsIgnoreCase("date")
						&& getMethod.invoke(source) != null) {
					aRow.createCell(cellCount++).setCellValue((Date) getMethod.invoke(source));
					aRow.getCell(cellCount - 1).setCellStyle(style);
				} else if (getMethod.invoke(source) != null) {
					aRow.createCell(cellCount++).setCellValue(getMethod.invoke(source).toString());
					aRow.getCell(cellCount - 1).setCellStyle(style);
				} else {
					aRow.createCell(cellCount++).setCellValue("");
				}
				aRow.getCell(cellCount - 1).setCellStyle(style);
			} catch (Exception e) {
				throw new IExcelFileGeneratorException(e.getMessage());
			}
		}
	}

	/* This Method is to get File for excel*/
	public File getFile() throws IOException {
		for (int columnNumber = 0; columnNumber < columnCount; columnNumber++) {

			sheet.autoSizeColumn(columnNumber);
		}
		try(FileOutputStream fileOut = new FileOutputStream(file);) {
			
			workbook.write(fileOut);
		} catch (Exception e) {
		}
		return file;
	}

	/* By this method u can create various objects of Apache POI*/
	public HSSFWorkbook getWorkbook() {
		return workbook;
	}

	public Integer getRowCount() {
		return rowCount;
	}

	public void groupRows(int firstRowNumber) {
		sheet.groupRow(firstRowNumber, rowCount - 1);
	}

	public void addBlankRow() {
		rowCount++;
	}

	/* This method allows you to create new sheet within the workbook*/
	public void createNewSheet(String sheetName) {
		sheetNameSet.put(this.sheet.getSheetName(), rowCount);
		Boolean sheetNameCheck = sheetNameSet.containsKey(sheetName);
		if (sheetNameCheck) {
			throw new IExcelFileGeneratorException(String.format("Sheet with %s name already exists", sheetName));
		}
		sheetNameSet.put(sheetName, 0);
		this.sheet = workbook.createSheet(sheetName);
		this.rowCount = sheetNameSet.get(sheetName);
	}

	/* This method will help you to switch between sheets */
	public void changeSheet(String sheetName) {
		sheetNameSet.put(sheetName, rowCount);
		this.sheet = workbook.getSheet(sheetName);
		if (sheet == null) {
			throw new IExcelFileGeneratorException(String.format("Sheet with %s name does not exists", sheetName));
		}
		this.rowCount = sheetNameSet.get(sheetName);
	}
	
	public String currentSheetName() {
		return this.sheet.getSheetName();
	}
}
