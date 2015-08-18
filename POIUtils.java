import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

public class POIUtils {

	/**
	 * Helps in obtaining cell for given row, if it does not exits it will create
	 * 
	 * @param row - Row
	 * @param cellNo - cell number
	 * @return Cell
	 * @author Venkata Rajeev Mandava
	 */
	public static Cell getCell(final Row row, final int cellNo) {
		Cell cell = row.getCell(cellNo);
		if (null == cell) {
			cell = row.createCell(cellNo);
		}
		return cell;
	}

	/**
	 * Helps in obtaining row for given sheet, if it does not exits it will create
	 * 
	 * @param sheet - Sheet
	 * @param rowNo - row number
	 * @return Row
	 * @author Venkata Rajeev Mandava
	 */
	public static Row getRow(final Sheet sheet, final int rowNo) {
		Row row = sheet.getRow(rowNo);
		if (null == row) {
			row = sheet.createRow(rowNo);
		}
		return row;
	}

	/**
	 * Helps in obtaining the next available header cell
	 * 
	 * @param sheet - Sheet
	 * @return Cell
	 * @author Venkata Rajeev Mandava
	 */
	public static Cell getNextAvailableHeaderCell(final Sheet sheet) {
		final Row headerRow = getRow(sheet, 0);
		Cell headerCell = null;
		int i = 0;
		do {
			headerCell = getCell(headerRow, i++);
		} while (!StringUtils.isEmpty(headerCell.getStringCellValue()));
		return headerCell;
	}

	/**
	 * Helps in obtaining the next available row
	 * 
	 * @param sheet
	 * @return Row
	 * @author Venkata Rajeev Mandava
	 */
	public static Row getNextAvailableRow(final Sheet sheet) {
		Row row = null;
		int i = 0;
		do {
			row = getRow(sheet, i++);
		} while (!StringUtils.isEmpty(getCell(row, 0).getStringCellValue()));
		return row;
	}

	/**
	 * Helps in populating the column of the given cell with the given data and style
	 * 
	 * @param cell - Cell which will be taken as reference
	 * @param columnData - List<String> is the data that needs to be populated
	 * @param cellStyle - style to be applied
	 * @author Venkata Rajeev Mandava
	 */
	public static void populateColumnOfCell(final Cell cell, final List<String> columnData, final CellStyle cellStyle) {
		final Sheet sheet = cell.getSheet();
		final int columnNumber = cell.getColumnIndex();
		for (int i = cell.getRowIndex(), j = 0; j < columnData.size(); i++, j++) {
			final Cell tempCell = getCell(getRow(sheet, i), columnNumber);
			tempCell.setCellStyle(cellStyle);
			tempCell.setCellValue(columnData.get(i));
		}
		sheet.autoSizeColumn(columnNumber);
	}

	/**
	 * Helps in populating the row of the given Row with the given data and style
	 * 
	 * @param row - Row which will be populated
	 * @param data - List<String> is the data that needs to be populated
	 * @param style - style to be applied
	 * @author Venkata Rajeev Mandava
	 */
	public static void populateRowWithList(final Row row, final List<String> data, final CellStyle style) {
		for (int i = 0; i < data.size(); i++) {
			final Cell tempCell = getCell(row, i);
			tempCell.setCellValue(data.get(i));
			tempCell.setCellStyle(style);
			tempCell.getSheet().autoSizeColumn(i);
		}
	}

	/**
	 * Helps in marking the Named Range with the given details with reference to the given Cell
	 * 
	 * @param cell - Reference of the cell from where the coordinates are considered
	 * @param rows - Number of rows to be covered
	 * @param columns - Number of columns to be covered
	 * @param rangeName - Name to be given for the Named Range
	 * @param skipfirst - Should skip the given Cell row or not
	 * @author Venkata Rajeev Mandava
	 */
	public static void markNamedRange(final Cell cell, final int rows, final int columns, final String rangeName,
			final boolean skipfirst) {
		final CellReference cellReference = new CellReference(cell);
		final String[] cellPosition = cellReference.getCellRefParts();

		final String cellRowRef = cellPosition[1];
		final int startCellRowRef = skipfirst ? Integer.parseInt(cellRowRef) + 1 : Integer.parseInt(cellRowRef);
		final int endCellRowRef = startCellRowRef + (skipfirst ? rows - 1 : rows);

		final String cellColRef = cellPosition[2];
		final int endCellColRef = CellReference.convertColStringToIndex(cellColRef) + columns;
		final String endCellColRefS = CellReference.convertNumToColString(endCellColRef);

		final String reference =
				cell.getSheet().getSheetName() + "!$" + cellColRef + "$" + startCellRowRef + ":$" + endCellColRefS + "$" + endCellRowRef;

		final Workbook workbook = cell.getSheet().getWorkbook();
		Name name = workbook.getName(rangeName);
		if (null == name) {
			name = workbook.createName();
			name.setNameName(rangeName);
		}
		name.setRefersToFormula(reference);
	}

	/**
	 * Helps in obtaining the common style used in dynamic sheets
	 * 
	 * @param workbook
	 * @return CellStyle
	 * @author Venkata Rajeev Mandava
	 */
	public static CellStyle getCellStyle(final Workbook workbook) {
		final CellStyle style = workbook.createCellStyle();
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setRightBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());
		return style;
	}

}
