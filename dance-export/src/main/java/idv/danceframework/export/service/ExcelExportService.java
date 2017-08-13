package idv.danceframework.export.service;

import java.lang.reflect.InvocationTargetException;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import idv.danceframework.core.ContextWrapper;
import idv.danceframework.core.SheetContext;

public class ExcelExportService {
	public static void writeDataToSheet(SheetContext context) throws Exception {

		checkRequiredInput(context);

		// write title row
		writeTitleRow(context);

		// write each row data
		writeEachRowData(context);
	}

	private static void checkRequiredInput(SheetContext context) {

		if (context == null) {
			throw new NullPointerException("Context can not be null.");
		}

		if (context.getSheet() == null) {
			throw new NullPointerException("Sheet is null !!! Please init sheet first.");
		}
	}

	private static void writeTitleRow(SheetContext context) {
		List<ContextWrapper> contextList = context.getContextList();
		if (CollectionUtils.isNotEmpty(contextList)) {

			HSSFSheet sheet = context.getSheet();

			Row currentRow = getRow(sheet,context.getCruuentRow());

			boolean hasTitle = false;

			int cellStart = 0;

			for (ContextWrapper cw : contextList) {
				if (cw.getTitle() != null) {
					if (!hasTitle) {
						hasTitle = true;
					}
					// create title cell
					Cell c = currentRow.createCell(cellStart);
					c.setCellValue(cw.getTitle());
				}
				cellStart++;
			}

			// if has title , new data will insert next row
			if (hasTitle) {
				context.addRow();
			}
		} else {
			throw new NullPointerException("Must define context wrapper !!!");
		}
	}

	private static Row getRow(HSSFSheet sheet,int rowIndex){
		Row row = sheet.getRow(rowIndex);
		if(row == null){
			row = sheet.createRow(rowIndex);
		}
		return row;
	}
	
	private static void writeEachRowData(SheetContext context)
			throws IllegalAccessException, InvocationTargetException, NoSuchMethodException {

		// for each data
		List list = context.getDataList();
		if (CollectionUtils.isNotEmpty(list)) {

			HSSFSheet sheet = context.getSheet();

			Row currentRow = getRow(sheet,context.getCruuentRow());

			int cellStart = 0;

			for (Object obj : list) {
				// restart cell index
				cellStart = 0;
				List<ContextWrapper> contextList = context.getContextList();
				for (ContextWrapper cw : contextList) {
					String fieldName = cw.getField();
					if (fieldName != null) {
						Object fieldValue = PropertyUtils.getProperty(obj, fieldName);
						if (fieldValue == null) {
							throw new IllegalArgumentException(
									obj.getClass().getSimpleName() + " do not has field name : " + fieldName);
						} else {
							// write data cell
							Cell cell = currentRow.createCell(cellStart);
							writeCellValue(cell, fieldValue);
						}
					}
					cellStart++;
				}
				context.addRow();
				currentRow = getRow(sheet,context.getCruuentRow());
			}

		} else {
			throw new NullPointerException("No data !!!");
		}
	}

	private static void writeCellValue(Cell cell, Object fieldValue) {

		if (fieldValue instanceof Boolean) {
			cell.setCellValue((Boolean) fieldValue);
		} else if (fieldValue instanceof Calendar) {
			cell.setCellValue((Calendar) fieldValue);
		} else if (fieldValue instanceof Date) {
			cell.setCellValue((Date) fieldValue);
		} else if (fieldValue instanceof Double) {
			cell.setCellValue((Double) fieldValue);
		} else {
			cell.setCellValue(fieldValue.toString());
		}
	}
}
