package idv.danceframework.core;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;

public class SheetContext {
	private HSSFSheet sheet;
	private List<ContextWrapper> contextList = new ArrayList<ContextWrapper>();
	private List dataList = new ArrayList();
	private Class dataClass;
	private int cruuentRow = 0;

	public HSSFSheet getSheet() {
		return sheet;
	}

	public void setSheet(HSSFSheet sheet) {
		this.sheet = sheet;
	}

	public List<ContextWrapper> getContextList() {
		return contextList;
	}

	public void setContextList(List<ContextWrapper> contextList) {
		this.contextList = contextList;
	}

	public List getDataList() {
		return dataList;
	}

	public void setDataList(List dataList) {
		this.dataList = dataList;
	}

	public Class getDataClass() {
		return dataClass;
	}

	public void setDataClass(Class dataClass) {
		this.dataClass = dataClass;
	}

	public int getCruuentRow() {
		return cruuentRow;
	}

	public void setCruuentRow(int cruuentRow) {
		this.cruuentRow = cruuentRow;
	}

	public void addRow() {
		this.cruuentRow++;
	}
}
