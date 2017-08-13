package idv.danceframework.example;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import idv.danceframework.core.ContextWrapper;
import idv.danceframework.core.SheetContext;
import idv.danceframework.example.User;
import idv.danceframework.export.service.ExcelExportService;

public class TestCase {

	public static void main(String[] args) {
		FileOutputStream fos;
		POIFSFileSystem fs;
		HSSFWorkbook wb;

		try {
			String filePath = "abc.xls";
			java.io.File f = new java.io.File(filePath);

			fos = new FileOutputStream(f);

			wb = new HSSFWorkbook();
			HSSFSheet sheet = wb.createSheet();

			// for test
			SheetContext context = new SheetContext();
			
			context.setSheet(sheet);
			
			context.getContextList().add(new ContextWrapper("姓名", "name"));
			context.getContextList().add(new ContextWrapper("分數", "grade"));

			List<User> userList = new ArrayList<User>();
			User u = new User();
			u.setName("ABC");
			u.setGrade(100);
			userList.add(u);

			u = new User();
			u.setName("Micheal");
			u.setGrade(50);
			userList.add(u);

			context.setDataList(userList);
			// write data
			ExcelExportService.writeDataToSheet(context);

			// end
			wb.write(fos);
			fos.flush();
			fos.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
