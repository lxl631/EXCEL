package mypoi;

import java.io.File;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 * Created by 李小龙 on 2017/8/26 8:50.
 */
public class OprateExcel {

	public static void main(String[] args) throws Exception {
		invokeTheTask();
	}

	static void invokeTheTask() {
		String path = "C:\\Users\\Administrator\\Desktop\\客户类-系统筛选未加工数据";
		File file = new File(path);
		File[] files = file.listFiles();
		for (File f : files) {
			if (f.isDirectory()) {
				String fileName = f.getName();
				String filePath = path + "\\" + fileName;
				new Thread(new RunTask(filePath)).start();
			}
		}

	}

}

class ExcelBean {

	private Row row;

	public Row getRow() {
		return row;
	}

	public void setRow(Row row) {
		this.row = row;
	}

	private int endRow;

	public int getEndRow() {
		return endRow;
	}

	public void setEndRow(int endRow) {
		this.endRow = endRow;
	}

	public int getNowRow() {
		return nowRow;
	}

	public void setNowRow(int nowRow) {
		this.nowRow = nowRow;
	}

	private int nowRow;

}

class RunTask implements Runnable {

	private String sourcePath = "";

	File sourceFile = null;
	File targetFile = null;

	public RunTask(String sourcePath) {
		this.sourcePath = sourcePath;
		sourceFile = new File(sourcePath + "/多个客户同一个联系方式.xls");
		String parentName = sourceFile.getParentFile().getName();

		String targetPath = "C:\\Users\\Administrator\\Desktop\\target\\" + parentName;

		File parentFile = new File(targetPath);
		if (!parentFile.exists())
			parentFile.mkdirs();
		targetFile = new File(targetPath + "\\多个客户同一个联系方式.xls");

	}

	@Override
	public void run() {

		try {
			readFile();

		} catch (Exception e) {
			System.err.println("执行" + this.sourcePath + "，时候出现异常，原因：" + e.toString());
		}
	}

	void readFile() throws Exception {

		NPOIFSFileSystem fs = new NPOIFSFileSystem(sourceFile);

		ConcurrentHashMap map = new ConcurrentHashMap();

		try (HSSFWorkbook wb = new HSSFWorkbook(fs.getRoot(), true)) {
			HSSFSheet sheet = wb.getSheetAt(0);
			int firstRowNum = sheet.getFirstRowNum();
			int lastRowNum = sheet.getLastRowNum();
			for (int i = firstRowNum; i <= lastRowNum; i++) {
				HSSFRow row = sheet.getRow(i);
				HSSFCell cellTel = row.getCell(6);
				HSSFCell cellIdCard = row.getCell(8);
				String seperator = "--";
				String key = cellTel + seperator + cellIdCard;

				ExcelBean bean = new ExcelBean();
				bean.setNowRow(i + 1);
				bean.setEndRow(lastRowNum + 1);
				bean.setRow(row);
				if (map.containsKey(key)) {
					List list = (List) map.get(key);
					list.add(bean);
				} else {
					List list = new ArrayList();
					list.add(bean);
					map.put(key, list);
				}
			}
			fs.close();

			Set set = map.entrySet();
			Iterator itr = set.iterator();
			while (itr.hasNext()) {
				Map.Entry entry = (Map.Entry) itr.next();
				List value = (List) entry.getValue();
				if (value.size() > 1) {
					shuffleData(value, wb, targetFile);
				}
			}
			wb.write(targetFile);

			fs.close();

			removeEmpty(targetFile);
		}

		System.err.println(this.sourcePath + "执行完毕。。");

	}

	void shuffleData(List value, HSSFWorkbook wb, File targetFile) throws Exception {
		for (int i = 1; i < value.size(); i++) {
			HSSFSheet sheet = wb.getSheetAt(0);
			ExcelBean bean = (ExcelBean) value.get(i);
			Row row = bean.getRow();
			sheet.removeRow(row);
		}
	}

	void removeEmpty(File targetFile) throws Exception {
		NPOIFSFileSystem fs = new NPOIFSFileSystem((targetFile));
		try (HSSFWorkbook wb = new HSSFWorkbook(fs.getRoot(), true)) {
			HSSFSheet sheet = wb.getSheetAt(0);
			for (int i = 0; i < sheet.getLastRowNum(); i++) {
				Row row = sheet.getRow(i);
				if (isEmpty(row)) {
					sheet.shiftRows(i + 1, sheet.getLastRowNum(), -1);
					i--;
				}
			}
			wb.write((targetFile));
			fs.close();
		}
	}

	boolean isEmpty(Row row) {
		if (row == null)
			return true;

		Cell cell1 = row.getCell(1);
		Cell cell8 = row.getCell(8);
		if (cell1 == null && cell8 == null) {
			return true;
		}

		String v1 = cell1.getStringCellValue();
		String v8 = cell8.getStringCellValue();
		if ((v1 == null || "".equals(v1)) && (v8 == null || "".equals(v8))) {
			return true;
		}
		return false;
	}

}