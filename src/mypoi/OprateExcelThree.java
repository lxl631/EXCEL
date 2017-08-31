package mypoi;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.concurrent.ConcurrentHashMap;


public class OprateExcelThree {
 

    public static void main(String[] args) throws Exception {
        invokeTheTask();
    }

    static void invokeTheTask() {
        String path = "C:\\Users\\Administrator\\Desktop\\客户类-系统筛选未加工数据";
        File file = new File(path);
        File[] files = file.listFiles();
        int index = 1;
        for (File f : files) {
            if (f.isDirectory()) {
                String fileName = f.getName();
                String filePath = path + "\\" + fileName;
                new Thread(new RunTaskThree(filePath, index++)).start();
            }
        }
    }


}


class RunTaskThree implements Runnable {
    private String sourcePath = "";
    File sourceFile = null;
    File targetFile = null;
    int index = 0;

    public RunTaskThree(String sourcePath, int index) {
        this.index = index;
        this.sourcePath = sourcePath;
        sourceFile = new File(sourcePath + "/多个客户同一住所地.xls");
        String parentName = sourceFile.getParentFile().getName();

        String targetPath = "C:\\Users\\Administrator\\Desktop\\target1\\" + parentName;

        File parentFile = new File(targetPath);
        if (!parentFile.exists())
            parentFile.mkdirs();
        targetFile = new File(targetPath + "\\多个客户同一住所地.xls");
    }

    @Override
    public void run() {
        try {
            readFile();
        } catch (Exception e) {
            System.err.println("执行" + this.sourcePath + "，时候出现异常，原因：" + e.toString());
            e.printStackTrace();
        }
    }

    void readFile() throws Exception {
        NPOIFSFileSystem fs = new NPOIFSFileSystem(sourceFile);
        ConcurrentHashMap map = new ConcurrentHashMap();
        long begint = System.currentTimeMillis();

        try (HSSFWorkbook wb = new HSSFWorkbook(fs.getRoot(), true)) {

            System.err.println("线程：" + index + "开始执行，对应文件：" + this.sourcePath + "，使用时间：" + (System.currentTimeMillis() - begint) / 1000 + "S");


            SXSSFWorkbook wbnew = new SXSSFWorkbook(-1); // turn off auto-flushing and accumulate all rows in memory
            Sheet sh = wbnew.createSheet();
            FileOutputStream out = new FileOutputStream(targetFile);
            // 创建单元格样式
            CellStyle style = wbnew.createCellStyle();
            HSSFFont font = wb.createFont();
            font.setFontName("Arial");
            style.setFont(font);


            HSSFSheet sheet = wb.getSheetAt(0);
            int firstRowNum = sheet.getFirstRowNum();
            int lastRowNum = sheet.getLastRowNum();
            int writeIndex = 0;
            for (int i = firstRowNum; i <= lastRowNum; i++) {
                HSSFRow row = sheet.getRow(i);

                HSSFCell workPlace = row.getCell(5);
                HSSFCell cellIdCard = row.getCell(8);
                HSSFCell workPlaceOrHome = row.getCell(18);
                String seperator = "--";
                String key = workPlace + seperator + cellIdCard + seperator + workPlaceOrHome;

                if (!map.containsKey(key)) {
                    int rownum = writeIndex;
                    Row rownew = sh.createRow(rownum);
                    int startCell = row.getFirstCellNum();
                    int endCell = row.getLastCellNum();
                    for (int j = startCell; j < endCell; j++) {
                        HSSFCell source = row.getCell(j);
                        Cell cell = rownew.createCell(j);
                        cell.setCellStyle(style);
                        cell.setCellValue(source.getStringCellValue());
                    }

                    if (rownum % 100 == 0) {
                        ((SXSSFSheet) sh).flushRows(100); // retain 100 last rows and flush all others
                    }
                    writeIndex++;
                    map.put(key, "");
                }
            }
            wbnew.write(out);
            out.close();
            wbnew.dispose();
            fs.close();
        }


        System.err.println("--------------线程：" + index + "执行完毕，对应路径：" + this.sourcePath + "，使用时间：" + (System.currentTimeMillis() - begint) / 1000 + "S" + "，文件大小：" + targetFile.length() / 1024 / 1024);

    }

}