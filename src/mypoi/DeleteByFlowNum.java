package mypoi;


import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;

public class DeleteByFlowNum {


    public static void main(String[] args) {

        invokeTheTask();

    }


    static void invokeTheTask() {
        String path = "C:\\Users\\Administrator\\Desktop\\大额漏报待查－已处理\\";
        File file = new File(path);
        File[] files = file.listFiles();
        int index = 1;
        for (File f : files) {
            File[] files1 = f.listFiles();
            for (File f1 : files1) {
                if (!f1.isDirectory()) {
                    String tarPath = f1.getPath();
//                new Thread(new FlowNumTask(tarPath, index++)).start();
                    new FlowNumTask(tarPath, index++).run();
                }

            }
        }
    }


}


class FlowNumTask implements Runnable {
    private String sourcePath = "";
    File sourceFile = null;
    File targetFile = null;
    int index = 0;

    public FlowNumTask(String sourcePath, int index) {

        this.index = index;
        this.sourcePath = sourcePath;
        sourceFile = new File(sourcePath);

        String fileName = sourceFile.getName();

        String parentName = sourceFile.getParentFile().getName();

        String parentPath = "C:\\Users\\Administrator\\Desktop\\target1\\";
        File pFile = new File(parentPath + parentName);
        if (!pFile.exists())
            pFile.mkdirs();
        targetFile = new File(pFile.getPath() + File.separator + fileName);

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
        System.err.println("执行：" + this.sourcePath + "--------------------");
        NPOIFSFileSystem fs = new NPOIFSFileSystem(sourceFile);
        ConcurrentHashMap fundSourceMap = new ConcurrentHashMap();
        ConcurrentHashMap flowNumMap = new ConcurrentHashMap();
        ConcurrentHashMap removeRowNum = new ConcurrentHashMap();

        try (HSSFWorkbook wb = new HSSFWorkbook(fs.getRoot(), true)) {

            SXSSFWorkbook wbnew = new SXSSFWorkbook(-1); // turn off auto-flushing and accumulate all rows in memory
            Sheet sh = wbnew.createSheet();
            FileOutputStream out = new FileOutputStream(targetFile);


            HSSFSheet sheet = wb.getSheetAt(0);
            int firstRowNum = sheet.getFirstRowNum();
            int lastRowNum = sheet.getLastRowNum();
            for (int i = firstRowNum; i <= lastRowNum; i++) {
                HSSFRow row = sheet.getRow(i);
                HSSFCell fundSource = row.getCell(17);
                HSSFCell flowNum = row.getCell(21);
                String cellValue = fundSource.toString();
                String flowNumVal = flowNum.toString();
                if (cellValue.contains("冲账：")) {
                    String[] vals = cellValue.split("流水号：");
                    String flowNumSource = vals[1];
                    fundSourceMap.put(flowNumSource, i + "");
                }
                flowNumMap.put(flowNumVal, i + "");
            }


            Set fundES = fundSourceMap.entrySet();
            Iterator fundItr = fundES.iterator();
            while (fundItr.hasNext()) {
                Map.Entry entry = (Map.Entry) fundItr.next();
                String key = (String) entry.getKey();
                String rowNum = (String) entry.getValue();
                if (flowNumMap.containsKey(key)) {
                    removeRowNum.put(rowNum, rowNum);
                    String flowRowNum = (String) flowNumMap.get(key);
                    removeRowNum.put(flowRowNum, flowRowNum);
                    System.err.println("\t\t\t\t\t第 " + (Integer.valueOf(rowNum) + 1) + "行，冲账号：" + key + "，对应流水号在第" + (Integer.valueOf(flowRowNum) + 1) + "行");
                } else {
                    System.err.println("\t\t\t\t\t\t\t\t\t\t第：" + (Integer.valueOf(rowNum) + 1) + "行，冲账号：" + key + " 不存在对应数据");
                }
            }


            int writeIndex = 0;
            for (int i = firstRowNum; i <= lastRowNum; i++) {
                HSSFRow row = sheet.getRow(i);
                String is = i + "";
                if (!removeRowNum.containsKey(is)) {
                    int rownum = writeIndex;
                    Row rownew = sh.createRow(rownum);
                    int startCell = row.getFirstCellNum();
                    int endCell = row.getLastCellNum();
                    for (int j = startCell; j < endCell; j++) {
                        HSSFCell source = row.getCell(j);
                        Cell cell = rownew.createCell(j);
//                        cell.setCellValue(source.getStringCellValue());
                        cell.setCellValue(source.toString());
                    }

                    if (rownum % 100 == 0) {
                        ((SXSSFSheet) sh).flushRows(100); // retain 100 last rows and flush all others
                    }
                    writeIndex++;
                }
            }


            wbnew.write(out);
            out.close();
            wbnew.dispose();
            fs.close();
        }

    }

}

