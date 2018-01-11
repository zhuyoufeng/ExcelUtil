package com.sargeraswang.util.ExcelUtil;

import org.apache.poi.util.IOUtils;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.*;

public class TestExportBean {

    @Test
    public void exportXls() throws IOException {

        List<ExcelHeaderCell> excelHeaderCellList = new ArrayList<>();
        excelHeaderCellList.add(ExcelHeaderCell.createExcelHeaderCell("姓名", "productType", 0, 1, 0, 0));
        excelHeaderCellList.add(ExcelHeaderCell.createExcelHeaderCell("年龄", "productName", 0, 1, 1, 1));
        excelHeaderCellList.add(ExcelHeaderCell.createExcelHeaderCell("班级", "totalGuideFundInvested", 0, 1, 2, 2));
        excelHeaderCellList.add(ExcelHeaderCell.createExcelHeaderCell("爱好", "totalSocialCapitalInvested", 0, 1, 3, 3));

        excelHeaderCellList.add(ExcelHeaderCell.createExcelHeaderCell("随便写写1", 0, 0, 4, 5));
        excelHeaderCellList.add(ExcelHeaderCell.createExcelHeaderCell("001", "guideFundInvested", 1, 4));
        excelHeaderCellList.add(ExcelHeaderCell.createExcelHeaderCell("002", "actualGuideFundInvested", 1, 5));

        excelHeaderCellList.add(ExcelHeaderCell.createExcelHeaderCell("随便写写2", 0, 0, 6, 7));
        excelHeaderCellList.add(ExcelHeaderCell.createExcelHeaderCell("001", "socialCapitalInvested", 1, 6));
        excelHeaderCellList.add(ExcelHeaderCell.createExcelHeaderCell("002", "actualSocialCapitalInvested", 1, 7));

        excelHeaderCellList.add(ExcelHeaderCell.createExcelHeaderCell("呵呵哒", "guideFundPercentage", 0, 1, 8, 8));
        excelHeaderCellList.add(ExcelHeaderCell.createExcelHeaderCell("哒哒呵", "approvedApplications", 0, 1, 9, 9));


        java.io.File file = new File("D:/test/temp1.xls");

        java.io.OutputStream outputStream = null;
        try {
            file.createNewFile();
            outputStream = new FileOutputStream(file);
            ExcelExportUtil.exportExcel("学生", "学生汇总表", excelHeaderCellList, 10, new ArrayList<>(), outputStream, "yyyy-MM-dd");
            outputStream.flush();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            IOUtils.closeQuietly(outputStream);
        }
    }
}
