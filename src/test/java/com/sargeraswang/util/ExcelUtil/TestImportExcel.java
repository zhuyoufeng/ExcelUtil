/**
 * @author SargerasWang
 */
package com.sargeraswang.util.ExcelUtil;

import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.Arrays;
import java.util.Collection;
import java.util.Date;
import java.util.Map;

/**
 * 测试导入Excel 97/2003
 */
public class TestImportExcel {

    @Test
    public void importXls() throws FileNotFoundException {
        File f = new File("src/test/resources/test.xls");

        ExcelLogs logs = new ExcelLogs();
        Collection<Map> importExcel = ExcelImportUtil.importExcel(f, Map.class, 0, logs, 0);

        for (Map m : importExcel) {
            System.out.println(m);
        }
    }

    @Test
    public void importXlsx() throws FileNotFoundException {
        File f = new File("src/test/resources/test.xlsx");

        ExcelLogs logs = new ExcelLogs();
        Collection<Map> importExcel = ExcelImportUtil.importExcel(f, Map.class, 0, logs, 0);

        for (Map m : importExcel) {
            System.out.println(m);
        }
    }

    @Test
    public void importBean() throws FileNotFoundException {
        File f = new File("src/test/resources/test2.xlsx");

        ExcelLogs logs = new ExcelLogs();
        Collection<Student> importExcel = ExcelImportUtil.importExcel(f, Student.class, 0, logs, 2);

        for (Student m : importExcel) {
            System.out.println(m);
        }
    }


}
