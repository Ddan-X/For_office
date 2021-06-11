/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package GUI;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author luod1
 */
public class ReadExcel {

    public static File chooseFile;

    public ArrayList<String> dayM12 = new ArrayList<>();
    public ArrayList<String> dayM34 = new ArrayList<>();

    public ArrayList<String> dayP56 = new ArrayList<>();
    public ArrayList<String> dayP78 = new ArrayList<>();

    public ArrayList<String> onlynameM12 = new ArrayList<>();
    public ArrayList<String> onlynameM34 = new ArrayList<>();

    public ArrayList<String> onlynameP56 = new ArrayList<>();
    public ArrayList<String> onlynameP78 = new ArrayList<>();

    public ArrayList<String> levels = new ArrayList<>();
    public ArrayList<String> levelName = new ArrayList<>();

    public ArrayList<String> dayML12 = new ArrayList<>();
    public ArrayList<String> dayPL56 = new ArrayList<>();
    public ArrayList<String> dayML34 = new ArrayList<>();
    public ArrayList<String> dayPL78 = new ArrayList<>();

    public ArrayList<String> onlynameML12 = new ArrayList<>();
    public ArrayList<String> onlynamePL56 = new ArrayList<>();

    public ArrayList<String> onlynameML34 = new ArrayList<>();
    public ArrayList<String> onlynamePL78 = new ArrayList<>();

    public ArrayList<String> getDayM12() {
        return dayM12;
    }

    public ArrayList<String> getDayM34() {
        return dayM34;
    }

    public ArrayList<String> getDayP56() {
        return dayP56;
    }

    public ArrayList<String> getDayP78() {
        return dayP78;
    }

    public ArrayList<String> getOnlynameM12() {
        return onlynameM12;
    }

    public ArrayList<String> getOnlynameM34() {
        return onlynameM34;
    }

    public ArrayList<String> getOnlynameP56() {
        return onlynameP56;
    }

    public ArrayList<String> getOnlynameP78() {
        return onlynameP78;
    }

    public ArrayList<String> getDayML12() {
        return dayML12;
    }

    public ArrayList<String> getDayPL56() {
        return dayPL56;
    }

    public ArrayList<String> getDayML34() {
        return dayML34;
    }

    public ArrayList<String> getDayPL78() {
        return dayPL78;
    }

    public ArrayList<String> getOnlynameML12() {
        return onlynameML12;
    }

    public ArrayList<String> getOnlynamePL56() {
        return onlynamePL56;
    }

    public ArrayList<String> getOnlynameML34() {
        return onlynameML34;
    }

    public ArrayList<String> getOnlynamePL78() {
        return onlynamePL78;
    }

    

    public ArrayList<String> getLevels() {
        return levels;
    }

    public ArrayList<String> getLevelName() {
        return levelName;
    }

    public FileInputStream getFile() throws FileNotFoundException {
        JFileChooser chooser = new JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel file", "xls", "xlsx");
        chooser.setFileFilter(filter);

        int returnVal = chooser.showOpenDialog(null);

        chooseFile = chooser.getSelectedFile();

        FileInputStream inputstream = new FileInputStream(chooseFile);
        return inputstream;

    }

    public XSSFSheet readExcel(FileInputStream inputstream) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook(inputstream);

        XSSFSheet sheet = workbook.getSheetAt(0);//frist sheet

        return sheet;

    }

    public void SearchDay(ArrayList<String> name, int rows, int cols, XSSFSheet sheet, String day) {

        for (int r = 0; r <= rows; r++) {
            XSSFRow row = sheet.getRow(0);

            for (int c = 0; c < cols; c++) {
                XSSFCell cell = row.getCell(c);
                String s = cell.getStringCellValue();
                if (s.contains(day)) {
                    XSSFRow rowss = sheet.getRow(r);
                    XSSFCell cellss = rowss.getCell(c);
                    String ss = cellss.getStringCellValue();

                    XSSFRow rowsss = sheet.getRow(1);
                    XSSFCell cellsss = rowsss.getCell(c);
                    String sss = cellsss.getStringCellValue();
                    XSSFCell cells = rowss.getCell(0);
                    String ssss = cells.getStringCellValue();

                    if (sss.contains("一") || sss.contains("二")) {
                        for (String i : name) {
                            if (ss.contains(i)) {
                                String m12 = ssss + "/" + sss + "节/" + i; // ss 字符串 包括姓名和课程名， i 是姓名
                                dayM12.add(m12);

                                onlynameM12.add(i);
                            }

                        }

                    }
                    if (sss.contains("三") || sss.contains("四")) {
                        for (String i : name) {
                            if (ss.contains(i)) {
                                String m12 = ssss + "/" + sss + "节/" + i; // ss 字符串 包括姓名和课程名， i 是姓名
                                dayM34.add(m12);

                                onlynameM34.add(i);
                            }

                        }

                    }
                    if (sss.contains("五") || sss.contains("六") ) {
                        for (String i : name) {

                            if (ss.contains(i)) {
                                String m12 = ssss + "/" + sss + "节/" + i;
                                //System.out.println(sss);
                                //System.out.println(ss);
                                dayP56.add(m12);
                                onlynameP56.add(i);
                            }

                        }

                    }
                    if (sss.contains("七") || sss.contains("八")) {
                        for (String i : name) {

                            if (ss.contains(i)) {
                                String m12 = ssss + "/" + sss + "节/" + i;
                                //System.out.println(sss);
                                //System.out.println(ss);
                                dayP78.add(m12);
                                onlynameP78.add(i);
                            }

                        }

                    }
                }
            }

            //System.out.println(cellsss.getStringCellValue());
        }
    }

    public void SearchLevel(ArrayList<String> name, int rows, int cols, XSSFSheet sheet, String level) {
        for (int r = 0; r <= rows; r++) {
            XSSFRow row = sheet.getRow(r);
            for (int c = 0; c < cols; c++) {
                XSSFCell cell = row.getCell(c);
                XSSFCell cellx = row.getCell(0);
                String s = cellx.getStringCellValue();

                if (s.contains(level)) {

                    String ss = cell.getStringCellValue(); //ss 字符串 包括姓名和课程名
                    XSSFRow rowss = sheet.getRow(0);
                    XSSFCell cellss = rowss.getCell(c);

                    XSSFRow rowss1 = sheet.getRow(1);
                    XSSFCell cellss1 = rowss1.getCell(c);

                    for (String n : name) {
                        if (ss.contains(n)) {
                            String l = cellss.getStringCellValue() + "/" + cellss1.getStringCellValue() + "/" + n;
                            levels.add(l);
                            levelName.add(n);
                        }
                    }

                }

            }
        }

    }

    public void SearchDayLevel(ArrayList<String> name, int rows, int cols, XSSFSheet sheet, String day, String level) {

        for (int r = 0; r <= rows; r++) {
            XSSFRow row = sheet.getRow(0);

            for (int c = 0; c < cols; c++) {
                XSSFCell cell = row.getCell(c);
                String s = cell.getStringCellValue();
                if (s.contains(day)) {
                    XSSFRow rowss = sheet.getRow(r);
                    XSSFCell cellss = rowss.getCell(c);
                    String ss = cellss.getStringCellValue();

                    XSSFRow rowsss = sheet.getRow(1);
                    XSSFCell cellsss = rowsss.getCell(c);
                    String sss = cellsss.getStringCellValue();
                    XSSFCell cells = rowss.getCell(0);
                    String ssss = cells.getStringCellValue();

                    if (ssss.contains(level)) {

                        if (sss.contains("一") || sss.contains("二") ) {
                            name.stream().filter((i) -> (ss.contains(i))).forEachOrdered((i) -> {
                                String m12 = ssss + "/" + sss + "节/" + i;
                                //System.out.println(sss);
                                //System.out.println(ss);
                                dayML12.add(m12);
                                onlynameML12.add(i);
                            });
                        }
                        if (sss.contains("三") || sss.contains("四")) {
                            name.stream().filter((i) -> (ss.contains(i))).forEachOrdered((i) -> {
                                String m12 = ssss + "/" + sss + "节/" + i;
                                //System.out.println(sss);
                                //System.out.println(ss);
                                dayML34.add(m12);
                                onlynameML34.add(i);
                            });
                        }
                        if (sss.contains("五") || sss.contains("六") ) {
                            name.stream().filter((i) -> (ss.contains(i))).forEachOrdered((i) -> {
                                String m12 = ssss + "/" + sss + "节/" + i;
                                //System.out.println(sss);
                                //System.out.println(ss);
                                dayPL56.add(m12);
                                onlynamePL56.add(i);
                            });
                        }
                        if ( sss.contains("七") || sss.contains("八")) {
                            name.stream().filter((i) -> (ss.contains(i))).forEachOrdered((i) -> {
                                String m12 = ssss + "/" + sss + "节/" + i;
                                //System.out.println(sss);
                                //System.out.println(ss);
                                dayPL78.add(m12);
                                onlynamePL78.add(i);
                            });
                        }
                    }

                    //System.out.println(cellsss.getStringCellValue());
                }
            }
        }
    }

    public ArrayList<String> deleteDuplicates(ArrayList<String> arr) {
        LinkedHashSet<String> hashSet = new LinkedHashSet<>(arr);
        ArrayList<String> listWithoutDuplicates = new ArrayList<>(hashSet);
        return listWithoutDuplicates;
    }

}
