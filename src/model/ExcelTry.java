/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package model;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import javax.swing.JFileChooser;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author luod1
 */
public class ExcelTry {

    public static File chooseFile;

    public static void main(String[] args) throws IOException {

        JFileChooser chooser = new JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel file", "xls", "xlsx");
        chooser.setFileFilter(filter);

        int returnVal = chooser.showOpenDialog(null);
        if (returnVal != JFileChooser.APPROVE_OPTION) {
            return;
        }

        chooseFile = chooser.getSelectedFile();

        FileInputStream inputstream = new FileInputStream(chooseFile);

        XSSFWorkbook workbook = new XSSFWorkbook(inputstream);

        XSSFSheet sheet = workbook.getSheetAt(0);//frist sheet

        int rows = sheet.getLastRowNum();
        int cols = sheet.getRow(1).getLastCellNum();

        List<String> name = new ArrayList();
        //name.add("程修远");
        name.add("林颖");
        name.add("刘畅");
        name.add("王媛");
        name.add("李发金");
        name.add("肖智涌");
        name.add("林梅");
        name.add("林南");
        name.add("陈超");
        name.add("陈忱");
        name.add("欧凡");
        // search by name
//        for (int r = 0; r <= rows; r++) {
//            XSSFRow row = sheet.getRow(r);
//            for (int c = 0; c < cols; c++) {
//                XSSFCell cell = row.getCell(c);
//                XSSFCell cellx = row.getCell(0);
//                String s = cellx.getStringCellValue();
//                
//                if (s.contains("19")) {
//                    
//                    String ss = cell.getStringCellValue();
//                    XSSFRow rowss = sheet.getRow(0);
//                    XSSFCell cellss = rowss.getCell(c);
//                    
//                    XSSFRow rowss1 = sheet.getRow(1);
//                    XSSFCell cellss1 = rowss1.getCell(c);
//                    
//                    for(String n : name){
//                        if(ss.contains(n)){
//                            System.out.println(cellss.getStringCellValue());
//                            System.out.println(cellss1.getStringCellValue());
//                            System.out.println(ss);
//                        }
//                    }
//                   
//
//                }
//
//            }
//        }
        ArrayList<String> mondayM = new ArrayList<>();
        ArrayList<String> mondayP = new ArrayList<>();
        //search by day

        ArrayList<String> onlyn = new ArrayList<>();
        ArrayList<String> onlyp = new ArrayList<>();

        for (int r = 0; r <= rows; r++) {
            XSSFRow row = sheet.getRow(0);

            for (int c = 0; c < cols; c++) {
                XSSFCell cell = row.getCell(c);
                String s = cell.getStringCellValue();
                if (s.contains("星期一")) {
                    XSSFRow rowss = sheet.getRow(r);
                    XSSFCell cellss = rowss.getCell(c);
                    String ss = cellss.getStringCellValue();

                    XSSFRow rowsss = sheet.getRow(1);
                    XSSFCell cellsss = rowsss.getCell(c);
                    String sss = cellsss.getStringCellValue();
                    XSSFCell cells = rowss.getCell(0);
                    String ssss = cells.getStringCellValue();

                    if (sss.contains("一") || sss.contains("二") || sss.contains("三") || sss.contains("四")) {
                        for (String i : name) {

                            if (ss.contains(i)) {
                                String m12 = ssss + "/" + sss + "节/" + i;
                                //System.out.println(sss);
                                //System.out.println(ss);
                                mondayM.add(m12);
                                onlyn.add(i);
                            }
                        }
                    }
                    if (sss.contains("五") || sss.contains("六") || sss.contains("七") || sss.contains("八")) {
                        for (String i : name) {

                            if (ss.contains(i)) {
                                String m12 = ssss + "/" + sss + "节/" + ss;
                                //System.out.println(sss);
                                //System.out.println(ss);
                                mondayP.add(m12);
                                onlyp.add(i);
                            }
                        }
                    }

                    //System.out.println(cellsss.getStringCellValue());
                }
            }
        }

        LinkedHashSet<String> hashSet = new LinkedHashSet<>(onlyn);
        ArrayList<String> listWithoutDuplicates = new ArrayList<>(hashSet);

        //System.out.println(onlyn);
        //System.out.println(listWithoutDuplicates);
        System.out.println(mondayM.toString());
        System.out.println("周一："+mondayP);

        //System.out.println(onlyp);
    }
}
