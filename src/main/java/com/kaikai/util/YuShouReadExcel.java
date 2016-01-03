package com.kaikai.util;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

/**
 * Created by Ricky on 2015/12/27.
 */
public class YuShouReadExcel {

//    public final static String EXCEL_PATH = "F:\\todo.xlsx";
    public final static String EXCEL_YING_SHOU_PATH = "E:\\kaikai\\yushou\\yushou.xlsx";
    public final static String EXCEL_ALL_COMPANY_PATH = "E:\\kaikai\\yushou\\allcompany.xlsx";
    public final static String EXCEL_CHONG_FEN_LEI_PATH = "E:\\kaikai\\yushou\\chongfenlei.xlsx";
    public final static String EXCEL_QI_CHU_PATH = "E:\\kaikai\\yushou\\qichu.xlsx";
    public final static String EXCEL_ZHANG_LIN_PATH = "E:\\kaikai\\yushou\\zhanglin.xlsx";
    public final static String OUTPUT_FILE_PERFIX = "F:";
    public static Map<String, Double> positiveNumMap = new HashMap<String, Double>();
    public static Map<String, Double> negativeNumMap = new HashMap<String, Double>();
    public static Map<String, String> companyIdNameMap = new HashMap<String, String>();
    public static Map<String, Double> chongFenLeiMap = new HashMap<String, Double>();
    public static Map<String, Double> qiChuMap = new HashMap<String, Double>();
    public static Map<String, List<Double>> zhangLinMap = new HashMap<String, List<Double>>();
    public static List<String> allCompanies = new ArrayList<String>();
    public static Set<List<String>> resultsForExport = new HashSet<List<String>>();
    public static Map<String, Double> positiveNumProcessedMap = new HashMap<String, Double>();
    public static Map<String, Double> negativeNumProcessedMap = new HashMap<String, Double>();


    public static void parseExcelYinShou() {

        XSSFWorkbook xwb = parseExcel(EXCEL_YING_SHOU_PATH);

        XSSFSheet sheet = xwb.getSheetAt(0);
        XSSFRow row;
        String cell;

        for(int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
//            System.out.println(row.getCell(15) + "\t" + row.getCell(18));
            String idTemp = row.getCell(15).toString().trim();
            double countTemp = Double.parseDouble(row.getCell(18).toString().trim());
            if("".equals(idTemp)) {
                idTemp = "外币评估调整";
            }
            if(countTemp > 0) {
                if(positiveNumMap.get(idTemp) != null) {
                    positiveNumMap.put(idTemp, positiveNumMap.get(idTemp) + countTemp);
                } else {
                    positiveNumMap.put(idTemp, countTemp);
                }
            } else {
                if(negativeNumMap.get(idTemp) != null) {
                    negativeNumMap.put(idTemp, negativeNumMap.get(idTemp) + countTemp);
                } else {
                    negativeNumMap.put(idTemp, countTemp);
                }
            }
        }
    }

    public static void parseExcelAllCompany() {
        XSSFWorkbook xwb = parseExcel(EXCEL_ALL_COMPANY_PATH);

        XSSFSheet sheet = xwb.getSheetAt(0);
        XSSFRow row;
        String cell;
        for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            String idTemp = row.getCell(0).toString().trim();
//            System.out.println(idTemp + '\t' + idTemp.replace(".0", ""));
            String companyNameTemp = row.getCell(1).toString().trim();
            companyIdNameMap.put(idTemp.replace(".0", ""), companyNameTemp);
        }
    }

    public static void parseExcelChongFenLei() {
        XSSFWorkbook xwb = parseExcel(EXCEL_CHONG_FEN_LEI_PATH);

        XSSFSheet sheet = xwb.getSheetAt(0);
        XSSFRow row;
        for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            String idTemp = row.getCell(0).toString().trim();
            String countTemp = row.getCell(1) == null ? "0" : row.getCell(1).toString().trim();
            if(countTemp.equals("0") || countTemp.equals("")) {
                chongFenLeiMap.put(idTemp, 0d);
            } else {
                chongFenLeiMap.put(idTemp, Double.parseDouble(countTemp));
            }
        }
    }

    public static void parseExcelQiChu() {
        XSSFWorkbook xwb = parseExcel(EXCEL_QI_CHU_PATH);
        XSSFSheet sheet = xwb.getSheetAt(0);
        XSSFRow row;
        for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
//            System.out.println(i);
//            System.out.println(row.getCell(0) + "\t" + row.getCell(1));
            String idTemp = row.getCell(0).toString().trim();
            String countTemp = row.getCell(1) == null ? "" : row.getCell(1).toString().trim();
            if(countTemp.equals("0") || countTemp.equals("")) {
                qiChuMap.put(idTemp, 0d);
            } else {
                qiChuMap.put(idTemp, Double.parseDouble(countTemp));
            }
        }
    }


    public static void parseExcelZhangLin() {
        XSSFWorkbook xwb = parseExcel(EXCEL_ZHANG_LIN_PATH);
        XSSFSheet sheet = xwb.getSheetAt(0);
        XSSFRow row;
        for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            String idTemp = row.getCell(0).toString().trim();
            List<Double> countTemps = new ArrayList<Double>();
            for(int j = 1; j <= 4; j++) {
                String countTemp = "0";
                if(!(row.getCell(j) == null)) {
                    countTemp  = row.getCell(j).toString().trim().equals("") ? "0" :  row.getCell(j).toString().trim();
                }
                countTemps.add(Double.parseDouble(countTemp));
            }
            zhangLinMap.put(idTemp, countTemps);
        }
    }

    private static XSSFWorkbook parseExcel(String path) {
        XSSFWorkbook xwb = null;
        try {
            InputStream is = new FileInputStream(path);
            xwb = new XSSFWorkbook(is);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return xwb;

    }

    public static void exportExcel() {

        // 第一步，创建一个webbook，对应一个Excel文件
        HSSFWorkbook wb = new HSSFWorkbook();
        // 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
        HSSFSheet sheet = wb.createSheet("应收");
        // 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制short
        HSSFRow row = sheet.createRow(0);
        // 第四步，创建单元格，并设置值表头 设置表头居中
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 创建一个居中格式

        HSSFCell cell = row.createCell(0);
        cell.setCellValue("预收对象名称");
        cell.setCellStyle(style);
        cell = row.createCell(1);
        cell.setCellValue("期初余额");
        cell.setCellStyle(style);
        cell = row.createCell(2);
        cell.setCellValue("本期增加");
        cell.setCellStyle(style);
        cell = row.createCell(3);
        cell.setCellValue("本期减少");
        cell.setCellStyle(style);
        cell = row.createCell(4);
        cell.setCellValue("重分类");
        cell.setCellStyle(style);
        cell = row.createCell(5);
        cell.setCellValue("调整后减少");
        cell.setCellStyle(style);
        cell = row.createCell(6);
        cell.setCellValue("期末余额");
        cell.setCellStyle(style);


        for(int i = 0; i < allCompanies.size(); i++ ) {
            String companyName = allCompanies.get(i);
            double qiChuYuE = qiChuMap.get(companyName) == null ? 0 : qiChuMap.get(companyName);
            double zengJia = positiveNumProcessedMap.get(companyName) == null ? 0 : positiveNumProcessedMap.get(companyName) * -1;
            double jianShao = negativeNumProcessedMap.get(companyName) == null ? 0 : negativeNumProcessedMap.get(companyName);
            double chongFenLei = chongFenLeiMap.get(companyName) == null ? 0 : chongFenLeiMap.get(companyName);
            List<Double> zhanglinTemp = zhangLinMap.get(companyName);

            double tiaoZhengHouJianShao = jianShao + chongFenLei;
            double qiMouYuE = qiChuYuE + zengJia - tiaoZhengHouJianShao;

            row = sheet.createRow(i + 1);
            row.createCell(0).setCellValue(companyName);
            row.createCell(1).setCellValue(qiChuYuE);
            row.createCell(2).setCellValue(zengJia);
            row.createCell(3).setCellValue(jianShao);
            row.createCell(4).setCellValue(chongFenLei);
            row.createCell(5).setCellValue(tiaoZhengHouJianShao);
            row.createCell(6).setCellValue(qiMouYuE);
        }

        try
        {
            FileOutputStream fout = new FileOutputStream("E:/kaikai/yushou/预收.xls");
            wb.write(fout);
            fout.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }

    }

    public static void listAllCompanies() {
        //转换id为公司名字
        Set<String> pIds = positiveNumMap.keySet();
        for(String s : pIds) {
            if(companyIdNameMap.get(s) != null) {
                positiveNumProcessedMap.put(companyIdNameMap.get(s), positiveNumMap.get(s));
            } else {
                positiveNumProcessedMap.put(s, positiveNumMap.get(s));
            }
        }
        Set<String> nIds = negativeNumMap.keySet();
        for(String s : nIds) {
            if(companyIdNameMap.get(s) != null) {
                negativeNumProcessedMap.put(companyIdNameMap.get(s), negativeNumMap.get(s));
            } else {
                negativeNumProcessedMap.put(s, negativeNumMap.get(s));
            }
        }

        //列出所有公司
        pIds = positiveNumProcessedMap.keySet();
        for(String s : pIds) {
            if(!allCompanies.contains(s))
                allCompanies.add(s);
        }
        nIds = negativeNumProcessedMap.keySet();
        for(String s : nIds) {
            if(!allCompanies.contains(s))
                allCompanies.add(s);
        }
        Set<String> qIds = qiChuMap.keySet();
        for(String s : qIds) {
            if(!allCompanies.contains(s))
                allCompanies.add(s);
        }

        System.out.println(allCompanies.size());
    }


    public static void main(String [] args) {
        parseExcelYinShou();
        parseExcelAllCompany();
        parseExcelChongFenLei();
        parseExcelQiChu();
//        parseExcelZhangLin();

        //转换id为公司名字
        listAllCompanies();

        exportExcel();

    }



}
