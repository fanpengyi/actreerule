package com.bigdata.jianghu.utils;

import com.bigdata.jianghu.action.InterruptibleCharSequence;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelTest {

    public static void main(String[] args) {
        List<ArticleBean> list = new ArrayList<ArticleBean>();
        ExcelTest excelTest=new ExcelTest();
        Workbook wb = excelTest.getExcel("F:\\match\\article0.xls");
        if(wb==null)
            System.out.println("文件读入出错");
        else {
            list=excelTest.analyzeExcel(wb);
        }

    }

    public Workbook getExcel(String filePath){
        Workbook wb=null;
        File file=new File(filePath);
        if(!file.exists()){
            System.out.println("文件不存在");
            wb=null;
        }
        else {
            String fileType=filePath.substring(filePath.lastIndexOf("."));//获得后缀名
            try {
                InputStream is = new FileInputStream(filePath);
                if(".xls".equals(fileType)){
                    wb = new HSSFWorkbook(is);
                }else if(".xlsx".equals(fileType)){
                    wb = new XSSFWorkbook(is);
                }else{
                    System.out.println("格式不正确");
                    wb=null;
                }
            }catch (Exception e){
                e.printStackTrace();
            }
        }
        return wb;
    }

    public List<ArticleBean> analyzeExcel(Workbook wb){
        List<ArticleBean> list = new ArrayList<ArticleBean>();
        Sheet sheet=wb.getSheetAt(0);//读取sheet(从0计数)
        int rowNum=sheet.getLastRowNum();//读取行数(从0计数)
        for(int i=1;i<=rowNum;i++){
            ArticleBean bean = new ArticleBean();
            Row row=sheet.getRow(i);//获得行
            int colNum=row.getLastCellNum();//获得当前行的列数
            Cell cell1 = row.getCell(1);
            String content = cell1.toString();
            bean.setId(1);
            bean.setContent(content);
            list.add(bean);
//            Cell cell=row.getCell(j);
//            for(int j=0;j<colNum;j++){
//                Cell cell=row.getCell(j);//获取单元格
//                if(cell==null)
//                    System.out.print("null     ");
//                else
//                    System.out.print(cell.toString()+"     ");
//            }
            System.out.println();
        }
        return list;
    }

    public static boolean isContainChinese(String str, String ze) {
//        Pattern p = Pattern.compile(ze);
////        Matcher m = p.matcher(str);
////        if (m.find()) {
////            return true;
////        }
        Pattern pattern = Pattern.compile(ze);
        Matcher matcher = pattern.matcher(new InterruptibleCharSequence(str));
        pattern.matcher(new InterruptibleCharSequence(str));
        return matcher.matches();

//        return false;
    }
}

