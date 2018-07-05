import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;

public class Util {

    //读取excel表格的第一列，返回一个包含所有url的一维数组
    public static String[] readExcel() {
        try {
            // 获取Excel模板文件
            File file = new File("D:\\read.xlsx");
            // 读取Excel模板
            XSSFWorkbook wb = new XSSFWorkbook(file);
            // 读取了模板内sheet的内容
            XSSFSheet sheet = wb.getSheetAt(0);
            // 在相应的单元格进行（读取）赋值 行列分别从0开始
            int rowNub = sheet.getLastRowNum();
            XSSFRow xssfRow = null;
            XSSFCell xssfCell = null;
            String []urlData = new String[rowNub+1];
            for (int i=0; i<=rowNub; i++) {
                xssfRow = sheet.getRow(i);
                xssfCell = xssfRow.getCell(0);
                urlData[i] = xssfCell.getStringCellValue();
            }
            wb.close();
            return urlData;
        } catch (InvalidFormatException inv) {
            System.out.println(inv);
        } catch (IOException ioe) {
            System.out.println(ioe);
        }  catch (Exception e) {
            System.out.println(e);
        }
        return new String[0];
    }

    //向excel表格里写入数据，按行写入，每行写入四个数据，data为二维数组
    public static void writeExcel(String [][]data) {
        try {
            FileOutputStream out = new FileOutputStream("D:\\crawlerData.xlsx");

            XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
            XSSFSheet xssfSheet = xssfWorkbook.createSheet();
            XSSFRow xssfRow= null;
            int dataNum = data.length;
            for (int i=0; i<dataNum; i++) {
                xssfRow = xssfSheet.createRow(i);
                xssfRow.createCell(0).setCellValue(data[i][0]);
                xssfRow.createCell(1).setCellValue(data[i][1]);
                xssfRow.createCell(2).setCellValue(data[i][2]);
                xssfRow.createCell(3).setCellValue(data[i][3]);
                xssfRow.createCell(4).setCellValue(data[i][4]);
            }
            xssfWorkbook.write(out);
            out.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    //爬取指定url的所需内容，当前爬取title、description、canonical、keywords,返回一维数组
    public static String[] GetUrlData(String url) {
        //解析Url获取Document对象
        Document document = null;
        try {
            document = Jsoup.connect(url).get();
        } catch (IOException e) {
            e.printStackTrace();
        }
        String title = document.title();
        Element description = document.select("meta[name=description]").first();
        Element canonical = document.select("link[rel=canonical]").first();
        Element keywords = document.select("meta[name=keywords]").first();
        String descriptionString = description.attr("content");
        String canonicalTextString = canonical.attr("href");
        String keywordsString = keywords.attr("content");
        String arr[] = new String[5];
        arr[0] = url;
        arr[1] = title;
        arr[2] = descriptionString;
        arr[3] = canonicalTextString;
        arr[4] = keywordsString;
        return arr;
    }
}
