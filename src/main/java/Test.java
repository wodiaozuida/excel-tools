import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;

public class Test {
    public static void main(String[] args) {
        readExcel();
    }

    public static void readExcel(){
        String oldfilePath = "C:\\may\\AAA.xlsx"; // 替换为你的Excel文件路径
        String newfilePath = "C:\\may\\宇宙超凶平哥手下为你倾力打造该表格.xlsx"; // 替换为你的Excel文件路径
        String sheetName = "Sheet1"; // 替换为你的Sheet名称
        int startRow = 1; // A2开始，索引从0开始

        try {
            FileInputStream file = new FileInputStream(new File(oldfilePath));
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheet(sheetName);

            for (int i = startRow; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell cell = row.createCell(1); // B列
                    cell.setCellValue(shortToLong(row.getCell(0).getStringCellValue()));
                }
            }

            FileOutputStream outputStream = new FileOutputStream(newfilePath);
            workbook.write(outputStream);
            file.close();
            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static String shortToLong(String url){
        try {
            URL urlObject = new URL(url);
            HttpURLConnection connection = (HttpURLConnection) urlObject.openConnection();
            connection.setInstanceFollowRedirects(false); // 禁止自动重定向
            int responseCode = connection.getResponseCode();

            if (responseCode >= 300 && responseCode < 400) {
                String redirectedUrl = connection.getHeaderField("Location");
                String[] parts = redirectedUrl.split("\\?");
//                return parts[0];
                return douyinUrl2(parts[0]);
            }else if(responseCode == 200){
                return douyinUrl(url);
            }

            connection.disconnect();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return "";
    }

    public static String douyinUrl(String input){

        String pattern = "https://www.douyin.com/user/[^/]+\\?modal_id=(\\d+)&.*";
        String replacement = "https://www.douyin.com/video/$1";

        return input.replaceAll(pattern, replacement);
    }

    public static String douyinUrl2(String input){
        String pattern = "https://www.iesdouyin.com/share/video/(\\d+)/";
        String replacement = "https://www.douyin.com/video/$1";

        return input.replaceAll(pattern, replacement);
    }

}