import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

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
            String prefix = "https://v.douyin.com";
            if(url.startsWith(prefix)){
                HttpURLConnection connection = (HttpURLConnection) urlObject.openConnection();
                connection.setInstanceFollowRedirects(false); // 禁止自动重定向
                int responseCode = connection.getResponseCode();

                if (responseCode >= 300 && responseCode < 400) {
                    String redirectedUrl = connection.getHeaderField("Location");
                    String[] parts = redirectedUrl.split("\\?");
                    return douyinUrl3(parts[0]);
                }else if(responseCode == 200){
                    return douyinAllUrl(url);
                }

                connection.disconnect();
            }else{
                HttpURLConnection connection = (HttpURLConnection) urlObject.openConnection();
                connection.setInstanceFollowRedirects(false); // 禁止自动重定向
                int responseCode = connection.getResponseCode();

                if (responseCode >= 300 && responseCode < 400) {
                    String redirectedUrl = connection.getHeaderField("Location");
                    String[] parts = redirectedUrl.split("\\?");
//                return parts[0];
                    return douyinUrl2(parts[0]);
                }else if(responseCode == 200){
                    return douyinAllUrl(url);
                }

                connection.disconnect();
            }
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

    public static String douyinUrl3(String input){
        String originalUrl = "https://www.iesdouyin.com/share/note/7360647663789870336/?region=CN&schema_type=37";

        // 提取数字部分
        Pattern pattern = Pattern.compile("/note/(\\d+)/");
        Matcher matcher = pattern.matcher(originalUrl);
        String number = "";
        if (matcher.find()) {
            number = matcher.group(1);
        }

        // 提取参数部分
        pattern = Pattern.compile("schema_type=(\\d+)");
        matcher = pattern.matcher(originalUrl);
        String schemaType = "";
        if (matcher.find()) {
            schemaType = matcher.group(1);
            // 拼接目标URL
            return "https://www.iesdouyin.com/share/video/" + number + "?schema_type=" + schemaType;
        }else{
            return shortToLong(input);
        }


    }

    public static String douyinAllUrl(String input){
//        String pattern1 = "https://www.douyin.com/user/[^/]+\\?modal_id=(\\d+)&.*";
        String pattern1 = "https://www.douyin.com/user/[^/]+\\?modal_id=(\\d+)";
        String pattern2 = "https://www.iesdouyin.com/share/video/(\\d+)/";

        // 创建Pattern对象
        Pattern p1 = Pattern.compile(pattern1);
        Pattern p2 = Pattern.compile(pattern2);

        // 创建Matcher对象
        Matcher m1 = p1.matcher(input);
        Matcher m2 = p2.matcher(input);

        if (m1.matches()) {
            return "https://www.douyin.com/video/" + m1.group(1);
        }

        // 如果匹配第二个模式，则执行替换
        if (m2.matches()) {
            return "https://www.douyin.com/video/" + m2.group(1);
        }

        return "";
    }

}
