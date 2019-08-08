package TestPackage;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;


public class Reader {

@SuppressWarnings({"unchecked","unchecked"})
    public static void main(String[] args) throws Exception{
        String filename = "your local path";

        List sheetData = new ArrayList();
        FileInputStream fis = null;
        try{
            fis = new FileInputStream(filename);
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFSheet sheet = workbook.getSheetAt(0);

            Iterator rows = ((XSSFSheet) sheet).rowIterator();
            while (rows.hasNext()){
                XSSFRow row = (XSSFRow) rows.next();
                Iterator cells = row.cellIterator();

                List data = new ArrayList();
                while (cells.hasNext()){
                    XSSFRow cell = (XSSFRow) cells.next();
                    data.add(cell);
            }
            sheetData.add(data);
        }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (fis != null) {
                fis.close();
            }
        }
        parseExcelData(sheetData);
        }

        private static void parseExcelData(List sheetData) throws Exception {
        java.sql.Connection conn;
        java.sql.Driver d = (java.sql.Driver) Class.forName("solid.jdbc.SolidDriver").newInstance();
        String sCon = "your url";
        conn = java.sql.DriverManager.getConnection(sCon);

        for (int i = 1; i < sheetData.size(); i++){
            List list = (List) sheetData.get(i);

            Cell to = (Cell) list.get(0);
            Cell from = (Cell) list.get(1);

            String bin = to.getRichStringCellValue().toString();
            String newTo = bin.substring(0,9) + "9999999999";

            bin = from.getRichStringCellValue().toString();
            String newFrom = bin.substring(0,9) + "0000000000";

            findDataInDatabase(newFrom,newTo,conn);
        }
        conn.close();
        }

        private static void findDataInDatabase(String from, String to, java.sql.Connection conn) throws Exception{

        java.sql.ResultSetMetaData meta;
        java.sql.Statement stmt;
        java.sql.ResultSet result;
        int i;

        String sQuery = "SELECT BIN_TO,BIN_FROM,BIN_REGRATE" +
                        "FROM BIN_ISSUERRANGES"              +
                        "WHERE SCH_ABBREV = 'VI'"            +
                        "AND BIN_REGRATE IS NULL"            +
                        "AND BIN_FROM >= '" + from + "'"     +
                        "AND BIN_TO >= '"   + to   + "'"     ;
        stmt = conn.createStatement();
        result = stmt.executeQuery(sQuery);
        meta = result.getMetaData();
        int cols = meta.getColumnCount();
        int cnt = 1;
            while (result.next()) {
                System.out.println("\nRow " + cnt + " : ");
                for (i = 1; i<= cols; i++) {
                    System.out.println(result.getString(i) + "\t");
                }
                cnt++;
            }
            stmt.close();

        }

}
