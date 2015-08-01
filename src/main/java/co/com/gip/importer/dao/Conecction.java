package co.com.gip.importer.dao;

import com.mongodb.*;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Barcode on 30/07/2015.
 */
public class Conecction {
    public static void main(String[] args) throws IOException {
        MongoClient mongoClient = new MongoClient("192.168.1.28", 27017);
        DB db = mongoClient.getDB("GIP");
        DBCollection coll = db.getCollection("excelconfigcontacts");
        DBObject myDoc = coll.findOne();
        BasicDBList fields = (BasicDBList) myDoc.get("fields");
        List<DBObject> res = new ArrayList<DBObject>();

        for (Object el : fields) res.add((DBObject) el);
        HSSFWorkbook wb = new HSSFWorkbook();
        Sheet data = wb.createSheet("Data");
        Row colorsRow = data.createRow((short) 0);
        Row row = data.createRow((short) 1);
        for (int i = 0; i < res.size(); i++) {

            Sheet sheet = wb.getSheet("Data");

            String value = (String) res.get(i).get("name"),
                    type = (String) res.get(i).get("type"),
                    order = res.get(i).get("order").toString();


            CellStyle style = wb.createCellStyle();
            Cell cell = row.createCell((short) i);
            Cell info = colorsRow.createCell(i);
            sheet.setColumnWidth(i, value.length() < 10 ? value.length() * 356 : value.length() * 256);
            // Cell styles for list and dependent Lists and creation of other sheets for list values
            if (type.equals("List")) {

                info.setCellValue(order);
                style.setFillBackgroundColor(HSSFColor.LIGHT_GREEN.index);
                style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                style.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);


                //sheets list values
                Sheet List = wb.createSheet(order);
                Row headers = List.createRow((short) 0);
                Cell CHid = headers.createCell((short) 0);
                Cell CHValue = headers.createCell((short) 1);
                CHid.setCellValue("ID");
                CHValue.setCellValue("Valores");

                // Values
                BasicDBList listValues = (BasicDBList) res.get(i).get("listValues");
                BasicDBObject[] lightArr = listValues.toArray(new BasicDBObject[0]);
                int j = 1;
                for (BasicDBObject dbObj : lightArr) {
                    Row dataList = List.createRow((short) j);
                    Cell IDList = dataList.createCell((short) 0);
                    Cell ValueList = dataList.createCell((short) 1);
                    IDList.setCellValue((Integer) dbObj.get("value"));
                    ValueList.setCellValue((String) dbObj.get("label"));
                    j++;
                }
            }
            if (type.equals("DependentList")) {

                info.setCellValue(order);
                String listName = res.get(i).get("listName").toString();
                style.setFillBackgroundColor(HSSFColor.LIGHT_BLUE.index);
                style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                style.setFillForegroundColor(HSSFColor.LIGHT_BLUE.index);

                //sheets list values
                Sheet List = wb.createSheet(order);
                Row headers = List.createRow((short) 0);
                Cell CHid = headers.createCell((short) 0);
                Cell CHValue = headers.createCell((short) 1);
                CHid.setCellValue("ID");
                CHValue.setCellValue("Valores");

                DBCollection lists = db.getCollection("lists");
                BasicDBObject query = new BasicDBObject("name", listName);
                DBCursor ListValuesDB = lists.find(query);
                try {
                    int k = 1;
                    while (ListValuesDB.hasNext()) {
                        DBObject result = ListValuesDB.next();
                        Row dataList = List.createRow((short) k);
                        Cell IDList = dataList.createCell((short) 0);
                        Cell ValueList = dataList.createCell((short) 1);
                        IDList.setCellValue((String) result.get("key"));
                        ValueList.setCellValue((String) result.get("value"));
                        k++;
                    }
                } finally {
                    ListValuesDB.close();
                }
            }

            cell.setCellValue(value);
            info.setCellStyle(style);
            // END -- Cell styles for list and dependent Lists

        }


        FileOutputStream fileOut = new FileOutputStream("contactosBase.xls");
        wb.write(fileOut);
        fileOut.close();
    }
}
