package co.com.gip.importer.dao;

import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBObject;
import com.mongodb.MongoClient;
import com.mongodb.BasicDBList;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
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
        Workbook wb = new HSSFWorkbook();
        Sheet data = wb.createSheet("Data");
        Row colorsRow = data.createRow((short) 0);
        Row row = data.createRow((short) 1);
        for (int i = 0; i < res.size(); i++) {
            String value = (String) res.get(i).get("name"),
                    type = (String) res.get(i).get("type");
            CellStyle style = wb.createCellStyle();
            Cell cell = row.createCell((short) i);
            Cell info = colorsRow.createCell(i);
            data.setColumnWidth(i, value.length() < 10 ? value.length() * 356 : value.length() * 256);
            if (type.equals("List")) {
                style.setFillBackgroundColor(HSSFColor.LIGHT_YELLOW.index);
                style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            }
            if (type.equals("DependentList")) {
                style.setFillBackgroundColor(HSSFColor.LIGHT_ORANGE.index);
                style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            }
            cell.setCellValue(value);
            info.setCellStyle(style);
        }

        FileOutputStream fileOut = new FileOutputStream("workbook4.xls");
        wb.write(fileOut);
        fileOut.close();
    }
}
