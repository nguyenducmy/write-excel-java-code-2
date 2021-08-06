package WriteExcel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
 
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
 

	import org.apache.poi.ss.usermodel.Cell;
	import org.apache.poi.ss.usermodel.Row;
	import org.apache.poi.xssf.usermodel.XSSFSheet;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;
	import java.io.File;
	import java.io.FileOutputStream;
	import java.util.Map;
	import java.util.Set;
	import java.util.TreeMap;
	
	public class WriteExcel {
	    public static void main(String[] args) {
	        // Khởi tạo workbook cho tệp xlsx 
	        XSSFWorkbook workbook = new XSSFWorkbook();
	        // Khởi tạo một worksheet mới từ workbook 
	        XSSFSheet sheet = workbook.createSheet("student Details");
	        // Dữ liệu sẽ được ghi xuống file exel
	        Map<String, Object[]> data = new TreeMap<String, Object[]>();
	        data.put("1", new Object[]{"ID", "NAME", "LASTNAME"});
	        data.put("2", new Object[]{1, "Pankaj", "Kumar"});
	        data.put("3", new Object[]{2, "Prakashni", "Yadav"});
	        data.put("4", new Object[]{3, "Ayan", "Mondal"});
	        data.put("5", new Object[]{4, "Virat", "kohli"});
	        // Duyệt và thêm dữ liệu từng row
	        Set<String> keyset = data.keySet();
	        int rownum = 0;
	        for (String key : keyset) {
	            // this creates a new row in the sheet
	            Row row = sheet.createRow(rownum++);
	            Object[] objArr = data.get(key);
	            int cellnum = 0;
	            for (Object obj : objArr) {
	                Cell cell = row.createCell(cellnum++);
	                if (obj instanceof String)
	                    cell.setCellValue((String) obj);
	                else if (obj instanceof Integer)
	                    cell.setCellValue((Integer) obj);
	            }
	        }
	        try {
	            // ghi dữ liệu xuống file
	            FileOutputStream out = new FileOutputStream(new File("D://Java web/asfd.xlsx"));
	            workbook.write(out);
	            out.close();
	            System.out.println("gfgcontribute.xlsx written successfully on disk.");
	        } catch (Exception e) {
	            e.printStackTrace();
	        }
	    }
	}