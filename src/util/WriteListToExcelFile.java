package util;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import dto.W3CData;

/**
 * 해당 URL에서 테이블의 값을 읽어와서(Jsoup 라이브러리 사용) 데이터를 Excel로 출력(poi 라이브러리 사용)해준다.
 * @author ohdoking
 * 참고한 사이트 : http://www.journaldev.com/2562/apache-poi-tutorial
 *
 */
public class WriteListToExcelFile {

	public static void writeCountryListToFile(String fileName, List<String[]> countryList) throws Exception {
		Workbook workbook = null;

		if (fileName.endsWith("xlsx")) {
			workbook = new XSSFWorkbook();
		} else if (fileName.endsWith("xls")) {
			workbook = new HSSFWorkbook();
		} else {
			throw new Exception("invalid file name, should be xls or xlsx");
		}

		Sheet sheet = workbook.createSheet("Countries");
		
		
		
		
		Iterator<String[]> iterator = countryList.iterator();

		int rowIndex = 0;
		while (iterator.hasNext()) {
			String[] country = iterator.next();
			Row row = sheet.createRow(rowIndex++);
			for(int i = 0 ; i < country.length; i++){
				Cell cell0 = row.createCell(i);
				cell0.setCellValue(country[i]);
			}
		}

		// lets write the excel data to file now
		FileOutputStream fos = new FileOutputStream(fileName);
		workbook.write(fos);
		workbook.close();
		fos.close();
		System.out.println(fileName + " written successfully");
	}
	
	public static void setExcel(String url, String tableAndId ) throws Exception{
		
		List<W3CData> list = new ArrayList<W3CData>();

		ArrayList<String[]> data = new ArrayList<String[]>();
		try {
			Document doc = Jsoup.connect(url).cookie("JSESSIONID", "bacdLtf5EUpzGaOnEcQzv").get();
			
			Elements tableElements = doc.select(tableAndId);
			Elements tableHeaderEles = tableElements.select("tr th");
			
			int arraySize;
			if (tableHeaderEles.size() != 0){
				arraySize = tableHeaderEles.size();
				String[] s = new String[arraySize];
				for (int i = 0 ; i<arraySize ; i++) {
					s[i] = tableHeaderEles.get(i).text();
				}
				data.add(s);
			}
			else {
				Element tableHeaderEles2 = tableElements.select("tr").first();
				arraySize = tableHeaderEles2.getAllElements().select("td").size();
			}

			Elements tableRowElements = tableElements.select("tr td");
			int table = tableRowElements.size();
			for(int j = 0 ; j < tableRowElements.size() / arraySize  ; j++){
				String[] s2 = new String[arraySize];
				int num = 0;
				for (int i = j * arraySize; num < arraySize ; num++) {
					s2[num] = tableRowElements.get(i++).text();
				}
				num = 0;
				data.add(s2);
			}

		} catch (IOException e) {
			e.printStackTrace();
		}

		 WriteListToExcelFile.writeCountryListToFile("Countries.xlsx", data);
	}

	public static void main(String args[]) throws Exception {
		
		String url = "http://tcsv.ntcb.co.kr/admin_main.do";
		String tableAndId = "table.table-default";
		
		setExcel(url, tableAndId);
	}
}