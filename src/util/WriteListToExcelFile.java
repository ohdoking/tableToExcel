package util;

import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
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

	public static void writeCountryListToFile(String fileName, List<W3CData> countryList) throws Exception {
		Workbook workbook = null;

		if (fileName.endsWith("xlsx")) {
			workbook = new XSSFWorkbook();
		} else if (fileName.endsWith("xls")) {
			workbook = new HSSFWorkbook();
		} else {
			throw new Exception("invalid file name, should be xls or xlsx");
		}

		Sheet sheet = workbook.createSheet("Countries");

		Iterator<W3CData> iterator = countryList.iterator();

		int rowIndex = 0;
		while (iterator.hasNext()) {
			W3CData country = iterator.next();
			Row row = sheet.createRow(rowIndex++);
			Cell cell0 = row.createCell(0);
			cell0.setCellValue(country.getCompany());
			Cell cell1 = row.createCell(1);
			cell1.setCellValue(country.getContact());
			Cell cell2 = row.createCell(2);
			cell2.setCellValue(country.getCountry());
		}

		// lets write the excel data to file now
		FileOutputStream fos = new FileOutputStream(fileName);
		workbook.write(fos);
		fos.close();
		System.out.println(fileName + " written successfully");
	}

	public static void main(String args[]) throws Exception {
		
		String url = "http://www.w3schools.com/css/css_table.asp";
		String tableAndId = "table#customers";
		List<W3CData> list = new ArrayList<W3CData>();

		try {
			Document doc = Jsoup.connect(url).get();
			Elements tableElements = doc.select(tableAndId);
			Elements tableHeaderEles = tableElements.select("tr th");
			W3CData tmp = new W3CData();
			tmp.setCompany(tableHeaderEles.get(0).text());
			tmp.setContact(tableHeaderEles.get(1).text());
			tmp.setCountry(tableHeaderEles.get(2).text());
			list.add(tmp);

			Elements tableRowElements = tableElements.select("tr");
			for (int i = 0; i < tableRowElements.size(); i++) {
				Element row = tableRowElements.get(i);
				Elements rowItems = row.select("td");
				if (rowItems.size() != 0) {
					W3CData tmp2 = new W3CData();
					tmp2.setCompany(rowItems.get(0).text());
					tmp2.setContact(rowItems.get(1).text());
					tmp2.setCountry(rowItems.get(2).text());
					list.add(tmp2);
				}
			}

		} catch (IOException e) {
			e.printStackTrace();
		}

		 WriteListToExcelFile.writeCountryListToFile("Countries.xls", list);
	}
}