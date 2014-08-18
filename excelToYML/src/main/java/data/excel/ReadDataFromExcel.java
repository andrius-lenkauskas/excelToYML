package data.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

public class ReadDataFromExcel {

	public ReadDataFromExcel() {
	}

	public List<ExcelRow> readExcel(InputStream inputStream) {
		List<ExcelRow> itemsList = new ArrayList<ExcelRow>();
		try {
			// Get the workbook instance for excel
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			// Get first sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);
			// Get iterator to all the rows in current sheet
			Iterator<Row> rowIterator = sheet.iterator();
			Map<String, Integer> headerMap = getHeaderMapFromSheet(rowIterator);
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				ExcelRow excelRow = new ExcelRow();
				for (Map.Entry<String, Integer> entry : headerMap.entrySet()) {
					String cellValueAsString = getCellValueAsString(row.getCell(entry.getValue()));
					switch (entry.getKey()) {
					case "group_id":
						excelRow.setGroup_id(cellValueAsString);
						break;
					case "available":
						excelRow.setAvailable(cellValueAsString);
						break;
					case "store":
						excelRow.setStore(cellValueAsString);
						break;
					case "pickup":
						excelRow.setPickup(cellValueAsString);
						break;
					case "local_delivery_cost":
						excelRow.setLocal_delivery_cost(cellValueAsString);
						break;
					case "url":
						excelRow.setUrl(cellValueAsString);
						break;
					case "price":
						excelRow.setPrice(cellValueAsString);
						break;
					case "currencyId":
						excelRow.setCurrencyId(cellValueAsString);
						break;
					case "picture":
						excelRow.setPicture(cellValueAsString);
						break;
					case "picture2":
						excelRow.setPicture2(cellValueAsString);
						break;
					case "picture3":
						excelRow.setPicture3(cellValueAsString);
						break;
					case "title":
						excelRow.setTitle(cellValueAsString);
						break;
					case "country_of_origin":
						excelRow.setCountry_of_origin(cellValueAsString);
						break;
					case "description":
						excelRow.setDescription(cellValueAsString);
						break;
					case "vendor":
						excelRow.setVendor(cellValueAsString);
						break;
					case "delivery":
						excelRow.setDelivery(cellValueAsString);
						break;
					case "yandex_color_name":
						excelRow.setYandex_color_name(cellValueAsString);
						break;
					case "gender":
						excelRow.setGender(cellValueAsString);
						break;
					case "age":
						excelRow.setAge(cellValueAsString);
						break;
					case "material":
						excelRow.setMaterial(cellValueAsString);
						break;
					/*case "typePrefix":
						excelRow.setTypePrefix(cellValueAsString);*/
					case "model":
						excelRow.setModel(cellValueAsString);
						break;
					case "size":
						excelRow.setSize(cellValueAsString);
						break;
					case "size_unit":
						excelRow.setSize_unit(cellValueAsString);
						break;
					case "categoryId":
						excelRow.setCategoryId(cellValueAsString);
						break;
					case "market_category":
						excelRow.setMarket_category(cellValueAsString);
						break;
					}
				}
				itemsList.add(excelRow);
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
		return itemsList;
	}

	private String getCellValueAsString(Cell cell) {
		String cellAsString = "";
		if (cell != null) {
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_BOOLEAN:
				cellAsString = Boolean.toString(cell.getBooleanCellValue());
				break;
			case Cell.CELL_TYPE_NUMERIC:
				if ((cell.getNumericCellValue() % 1) == 0) {
					cellAsString = Long.toString((long) cell.getNumericCellValue());
				} else {
					cellAsString = Double.toString(cell.getNumericCellValue());
				}
				break;
			case Cell.CELL_TYPE_STRING:
				cellAsString = cell.getStringCellValue();
				break;
			case Cell.CELL_TYPE_BLANK:
				break;
			}
		}
		return cellAsString.trim();
	}

	private Map<String, Integer> getHeaderMapFromSheet(Iterator<Row> rowIterator) {
		Map<String, Integer> headerMap = new HashMap<String, Integer>();
		if (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				if (!cell.getStringCellValue().trim().isEmpty())
					headerMap.put(cell.getStringCellValue().trim(), cell.getColumnIndex());
			}
		}
		return headerMap;
	}
}
