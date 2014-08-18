package excelToYML;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.dom4j.Document;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.XMLWriter;

import data.excel.ExcelRow;
import data.excel.ReadDataFromExcel;
import data.xml.ConvertListToXMLDocument;

public class Controller {
	private static Controller instance;
	private List<ExcelRow> excelRowList;

	private Controller() {
		excelRowList = new ArrayList<ExcelRow>();
	}

	public static synchronized Controller getInstance() {
		if (instance == null) {
			instance = new Controller();
		}
		return instance;
	}

	public String getSaveFileNameFromPath(String path) {
		String xmlFilePath = ".xml";
		Matcher matcher = Pattern.compile(".*(?=\\.)").matcher(path);
		if (matcher.find()) {
			xmlFilePath = matcher.group() + xmlFilePath;
		}
		return xmlFilePath;
	}

	public void readExcelFile(String path) {
		try (FileInputStream file = new FileInputStream(new File(path))) {
			excelRowList = new ReadDataFromExcel().readExcel(file);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public void saveXML(String path) {
		XMLWriter output = null;
		try {
			Document document = new ConvertListToXMLDocument().convert(excelRowList);
			OutputFormat outFormat = OutputFormat.createPrettyPrint();
			outFormat.setEncoding("windows-1251");
			output = new XMLWriter(new FileOutputStream(path), outFormat);
			output.write(document);
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (output != null)
				try {
					output.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
		}
	}
}
