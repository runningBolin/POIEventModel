package com.bolin.poi;

import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

/**
 * 用于解析wookbook信息中的sheet信息;
 * @author bolin
 * @create 2018年5月10日
 *
 */
public class WorkbookHandler extends DefaultHandler{
	
	private DefaultSheetHandler sheetHandler;
	private Map<String, String> sheetNameRelIdMap = new HashMap<String, String>();
	
	public WorkbookHandler(DefaultSheetHandler sheetHandler){
		this.sheetHandler = sheetHandler;
	}
	
	/**
	 * 解析单个sheet
	 * @author bolin
	 * @create 2018年5月14日
	 * 
	 * @param filename	excel文件路径
	 * @param sheetName	要解析的sheet页名称
	 * @throws Exception
	 */
	public void parseSheet(String filename, String sheetName) throws Exception{
		OPCPackage pkg = OPCPackage.open(filename);
		XSSFReader xssfReader = new XSSFReader( pkg );
		InputStream workbookData = xssfReader.getWorkbookData();
		XMLReader workbookXmlParser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
		workbookXmlParser.setContentHandler(this);
		workbookXmlParser.parse(new InputSource(workbookData));
		
		String sheetRid = sheetNameRelIdMap.get(sheetName);
		if(sheetRid == null){
			throw new Exception("Sheet named ["+ sheetName +"] dosen't exist in workbook");
		}
		InputStream sheetData = xssfReader.getSheet(sheetRid);
		
		sheetHandler.setSheetData(sheetData);
		sheetHandler.setSharedStringsTable(xssfReader.getSharedStringsTable());
		sheetHandler.parse();
	}
	
	/**
	 * @author bolin
	 * @create 2018年5月10日
	 * 
	 * @param uri
	 * @param localName
	 * @param qName
	 * @param attributes
	 * @throws SAXException
	 */
	@Override
	public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
		if(localName.equalsIgnoreCase("sheet")){
			String hidden = attributes.getValue("state");
			if(hidden == null){
				String sheetName = attributes.getValue("name");
				String rId = attributes.getValue("r:id");
				sheetNameRelIdMap.put(sheetName, rId);
			}
		}
	}

}
