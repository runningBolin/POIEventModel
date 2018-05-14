package com.bolin.poi;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;

import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

/**
 * @author bolin
 * @create 2018年5月14日
 *
 */
public class DefaultSheetHandler extends DefaultHandler{

	protected InputStream sheetData;
	protected SharedStringsTable sharedStringsTable;
	
	protected HashMap<Integer, String> rowData = null;
	protected Integer rowIndex = null;
	protected Integer cellIndex = null;
	protected String cellValue = null;
	protected String cellReference = null;
	protected boolean cellValueIsString = false;
	
	public void parse() throws SAXException, IOException{
		XMLReader sheetParser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
		sheetParser.setContentHandler(this);
		InputSource sheetSource = new InputSource(sheetData);
		sheetParser.parse(sheetSource);
	}
	
	/**
	 * @author bolin
	 * @create 2018年5月11日
	 * 
	 * @param uri
	 * @param localName
	 * @param qName
	 * @param attributes
	 * @throws SAXException
	 */
	@Override
	public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
		
		if(qName.equals("row")){
			rowData = new HashMap<Integer, String>();
			rowIndex = Integer.parseInt(attributes.getValue("r"));
		}
		
		// c => cell
		if(qName.equals("c")) {
			// 获取单元格在sharedStringsTable中的索引 cell reference
			cellReference = attributes.getValue("r");
			cellIndex = parseIndex(cellReference);
			cellValueIsString = "s".equals(attributes.getValue("t"));
		}
		
		if(qName.equals("v")){
			cellValue = "";
		}
		
	}
	
	/**
	 * 
	 * @param uri
	 * @param localName
	 * @param qName
	 * @throws SAXException
	 */
	@Override
	public void endElement(String uri, String localName, String qName) throws SAXException {
		
		if(qName.equals("v")){
			if(cellValueIsString){
				int idx = Integer.parseInt(cellValue);
				cellValue = new XSSFRichTextString(sharedStringsTable.getEntryAt(idx)).toString();
			}
			rowData.put(cellIndex, cellValue);
		}
	}
	
	/**
	 * @author bolin
	 * @create 2018年5月11日
	 * 
	 * @param ch
	 * @param start
	 * @param length
	 * @throws SAXException
	 */
	@Override
	public void characters(char[] ch, int start, int length) throws SAXException {
		cellValue += new String(ch, start, length);
	}
	
	/**
	 * 解析excel列号，将字母列号转为10进制列号
	 * @param cellReference
	 * @return
	 */
	protected Integer parseIndex(String cellReference) {
		
		char[] chars = cellReference.toCharArray();
		
		String letters = "";
		
		for ( int i = 0; i < chars.length; i++ ) {
			//判断是大写字母,ascii码表中大写字母的十进制范围为65~90
			if((chars[i]) > 64 && (chars[i]) < 91){
				letters += chars[i];
			}
		}
		
		char[] letterChars = letters.toCharArray();
		
		Integer index = 0;
		for ( int j = letterChars.length; j > 0; j-- ) {
			index += ((chars[j-1]) - 64) * pow(26, letterChars.length-j);
		}
		
		return index;
	}
	
	/**
	 * 幂运算
	 * 
	 * @param index 指数
	 * @param power 次幂
	 * @return
	 */
	protected int pow(int index, int power){
		return (int)Math.pow(index, power);
	}

	
	public void setSheetData(InputStream sheetData) {
		this.sheetData = sheetData;
	}
	public void setSharedStringsTable(SharedStringsTable sharedStringsTable) {
		this.sharedStringsTable = sharedStringsTable;
	}
	
	
}
