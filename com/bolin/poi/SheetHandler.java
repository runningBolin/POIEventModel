package com.bolin.poi;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.xml.sax.SAXException;

/**
 * sheet页解析器，可一次性读取出excel内数据
 * @author bolin
 * @create 2018年5月11日
 *
 */
public class SheetHandler extends DefaultSheetHandler {
	
	//用于存放当前sheet页内所有数据
	private List<Map<Integer, String>> dataList = new ArrayList<Map<Integer, String>>();
	
	/**
	 * 获取sheet页内数据，每行数据为一个HashMap，key为列号，从1开始
	 * @author bolin
	 * @create 2018年5月14日
	 * 
	 * @return
	 */
	public List<Map<Integer, String>> getDataList(){
		return dataList;
	}
	
	/**
	 * @author bolin
	 * @create 2018年5月14日
	 * 
	 * @param uri
	 * @param localName
	 * @param qName
	 * @throws SAXException
	 */
	@Override
	public void endElement(String uri, String localName, String qName) throws SAXException {
		super.endElement(uri, localName, qName);
		
		if(qName.equals("row")){
			dataList.add(rowData);
		}
	}
	
}
