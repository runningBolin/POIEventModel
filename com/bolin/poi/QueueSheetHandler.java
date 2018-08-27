package com.bolin.poi;

import java.io.IOException;
import java.util.HashMap;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;

import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

/**
 * 基于BlockingQueue实现的Sheet解析器
 * @author bolin
 * @create 2018年5月14日
 *
 */
public class QueueSheetHandler extends DefaultSheetHandler {
	

	private boolean start = false;
	private boolean finished = false;
	private BlockingQueue<HashMap<Integer, String>> queue;
	
	
	public boolean finished(){
		//if document read end and the queue is not empty and end get data
		//so that the readed rows not equals the document rows
		return start && finished && queue.isEmpty();
	}
	
	public HashMap<Integer, String> readRowData() throws InterruptedException {
		return start ? queue.take() : null;
	}
	
	/**
	 * @author bolin
	 * @create 2018年5月14日
	 * 
	 * @throws SAXException
	 * @throws IOException
	 */
	@Override
	public void parse() throws SAXException, IOException {
		final XMLReader sheetParser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
		sheetParser.setContentHandler(this);
		final InputSource sheetSource = new InputSource(sheetData);
		Thread parseThread = new Thread(new Runnable() {
			public void run() {
				try {
					sheetParser.parse(sheetSource);
				}
				catch (IOException e) {
					e.printStackTrace();
				}
				catch (SAXException e) {
					e.printStackTrace();
				}
			}
		});
		parseThread.start();
	}
	
	/**
	 * @author bolin
	 * @create 2018年5月14日
	 * 
	 * @throws SAXException
	 */
	@Override
	public void startDocument() throws SAXException {
		queue = new ArrayBlockingQueue<HashMap<Integer,String>>(20);
		start = true;
	}
	
	/**
	 * @author bolin
	 * @create 2018年5月14日
	 * 
	 * @throws SAXException
	 */
	@Override
	public void endDocument() throws SAXException {
		finished = true;
	}

	/**
	 * 
	 * @param uri
	 * @param localName
	 * @param qName
	 * @throws SAXException
	 */
	@Override
	public void endElement(String uri, String localName, String qName) throws SAXException  {
		super.endElement(uri, localName, qName);
		
		if(qName.equals("row")){
			try {
				queue.put(rowData);
			}
			catch (InterruptedException e) {
				e.printStackTrace();
			}
		}
	}
	
}
