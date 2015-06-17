package com.things.excel.utils;

import java.io.InputStream;
import java.util.Iterator;
import java.util.Map;
import java.util.HashMap;
import java.io.FileOutputStream;
import java.util.Date;
import java.text.*;
import java.io.FileWriter;
import java.io.PrintWriter;

import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;


public class SerialDataFormater {
	
	private static Workbook wb = new XSSFWorkbook();
	private static Sheet sheet;
	private static Map<String, CellStyle> styles;
	private static int rownum = 1;
	private static boolean bIsHeader = true;
	private static boolean bIsEmptyRow = false;
	private static final ConstantManager constMgrForInput = ConstantManager.getManager("config.properties.inputs");
	
	//Version 2 - start
	private static final ConstantManager outterInputs = ConstantManager.getManager("userinput");
	private static boolean bFirstTextFileLine = true;
	private static FileWriter txtFileWriter;
	private static PrintWriter txtPrintWriter;
	
//	private static ArrayList rowData = new ArrayList();
	
    private static final String[] titles = {
    	constMgrForInput.getString("serialformater.output.field.name.01"),
    	constMgrForInput.getString("serialformater.output.field.name.02"),
    	constMgrForInput.getString("serialformater.output.field.name.03"),
    	constMgrForInput.getString("serialformater.output.field.name.04"),
    	constMgrForInput.getString("serialformater.output.field.name.05"),
    	constMgrForInput.getString("serialformater.output.field.name.06"),
    	constMgrForInput.getString("serialformater.output.field.name.07"),
    	constMgrForInput.getString("serialformater.output.field.name.08"),
    	constMgrForInput.getString("serialformater.output.field.name.09"),
    	constMgrForInput.getString("serialformater.output.field.name.10")
    };
    
    private static final DateFormat fFullDateFormater = 
    		new SimpleDateFormat("yyyyMMdd");
    private static final DateFormat fFullOutPutDateFormater = 
    		new SimpleDateFormat("yyyy.MM.dd");
    private static final DateFormat fSmallDateFormater =
    		new SimpleDateFormat("yyMMdd");
    private static Date ipDate;
    private static String szBatchNum;
    
    private static void CreateDocHeader () {
        styles = createStyles(wb);

        sheet = wb.createSheet("Formatted Data");
        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setLandscape(true);
        sheet.setFitToPage(true);
        sheet.setHorizontallyCenter(true);

        //header row
        Row headerRow = sheet.createRow(0);
        headerRow.setHeightInPoints(40);
        Cell headerCell;
        for (int i = 0; i < titles.length; i++) {
            headerCell = headerRow.createCell(i);
            headerCell.setCellValue(titles[i]);
            headerCell.setCellStyle(styles.get("header"));
        }
    }
        
	private static void processInputSheets(String filename) throws Exception {
		OPCPackage pkg = OPCPackage.open(filename, PackageAccess.READ);
		XSSFReader r = new XSSFReader( pkg );
		SharedStringsTable sst = r.getSharedStringsTable();
		
		XMLReader parser = fetchSheetParser(sst);

		Iterator<InputStream> sheets = r.getSheetsData();
		while(sheets.hasNext()) {
			System.out.println("Processing new sheet:\n");
			InputStream sheet = sheets.next();
			InputSource sheetSource = new InputSource(sheet);
			parser.parse(sheetSource);
			sheet.close();
			System.out.println("Done.");
		}
	}
	
	private static XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
		XMLReader parser =
			XMLReaderFactory.createXMLReader(
					"org.apache.xerces.parsers.SAXParser"
			);
		ContentHandler handler = new SheetHandler(sst);
		parser.setContentHandler(handler);
		return parser;
	}
	
	private static class SheetHandler extends DefaultHandler {
		private SharedStringsTable sst;
		private String lastContents;
		private boolean nextIsString;
		private int colCount = 1;
		private String szSerialNum;
		private String szGTIN;
		
		private SheetHandler(SharedStringsTable sst) {
			this.sst = sst;
		}
		
		public void startElement(String uri, String localName, String name,
				Attributes attributes) throws SAXException {
			
			if (name.equals("row")) {
				//Handle the header and the empty rows, even with cleaned up excel sheets,
				//there seem to be empty rows, these rows must be skipped
				if (bIsHeader || bIsEmptyRow)
					return;
				//Handle the regular rows, create cells at the start of the row
	            Row row = sheet.createRow(rownum);
	            for (int j = 0; j < titles.length; j++) {
	                Cell cell = row.createCell(j);
	                cell.setCellStyle(styles.get("cell"));
	            }
	            colCount = 0;
				return;
			}
			// c => cell
			if(name.equals("c")) {
				
				//identify the empty rows with the number of attributes,
				//empty rows have no v attribute, if a empty row is found set the flag
				//use this flag to skip the row processing
				if (attributes.getLength() < 2) {
					bIsEmptyRow = true;
					Row row = sheet.getRow(rownum);
					sheet.removeRow(row);
					return;
				}
				// Figure out if the value is an index in the SST
				String cellType = attributes.getValue("t");
				if(cellType != null && cellType.equals("s")) {
					nextIsString = true;
				} else {
					nextIsString = false;
				}
			}
			// Clear contents cache
			lastContents = "";
		}
		
		public void endElement(String uri, String localName, String name)
				throws SAXException {
			// Process the last contents as required.
			// Do now, as characters() may be called more than once
			if(nextIsString) {
				int idx = Integer.parseInt(lastContents);
				lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
            nextIsString = false;
			}

			// v => contents of a cell
			// Output after we've seen the string contents
			if(name.equals("v")) {
				//hard-code for 1st cell--> Serial Number
				             //2nd cell--> GTIN
				if (colCount == 0)
					szSerialNum = lastContents;
				else
					szGTIN = lastContents;
				colCount++;
			}
			else if (name.equals("row")) {
				//Handle end of empty row, unset flags; get ready for new row
				if (bIsHeader || bIsEmptyRow) {
					bIsHeader = false;
					bIsEmptyRow = false;
					return;
				}
				//Version 2 Text File writer
				//Write the first row
				if (bFirstTextFileLine) {
						try {
							txtFileWriter = new FileWriter (constMgrForInput.getString("serialformater.output.text.file.full.path"), false);
							txtPrintWriter = new PrintWriter (txtFileWriter);
							txtPrintWriter.println("#"+szGTIN+"#"+szBatchNum+"#"+fSmallDateFormater.format(ipDate));
						} catch (Exception e) {
							e.printStackTrace();
						}
						bFirstTextFileLine = false;
					}
				txtPrintWriter.println(szSerialNum);
				//process regular row, at end of row prepare output row on the sheet
				//currently hard-coded, need to find a way to get processing info in a 
				//Generic manner, e.g. implement processing rule framework??!!
				Row row = sheet.getRow(rownum);
				row.getCell(0).setCellValue(szSerialNum);
				row.getCell(1).setCellValue("(21)"+szSerialNum);
				row.getCell(2).setCellValue(szGTIN);
				row.getCell(3).setCellValue("(01)"+szGTIN);
				row.getCell(4).setCellValue(szBatchNum);
				row.getCell(5).setCellValue("(10)"+szBatchNum);
				row.getCell(6).setCellValue(fSmallDateFormater.format(ipDate));
				row.getCell(7).setCellValue("(17)"+fSmallDateFormater.format(ipDate));
				row.getCell(8).setCellValue(fFullOutPutDateFormater.format(ipDate));
				row.getCell(9).setCellValue("(17)"+fFullOutPutDateFormater.format(ipDate));
				rownum = rownum + 1;
				System.out.println("Got to "+rownum);
			}
	}

		public void characters(char[] ch, int start, int length)
				throws SAXException {
			lastContents += new String(ch, start, length);
		}
	}
 
    public static void main(String[] args) throws Exception {
    	//Get the constants
    	ipDate = fFullDateFormater.parse(
    			outterInputs.getString("expiry.data.YYYYMMDD"));
    	szBatchNum = outterInputs.getString("batch.number");
    	
        //create the output document header
    	CreateDocHeader();
    	
    	//read the input file
    	try {
    		processInputSheets(
    		constMgrForInput.getString("serialformater.input.xlsx.file.full.path"));
    	} catch (Exception e) {
    		e.printStackTrace();
    	}
    	

        //finally set column widths
        sheet.setDefaultColumnWidth (20);  //20 characters wide

        // Write the output file
        FileOutputStream out = new FileOutputStream(
						constMgrForInput.getString("serialformater.output.xlsx.file.full.path"));
        wb.write(out);
        out.close();
        txtPrintWriter.close();
    }

    /**
     * Create a library of cell styles
     */
    private static Map<String, CellStyle> createStyles(Workbook wb){
        Map<String, CellStyle> styles = new HashMap<String, CellStyle>();
        CellStyle style;
        Font titleFont = wb.createFont();
        titleFont.setFontHeightInPoints((short)18);
        titleFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFont(titleFont);
        styles.put("title", style);

        Font monthFont = wb.createFont();
        monthFont.setFontHeightInPoints((short)11);
        monthFont.setColor(IndexedColors.WHITE.getIndex());
        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setFont(monthFont);
        style.setWrapText(true);
        styles.put("header", style);

        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setWrapText(true);
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        styles.put("cell", style);

        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        styles.put("formula", style);

        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        styles.put("formula_2", style);

        return styles;
    }
}
