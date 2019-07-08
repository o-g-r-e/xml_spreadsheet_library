package wiley.xmlspreadsheetlibrary;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.List;

import javax.xml.stream.XMLStreamException;

import wiley.xmlobjectlibrary.XmlNode;
import wiley.xmlobjectlibrary.XmlObjectManager;

public class Workbook {
	
	class DocumentProperties {
		private XmlNode xmlNode;
		
		public DocumentProperties(XmlNode xmlNode) {
			this.xmlNode = xmlNode;
		}
		
		public DocumentProperties() {
			xmlNode = new XmlNode("DocumentProperties");
			xmlNode.addAttribute("xmlns", "urn:schemas-microsoft-com:office:office");
			
			xmlNode.addNode(new XmlNode("LastAuthor", ""));
			xmlNode.addNode(new XmlNode("Created", ""));
			xmlNode.addNode(new XmlNode("LastSaved", ""));
			xmlNode.addNode(new XmlNode("Version", ""));
		}
		
		public XmlNode getXmlNode() {
			return xmlNode;
		}

		@Override
		public String toString() {
			return xmlNode.toString();
		}
	}
	
	class OfficeDocumentSettings {
		private XmlNode xmlNode;
		
		public OfficeDocumentSettings(XmlNode xmlNode) {
			this.xmlNode = xmlNode;
		}
		
		public OfficeDocumentSettings() {
			xmlNode = new XmlNode("OfficeDocumentSettings");
			xmlNode.addAttribute("xmlns", "urn:schemas-microsoft-com:office:office");
			
			xmlNode.addNode(new XmlNode("AllowPNG"));
		}
		
		public XmlNode getXmlNode() {
			return xmlNode;
		}
		
		@Override
		public String toString() {
			return xmlNode.toString();
		}
	}
	
	class ExcelWorkbook {
		private XmlNode xmlNode;
		
		public ExcelWorkbook(XmlNode xmlNode) {
			this.xmlNode = xmlNode;
		}
		
		public ExcelWorkbook() {
			xmlNode = new XmlNode("ExcelWorkbook");
			xmlNode.addAttribute("xmlns", "urn:schemas-microsoft-com:office:excel");
			
			xmlNode.addNode(new XmlNode("WindowHeight", "12300"));
			xmlNode.addNode(new XmlNode("WindowWidth", "28800"));
			xmlNode.addNode(new XmlNode("WindowTopX", "0"));
			xmlNode.addNode(new XmlNode("WindowTopY", "0"));
			xmlNode.addNode(new XmlNode("ProtectStructure", "False"));
			xmlNode.addNode(new XmlNode("ProtectWindows", "False"));
		}
		
		public XmlNode getXmlNode() {
			return xmlNode;
		}
		
		@Override
		public String toString() {
			return xmlNode.toString();
		}
	}
	
	private DocumentProperties documentProperties;
	private OfficeDocumentSettings officeDocumentSettings;
	private ExcelWorkbook excelWorkbook;
	private List<CellStyle> styles;
	private List<Sheet> sheets;
	private XmlObjectManager xmlObjectManager;
	
	public Workbook() {
		
		XmlNode workbookNode = new XmlNode("Workbook", "");
		workbookNode.addAttribute("xmlns", "urn:schemas-microsoft-com:office:spreadsheet");
		workbookNode.addAttribute("xmlns:o", "urn:schemas-microsoft-com:office:office");
		workbookNode.addAttribute("xmlns:x", "urn:schemas-microsoft-com:office:excel");
		workbookNode.addAttribute("xmlns:ss", "urn:schemas-microsoft-com:office:spreadsheet");
		workbookNode.addAttribute("xmlns:html", "http://www.w3.org/TR/REC-html40");
		
		documentProperties = new DocumentProperties();
		officeDocumentSettings = new OfficeDocumentSettings();
		excelWorkbook = new ExcelWorkbook();
		
		workbookNode.addNode(documentProperties.getXmlNode());
		workbookNode.addNode(officeDocumentSettings.getXmlNode());
		workbookNode.addNode(excelWorkbook.getXmlNode());
		workbookNode.addNode(new XmlNode("Styles"));
		
		xmlObjectManager = new XmlObjectManager("1.0", "UTF-8", workbookNode);
		
		sheets = new ArrayList<>();
		styles = new ArrayList<CellStyle>();
	}
	
	public Workbook(String xmlFilePath) throws XMLStreamException, FileNotFoundException {
		xmlObjectManager = new XmlObjectManager(new File(xmlFilePath));
		styles = new ArrayList<CellStyle>();
		
		XmlNode[] styleNodes = xmlObjectManager.getRootNode().getFirstChildByName("Styles").getAllChildsByName("Style");
		
		for(XmlNode styleNode : styleNodes) {
			styles.add(new CellStyle(styleNode));
		}
		
		sheets = new ArrayList<Sheet>();
		
		XmlNode[] workSheets = xmlObjectManager.getRootNode().getAllChildsByName("Worksheet");
		
		for(XmlNode ws : workSheets) {
			sheets.add(new Sheet(ws));
		}
	}
	
	public Sheet getSheetByIndex(int sheetIndex) {
		if(sheetIndex < 0 || sheetIndex > sheets.size()-1) {
			return null;
		}
		return sheets.get(sheetIndex);
	}
	
	public Sheet getSheetByName(String sheetName) {
		for(Sheet worksheet : sheets) {
			if(worksheet.getName().equals(sheetName)) {
				return worksheet;
			}
		}
		return null;
	}
	
	@Override
	public String toString() {
		for(Sheet sheet : sheets) {
			sheet.updateExpandetColumnsCount();
			sheet.updateExpandetRowsCount();
		}
		return xmlObjectManager.toString(new String[]{"Cell"});
	}
	
	private String generateStyleId(int stylesCount) {
		return "s"+(15+stylesCount);
	}
	
	public CellStyle createStyle() {
		CellStyle newStyle = new CellStyle();
		styles.add(newStyle);
		newStyle.setId(generateStyleId(styles.size()));
		xmlObjectManager.getRootNode().getFirstChildByName("Styles").addNode(newStyle.getStyleNode());
		return newStyle;
	}
	
	public CellStyle getStyleById(String id) {
		for(CellStyle style : styles) {
			if(style.getId().equals(id)) {
				return style;
			}
		}
		return null;
	}
	
	public Font createFont() {
		Font newFont = new Font();
		return newFont;
	}
	
	public Sheet createSheet(String sheetName) {
		Sheet newSheet = new Sheet(sheetName);
		sheets.add(newSheet);
		xmlObjectManager.getRootNode().addNode(newSheet.getSheetNode());
		return newSheet;
	}
	
	public void write(/*String content,*/ String filePath) throws IOException {
		File file = new File(filePath);
		
		if(!file.exists()) {
			file.getParentFile().mkdirs(); 
			file.createNewFile();
		}
		
		PrintWriter writer = new PrintWriter(file, "UTF-8");
		writer.print(this.toString());
		writer.close();
	}
}
