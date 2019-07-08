package wiley.xmlspreadsheetlibrary;

import wiley.xmlobjectlibrary.XmlNode;

public class Cell {
	private XmlNode cellNode;
	
	protected Cell(XmlNode cellNode) {
		if(cellNode == null) {
			return;
		}
		this.cellNode = cellNode;
	}
	
	protected Cell() {
		cellNode = new XmlNode("Cell");
		XmlNode data = new XmlNode("Data");
		data.addAttribute("ss:Type", "String");
		cellNode.addNode(data);
	}
	
	protected Cell(String value) {
		cellNode = new XmlNode("Cell");
		XmlNode data = new XmlNode("Data", value);
		data.addAttribute("ss:Type", "String");
		cellNode.addNode(data);
	}

	public XmlNode getCellNode() {
		return cellNode;
	}
	
	public void setStyle(CellStyle style) {
		if(cellNode.containsAttribute("ss:StyleID")) {
			cellNode.setAttribute("ss:StyleID", style.getId());
		} else {
			cellNode.addAttribute("ss:StyleID", style.getId());
		}
	}
	
	public String getValue() {
		XmlNode dataNode = cellNode.getFirstChild();
		if(dataNode == null) {
			return null;
		}
		return dataNode.getTextContent();
	}
	
	public void setValue(String value) {
		XmlNode dataNode = cellNode.getFirstChild();
		if(dataNode == null) {
			return;
		}
		dataNode.setTextContent(value);
	}
	
	public String getType() {
		return cellNode.getAttributeByName("ss:Type").getValue();
	}
	
	@Override
	public String toString() {
		return cellNode.toString();
	}
}
