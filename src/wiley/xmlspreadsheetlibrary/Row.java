package wiley.xmlspreadsheetlibrary;

import wiley.xmlobjectlibrary.XmlNode;

public class Row {
	private XmlNode rowNode;
	
	protected Row(XmlNode rowNode) {
		this.rowNode = rowNode;
	}
	
	protected Row() {
		rowNode = new XmlNode("Row");
	}

	protected XmlNode getRowNode() {
		return rowNode;
	}
	
	public void setStyle(CellStyle style) {
		if(rowNode.containsAttribute("ss:StyleID")) {
			rowNode.setAttribute("ss:StyleID", style.getId());
		} else {
			rowNode.addAttribute("ss:StyleID", style.getId());
		}
	}
	
	public Cell createCell() {
		Cell newCell = new Cell();
		rowNode.addNode(newCell.getCellNode());
		return newCell;
	}
	
	public Cell getCell(int i) {
		XmlNode cellNode = rowNode.getNodeByIndex(i);
		if(cellNode == null) {
			return null;
		}
		return new Cell(cellNode);
	}
	
	@Override
	public String toString() {
		return rowNode.toString();
	}
}
