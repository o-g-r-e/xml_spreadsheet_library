package wiley.xmlspreadsheetlibrary;

import wiley.xmlobjectlibrary.XmlNode;

public class Sheet {
	
	private XmlNode sheetNode;
	private XmlNode tableNode;
	private int rowsCount = 0;
	
	protected Sheet(String sheetName) {
		sheetNode = new XmlNode("Worksheet");
		sheetNode.addAttribute("ss:Name", sheetName);
		
		tableNode = new XmlNode("Table");
		tableNode.addAttribute("ss:ExpandedColumnCount", "1");
		tableNode.addAttribute("ss:ExpandedRowCount", "1");
		tableNode.addAttribute("x:FullColumns", "1");
		tableNode.addAttribute("x:FullRows", "1");
		tableNode.addAttribute("ss:DefaultRowHeight", "15");
		
		sheetNode.addNode(tableNode);
		
		XmlNode worksheetOptions = new XmlNode("WorksheetOptions");
		worksheetOptions.addAttribute("xmlns", "urn:schemas-microsoft-com:office:excel");
		
		XmlNode pageSetup = new XmlNode("PageSetup");
		
		XmlNode header = new XmlNode("Header");
		header.addAttribute("x:Margin", "0.3");
		
		XmlNode footer = new XmlNode("Footer");
		footer.addAttribute("x:Margin", "0.3");
		
		XmlNode pageMargins = new XmlNode("PageMargins");
		footer.addAttribute("x:Bottom", "0.75");
		footer.addAttribute("x:Left", "0.7");
		footer.addAttribute("x:Right", "0.7");
		footer.addAttribute("x:Top", "0.75");
		
		pageSetup.addNode(header);
		pageSetup.addNode(footer);
		pageSetup.addNode(pageMargins);
		
		worksheetOptions.addNode(pageSetup);
		
		sheetNode.addNode(worksheetOptions);
	}
	
	protected Sheet(XmlNode sheetNode) {
		this.sheetNode = sheetNode;
		tableNode = this.sheetNode.getFirstChildByName("Table");
		rowsCount = tableNode.getAllChildsByName("Row").length;
	}
	
	public void setColumnWidth(int columnIndex, double width) {
		
		if(sheetNode == null || columnIndex < 0 || width < 0.0) {
			return;
		}
		
		//XmlNode tableNode = sheetNode.getFirstChildByName("Table");
		
		if(tableNode == null) {
			return;
		}
		
		int columnLength = tableNode.getAllChildsByName("Column").length;
		
		XmlNode columnNode = new XmlNode("Column");
		columnNode.addAttribute("ss:AutoFitWidth", "0");
		columnNode.addAttribute("ss:Width", String.valueOf(width));
		
		if(columnLength == 0) {
			/*int expandedColumnCount = Integer.parseInt(tableNode.getAttributeByName("ss:ExpandedColumnCount").getValue());
			if(columnIndex > (expandedColumnCount-1)) {
				tableNode.setAttribute("ss:ExpandedColumnCount", String.valueOf(columnIndex+1));
			}*/
			int expandedColumnCount = computeExpandedRowCount();
			tableNode.setAttribute("ss:ExpandedColumnCount", String.valueOf(expandedColumnCount));
			
			tableNode.addNode(columnNode);
			return;
		}	
		
		int currentColumnIndex = 0;
		XmlNode prependNode = null;
		int prependIndex = 0;
		int i = 0;
		
		for(XmlNode currentColumnNode : tableNode.getAllChildsByName("Column")) {
			if(currentColumnNode.containsAttribute("ss:Index")) {
				currentColumnIndex = Integer.parseInt(currentColumnNode.getAttributeByName("ss:Index").getValue());
			} else {
				currentColumnIndex++;
			}
			
			if(currentColumnIndex > columnIndex) {
				if((currentColumnIndex - columnIndex) < 3) {
					if(currentColumnNode.containsAttribute("ss:Index")) {
						currentColumnNode.removeAttribute("ss:Index");
					}
				}
				
				if((columnIndex - prependIndex) < 0) {
					prependNode.setAttribute("ss:Width", String.valueOf(width));
					return;
				} else if((columnIndex - prependIndex) == 0) {
					if(prependNode.containsAttribute("ss:Index")) {
						prependNode.removeAttribute("ss:Index");
					}
				} else if((columnIndex - prependIndex) > 0) {
					columnNode.addAttribute("ss:Index", String.valueOf(columnIndex+1));
				}
				
				break;
			}
			
			prependNode = currentColumnNode;
			prependIndex = currentColumnIndex;
			i++;
		}
		
		if(columnIndex > currentColumnIndex-1) {
			XmlNode lastColumnNode = tableNode.getAllChildsByName("Column")[columnLength-1];
			if(lastColumnNode.containsAttribute("ss:Index")) {
				if(columnIndex - Integer.parseInt(lastColumnNode.getAttributeByName("ss:Index").getValue()) > 0) {
					columnNode.addAttribute("ss:Index", String.valueOf(columnIndex+1));
				}
			} else {
				if(columnIndex - currentColumnIndex > 0) {
					columnNode.addAttribute("ss:Index", String.valueOf(columnIndex+1));
				}
			}
		} else if(columnIndex == currentColumnIndex-1) {
			XmlNode lastColumnNode = tableNode.getAllChildsByName("Column")[columnLength-1];
			if(lastColumnNode.containsAttribute("ss:Index")) {
				lastColumnNode.setAttribute("ss:Width", String.valueOf(width));
			}
		}
		
		if((columnLength-i) == 0) {
			/*int expandedColumnCount = Integer.parseInt(tableNode.getAttributeByName("ss:ExpandedColumnCount").getValue());
			if(columnIndex > (expandedColumnCount-1)) {
				tableNode.setAttribute("ss:ExpandedColumnCount", String.valueOf(columnIndex+1));
			}*/
			int expandedColumnCount = computeExpandedRowCount();
			tableNode.setAttribute("ss:ExpandedColumnCount", String.valueOf(expandedColumnCount));
			tableNode.addNode(columnNode);
		} else if((columnLength-i) > 0) {
			/*int expandedColumnCount = Integer.parseInt(tableNode.getAttributeByName("ss:ExpandedColumnCount").getValue());
			if(columnIndex > (expandedColumnCount-1)) {
				tableNode.setAttribute("ss:ExpandedColumnCount", String.valueOf(columnIndex+1));
			}*/
			int expandedColumnCount = computeExpandedRowCount();
			tableNode.setAttribute("ss:ExpandedColumnCount", String.valueOf(expandedColumnCount));
			tableNode.addNode(columnNode, i+1);
		}
	}
	
	public Row createRow() {
		Row row = new Row();
		tableNode.addNode(row.getRowNode());
		rowsCount++;
		tableNode.getAttributeByName("ss:ExpandedRowCount").setValue(String.valueOf(rowsCount));
		
		return row;
	}
	
	private int computeExpandedRowCount() {
		//int n1 = computeColsCount();
		int n2 = 0;
		for(XmlNode currentColumnNode : tableNode.getAllChildsByName("Row")) {
			if(currentColumnNode.containsAttribute("ss:Index")) {
				n2 = Integer.parseInt(currentColumnNode.getAttributeByName("ss:Index").getValue());
			} else {
				n2++;
			}
		}
		
		/*if(n1 > n2 ) {
			return n1;
		}*/

		return n2;
	}
	
	private int computeColsCount() {
		int maxColsCount = 0;
		XmlNode[] rows = tableNode.getAllChildsByName("Row");
		for (int i = 0; i < rowsCount; i++) {
			int rowLength = rows[i].getAllChildsByName("Cell").length;
			if(rows[i].containsAttribute("ss:Index")) {
				rowLength = Integer.parseInt(rows[i].getAttributeByName("ss:Index").getValue());
			}
			if(rowLength > maxColsCount) {
				maxColsCount = rowLength;
			}
		}
		
		return maxColsCount;
	}
	
	public void updateExpandetColumnsCount() {
		tableNode.setAttribute("ss:ExpandedColumnCount", computeColsCount()+"");
	}
	
	public void updateExpandetRowsCount() {
		tableNode.setAttribute("ss:ExpandedRowCount", computeExpandedRowCount()+"");
	}
	
	public Row getRow(int rowIndex) {
		XmlNode rowNode = tableNode.getAllChildsByName("Row")[rowIndex];
		if(rowNode == null) {
			return null;
		}
		return new Row(rowNode);
	}

	protected XmlNode getSheetNode() {
		return sheetNode;
	}

	public String getName() {
		return sheetNode.getAttributeByName("ss:Name").getValue();
	}

	public int getRowsCount() {
		return rowsCount;
	}
	
	@Override
	public String toString() {
		return sheetNode.toString();
	}
}
