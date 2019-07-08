package wiley.xmlspreadsheetlibrary;

import wiley.xmlobjectlibrary.XmlNode;

public class CellStyle {
	private XmlNode styleNode;
	
	protected CellStyle() {
		this.styleNode = new XmlNode("Style");
	}
	
	protected CellStyle(XmlNode styleNode) {
		this.styleNode = styleNode;
	}

	public void setFont(Font font) {
		if(this.styleNode == null) {
			return;
		}
		
		XmlNode fontNode = styleNode.getFirstChildByName("Font");
		
		if(fontNode != null) {
			styleNode.removeFirstChildByName("Font");
		}
		
		styleNode.addNode(font.getFontNode());
	}

	public void setBackgroundColor(String backgroundColor) {
		if(this.styleNode == null) {
			return;
		}
		
		XmlNode interiorNode = styleNode.getFirstChildByName("Interior");
		
		if(interiorNode == null) {
			interiorNode = new XmlNode("Interior");
			interiorNode.addAttribute("ss:Color", backgroundColor);
			interiorNode.addAttribute("ss:Pattern", "Solid");
		} else {
			if(interiorNode.containsAttribute("ss:Color")) {
				interiorNode.setAttribute("ss:Color", backgroundColor);
			} else {
				interiorNode.addAttribute("ss:Color", backgroundColor);
			}
		}
		
		styleNode.addNode(interiorNode);
	}
	
	public void setId(String id) {
		if(styleNode.containsAttribute("ss:ID")) {
			styleNode.setAttribute("ss:ID", id);
		} else {
			styleNode.addAttribute("ss:ID", id);
		}
	}
	
	public String getId() {
		if(styleNode == null) {
			return null;
		}
		
		return styleNode.getAttributeByName("ss:ID")==null?null:styleNode.getAttributeByName("ss:ID").getValue();
	}

	protected XmlNode getStyleNode() {
		return styleNode;
	}
	
	@Override
	public String toString() {
		return styleNode.toString();
	}
}
