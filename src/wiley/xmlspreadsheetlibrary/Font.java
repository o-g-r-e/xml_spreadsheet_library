package wiley.xmlspreadsheetlibrary;

import java.awt.Color;

import wiley.xmlobjectlibrary.XmlNode;

public class Font {
	private XmlNode fontNode;
	
	protected Font() {
		this.fontNode = new XmlNode("Font");
	}

	public void setFontName(String fontName) {
		putFontNodeAttribute("ss:FontName", fontName);
	}

	public void setSize(String size) {
		putFontNodeAttribute("ss:Size", size);
	}

	public void setBold(boolean bold) {
		if(bold) {
			putFontNodeAttribute("ss:Bold", "1");
		} else {
			if(fontNode.containsAttribute("ss:Bold")) {
				fontNode.removeAttribute("ss:Bold");
			}
		}
	}
	
	public void setItalic(boolean italic) {
		if(italic) {
			putFontNodeAttribute("ss:Italic", "1");
		} else {
			if(fontNode.containsAttribute("ss:Italic")) {
				fontNode.removeAttribute("ss:Italic");
			}
		}
	}

	public void setColor(String color) {
		putFontNodeAttribute("ss:Color", color);
	}
	
	public void setColor(Color color) {
		putFontNodeAttribute("ss:Color", String.format("#%02x%02x%02x", color.getRed(), color.getGreen(), color.getBlue()));
	}
	
	private void putFontNodeAttribute(String attrName, String attrValue) {
		if(fontNode.containsAttribute(attrName)) {
			fontNode.setAttribute(attrName, attrValue);
		} else {
			fontNode.addAttribute(attrName, attrValue);
		}
	}

	public XmlNode getFontNode() {
		return fontNode;
	}
}
