package ooxml;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.UUID;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.commons.io.IOUtils;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPicture;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.xml.sax.SAXException;

/**
 * @author Son
 *
 */
public class PoiTest {

	public static final String XML = "<v:shape alt=\"Microsoft Office Signature Line...\" style=\"width:192pt;height:96pt\"><v:imagedata o:title=\"\"/>"
			+ "<o:lock v:ext=\"edit\" ungrouping=\"t\" rotation=\"t\" cropping=\"t\" verticies=\"t\" text=\"t\" grouping=\"t\"/>"
			+ "<o:signatureline v:ext=\"edit\" id=\"{3A7233FE-85B4-42EA-9F7A-C8F11A0BC5A7}\" provid=\"{00000000-0000-0000-0000-000000000000}\" o:suggestedsigner=\"aaaa\" o:suggestedsigner2=\"bbbb\" o:suggestedsigneremail=\"ddd@sample.com\" allowcomments=\"t\" issignatureline=\"t\"/>"
			+ "</v:shape>";
	
	public static final String SCHEMA_O = "urn:schemas-microsoft-com:office:office";
	public static final String SCHEMA_V = "urn:schemas-microsoft-com:vml";
	public static final String SCHEMA_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
	
	@Test
	public void test() throws IOException, ParserConfigurationException, SAXException, InvalidFormatException {
		InputStream docxInputStream = new FileInputStream("C:\\Users\\Son\\Documents\\aaa.docx");
		InputStream signatueInputStream = new FileInputStream("C:\\Users\\Son\\Documents\\signature.png");
		OutputStream os = new FileOutputStream("C:\\Users\\Son\\Documents\\" + UUID.randomUUID().toString() + ".docx");
		
		XWPFDocument doc = new XWPFDocument(docxInputStream);
		String pictureDataRelationId = doc.addPictureData(IOUtils.toByteArray(signatueInputStream), org.apache.poi.xwpf.usermodel.Document.PICTURE_TYPE_PNG);
		XWPFParagraph para = doc.getLastParagraph();
		XWPFRun signatureRun = para.createRun();
		CTPicture pict = signatureRun.getCTR().addNewPict();
		Node pictDomNode = pict.getDomNode();
		Document docNode = pictDomNode.getOwnerDocument();
		
		Element shapeEl = docNode.createElementNS(SCHEMA_V, "shape");
		shapeEl.setAttribute("alt", "Microsoft Office Signature Line...");
		shapeEl.setAttribute("style", "width:192pt;height:96pt");
		
		Element imagedataEl = docNode.createElementNS(SCHEMA_V, "imagedata");
		//r:id="rId5"
		imagedataEl.setAttributeNS(SCHEMA_R, "id", pictureDataRelationId);
		imagedataEl.setAttributeNS(SCHEMA_O, "title", "");
		shapeEl.appendChild(imagedataEl);
		
		Element lockEl = docNode.createElementNS(SCHEMA_O, "lock");
		lockEl.setAttributeNS(SCHEMA_V, "ext", "edit");
		lockEl.setAttribute("ungrouping", "t");
		lockEl.setAttribute("rotation", "t");
		lockEl.setAttribute("cropping", "t");
		lockEl.setAttribute("verticies", "t");
		lockEl.setAttribute("text", "t");
		lockEl.setAttribute("grouping", "t");
		shapeEl.appendChild(lockEl);

		Element signatureLineEl = docNode.createElementNS(SCHEMA_O, "signatureline");
		signatureLineEl.setAttributeNS(SCHEMA_V, "ext", "edit");
		signatureLineEl.setAttribute("id", "{" + UUID.randomUUID().toString().toUpperCase() + "}");
		signatureLineEl.setAttribute("provid", "{00000000-0000-0000-0000-000000000000}");
		signatureLineEl.setAttributeNS(SCHEMA_O, "suggestedsigner", "Nguyen Duc Son");
		signatureLineEl.setAttributeNS(SCHEMA_O, "suggestedsigner2", "Nguyen Duc Son2");
		signatureLineEl.setAttributeNS(SCHEMA_O, "suggestedsigneremail", "sonnd1988@gmail.com");
		signatureLineEl.setAttribute("allowcomments", "t");
		signatureLineEl.setAttribute("issignatureline", "t");
		shapeEl.appendChild(signatureLineEl);
		
		pictDomNode.appendChild(shapeEl);
		
		para.createRun().addPicture(signatueInputStream, PictureType.PNG.ordinal(), "signature.png", 200, 160);
		
		doc.write(os);
		
		doc.close();
		IOUtils.closeQuietly(signatueInputStream);
		IOUtils.closeQuietly(docxInputStream);
		IOUtils.closeQuietly(os);
	}
}
