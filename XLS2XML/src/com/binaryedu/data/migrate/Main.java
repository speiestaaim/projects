package com.binaryedu.data.migrate;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.w3c.dom.CDATASection;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Text;

import com.sun.org.apache.xml.internal.serialize.OutputFormat;
import com.sun.org.apache.xml.internal.serialize.XMLSerializer;

/**
 * Read input excel file for a test and generate
 * a corresponding XML.
 * 
 * @author parsingh
 *
 */
public class Main
{
	// This name is used to create proper image tags
	// e.g. <img src="http://server/TEST_NAME/image.jpg"
	private static String TEST_NAME = "CAT_Free_Test_07";
	
	
	public static void main(String[] args)
	{
		xls2xml();
	}

	@SuppressWarnings("unchecked")
	private static void xls2xml()
	{
		Document dom = null;
		HSSFWorkbook wb = null;

		try
		{
			wb = new HSSFWorkbook(new FileInputStream("C:/Temp/" + TEST_NAME + ".xls"));
		}
		catch (FileNotFoundException e)
		{
			e.printStackTrace();
		}
		catch (IOException e)
		{
			e.printStackTrace();
		}
		
		// Read data from first sheet of the workbook.
		HSSFSheet sheet = wb.getSheetAt(0);

		// get an instance of factory
		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
		try
		{
			// get an instance of builder
			DocumentBuilder db = dbf.newDocumentBuilder();

			// create an instance of DOM
			dom = db.newDocument();
		}
		catch (ParserConfigurationException e)
		{
			e.printStackTrace();
			System.exit(1);
		}

		// create the root element
		Element rootEle = dom.createElement("test");
		dom.appendChild(rootEle);

		// initialize row and column counter
		int r = 0, c;
		for (Iterator<HSSFRow> rit = (Iterator<HSSFRow>) sheet.rowIterator(); rit.hasNext(); r++)
		{
			// get next row
			HSSFRow row = rit.next();

			// create question node
			Element questionEle = dom.createElement("question");
			rootEle.appendChild(questionEle);

			// start reading from column 1
			c = 1;
			for (Iterator<HSSFCell> cit = (Iterator<HSSFCell>) row.cellIterator(); cit.hasNext(); c++)
			{
				HSSFCell cell = cit.next();

				// check which column we are currently reading
				switch (c)
				{
//					case 1:
//						questionEle.appendChild(getQuestionNumNode(dom, cell));
//						break;
					case 2:
						// this is the question text node
						questionEle.appendChild(getQuestionTextNode(dom, cell));
						break;
					case 3:
						// this is the options node
						questionEle.appendChild(getOptionsNode(dom, cell));
						break;
					case 4:
						// this is the options node
						questionEle.appendChild(getSectionNameNode(dom, cell));
						break;
					case 5:
						// this is the directions node
						questionEle.appendChild(getDirectionsNode(dom, cell));
						break;
					case 6:
						// this is the correct option node
						questionEle.appendChild(getCorrectOptionNode(dom, cell));
						break;
					default:
						break;
				}
			}
		}

		printToFile(dom);
	}

	/**
	 * Parse question number from excel cell.
	 * 
	 * @param dom The XML Document
	 * @param cell The Excel cell currently under review
	 * @return
	 */
	private static Element getQuestionNumNode(Document dom, HSSFCell cell)
	{
		Element questionNumEle = dom.createElement("QuestionNumber");
		try
		{
			Text questionNumText = dom.createTextNode(String.valueOf((int) cell.getNumericCellValue()));
			questionNumEle.appendChild(questionNumText);
		}
		catch (Exception e)
		{
			System.err.println("Unable to parse cell: " + cell.getRowIndex() + "," + cell.getColumnIndex());
			System.err.println(cell.toString());
			e.printStackTrace();
		}
		return questionNumEle;
	}

	private static Element getCorrectOptionNode(Document dom, HSSFCell cell)
	{
		Element element = dom.createElement("CorrectOption");
		try
		{
			CDATASection text = dom.createCDATASection(String.valueOf((int) cell.getNumericCellValue()));
			element.appendChild(text);
		}
		catch (Exception e)
		{
			System.err.println("Unable to parse cell: " + cell.getRowIndex() + "," + cell.getColumnIndex());
			System.err.println(cell.toString());
			e.printStackTrace();
		}
		return element;
	}

	private static Element getQuestionTextNode(Document dom, HSSFCell cell)
	{
		Element questionEle = dom.createElement("QuestionText");
		try
		{
			String questionString = cell.getRichStringCellValue().toString();
			
			questionString = replaceImageTags(questionString);
			
			CDATASection questionText = dom.createCDATASection(questionString);
			questionEle.appendChild(questionText);
		}
		catch (Exception e)
		{
			System.err.println("Unable to parse cell: " + cell.getRowIndex() + "," + cell.getColumnIndex());
			System.err.println(cell.toString());
			e.printStackTrace();
		}
		
		return questionEle;
	}
	
	private static Element getSectionNameNode(Document dom, HSSFCell cell)
	{
		Element questionEle = dom.createElement("Section");
		try
		{
			String sectionString = cell.getRichStringCellValue().toString();
			
			CDATASection section = dom.createCDATASection(sectionString);
			questionEle.appendChild(section);
		}
		catch (Exception e)
		{
			System.err.println("Unable to parse cell: " + cell.getRowIndex() + "," + cell.getColumnIndex());
			System.err.println(cell.toString());
			e.printStackTrace();
		}
		
		return questionEle;
	}

	private static String replaceImageTags(String text)
	{
		while(true)
		{
			int startIndex = text.indexOf("{{");
			if(startIndex == -1)
			{
				return text;
			}
			
			int endIndex = text.indexOf("}}", startIndex);
			if(endIndex == -1)
			{
				return text;
			}
			
			String imageName = text.substring(startIndex, endIndex+2);
			String imageTag = "<br/><img src=\"http://binaryedu.com/images/" + TEST_NAME + "/" + imageName.substring(2, imageName.length()-2) + "\"/><br/>";
			
			text = text.replace(imageName, imageTag);
		}
	}
	
	private static Element getDirectionsNode(Document dom, HSSFCell cell)
	{
		Element element = dom.createElement("Directions");
		try
		{
			String directionsString = cell.getRichStringCellValue().toString();
			directionsString = replaceImageTags(directionsString);
			CDATASection text = dom.createCDATASection(directionsString);
			element.appendChild(text);
		}
		catch (Exception e)
		{
			System.err.println("Unable to parse cell: " + cell.getRowIndex() + "," + cell.getColumnIndex());
			System.err.println(cell.toString());
			e.printStackTrace();
		}
		return element;
	}

	private static Element getOptionsNode(Document dom, HSSFCell cell)
	{
		Element element = dom.createElement("options");

		String optionsText = cell.getRichStringCellValue().toString();
		if (optionsText == "")
		{
			return element;
		}

		String[] options = optionsText.split("\n");

		for (int i = 0; i < options.length; i++)
		{
			Element subElement = dom.createElement("option");
			
			try
			{
				String optionString = options[i];
				optionString = replaceImageTags(optionString);
				CDATASection subElementText = dom.createCDATASection(optionString);
				subElement.appendChild(subElementText);
			}
			catch (Exception e)
			{
				System.err.println("Unable to parse cell: " + cell.getRowIndex() + "," + cell.getColumnIndex());
				System.err.println(cell.toString());
				e.printStackTrace();
			}
			finally
			{
				element.appendChild(subElement);
			}
		}

		return element;
	}

	/**
	 * This method uses Xerces specific classes prints the XML document to file.
	 */
	private static void printToFile(Document dom)
	{
		try
		{
			// print
			OutputFormat format = new OutputFormat(dom);
			format.setIndenting(true);

			// to generate output to console use this serializer
			// XMLSerializer serializer = new XMLSerializer(System.out, format);

			// to generate a file output use fileoutputstream instead of
			// system.out
			XMLSerializer serializer = new XMLSerializer(new FileOutputStream(new File("C:/Temp/" + TEST_NAME + ".xml")), format);

			serializer.serialize(dom);

		}
		catch (IOException ie)
		{
			ie.printStackTrace();
		}
	}

}
