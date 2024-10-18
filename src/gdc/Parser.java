package gdc;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;

import javax.xml.namespace.QName;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.xmlbeans.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

public class Parser {

	private static int wordFrequency;
	private static String wordToFind;
	private static String fileLocation;
	private static String templatePath;
	private static String outputPath;

	public static void main(String[] args) {
		DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss");
		LocalDateTime now = LocalDateTime.now();
		
		System.out.println("::START::TEXTPARSER::DATE-TIME::" + (String) dtf.format(now) + "::");
		fileLocation = "C:\\SMARTMETER\\fy24\\04_word_search\\textparser\\files";//args[0];
		templatePath = "C:\\SMARTMETER\\fy24\\04_word_search\\textparser\\template\\textparser.xlsx";//args[1];
		outputPath = "C:\\SMARTMETER\\fy24\\04_word_search\\textparser\\output";//args[2];
		
		wordFrequency = 0; //initialize count of word

		Helper helper = new Helper();

		List<HashMap<String, Object>> wordsToFind = helper.getWordsToFind(templatePath); // look up each word to search
		for (HashMap<String, Object> rowOfWords : wordsToFind) {
			HashMap<String, Integer> hmOccurence = new HashMap<String, Integer>();
			
			if (rowOfWords != null) {
				if (rowOfWords.get(Helper.MAP_TOFIND_KEY_INDICATOR) != null && rowOfWords
						.get(Helper.MAP_TOFIND_KEY_INDICATOR).equals(Helper.MAP_TOFIND_VAL_INDICATOR_INCLUDED)) {// only if it is not marked as "x"
					
					wordToFind = (String) rowOfWords.get(Helper.MAP_TOFIND_KEY_VALUE);
					HashSet<Sourcefile> searchSpace = helper.getAllFiles(fileLocation);
					for (Sourcefile file : searchSpace) {
						if (wordToFind != null && !wordToFind.isEmpty()) {
							System.out.println(file.getFileName());
							wordFrequency = 0;
							String filepath = file.getFilePath();
							
							switch (file.getFileExtension()) {
							case xlsx:
								try {
									findInXlsx(filepath);
								} catch (IOException e) {
									e.printStackTrace();
								}
								break;
							case xls:
								try {
									findInXls(filepath);
								} catch (IOException e) {
									e.printStackTrace();
								}
								break;
							case docx:
								try {
									findInDocx(filepath);
								} catch (Exception e) {
									e.printStackTrace();
								}
								break;
							case doc:
								break;
							}
							
							hmOccurence.put(file.getFileName(), wordFrequency);
						} else {
							hmOccurence.put(file.getFileName(), null);
						}
					}
				}
			}
			rowOfWords.put(Helper.MAP_TOFIND_KEY_OCCURENCE, hmOccurence);
		}

		helper.printToTemplate(wordsToFind, templatePath, now, outputPath);
		System.out.println("::END::TEXTPARSER::DATE-TIME::" + (String) dtf.format(now) + "::");
	}

	private static void findInXls(String filepath) throws IOException {
		System.out.println("STARTING METHOD::findInXls()");
		
		FileInputStream fis = new FileInputStream(filepath);
		POIFSFileSystem fs = new POIFSFileSystem(fis);
		HSSFWorkbook xlswb = new HSSFWorkbook(fs);

		Iterator<Sheet> sheetIterator = xlswb.iterator();
		while (sheetIterator.hasNext()) {
			HSSFSheet currentSheet = (HSSFSheet) sheetIterator.next();
			Iterator<Row> rowIterator = currentSheet.iterator();

			// texts from cells
			while (rowIterator.hasNext()) {
				Row currentRow = rowIterator.next();
				Iterator<Cell> cellIterator = currentRow.iterator();

				while (cellIterator.hasNext()) {
					Cell currentCell = cellIterator.next();
					String strValue = "";
					
					switch (currentCell.getCellType()) {
					case NUMERIC:
						strValue = String.valueOf(currentCell.getNumericCellValue());
						break;
					case STRING:
						strValue = currentCell.getStringCellValue();
						break;
					case FORMULA:
						DataFormatter dataFormatter = new DataFormatter(new java.util.Locale("en", "US"));
						dataFormatter.setUseCachedValuesForFormulaCells(true);
						strValue = dataFormatter.formatCellValue(currentCell);
						break;
					default:
						break;
					}
					wordFrequency += countOccurence(strValue);
				}
			}
			// text from shapes
			HSSFPatriarch patriarch = currentSheet.getDrawingPatriarch();
			if (patriarch != null) {
				countInShapeXls(patriarch);
			}
		}
		fis.close();
		fs.close();
		xlswb.close();
		System.out.println("ENDING METHOD::findInXls()");
	}

	private static void countInShapeXls(ShapeContainer<HSSFShape> container) {
		System.out.println("STARTING METHOD::countInShapeXls()");
		if (container != null) {
			for (HSSFShape shape : container) {
				if (shape instanceof HSSFShapeGroup) {
					HSSFShapeGroup shapeGroup = (HSSFShapeGroup) shape;
					for (HSSFShape shapeInside : shapeGroup) {
						if (shapeInside instanceof HSSFShapeGroup) {
							HSSFShapeGroup innerGroup = (HSSFShapeGroup) shapeInside;
							countInShapeXls(innerGroup);
						} else if (shapeInside instanceof HSSFTextbox) {
							HSSFTextbox textboxShape = (HSSFTextbox) shapeInside;
							wordFrequency += countOccurence(textboxShape.getString().getString());

						} else if (shapeInside instanceof HSSFPolygon) {
							HSSFPolygon polygonShape = (HSSFPolygon) shapeInside;
							wordFrequency += countOccurence(polygonShape.getString().getString());

						} else if (shapeInside instanceof HSSFPicture) {
							HSSFPicture picShape = (HSSFPicture) shapeInside;
							wordFrequency += countOccurence(picShape.getString().getString());

						} else if (shapeInside instanceof HSSFCombobox) {
							HSSFCombobox comboShape = (HSSFCombobox) shapeInside;
							wordFrequency += countOccurence(comboShape.getString().getString());

						} else if (shapeInside instanceof HSSFSimpleShape) {
							HSSFSimpleShape simpleShape = (HSSFSimpleShape) shapeInside;
							HSSFRichTextString richStr;
							try {
								richStr = simpleShape.getString();
								if (richStr != null) {
									wordFrequency += countOccurence(richStr.getString());
								}
							} catch (NullPointerException e) {
								e.printStackTrace();
							}
						}
					}

				} else if (shape instanceof HSSFTextbox) {
					HSSFTextbox textboxShape = (HSSFTextbox) shape;
					wordFrequency += countOccurence(textboxShape.getString().getString());

				} else if (shape instanceof HSSFPolygon) {
					HSSFPolygon polygonShape = (HSSFPolygon) shape;
					wordFrequency += countOccurence(polygonShape.getString().getString());

				} else if (shape instanceof HSSFPicture) {
					HSSFPicture picShape = (HSSFPicture) shape;
					wordFrequency += countOccurence(picShape.getString().getString());

				} else if (shape instanceof HSSFCombobox) {
					HSSFCombobox comboShape = (HSSFCombobox) shape;
					wordFrequency += countOccurence(comboShape.getString().getString());

				} else if (shape instanceof HSSFSimpleShape) {
					HSSFSimpleShape simpleShape = (HSSFSimpleShape) shape;
					HSSFRichTextString richStr;
					try {
						richStr = simpleShape.getString();
						if (richStr != null) {
							wordFrequency += countOccurence(richStr.getString());
						}
					} catch (NullPointerException e) {
						e.printStackTrace();
					}
				}
			}
		}
		System.out.println("ENDING METHOD::countInShapeXls()");
	}

	private static void findInXlsx(String filepath) throws IOException {
		System.out.println("STARTING METHOD::findInXlsx()");
		FileInputStream fis = new FileInputStream(filepath);
		XSSFWorkbook xlswb = new XSSFWorkbook(fis);

		Iterator<Sheet> sheetIterator = xlswb.iterator();
		while (sheetIterator.hasNext()) {
			Sheet currentSheet = sheetIterator.next();
			Iterator<Row> rowIterator = currentSheet.iterator();

			// texts from cells
			while (rowIterator.hasNext()) {
				Row currentRow = rowIterator.next();
				Iterator<Cell> cellIterator = currentRow.iterator();
				while (cellIterator.hasNext()) {
					Cell currentCell = cellIterator.next();
					String strValue = "";
					
					switch (currentCell.getCellType()) {
					case NUMERIC:
						strValue = String.valueOf(currentCell.getNumericCellValue());
						break;
					case STRING:
						strValue = currentCell.getStringCellValue();
						break;
					case FORMULA:
						DataFormatter dataFormatter = new DataFormatter(new java.util.Locale("en", "US"));
						dataFormatter.setUseCachedValuesForFormulaCells(true);
						strValue = dataFormatter.formatCellValue(currentCell);
						break;
					default:
						break;
					}
					wordFrequency += countOccurence(strValue);
				}
			}
			// text from shapes
			XSSFDrawing drawing = (XSSFDrawing) currentSheet.getDrawingPatriarch();
			countInShapeXlsx(drawing);
		}
		fis.close();
		xlswb.close();
		System.out.println("ENDING METHOD::findInXlsx()");
	}

	private static void countInShapeXlsx(ShapeContainer<XSSFShape> container) {
		System.out.println("STARTING METHOD::countInShapeXlsx()");
		if (container != null) {
			for (XSSFShape shape : container) {
				if (shape instanceof XSSFConnector) {
					continue;

				} else if (shape instanceof XSSFGraphicFrame) {
					continue;

				} else if (shape instanceof XSSFPicture) {
					continue;

				} else if (shape instanceof XSSFShapeGroup) {
					XSSFShapeGroup shapeGroup = (XSSFShapeGroup) shape;
					countInShapeXlsx(shapeGroup); // recursion to iterate through the group

				} else if (shape instanceof XSSFSimpleShape) {
					XSSFSimpleShape simpleShape = (XSSFSimpleShape) shape;
					wordFrequency += countOccurence(simpleShape.getText());
				}
			}
		}
		System.out.println("ENDING METHOD::countInShapeXlsx()");
	}

	private static void findInDocx(String filepath) throws Exception {
		System.out.println("STARTING METHOD::findInDocx()");
		FileInputStream fis = new FileInputStream(filepath);
		XWPFDocument docx = new XWPFDocument(fis);

		// text in tables
		for (XWPFTable tbl : docx.getTables()) {
			for (XWPFTableRow row : tbl.getRows()) {
				for (XWPFTableCell cell : row.getTableCells()) {
					wordFrequency += countOccurence(cell.getText());
				}
			}
		}

		for (XWPFParagraph p : docx.getParagraphs()) {
			// text in body
			wordFrequency += countOccurence(p.getParagraphText());

			// text in text boxes
			findInDocxTxtbox(p);
		}

		fis.close();
		docx.close();
		System.out.println("ENDING METHOD::findInDocx()");
	}

	private static void findInDocxTxtbox(XWPFParagraph paragraph) {
		System.out.println("STARTING METHOD::findInDocxTxtbox()");
		XmlObject[] textBoxObjects = paragraph
				.getCTP()
				.selectPath(
						"declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'" + 
						"declare namespace wps='http://schemas.microsoft.com/office/word/2010/wordprocessingShape' .//*/wps:txbx/w:txbxContent"
							);

		for (int i = 0; i < textBoxObjects.length; i++) {
			XWPFParagraph embeddedPara = null;
			try {
				XmlObject[] paraObjects = textBoxObjects[i].selectChildren(new QName("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "p"));
				for (int j = 0; j < paraObjects.length; j++) {
					embeddedPara = new XWPFParagraph(CTP.Factory.parse(paraObjects[j].xmlText()), paragraph.getBody());
					// paragraphs inside the text box
					wordFrequency += countOccurence(embeddedPara.getParagraphText());
				}
			} catch (XmlException e) {
				e.printStackTrace();
			}
		}
		System.out.println("ENDING METHOD::findInDocxTxtbox()");
	}

	private static int countOccurence(String from) {
		int occurence = 0;

		int index = 0;
		int length = wordToFind.length();

		while ((index = from.indexOf(wordToFind, index)) != -1) {
			index += length;
			occurence++;
		}

		return occurence;
	}
}
