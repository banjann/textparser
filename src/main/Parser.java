// VERSION		AUTHOR			DATE
// 001			Naparota GDC    October 2024 (initial creation)

package main;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;

import javax.xml.namespace.QName;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.xmlbeans.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

public class Parser {
	
	private final static org.apache.logging.log4j.Logger logger = org.apache.logging.log4j.LogManager.getLogger(Parser.class);

	private static int wordFrequency;
	private static String wordToFind;
	private static String fileLocation;
	private static String templatePath;
	private static String outputPath;
	private static String includeLocation;

	public static void main(String[] args) {
		logger.info("START-TEXTPARSER");

		LocalDateTime startTime = LocalDateTime.now();

		fileLocation = args[0];
		templatePath = args[1];
		outputPath = args[2];
		includeLocation = args[3];
		wordFrequency = 0;

		// for testing
//		fileLocation = "C:\\SMARTMETER\\fy24\\04_word_search\\textparser\\files";
//		templatePath = "C:\\SMARTMETER\\fy24\\04_word_search\\textparser\\template\\textparser.xlsx";
//		outputPath = "C:\\SMARTMETER\\fy24\\04_word_search\\textparser\\output";
//		includeLocation = "yes";
//		wordFrequency = 0;

		Helper helper = new Helper();

		List<HashMap<String, Object>> wordsToFind = helper.getWordsToFind(templatePath); // look up each word to search

		// count the number of words that need to be searched (for progress display)
		double numberOfwordsToSearch = 0;
		for (HashMap<String, Object> rowOfWords : wordsToFind) {
			if (rowOfWords != null) {
				if (rowOfWords.get(Helper.MAP_TOFIND_KEY_INDICATOR) != null && 
					rowOfWords.get(Helper.MAP_TOFIND_KEY_INDICATOR).equals(Helper.MAP_TOFIND_VAL_INDICATOR_INCLUDED)) {
					numberOfwordsToSearch++;
				}
			}
		}

		// searching the words
		double counterOfSearchedWords = 0;
		HashSet<Integer> uniqueProgress = new HashSet<Integer>();
		System.out.println("===============================================================");
		System.out.println("進捗率：");
		for (HashMap<String, Object> rowOfWords : wordsToFind) {
			HashMap<String, Integer> hmOccurence = new HashMap<String, Integer>();

			if (rowOfWords != null) {
				if (rowOfWords.get(Helper.MAP_TOFIND_KEY_INDICATOR) != null && 
					rowOfWords.get(Helper.MAP_TOFIND_KEY_INDICATOR).equals(Helper.MAP_TOFIND_VAL_INDICATOR_INCLUDED)) { // only if it is not marked as "x"

					wordToFind = (String) rowOfWords.get(Helper.MAP_TOFIND_KEY_VALUE);
					HashSet<Sourcefile> searchSpace = helper.getAllFiles(fileLocation);

					for (Sourcefile file : searchSpace) {
						if (wordToFind != null && !wordToFind.isEmpty()) {
							wordFrequency = 0;
							String filepath = file.getFilePath();
							HashMap<String, Object> hmSheetOfWord = new HashMap<String, Object>();

							switch (file.getFileExtension()) {
							case xlsx:
								try {
									findInXlsx(filepath, hmSheetOfWord);
								} catch (IOException e) {
									logger.error("Error while processing " + file.getFileName(), e);
								}
								break;
							case xls:
								try {
									findInXls(filepath, hmSheetOfWord);
								} catch (IOException e) {
									logger.error("Error while processing " + file.getFileName(), e);
								}
								break;
							case docx:
								try {
									findInDocx(filepath);
								} catch (IOException e) {
									logger.error("Error while processing " + file.getFileName(), e);
								}
								break;
							case doc:
								break;
							}

							String tempKey = "ファイル名 : " + file.getFileName() + "\n" + "シート : " + hmSheetOfWord.toString();
							if (!hmSheetOfWord.isEmpty() && includeLocation.toLowerCase().equals("yes") && tempKey.length() < 32767) {
								hmOccurence.put(tempKey, wordFrequency);
							} else {
								tempKey = "ファイル名 : " + file.getFileName();
								hmOccurence.put(tempKey, wordFrequency);
							}

						} else {
							hmOccurence.put(Helper.OUTPUT_SHEET_EMPTY_COL_B, null);
						}
					}

					counterOfSearchedWords++;
					if (numberOfwordsToSearch != 0) {
						int progress = (int) ((counterOfSearchedWords / numberOfwordsToSearch) * 100);
						if (uniqueProgress.add(progress)) {
							System.out.print(" " + progress + "%");
						}
					} else {
						System.out.println("検索する言葉がない");
					}
				}
			}
			rowOfWords.put(Helper.MAP_TOFIND_KEY_OCCURENCE, hmOccurence);
		}

		System.out.println();
		System.out.println("===============================================================");
		helper.printToTemplate(wordsToFind, templatePath, startTime, outputPath);

		logger.info("END-TEXTPARSER");
	}

	private static void findInXls(String filepath, HashMap<String, Object> hmSheetLocation) throws IOException {
		FileInputStream fis = new FileInputStream(filepath);
		POIFSFileSystem fs = new POIFSFileSystem(fis);
		HSSFWorkbook xlswb = new HSSFWorkbook(fs);

		Iterator<Sheet> sheetIterator = xlswb.iterator();
		while (sheetIterator.hasNext()) {
			ArrayList<String> listLocationOfWords = new ArrayList<String>();

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
					int freqOfWord = countOccurence(strValue);
					if (freqOfWord > 0) {
						listLocationOfWords.add("cell: " + intToAlphabet(currentCell.getColumnIndex()) + String.valueOf(currentCell.getRowIndex() + 1));
					}
					wordFrequency += freqOfWord;
				}
			}
			// text from shapes
			HSSFPatriarch patriarch = currentSheet.getDrawingPatriarch();
			if (patriarch != null) {
				countInShapeXls(patriarch, listLocationOfWords);
			}

			if (listLocationOfWords.size() > 0) {
				hmSheetLocation.put(currentSheet.getSheetName(), listLocationOfWords);
			}
		}
		fis.close();
		fs.close();
		xlswb.close();
	}

	private static void countInShapeXls(ShapeContainer<HSSFShape> container, ArrayList<String> listLocationOfWords) {
		if (container != null) {
			for (HSSFShape shape : container) {
				String strTextInShape = "";
				if (shape instanceof HSSFShapeGroup) {
					HSSFShapeGroup shapeGroup = (HSSFShapeGroup) shape;
					for (HSSFShape shapeInside : shapeGroup) {
						if (shapeInside instanceof HSSFShapeGroup) {
							HSSFShapeGroup innerGroup = (HSSFShapeGroup) shapeInside;
							countInShapeXls(innerGroup, listLocationOfWords);

						} else if (shapeInside instanceof HSSFTextbox) {
							HSSFTextbox textboxShape = (HSSFTextbox) shapeInside;
							HSSFRichTextString richStr;
							try {
								richStr = textboxShape.getString();
								if (richStr != null) {
									strTextInShape = richStr.getString();
								}
							} catch (NullPointerException e) {
								//e.printStackTrace(); implement logger
							}

						} else if (shapeInside instanceof HSSFPolygon) {
							HSSFPolygon polygonShape = (HSSFPolygon) shapeInside;
							strTextInShape = polygonShape.getString().getString();

						} else if (shapeInside instanceof HSSFPicture) {
							HSSFPicture picShape = (HSSFPicture) shapeInside;
							strTextInShape = picShape.getString().getString();

						} else if (shapeInside instanceof HSSFCombobox) {
							HSSFCombobox comboShape = (HSSFCombobox) shapeInside;
							strTextInShape = comboShape.getString().getString();

						} else if (shapeInside instanceof HSSFSimpleShape) {
							HSSFSimpleShape simpleShape = (HSSFSimpleShape) shapeInside;
							HSSFRichTextString richStr;
							try {
								richStr = simpleShape.getString();
								if (richStr != null) {
									strTextInShape = richStr.getString();
								}
							} catch (NullPointerException e) {
								//e.printStackTrace(); implement logger
							}
						}
						int occurenceOfWord = countOccurence(strTextInShape);
						if (occurenceOfWord > 0) {
							listLocationOfWords.add("txtbox: " + shapeInside.getShapeName());
						}
						wordFrequency += occurenceOfWord;
					}

				} else if (shape instanceof HSSFTextbox) {
					HSSFTextbox textboxShape = (HSSFTextbox) shape;
					HSSFRichTextString richStr;
					try {
						richStr = textboxShape.getString();
						if (richStr != null) {
							strTextInShape = richStr.getString();
						}
					} catch (NullPointerException e) {
						//e.printStackTrace(); implement logger
					}

				} else if (shape instanceof HSSFPolygon) {
					HSSFPolygon polygonShape = (HSSFPolygon) shape;
					strTextInShape = polygonShape.getString().getString();

				} else if (shape instanceof HSSFPicture) {
					HSSFPicture picShape = (HSSFPicture) shape;
					strTextInShape = picShape.getString().getString();

				} else if (shape instanceof HSSFCombobox) {
					HSSFCombobox comboShape = (HSSFCombobox) shape;
					strTextInShape = comboShape.getString().getString();

				} else if (shape instanceof HSSFSimpleShape) {
					HSSFSimpleShape simpleShape = (HSSFSimpleShape) shape;
					HSSFRichTextString richStr;
					try {
						richStr = simpleShape.getString();
						if (richStr != null) {
							strTextInShape = richStr.getString();
						}
					} catch (NullPointerException e) {
						//e.printStackTrace(); implement logger
					}
				}
				int occurenceOfWord = countOccurence(strTextInShape);
				if (occurenceOfWord > 0) {
					listLocationOfWords.add("txtbox: " + shape.getShapeName());
				}
				wordFrequency += occurenceOfWord;
			}
		}
	}

	private static void findInXlsx(String filepath, HashMap<String, Object> hmSheetLocation) throws IOException {
		FileInputStream fis = new FileInputStream(filepath);

		ZipSecureFile.setMinInflateRatio(0);// no zip bomb possibility, handle large files
		XSSFWorkbook xlswb = new XSSFWorkbook(fis);

		if (xlswb != null) {
			Iterator<Sheet> sheetIterator = xlswb.iterator();

			while (sheetIterator.hasNext()) {
				ArrayList<String> listLocationOfWords = new ArrayList<String>();

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
						int freqOfWord = countOccurence(strValue);
						if (freqOfWord > 0) {
							listLocationOfWords.add("cell: " + intToAlphabet(currentCell.getColumnIndex()) + String.valueOf(currentCell.getRowIndex() + 1));
						}
						wordFrequency += freqOfWord;
					}
				}
				// text from shapes
				XSSFDrawing drawing = (XSSFDrawing) currentSheet.getDrawingPatriarch();
				countInShapeXlsx(drawing, listLocationOfWords);

				if (listLocationOfWords.size() > 0) {
					hmSheetLocation.put(currentSheet.getSheetName(), listLocationOfWords);
				}
			}
		}

		xlswb.close();
		fis.close();
	}

	private static void countInShapeXlsx(ShapeContainer<XSSFShape> container, ArrayList<String> listLocationOfWords) {
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
					countInShapeXlsx(shapeGroup, listLocationOfWords); // recursion to iterate through the group

				} else if (shape instanceof XSSFSimpleShape) {
					XSSFSimpleShape simpleShape = (XSSFSimpleShape) shape;

					int freqOfWord = countOccurence(simpleShape.getText());
					if (freqOfWord > 0) {
						listLocationOfWords.add("txtbox: " + simpleShape.getShapeName());
					}
					wordFrequency += freqOfWord;
				}
			}
		}
	}

	private static void findInDocx(String filepath) throws IOException {
		FileInputStream fis = new FileInputStream(filepath);

		ZipSecureFile.setMinInflateRatio(0);// no zip bomb possibility, handle large files
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
			wordFrequency += countOccurence(p.getText());

			// text in text boxes
			findInDocxTxtbox(p);
		}

		fis.close();
		docx.close();
	}

	private static void findInDocxTxtbox(XWPFParagraph paragraph) {
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
				logger.error("Error while getting words from textbox in docx.", e);
			}
		}
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

	private static String intToAlphabet(int i) {
		if (i < 0) {
			return "-" + intToAlphabet(-i - 1);
		}

		int quot = i / 26;
		int rem = i % 26;
		char letter = (char) ((int) 'A' + rem);
		if (quot == 0) {
			return "" + letter;
		} else {
			return intToAlphabet(quot - 1) + letter;
		}
	}
}
