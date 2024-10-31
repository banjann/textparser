package gdc;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Helper {

	public final static String MAP_TOFIND_KEY_INDICATOR = "tag";
	public final static String MAP_TOFIND_KEY_VALUE = "value";
	public final static String MAP_TOFIND_KEY_OCCURENCE = "occurence";

	public final static String MAP_TOFIND_VAL_INDICATOR_INCLUDED = "o";
	public final static String MAP_TOFIND_VAL_INDICATOR_UNINCLUDED = "x";

	private final String TEMPLATE_INPUT_SHEET = "ゆらぎ用語";
	private final String TEMPLATE_OUTPUT_SHEET = "出力";
	private final String OUTPUT_SHEET_COLUMN_A = "スキップフラグ";
	private final String OUTPUT_SHEET_COLUMN_B = "類似表現";
	private final String OUTPUT_SHEET_COL_FILENAME = "結果";
	private final String OUTPUT_SHEET_COL_FREQUENCY = "検出数";

	public final static String OUTPUT_SHEET_EMPTY_COL_B = "B列が空";

	public HashSet<Sourcefile> getAllFiles(String fileLocation) {
		HashSet<Sourcefile> files = new HashSet<Sourcefile>();

		if (fileLocation != null && !fileLocation.isEmpty()) {
			files.addAll(getXlsx(fileLocation));
			files.addAll(getXls(fileLocation));
			files.addAll(getDoc(fileLocation));
			files.addAll(getDocx(fileLocation));
		}

		return files;
	}

	private HashSet<Sourcefile> getXlsx(String fileLocation) {
		HashSet<Sourcefile> setPaths = new HashSet<Sourcefile>();

		File dir = new File(fileLocation);
		File[] files = dir.listFiles(new FilenameFilter() {
			public boolean accept(File dir, String name) {
				return name.toLowerCase().endsWith(".xlsx");
			}
		});

		if (files != null && files.length > 0) {
			for (File file : files) {
				Sourcefile thesource = new Sourcefile(file.getName(), Sourcefile.Filetype.xlsx, file.getAbsolutePath());
				setPaths.add(thesource);
			}
		}

		return setPaths;
	}

	private HashSet<Sourcefile> getXls(String fileLocation) {
		HashSet<Sourcefile> setPaths = new HashSet<Sourcefile>();

		File dir = new File(fileLocation);
		File[] files = dir.listFiles(new FilenameFilter() {
			public boolean accept(File dir, String name) {
				return name.toLowerCase().endsWith(".xls");
			}
		});

		if (files != null && files.length > 0) {
			for (File file : files) {
				Sourcefile thesource = new Sourcefile(file.getName(), Sourcefile.Filetype.xls, file.getAbsolutePath());
				setPaths.add(thesource);
			}
		}

		return setPaths;
	}

	private HashSet<Sourcefile> getDoc(String fileLocation) {
		HashSet<Sourcefile> setPaths = new HashSet<Sourcefile>();

		File dir = new File(fileLocation);
		File[] files = dir.listFiles(new FilenameFilter() {
			public boolean accept(File dir, String name) {
				return name.toLowerCase().endsWith(".doc");
			}
		});

		if (files != null && files.length > 0) {
			for (File file : files) {
				Sourcefile thesource = new Sourcefile(file.getName(), Sourcefile.Filetype.doc, file.getAbsolutePath());
				setPaths.add(thesource);
			}
		}

		return setPaths;
	}

	private HashSet<Sourcefile> getDocx(String fileLocation) {
		HashSet<Sourcefile> setPaths = new HashSet<Sourcefile>();

		File dir = new File(fileLocation);
		File[] files = dir.listFiles(new FilenameFilter() {
			public boolean accept(File dir, String name) {
				return name.toLowerCase().endsWith(".docx");
			}
		});

		if (files != null && files.length > 0) {
			for (File file : files) {
				Sourcefile thesource = new Sourcefile(file.getName(), Sourcefile.Filetype.docx, file.getAbsolutePath());
				setPaths.add(thesource);
			}
		}

		return setPaths;
	}

	public List<HashMap<String, Object>> getWordsToFind(String filepath) {
		List<HashMap<String, Object>> listWords = new ArrayList<HashMap<String, Object>>();

		try {
			FileInputStream fis = new FileInputStream(filepath);
			XSSFWorkbook xlswb = new XSSFWorkbook(fis);

			XSSFSheet sheetOfWords = xlswb.getSheet(TEMPLATE_INPUT_SHEET);
			Iterator<Row> rowIterator = sheetOfWords.iterator();
			int rowCount = 1;
			while (rowIterator.hasNext()) {
				HashMap<String, Object> hmRow = new HashMap<String, Object>();

				Row currentRow = rowIterator.next();
				if (rowCount > 1) {
					Iterator<Cell> cellIterator = currentRow.iterator();

					int colCount = 1;
					while (colCount <= 2 && cellIterator.hasNext()) {
						Cell currentCell = cellIterator.next();
						String strValue = "";

						switch (currentCell.getCellType()) {
						case NUMERIC:
							strValue = String.valueOf(currentCell.getNumericCellValue());
							break;
						default:
							strValue = currentCell.getStringCellValue();
							break;
						}
						if (colCount == 1) {
							if (strValue.toLowerCase().equals("x")) {
								hmRow.put(MAP_TOFIND_KEY_INDICATOR, MAP_TOFIND_VAL_INDICATOR_UNINCLUDED);
							} else if (strValue.toLowerCase().equals("o")){
								hmRow.put(MAP_TOFIND_KEY_INDICATOR, MAP_TOFIND_VAL_INDICATOR_INCLUDED);
							}
						} else if (colCount == 2) {
							hmRow.put(MAP_TOFIND_KEY_VALUE, strValue);
						}
						colCount++;
					}
					listWords.add(hmRow);
				}
				rowCount++;
			}
			xlswb.close();

		} catch (IOException e) {
			e.printStackTrace();
		}
		return listWords;
	}

	@SuppressWarnings("unchecked")
	public void printToTemplate(List<HashMap<String, Object>> words, String templateLocation, LocalDateTime now, String outPath) {
		try {
			FileInputStream fis = new FileInputStream(templateLocation);
			XSSFWorkbook xlswb = new XSSFWorkbook(fis);
			XSSFSheet outputSheet = xlswb.getSheet(TEMPLATE_OUTPUT_SHEET);

			// clear contents of output sheet
			for (int i = outputSheet.getLastRowNum(); i >= outputSheet.getFirstRowNum(); i--) {
				if (outputSheet.getRow(i) != null) {
					outputSheet.removeRow(outputSheet.getRow(i));
				}
		    }

			XSSFRow row = outputSheet.createRow(0); //create row for the inputs

			// print label for first row
			row.createCell(0).setCellValue(OUTPUT_SHEET_COLUMN_A);
			row.createCell(1).setCellValue(OUTPUT_SHEET_COLUMN_B);

			// print words and details
			if (words != null && words.size() > 0) {
				for (int i = 0; i < words.size(); i++) {
					HashMap<String, Object> currentWordDetails = words.get(i);
					String tagging = (String) currentWordDetails.get(MAP_TOFIND_KEY_INDICATOR);
					String theWord = (String) currentWordDetails.get(MAP_TOFIND_KEY_VALUE);
					Map<String, Integer> sortedExistenceMap = new TreeMap<>(
							(HashMap<String, Integer>) currentWordDetails.get(MAP_TOFIND_KEY_OCCURENCE)); // TreeMap automatically sorts (naturally) String keys

					// print to output sheet
					row = outputSheet.createRow(i + 1);
					row.createCell(0).setCellValue(tagging);
					row.createCell(1).setCellValue(theWord);

					if (sortedExistenceMap != null && !sortedExistenceMap.isEmpty()) {
						Set<String> keys = sortedExistenceMap.keySet();
						int colCounter = 0;
						for (String key : keys) {
							int indexOfFile = 2 + colCounter;

							row.createCell(indexOfFile).setCellValue(key); // print filename

							if (sortedExistenceMap.get(key) != null) {
								row.createCell(indexOfFile + 1).setCellValue(sortedExistenceMap.get(key)); // print frequency
							} else {
								row.createCell(indexOfFile + 1).setCellValue(OUTPUT_SHEET_EMPTY_COL_B); // print null string
							}

							// label headers of files and frequency
							outputSheet.getRow(0).createCell(indexOfFile).setCellValue(OUTPUT_SHEET_COL_FILENAME);
							outputSheet.getRow(0).createCell(indexOfFile + 1).setCellValue(OUTPUT_SHEET_COL_FREQUENCY);

							colCounter += 2;
						}
					}
				}
			}
			fis.close();

			DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss");
			String tempFileName = outPath + "\\textparser_" + (String) dtf.format(now) + ".xlsx";
			FileOutputStream fos =new FileOutputStream(new File(tempFileName));

			xlswb.write(fos);
			xlswb.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
